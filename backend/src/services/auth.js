import bcrypt from "bcryptjs";
import jwt from "jsonwebtoken";
import jwksRsa from "jwks-rsa";
import { v4 as uuidv4 } from "uuid";
import { prisma } from "./db.js";
import { createError } from "./validation.js";

const SESSION_COOKIE_NAME = "m365_toolbox_session";
const SESSION_TTL_SECONDS = Number(process.env.AUTH_SESSION_TTL_SECONDS || 60 * 60 * 8);
const DEFAULT_ADMIN_USERNAME = "admin";
const DEFAULT_ADMIN_PASSWORD = "admin";
const ROLE_RANK = {
  restricted_user: 1,
  privileged_user: 2,
  administrator: 3
};
const VALID_ROLES = new Set(Object.keys(ROLE_RANK));

function getSessionSecret() {
  return (
    process.env.AUTH_SESSION_SECRET ||
    process.env.ARTIFACT_TOKEN_SECRET ||
    "m365-toolbox-development-session-secret"
  );
}

function normalizeUser(user) {
  if (!user) {
    return null;
  }

  return {
    id: user.id,
    username: user.username,
    displayName: user.displayName,
    email: user.email,
    authProvider: user.authProvider,
    entraObjectId: user.entraObjectId,
    role: user.role,
    mustChangePassword: Boolean(user.mustChangePassword),
    lastLoginAt: user.lastLoginAt ? user.lastLoginAt.toISOString() : null
  };
}

function normalizeMicrosoftConfig(config) {
  return {
    enabled: Boolean(config?.enabled),
    tenantId: config?.tenantId || "",
    clientId: config?.clientId || "",
    apiClientId: config?.apiClientId || "",
    authorityUrl: config?.authorityUrl || ""
  };
}

function validateRole(role) {
  if (!VALID_ROLES.has(role)) {
    throw createError("Assigned role is invalid.", 400);
  }
  return role;
}

export function getRoleRank(role) {
  return ROLE_RANK[role] || 0;
}

export function userCanRunScript(user, script) {
  if (!user) {
    return false;
  }
  if (user.role === "administrator" || user.role === "privileged_user") {
    return true;
  }
  return script?.mode !== "remediation" && !script?.approvalRequired;
}

export function setSessionCookie(res, user) {
  const token = jwt.sign(
    {
      sub: user.id,
      role: user.role,
      authProvider: user.authProvider
    },
    getSessionSecret(),
    {
      expiresIn: SESSION_TTL_SECONDS,
      issuer: "m365-toolbox"
    }
  );

  res.cookie(SESSION_COOKIE_NAME, token, {
    httpOnly: true,
    sameSite: "lax",
    secure: process.env.NODE_ENV === "production",
    maxAge: SESSION_TTL_SECONDS * 1000,
    path: "/"
  });
}

export function clearSessionCookie(res) {
  res.clearCookie(SESSION_COOKIE_NAME, {
    httpOnly: true,
    sameSite: "lax",
    secure: process.env.NODE_ENV === "production",
    path: "/"
  });
}

export async function getUserFromRequest(req) {
  const token = req.cookies?.[SESSION_COOKIE_NAME];
  if (!token) {
    return null;
  }

  let payload;
  try {
    payload = jwt.verify(token, getSessionSecret(), {
      issuer: "m365-toolbox"
    });
  } catch {
    return null;
  }

  const userId = typeof payload?.sub === "string" ? payload.sub.trim() : "";
  if (!userId) {
    return null;
  }

  try {
    const user = await prisma.user.findUnique({ where: { id: userId } });
    return normalizeUser(user);
  } catch {
    return null;
  }
}

export async function ensureDefaultAdmin() {
  const existing = await prisma.user.findUnique({
    where: { username: DEFAULT_ADMIN_USERNAME }
  });

  if (existing) {
    return normalizeUser(existing);
  }

  const passwordHash = await bcrypt.hash(DEFAULT_ADMIN_PASSWORD, 12);
  const user = await prisma.user.create({
    data: {
      id: uuidv4(),
      username: DEFAULT_ADMIN_USERNAME,
      passwordHash,
      displayName: "Default Administrator",
      authProvider: "local",
      role: "administrator",
      mustChangePassword: true
    }
  });
  return normalizeUser(user);
}

export async function loginLocal(username, password) {
  const normalizedUsername = String(username || "").trim();
  if (!normalizedUsername || !password) {
    throw createError("Username and password are required.", 400);
  }

  const user = await prisma.user.findUnique({
    where: { username: normalizedUsername }
  });

  if (!user || user.authProvider !== "local" || !user.passwordHash) {
    throw createError("Invalid username or password.", 401);
  }

  const passwordMatches = await bcrypt.compare(String(password), user.passwordHash);
  if (!passwordMatches) {
    throw createError("Invalid username or password.", 401);
  }

  const updated = await prisma.user.update({
    where: { id: user.id },
    data: { lastLoginAt: new Date() }
  });
  return normalizeUser(updated);
}

export async function changeLocalPassword(userId, currentPassword, newPassword) {
  const nextPassword = String(newPassword || "");
  if (nextPassword.length < 8) {
    throw createError("New password must be at least 8 characters long.", 400);
  }

  const user = await prisma.user.findUnique({ where: { id: userId } });
  if (!user || user.authProvider !== "local" || !user.passwordHash) {
    throw createError("Password changes are only available for local users.", 400);
  }

  const currentMatches = await bcrypt.compare(String(currentPassword || ""), user.passwordHash);
  if (!currentMatches) {
    throw createError("Current password is incorrect.", 401);
  }

  const passwordHash = await bcrypt.hash(nextPassword, 12);
  const updated = await prisma.user.update({
    where: { id: user.id },
    data: {
      passwordHash,
      mustChangePassword: false
    }
  });
  return normalizeUser(updated);
}

export async function getMicrosoftAuthConfig() {
  const config = await prisma.microsoftAuthConfig.findUnique({
    where: { id: "singleton" }
  });
  return normalizeMicrosoftConfig(config);
}

export async function getPublicMicrosoftAuthConfig() {
  const config = await getMicrosoftAuthConfig();
  const scope = config.apiClientId ? `api://${config.apiClientId}/access_as_user` : "";
  return {
    enabled: config.enabled,
    tenantId: config.tenantId,
    clientId: config.clientId,
    apiClientId: config.apiClientId,
    authorityUrl: config.authorityUrl,
    scope
  };
}

export async function updateMicrosoftAuthConfig(input = {}) {
  const enabled = Boolean(input.enabled);
  const tenantId = String(input.tenantId || "").trim();
  const clientId = String(input.clientId || "").trim();
  const apiClientId = String(input.apiClientId || "").trim();
  const authorityUrl = String(input.authorityUrl || "").trim();

  if (enabled && (!tenantId || !clientId || !apiClientId)) {
    throw createError("Tenant ID, frontend client ID, and API audience are required when Microsoft login is enabled.", 400);
  }

  const config = await prisma.microsoftAuthConfig.upsert({
    where: { id: "singleton" },
    update: { enabled, tenantId, clientId, apiClientId, authorityUrl },
    create: { id: "singleton", enabled, tenantId, clientId, apiClientId, authorityUrl }
  });
  return normalizeMicrosoftConfig(config);
}

function normalizeMapping(mapping) {
  return {
    id: mapping.id,
    groupName: mapping.groupName,
    groupId: mapping.groupId,
    assignedRole: mapping.assignedRole,
    createdAt: mapping.createdAt.toISOString(),
    updatedAt: mapping.updatedAt.toISOString()
  };
}

function normalizeEmail(email) {
  const value = String(email || "").trim();
  if (!value) {
    return null;
  }
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(value)) {
    throw createError("Email address is invalid.", 400);
  }
  return value;
}

function normalizeLocalUserInput(input = {}, { requirePassword = false } = {}) {
  const username = String(input.username || "").trim();
  const displayName = String(input.displayName || "").trim() || null;
  const email = normalizeEmail(input.email);
  const role = validateRole(input.role);
  const password = String(input.password || "");

  if (!username) {
    throw createError("Username is required.", 400);
  }
  if (!/^[A-Za-z0-9._@-]{3,80}$/.test(username)) {
    throw createError("Username must be 3-80 characters and can include letters, numbers, dots, underscores, at signs, and hyphens.", 400);
  }
  if ((requirePassword || password) && password.length < 8) {
    throw createError("Password must be at least 8 characters long.", 400);
  }

  return { username, displayName, email, role, password };
}

async function ensureAnotherAdministratorExists(userId) {
  const administratorCount = await prisma.user.count({
    where: {
      role: "administrator",
      NOT: { id: userId }
    }
  });
  if (administratorCount < 1) {
    throw createError("At least one other administrator account is required.", 400);
  }
}

export async function listUsers() {
  const users = await prisma.user.findMany({
    orderBy: [{ authProvider: "asc" }, { username: "asc" }]
  });
  return users.map(normalizeUser);
}

export async function createLocalUser(input) {
  const data = normalizeLocalUserInput(input, { requirePassword: true });
  const passwordHash = await bcrypt.hash(data.password, 12);
  try {
    const user = await prisma.user.create({
      data: {
        id: uuidv4(),
        username: data.username,
        passwordHash,
        displayName: data.displayName,
        email: data.email,
        authProvider: "local",
        role: data.role,
        mustChangePassword: Boolean(input.mustChangePassword)
      }
    });
    return normalizeUser(user);
  } catch (error) {
    if (error?.code === "P2002") {
      throw createError("A user with that username, email, or Microsoft object ID already exists.", 409);
    }
    throw error;
  }
}

export async function updateUser(id, input) {
  const existing = await prisma.user.findUnique({ where: { id } });
  if (!existing) {
    throw createError("User not found.", 404);
  }

  const role = existing.authProvider === "local" ? validateRole(input.role || existing.role) : existing.role;
  if (existing.role === "administrator" && role !== "administrator") {
    await ensureAnotherAdministratorExists(existing.id);
  }

  const updateData = {
    displayName: String(input.displayName || "").trim() || null,
    email: normalizeEmail(input.email),
    role
  };

  if (existing.authProvider === "local") {
    const username = String(input.username || "").trim();
    if (!username) {
      throw createError("Username is required.", 400);
    }
    if (!/^[A-Za-z0-9._@-]{3,80}$/.test(username)) {
      throw createError("Username must be 3-80 characters and can include letters, numbers, dots, underscores, at signs, and hyphens.", 400);
    }
    updateData.username = username;

    updateData.mustChangePassword = Boolean(input.mustChangePassword);

    const password = String(input.password || "");
    if (password) {
      if (password.length < 8) {
        throw createError("Password must be at least 8 characters long.", 400);
      }
      updateData.passwordHash = await bcrypt.hash(password, 12);
    }
  }

  try {
    const user = await prisma.user.update({
      where: { id },
      data: updateData
    });
    return normalizeUser(user);
  } catch (error) {
    if (error?.code === "P2002") {
      throw createError("A user with that username, email, or Microsoft object ID already exists.", 409);
    }
    if (error?.code === "P2025") {
      throw createError("User not found.", 404);
    }
    throw error;
  }
}

export async function deleteUser(id, currentUserId) {
  if (id === currentUserId) {
    throw createError("You cannot remove your own account while signed in.", 400);
  }

  const existing = await prisma.user.findUnique({ where: { id } });
  if (!existing) {
    throw createError("User not found.", 404);
  }
  if (existing.role === "administrator") {
    await ensureAnotherAdministratorExists(existing.id);
  }

  await prisma.user.delete({ where: { id } });
}

function normalizeMappingInput(input = {}) {
  const groupName = String(input.groupName || "").trim();
  const groupId = String(input.groupId || "").trim() || null;
  const assignedRole = validateRole(input.assignedRole);
  if (!groupName) {
    throw createError("Entra group name is required.", 400);
  }
  return { groupName, groupId, assignedRole };
}

export async function listGroupRoleMappings() {
  const mappings = await prisma.entraGroupRoleMapping.findMany({
    orderBy: [{ assignedRole: "asc" }, { groupName: "asc" }]
  });
  return mappings.map(normalizeMapping);
}

export async function createGroupRoleMapping(input) {
  const data = normalizeMappingInput(input);
  try {
    const mapping = await prisma.entraGroupRoleMapping.create({
      data: { id: uuidv4(), ...data }
    });
    return normalizeMapping(mapping);
  } catch (error) {
    if (error?.code === "P2002") {
      throw createError("That Entra group object ID is already mapped.", 409);
    }
    throw error;
  }
}

export async function updateGroupRoleMapping(id, input) {
  const data = normalizeMappingInput(input);
  try {
    const mapping = await prisma.entraGroupRoleMapping.update({
      where: { id },
      data
    });
    return normalizeMapping(mapping);
  } catch (error) {
    if (error?.code === "P2002") {
      throw createError("That Entra group object ID is already mapped.", 409);
    }
    if (error?.code === "P2025") {
      throw createError("Group mapping not found.", 404);
    }
    throw error;
  }
}

export async function deleteGroupRoleMapping(id) {
  try {
    await prisma.entraGroupRoleMapping.delete({ where: { id } });
  } catch (error) {
    if (error?.code === "P2025") {
      throw createError("Group mapping not found.", 404);
    }
    throw error;
  }
}

function buildAuthority(config) {
  const base = config.authorityUrl || `https://login.microsoftonline.com/${config.tenantId}`;
  return base.replace(/\/+$/, "");
}

function getTokenKey(config) {
  const jwksClient = jwksRsa({
    cache: true,
    rateLimit: true,
    jwksRequestsPerMinute: 10,
    jwksUri: `${buildAuthority(config)}/discovery/v2.0/keys`
  });

  return (header, callback) => {
    jwksClient.getSigningKey(header.kid, (error, key) => {
      if (error) {
        callback(error);
        return;
      }
      callback(null, key.getPublicKey());
    });
  };
}

function verifyMicrosoftJwt(token, config) {
  return new Promise((resolve, reject) => {
    jwt.verify(
      token,
      getTokenKey(config),
      {
        algorithms: ["RS256"],
        audience: [config.apiClientId, `api://${config.apiClientId}`],
        issuer: [
          `https://login.microsoftonline.com/${config.tenantId}/v2.0`,
          `https://sts.windows.net/${config.tenantId}/`
        ]
      },
      (error, decoded) => {
        if (error) {
          reject(error);
          return;
        }
        resolve(decoded);
      }
    );
  });
}

function getClaimValues(value) {
  if (!value) {
    return [];
  }
  if (Array.isArray(value)) {
    return value.map((entry) => String(entry).trim()).filter(Boolean);
  }
  return [String(value).trim()].filter(Boolean);
}

function roleFromMappings(claims, mappings) {
  if (claims?._claim_names?.groups) {
    throw createError("Microsoft token uses group overage claims. Configure the app registration to emit role/group claims for this API or reduce group overage before login.", 403);
  }

  const groupIds = new Set(getClaimValues(claims.groups).map((value) => value.toLowerCase()));
  const groupNames = new Set([
    ...getClaimValues(claims.groupNames),
    ...getClaimValues(claims.groups).filter((value) => !/^[0-9a-f-]{36}$/i.test(value))
  ].map((value) => value.toLowerCase()));

  let selectedRole = null;
  for (const mapping of mappings) {
    const idMatch = mapping.groupId && groupIds.has(mapping.groupId.toLowerCase());
    const nameMatch = !mapping.groupId && groupNames.has(mapping.groupName.toLowerCase());
    if (!idMatch && !nameMatch) {
      continue;
    }
    if (!selectedRole || getRoleRank(mapping.assignedRole) > getRoleRank(selectedRole)) {
      selectedRole = mapping.assignedRole;
    }
  }

  if (!selectedRole) {
    throw createError("Microsoft login succeeded, but no Entra group mapping matched an authorized Toolbox role.", 403);
  }
  return selectedRole;
}

export async function loginMicrosoft(token) {
  if (!token) {
    throw createError("Microsoft token is required.", 400);
  }

  const config = await getMicrosoftAuthConfig();
  if (!config.enabled) {
    throw createError("Microsoft login is not enabled.", 403);
  }
  if (!config.tenantId || !config.apiClientId) {
    throw createError("Microsoft authentication is not fully configured.", 500);
  }

  let claims;
  try {
    claims = await verifyMicrosoftJwt(token, config);
  } catch (error) {
    throw createError(`Microsoft token validation failed: ${error.message}`, 401);
  }

  const mappings = await prisma.entraGroupRoleMapping.findMany();
  const role = roleFromMappings(claims, mappings);
  const entraObjectId = claims.oid || claims.sub;
  const email = claims.preferred_username || claims.email || claims.upn || null;
  const displayName = claims.name || email || "Microsoft user";
  const username = email || entraObjectId;

  const user = await prisma.user.upsert({
    where: { entraObjectId },
    update: {
      username,
      displayName,
      email,
      role,
      lastLoginAt: new Date()
    },
    create: {
      id: uuidv4(),
      username,
      displayName,
      email,
      authProvider: "microsoft",
      entraObjectId,
      role,
      mustChangePassword: false,
      lastLoginAt: new Date()
    }
  });

  return normalizeUser(user);
}

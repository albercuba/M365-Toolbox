import crypto from "node:crypto";

const AUTH_MODE = (process.env.AUTH_MODE || "optional").toLowerCase();
const SESSION_COOKIE_NAME = "m365_toolbox_session";
const SESSION_TTL_MS = Math.max(1, Number(process.env.SESSION_TTL_HOURS || 12)) * 60 * 60 * 1000;
const APP_BASE_URL = process.env.APP_BASE_URL || process.env.FRONTEND_ORIGIN?.split(",")[0]?.trim() || "http://localhost:5173";
const ENTRA_TENANT_ID = process.env.ENTRA_TENANT_ID || "common";
const ENTRA_CLIENT_ID = process.env.ENTRA_CLIENT_ID || "";
const ENTRA_CLIENT_SECRET = process.env.ENTRA_CLIENT_SECRET || "";
const CALLBACK_PATH = "/api/auth/callback";
const REDIRECT_URI = `${APP_BASE_URL.replace(/\/$/, "")}${CALLBACK_PATH}`;
const sessionStore = new Map();
const stateStore = new Map();

function pruneExpiredSessions() {
  const now = Date.now();
  for (const [id, session] of sessionStore.entries()) {
    if (session.expiresAt <= now) {
      sessionStore.delete(id);
    }
  }

  for (const [state, value] of stateStore.entries()) {
    if (value.expiresAt <= now) {
      stateStore.delete(state);
    }
  }
}

setInterval(pruneExpiredSessions, 15 * 60 * 1000).unref();

function toBase64Url(input) {
  return Buffer.from(input)
    .toString("base64")
    .replace(/\+/g, "-")
    .replace(/\//g, "_")
    .replace(/=+$/g, "");
}

function decodeJwtPayload(token) {
  const parts = token.split(".");
  if (parts.length < 2) {
    return {};
  }

  try {
    const payload = parts[1].replace(/-/g, "+").replace(/_/g, "/");
    return JSON.parse(Buffer.from(payload, "base64").toString("utf8"));
  } catch {
    return {};
  }
}

function parseCookieHeader(cookieHeader = "") {
  return cookieHeader
    .split(";")
    .map((part) => part.trim())
    .filter(Boolean)
    .reduce((acc, part) => {
      const separator = part.indexOf("=");
      if (separator === -1) {
        return acc;
      }
      const key = part.slice(0, separator);
      const value = part.slice(separator + 1);
      acc[key] = decodeURIComponent(value);
      return acc;
    }, {});
}

function buildCookieOptions(maxAgeMs) {
  const secure = APP_BASE_URL.startsWith("https://");
  return [
    `${SESSION_COOKIE_NAME}=`,
    "Path=/",
    "HttpOnly",
    "SameSite=Lax",
    secure ? "Secure" : null,
    `Max-Age=${Math.floor(maxAgeMs / 1000)}`
  ].filter(Boolean);
}

function setSessionCookie(res, sessionId, maxAgeMs = SESSION_TTL_MS) {
  const cookie = buildCookieOptions(maxAgeMs);
  cookie[0] = `${SESSION_COOKIE_NAME}=${encodeURIComponent(sessionId)}`;
  res.setHeader("Set-Cookie", cookie.join("; "));
}

function clearSessionCookie(res) {
  const cookie = buildCookieOptions(0);
  cookie[0] = `${SESSION_COOKIE_NAME}=`;
  res.setHeader("Set-Cookie", cookie.join("; "));
}

function createSession(user) {
  const id = crypto.randomUUID();
  const session = {
    id,
    user,
    createdAt: new Date().toISOString(),
    expiresAt: Date.now() + SESSION_TTL_MS
  };
  sessionStore.set(id, session);
  return session;
}

function getProviderConfigured() {
  return Boolean(ENTRA_CLIENT_ID && ENTRA_CLIENT_SECRET);
}

export function getAuthConfig() {
  const providerConfigured = getProviderConfigured();
  const effectiveMode = AUTH_MODE === "disabled" ? "disabled" : providerConfigured ? AUTH_MODE : "disabled";

  return {
    mode: effectiveMode,
    requestedMode: AUTH_MODE,
    providerConfigured,
    loginEnabled: providerConfigured && (AUTH_MODE === "optional" || AUTH_MODE === "required"),
    allowAnonymous: effectiveMode === "disabled" || effectiveMode === "optional",
    anonymousWarning: "Audit trail, per-user accountability, and user-specific approvals are not available without login."
  };
}

export function attachSession(req, _res, next) {
  const config = getAuthConfig();
  const cookies = parseCookieHeader(req.headers.cookie || "");
  const sessionId = cookies[SESSION_COOKIE_NAME];
  const existing = sessionId ? sessionStore.get(sessionId) : null;

  if (existing && existing.expiresAt > Date.now()) {
    existing.expiresAt = Date.now() + SESSION_TTL_MS;
    req.toolboxSession = existing;
    req.toolboxUser = existing.user;
  } else if (config.mode === "disabled") {
    req.toolboxSession = {
      id: "disabled-auth",
      user: {
        id: "anonymous",
        displayName: "Anonymous",
        username: "anonymous",
        authType: "anonymous",
        isAnonymous: true
      }
    };
    req.toolboxUser = req.toolboxSession.user;
  } else {
    req.toolboxSession = null;
    req.toolboxUser = null;
  }

  next();
}

export function requireToolboxAccess(req, res, next) {
  const config = getAuthConfig();
  if (config.mode === "disabled") {
    next();
    return;
  }

  if (req.toolboxUser) {
    next();
    return;
  }

  res.status(401).json({ message: "Sign in or continue without login to access the toolbox." });
}

export function getSessionInfo(req) {
  const config = getAuthConfig();
  return {
    config,
    session: req.toolboxUser
      ? {
          isAuthenticated: !req.toolboxUser.isAnonymous,
          isAnonymous: Boolean(req.toolboxUser.isAnonymous),
          user: req.toolboxUser
        }
      : null
  };
}

export function startAnonymousSession(req, res) {
  const config = getAuthConfig();
  if (!config.allowAnonymous) {
    const error = new Error("Anonymous access is not allowed in the current authentication mode.");
    error.statusCode = 403;
    throw error;
  }

  const session = createSession({
    id: `anonymous-${crypto.randomUUID()}`,
    displayName: "Anonymous",
    username: "anonymous",
    authType: "anonymous",
    isAnonymous: true
  });
  setSessionCookie(res, session.id);
  req.toolboxSession = session;
  req.toolboxUser = session.user;
  return getSessionInfo(req);
}

export function logoutSession(req, res) {
  const cookies = parseCookieHeader(req.headers.cookie || "");
  const sessionId = cookies[SESSION_COOKIE_NAME];
  if (sessionId) {
    sessionStore.delete(sessionId);
  }
  clearSessionCookie(res);
}

export function createMicrosoftLoginUrl(req) {
  const config = getAuthConfig();
  if (!config.loginEnabled) {
    const error = new Error("Microsoft Entra login is not configured.");
    error.statusCode = 400;
    throw error;
  }

  const state = crypto.randomUUID();
  stateStore.set(state, {
    expiresAt: Date.now() + 10 * 60 * 1000,
    returnTo: req.query.returnTo || "/"
  });

  const url = new URL(`https://login.microsoftonline.com/${ENTRA_TENANT_ID}/oauth2/v2.0/authorize`);
  url.searchParams.set("client_id", ENTRA_CLIENT_ID);
  url.searchParams.set("response_type", "code");
  url.searchParams.set("redirect_uri", REDIRECT_URI);
  url.searchParams.set("response_mode", "query");
  url.searchParams.set("scope", "openid profile email User.Read");
  url.searchParams.set("state", state);
  url.searchParams.set("prompt", "select_account");
  return url.toString();
}

async function exchangeAuthorizationCode(code) {
  const response = await fetch(`https://login.microsoftonline.com/${ENTRA_TENANT_ID}/oauth2/v2.0/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      client_id: ENTRA_CLIENT_ID,
      client_secret: ENTRA_CLIENT_SECRET,
      code,
      redirect_uri: REDIRECT_URI,
      grant_type: "authorization_code"
    })
  });

  const data = await response.json();
  if (!response.ok) {
    const error = new Error(data.error_description || data.error || "Failed to exchange authorization code.");
    error.statusCode = 400;
    throw error;
  }

  return data;
}

export async function completeMicrosoftLogin(req, res) {
  const { code, state, error, error_description: errorDescription } = req.query;
  const stateEntry = stateStore.get(state);
  if (state) {
    stateStore.delete(state);
  }

  if (error) {
    return `${APP_BASE_URL.replace(/\/$/, "")}/?authError=${encodeURIComponent(errorDescription || String(error))}`;
  }

  if (!stateEntry || stateEntry.expiresAt <= Date.now()) {
    return `${APP_BASE_URL.replace(/\/$/, "")}/?authError=${encodeURIComponent("The login request expired. Please try again.")}`;
  }

  const tokenData = await exchangeAuthorizationCode(code);
  const idTokenClaims = decodeJwtPayload(tokenData.id_token || "");
  const session = createSession({
    id: idTokenClaims.oid || idTokenClaims.sub || crypto.randomUUID(),
    displayName: idTokenClaims.name || idTokenClaims.preferred_username || "Microsoft Entra user",
    username: idTokenClaims.preferred_username || idTokenClaims.email || "unknown",
    tenantId: idTokenClaims.tid || null,
    authType: "entra",
    isAnonymous: false
  });

  setSessionCookie(res, session.id);
  return `${APP_BASE_URL.replace(/\/$/, "")}/${stateEntry.returnTo ? stateEntry.returnTo.replace(/^\//, "") : ""}`;
}

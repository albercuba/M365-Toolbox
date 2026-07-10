import { getRoleRank, getUserFromRequest, userCanRunScript } from "../services/auth.js";
import { getScript } from "../services/scriptRunner.js";

export async function attachUser(req, _res, next) {
  try {
    req.user = await getUserFromRequest(req);
    next();
  } catch (error) {
    next(error);
  }
}

export function requireAuth(req, res, next) {
  if (!req.user) {
    res.status(401).json({ message: "Authentication is required." });
    return;
  }
  next();
}

export function requireRole(...roles) {
  const minimumRank = Math.min(...roles.map((role) => getRoleRank(role)).filter(Boolean));
  return (req, res, next) => {
    if (!req.user) {
      res.status(401).json({ message: "Authentication is required." });
      return;
    }
    if (getRoleRank(req.user.role) < minimumRank) {
      res.status(403).json({ message: "Insufficient role for this operation." });
      return;
    }
    next();
  };
}

const PASSWORD_CHANGE_ALLOWED_PATHS = new Set([
  "/api/auth/me",
  "/api/auth/logout",
  "/api/auth/change-password"
]);

export function enforcePasswordChange(req, res, next) {
  if (!req.user?.mustChangePassword || req.method === "OPTIONS") {
    next();
    return;
  }

  if (PASSWORD_CHANGE_ALLOWED_PATHS.has(req.path)) {
    next();
    return;
  }

  res.status(403).json({ message: "You must change your password before using other Toolbox features." });
}

export function requireCanRunScript(req, res, next) {
  if (!req.user) {
    res.status(401).json({ message: "Authentication is required." });
    return;
  }

  let script;
  try {
    script = getScript(req.params.id);
  } catch (error) {
    res.status(error.statusCode || 404).json({ message: error.message });
    return;
  }

  if (!userCanRunScript(req.user, script)) {
    res.status(403).json({ message: "Restricted users cannot run remediation or high-impact scripts." });
    return;
  }

  req.script = script;
  next();
}

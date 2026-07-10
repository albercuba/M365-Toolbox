import express from "express";
import cors from "cors";
import cookieParser from "cookie-parser";
import morgan from "morgan";
import { authRouter } from "./routes/auth.js";
import { authSettingsRouter } from "./routes/authSettings.js";
import { scriptsRouter } from "./routes/scripts.js";
import { attachUser, enforcePasswordChange } from "./middleware/auth.js";
import { assertSessionSecretConfiguration, ensureDefaultAdmin } from "./services/auth.js";
import { ensureDatabaseReady } from "./services/db.js";

assertSessionSecretConfiguration();

const app = express();
app.set("etag", false);
const port = Number(process.env.PORT || 3001);
const frontendOrigin = process.env.FRONTEND_ORIGIN || "";
const allowPrivateNetworkOrigins = process.env.CORS_ALLOW_PRIVATE_NETWORK === "true";

function normalizeConfiguredOrigin(origin) {
  const trimmedOrigin = origin.trim();
  if (!trimmedOrigin) {
    return "";
  }

  try {
    const url = new URL(trimmedOrigin);
    if (url.protocol !== "http:" && url.protocol !== "https:") {
      return "";
    }

    return url.origin;
  } catch {
    return trimmedOrigin.replace(/\/+$/, "");
  }
}

const allowedOrigins = frontendOrigin
  ? frontendOrigin
      .split(",")
      .map((origin) => normalizeConfiguredOrigin(origin))
      .filter(Boolean)
  : [];

function isAllowedToolboxOrigin(origin) {
  if (!origin) {
    return true;
  }

  const normalizedOrigin = normalizeConfiguredOrigin(origin);

  if (allowedOrigins.includes(normalizedOrigin)) {
    return true;
  }

  try {
    const url = new URL(normalizedOrigin);
    if (url.protocol !== "http:" && url.protocol !== "https:") {
      return false;
    }

    const host = url.hostname.toLowerCase();
    const isLocalHost =
      host === "localhost" ||
      host === "127.0.0.1" ||
      host === "::1";

    const isPrivateIpv4 =
      /^10\./.test(host) ||
      /^192\.168\./.test(host) ||
      /^172\.(1[6-9]|2\d|3[0-1])\./.test(host);

    return isLocalHost || (allowPrivateNetworkOrigins && isPrivateIpv4);
  } catch {
    return false;
  }
}

app.use(
  cors({
    credentials: true,
    origin(origin, callback) {
      if (isAllowedToolboxOrigin(origin)) {
        callback(null, true);
        return;
      }

      callback(new Error(`Origin '${origin}' is not allowed.`));
    }
  })
);
app.use(cookieParser());
app.use(express.json({ limit: "1mb" }));
app.use(morgan("dev"));
app.use((req, res, next) => {
  if (req.path.startsWith("/api/")) {
    res.setHeader("Cache-Control", "no-store, no-cache, must-revalidate, private");
    res.setHeader("Pragma", "no-cache");
    res.setHeader("Expires", "0");
  }
  next();
});
app.use(attachUser);
app.use(enforcePasswordChange);

app.get("/api/health", (_req, res) => {
  res.json({
    name: "M365 Toolbox API",
    status: "ok",
    tagline: "M365 Toolbox - Web-based PowerShell operations for Microsoft 365"
  });
});

app.use("/api", authRouter);
app.use("/api", authSettingsRouter);
app.use("/api", scriptsRouter);

app.use("/api", (_req, res) => {
  res.status(404).json({ message: "API route not found." });
});

app.use((error, _req, res, _next) => {
  console.error(error);
  res.status(error.statusCode || 500).json({ message: error.message || "Unexpected server error." });
});

await ensureDatabaseReady();
await ensureDefaultAdmin();

app.listen(port, () => {
  console.log(`M365 Toolbox API listening on port ${port}`);
});

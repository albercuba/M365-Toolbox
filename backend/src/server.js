import express from "express";
import cors from "cors";
import morgan from "morgan";
import { scriptsRouter } from "./routes/scripts.js";
import { ensureDatabaseReady } from "./services/db.js";

const app = express();
const port = Number(process.env.PORT || 3001);
const frontendOrigin = process.env.FRONTEND_ORIGIN || "";

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

    return isLocalHost || isPrivateIpv4;
  } catch {
    return false;
  }
}

app.use(
  cors({
    origin(origin, callback) {
      if (isAllowedToolboxOrigin(origin)) {
        callback(null, true);
        return;
      }

      callback(new Error(`Origin '${origin}' is not allowed.`));
    }
  })
);
app.use(express.json({ limit: "1mb" }));
app.use(morgan("dev"));

app.get("/api/health", (_req, res) => {
  res.json({
    name: "M365 Toolbox API",
    status: "ok",
    tagline: "M365 Toolbox - Web-based PowerShell operations for Microsoft 365"
  });
});

app.use("/api", scriptsRouter);

app.use("/api", (_req, res) => {
  res.status(404).json({ message: "API route not found." });
});

app.use((error, _req, res, _next) => {
  console.error(error);
  res.status(500).json({ message: error.message || "Unexpected server error." });
});

await ensureDatabaseReady();

app.listen(port, () => {
  console.log(`M365 Toolbox API listening on port ${port}`);
});

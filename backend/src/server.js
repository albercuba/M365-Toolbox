import express from "express";
import cors from "cors";
import morgan from "morgan";
import { scriptsRouter } from "./routes/scripts.js";

const app = express();
const port = Number(process.env.PORT || 3001);
const frontendOrigin = process.env.FRONTEND_ORIGIN || "";
const allowedOrigins = frontendOrigin
  ? frontendOrigin.split(",").map((origin) => origin.trim()).filter(Boolean)
  : [];

app.use(
  cors({
    origin(origin, callback) {
      if (!origin || allowedOrigins.length === 0 || allowedOrigins.includes(origin)) {
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

app.listen(port, () => {
  console.log(`M365 Toolbox API listening on port ${port}`);
});

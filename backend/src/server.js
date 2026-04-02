import express from "express";
import cors from "cors";
import morgan from "morgan";
import { scriptsRouter } from "./routes/scripts.js";

const app = express();
const port = Number(process.env.PORT || 3001);
const frontendOrigin = process.env.FRONTEND_ORIGIN || "http://localhost:5173";

app.use(cors({ origin: frontendOrigin }));
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


import { Router } from "express";
import {
  cancelRun,
  getRun,
  getRunArtifact,
  getRunArtifacts,
  getRunHtml,
  getRuns,
  getScript,
  listScripts,
  startRun
} from "../services/scriptRunner.js";
import { getSystemStatus } from "../services/healthStatus.js";

export const scriptsRouter = Router();

scriptsRouter.get("/scripts", (_req, res) => {
  res.json(listScripts());
});

scriptsRouter.get("/scripts/:id", (req, res) => {
  try {
    res.json(getScript(req.params.id));
  } catch (error) {
    res.status(404).json({ message: error.message });
  }
});

scriptsRouter.post("/scripts/:id/run", (req, res) => {
  try {
    const { approvalConfirmed, ...payload } = req.body || {};
    const run = startRun(req.params.id, payload, {
      approvalConfirmed: Boolean(approvalConfirmed)
    });
    res.status(202).json(run);
  } catch (error) {
    res.status(error.statusCode || 400).json({ message: error.message });
  }
});

scriptsRouter.get("/runs", (_req, res) => {
  res.json(getRuns());
});

scriptsRouter.get("/runs/:id", (req, res) => {
  const run = getRun(req.params.id);
  if (!run) {
    res.status(404).json({ message: "Run not found." });
    return;
  }

  res.json(run);
});

scriptsRouter.post("/runs/:id/cancel", (req, res) => {
  try {
    res.json(cancelRun(req.params.id));
  } catch (error) {
    res.status(error.statusCode || 400).json({ message: error.message });
  }
});

scriptsRouter.get("/runs/:id/artifacts", (req, res) => {
  try {
    res.json(getRunArtifacts(req.params.id));
  } catch (error) {
    res.status(error.statusCode || 400).json({ message: error.message });
  }
});

scriptsRouter.get("/runs/:id/artifacts/:artifactId", (req, res) => {
  try {
    const artifact = getRunArtifact(req.params.id, req.params.artifactId);
    res.download(artifact.path, artifact.name);
  } catch (error) {
    res.status(error.statusCode || 400).json({ message: error.message });
  }
});

scriptsRouter.get("/runs/:id/html", (req, res) => {
  const htmlReport = getRunHtml(req.params.id);
  if (!htmlReport) {
    res.status(404).json({ message: "HTML report not found for this run." });
    return;
  }

  res.type("html").send(htmlReport.content);
});

scriptsRouter.get("/status", async (_req, res, next) => {
  try {
    res.json(await getSystemStatus());
  } catch (error) {
    next(error);
  }
});

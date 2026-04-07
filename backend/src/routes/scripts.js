import { Router } from "express";
import { getRun, getRunHtml, getRuns, getScript, listScripts, startRun } from "../services/scriptRunner.js";

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
    const run = startRun(req.params.id, req.body || {});
    res.status(202).json(run);
  } catch (error) {
    res.status(400).json({ message: error.message });
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

scriptsRouter.get("/runs/:id/html", (req, res) => {
  const htmlReport = getRunHtml(req.params.id);
  if (!htmlReport) {
    res.status(404).json({ message: "HTML report not found for this run." });
    return;
  }

  res.type("html").send(htmlReport.content);
});

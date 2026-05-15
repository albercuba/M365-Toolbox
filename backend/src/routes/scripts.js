import { Router } from "express";
import { requireAuth, requireCanRunScript, requireRole } from "../middleware/auth.js";
import {
  buildArtifactArchive,
  cancelRun,
  clearRunArtifacts,
  getRun,
  getRunArtifact,
  getRunArtifacts,
  getRunHtml,
  getQueueStatus,
  getRuns,
  getScript,
  listScripts,
  startRun
} from "../services/scriptRunner.js";
import { verifyArtifactToken } from "../services/artifactTokens.js";
import { getSystemStatus } from "../services/healthStatus.js";
import {
  createCompany,
  deleteCompany,
  listCompanies,
  replaceCompanies,
  updateCompany
} from "../services/companyStore.js";

export const scriptsRouter = Router();

scriptsRouter.use(requireAuth);

scriptsRouter.get("/scripts", (_req, res) => {
  res.json(listScripts());
});

scriptsRouter.get("/companies", async (_req, res, next) => {
  try {
    res.json(await listCompanies());
  } catch (error) {
    next(error);
  }
});

scriptsRouter.post("/companies", requireRole("administrator"), async (req, res) => {
  try {
    res.status(201).json(await createCompany(req.body));
  } catch (error) {
    res.status(error.statusCode || 400).json({ message: error.message });
  }
});

scriptsRouter.put("/companies", requireRole("administrator"), async (req, res) => {
  try {
    res.json(await replaceCompanies(req.body?.companies || []));
  } catch (error) {
    res.status(error.statusCode || 400).json({ message: error.message });
  }
});

scriptsRouter.put("/companies/:id", requireRole("administrator"), async (req, res) => {
  try {
    res.json(await updateCompany(req.params.id, req.body));
  } catch (error) {
    res.status(error.statusCode || 400).json({ message: error.message });
  }
});

scriptsRouter.delete("/companies/:id", requireRole("administrator"), async (req, res) => {
  try {
    await deleteCompany(req.params.id);
    res.status(204).end();
  } catch (error) {
    res.status(error.statusCode || 400).json({ message: error.message });
  }
});

scriptsRouter.get("/scripts/:id", (req, res) => {
  try {
    res.json(getScript(req.params.id));
  } catch (error) {
    res.status(404).json({ message: error.message });
  }
});

scriptsRouter.post("/scripts/:id/run", requireCanRunScript, async (req, res) => {
  try {
    const { approvalConfirmed, ...payload } = req.body || {};
    const run = await startRun(req.params.id, payload, {
      approvalConfirmed: Boolean(approvalConfirmed),
      requestedBy: req.user.displayName || req.user.username,
      user: req.user
    });
    res.status(202).json(run);
  } catch (error) {
    res.status(error.statusCode || 400).json({ message: error.message });
  }
});

scriptsRouter.get("/runs", async (req, res) => {
  res.json(
    await getRuns({
      status: req.query.status,
      scriptId: req.query.scriptId,
      tenantId: req.query.tenantId,
      requestedBy: req.query.requestedBy,
      dateFrom: req.query.dateFrom,
      dateTo: req.query.dateTo,
      limit: req.query.limit,
      offset: req.query.offset
    })
  );
});

scriptsRouter.get("/runs/:id", async (req, res) => {
  const run = await getRun(req.params.id);
  if (!run) {
    res.status(404).json({ message: "Run not found." });
    return;
  }

  res.json(run);
});

scriptsRouter.post("/runs/:id/cancel", async (req, res) => {
  try {
    res.json(await cancelRun(req.params.id));
  } catch (error) {
    res.status(error.statusCode || 400).json({ message: error.message });
  }
});

scriptsRouter.get("/runs/:id/artifacts", async (req, res) => {
  try {
    res.json(await getRunArtifacts(req.params.id));
  } catch (error) {
    res.status(error.statusCode || 400).json({ message: error.message });
  }
});

scriptsRouter.delete("/runs/:id/artifacts", requireRole("administrator"), async (req, res) => {
  try {
    res.json(await clearRunArtifacts(req.params.id));
  } catch (error) {
    res.status(error.statusCode || 400).json({ message: error.message });
  }
});

scriptsRouter.get("/runs/:id/artifacts/:artifactId", async (req, res) => {
  try {
    if (
      !verifyArtifactToken(req.query.token, {
        runId: req.params.id,
        artifactId: req.params.artifactId,
        kind: "download"
      })
    ) {
      res.status(403).json({ message: "Artifact token is invalid or expired." });
      return;
    }
    const artifact = await getRunArtifact(req.params.id, req.params.artifactId);
    res.download(artifact.path, artifact.name);
  } catch (error) {
    res.status(error.statusCode || 400).json({ message: error.message });
  }
});

scriptsRouter.get("/runs/:id/package.zip", async (req, res) => {
  try {
    if (
      !verifyArtifactToken(req.query.token, {
        runId: req.params.id,
        kind: "bundle"
      })
    ) {
      res.status(403).json({ message: "Bundle token is invalid or expired." });
      return;
    }

    res.type("application/zip");
    res.setHeader("Content-Disposition", `attachment; filename="${req.params.id}-artifacts.zip"`);
    await buildArtifactArchive(req.params.id, res);
  } catch (error) {
    res.status(error.statusCode || 400).json({ message: error.message });
  }
});

scriptsRouter.get("/runs/:id/html", async (req, res) => {
  try {
    if (
      !verifyArtifactToken(req.query.token, {
        runId: req.params.id,
        kind: "html"
      })
    ) {
      res.status(403).json({ message: "HTML preview token is invalid or expired." });
      return;
    }

    const htmlReport = await getRunHtml(req.params.id);
    if (!htmlReport) {
      res.status(404).json({ message: "HTML report not found for this run." });
      return;
    }

    res.setHeader(
      "Content-Security-Policy",
      "default-src 'none'; script-src 'unsafe-inline'; style-src 'unsafe-inline' https://fonts.googleapis.com; font-src https://fonts.gstatic.com data:; img-src data: https: http:; sandbox allow-same-origin allow-scripts;"
    );
    res.setHeader("X-Content-Type-Options", "nosniff");
    res.type("html").send(htmlReport.content);
  } catch (error) {
    res.status(error.statusCode || 400).json({ message: error.message });
  }
});

scriptsRouter.get("/queue", async (_req, res) => {
  try {
    res.json(await getQueueStatus());
  } catch (error) {
    res.status(500).json({ message: error.message });
  }
});

scriptsRouter.get("/status", async (_req, res, next) => {
  try {
    res.json(await getSystemStatus());
  } catch (error) {
    next(error);
  }
});

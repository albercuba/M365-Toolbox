import path from "node:path";
import { v4 as uuidv4 } from "uuid";

export const STRUCTURED_EVENT_PREFIX = "::toolbox::";

export function nowIso() {
  return new Date().toISOString();
}

export function markRunActivity(run, timestamp = nowIso()) {
  run.updatedAt = timestamp;
  run.lastActivityAt = timestamp;
}

export function classifyLogLevel(stream, message) {
  if (stream === "stderr") {
    return "error";
  }

  if (/^\[\!\]/.test(message) || /error|failed/i.test(message)) {
    return "error";
  }
  if (/^\[\*\]/.test(message) || /warn/i.test(message)) {
    return "warn";
  }
  if (/^\[\+\]/.test(message)) {
    return "progress";
  }

  return "info";
}

export function addLogEntry(run, stream, message) {
  if (!message) {
    return;
  }

  const clean = message.replace(/\r/g, "").trim();
  if (!clean) {
    return;
  }

  const entry = {
    id: uuidv4(),
    timestamp: nowIso(),
    stream,
    level: classifyLogLevel(stream, clean),
    message: clean
  };

  run.logs.push(entry);
  markRunActivity(run, entry.timestamp);

  if (entry.level === "progress" || (stream === "stdout" && clean.length > 8)) {
    run.currentStep = clean.replace(/^\[[^\]]+\]\s*/, "");
  }

  if (entry.level === "error") {
    run.errorSummary = clean;
  }
}

export function applyStructuredEvent(run, event) {
  if (!event || typeof event !== "object") {
    return;
  }

  run.events.push({
    ...event,
    timestamp: event.timestamp || nowIso()
  });
  markRunActivity(run, event.timestamp || nowIso());

  if (event.type === "progress" && event.message) {
    run.currentStep = event.message;
  }

  if (event.type === "artifact" && event.path) {
    const artifactName = path.basename(event.path);
    const existing = run.artifacts.files.find((artifact) => artifact.path === event.path);
    if (!existing) {
      run.artifacts.files.push({
        id: artifactName,
        name: artifactName,
        path: event.path,
        type: event.kind || path.extname(artifactName).slice(1).toLowerCase() || "file",
        size: event.size || 0,
        createdAt: event.timestamp || nowIso()
      });
    }
    if ((event.kind || "").toLowerCase() === "html") {
      run.artifacts.htmlPath = event.path;
    }
  }
}

export function parseOutputChunk(run, stream, chunk) {
  const text = chunk.toString();
  const bufferProperty = stream === "stdout" ? "_stdoutBuffer" : "_stderrBuffer";
  run[bufferProperty] = (run[bufferProperty] || "") + text;

  const rawLines = run[bufferProperty].split(/\r?\n/);
  run[bufferProperty] = rawLines.pop() || "";

  for (const rawLine of rawLines) {
    if (!rawLine) {
      continue;
    }

    if (rawLine.startsWith(STRUCTURED_EVENT_PREFIX)) {
      try {
        const event = JSON.parse(rawLine.slice(STRUCTURED_EVENT_PREFIX.length));
        applyStructuredEvent(run, event);
      } catch {
        addLogEntry(run, stream, rawLine);
        run[stream] += `${rawLine}\n`;
      }
      continue;
    }

    run[stream] += `${rawLine}\n`;
    addLogEntry(run, stream, rawLine);
  }
}

export function flushOutputBuffers(run) {
  for (const [stream, bufferProperty] of [
    ["stdout", "_stdoutBuffer"],
    ["stderr", "_stderrBuffer"]
  ]) {
    const trailing = run[bufferProperty];
    if (!trailing) {
      continue;
    }

    if (trailing.startsWith(STRUCTURED_EVENT_PREFIX)) {
      try {
        const event = JSON.parse(trailing.slice(STRUCTURED_EVENT_PREFIX.length));
        applyStructuredEvent(run, event);
      } catch {
        run[stream] += trailing;
        addLogEntry(run, stream, trailing);
      }
    } else {
      run[stream] += trailing;
      addLogEntry(run, stream, trailing);
    }
    run[bufferProperty] = "";
  }
}

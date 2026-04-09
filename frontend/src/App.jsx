import { useEffect, useRef, useState } from "react";

const apiBase = "/api";

function stripAnsi(value) {
  if (!value) return "";
  return value.replace(/\u001b\[[0-9;]*m/g, "");
}

function extractDeviceCodePrompt(output) {
  const cleanOutput = stripAnsi(output);
  if (!cleanOutput) return null;

  const sentenceMatch = cleanOutput.match(
    /To sign in,\s*use a web browser to open the page\s*(https:\/\/[^\s]+)\s*and enter the code\s*([A-Z0-9-]+)\s*to authenticate\.?/i
  );
  if (sentenceMatch) {
    return {
      url: sentenceMatch[1],
      code: sentenceMatch[2]
    };
  }

  const urlMatch = cleanOutput.match(
    /https:\/\/(?:microsoft\.com\/devicelogin|login\.microsoft(?:online)?\.com\/device|login\.microsoftonline\.com\/common\/oauth2\/deviceauth)[^\s]*/i
  );
  const codeMatch = cleanOutput.match(/\b[A-Z0-9]{4,}(?:-[A-Z0-9]{4,})+\b|\b[A-Z0-9]{8,10}\b/);

  if (!urlMatch && !codeMatch) {
    return null;
  }

  return {
    url: urlMatch?.[0] || "https://microsoft.com/devicelogin",
    code: codeMatch?.[0] || ""
  };
}

async function parseApiResponse(response) {
  const contentType = response.headers.get("content-type") || "";

  if (contentType.includes("application/json")) {
    return response.json();
  }

  const text = await response.text();
  const compactText = text.replace(/\s+/g, " ").trim();
  throw new Error(
    `API returned ${response.status} ${response.statusText}: ${compactText.slice(0, 180) || "Empty response"}`
  );
}

function normalizeDefaults(fields) {
  return fields.reduce((acc, field) => {
    acc[field.id] = field.defaultValue ?? (field.type === "checkbox" ? false : "");
    return acc;
  }, {});
}

function formatDate(value) {
  if (!value) return "Pending";
  return new Date(value).toLocaleString();
}

function formatDuration(value) {
  if (!value && value !== 0) return "Pending";
  const totalSeconds = Math.max(0, Math.round(value / 1000));
  const minutes = Math.floor(totalSeconds / 60);
  const seconds = totalSeconds % 60;
  return `${minutes}m ${String(seconds).padStart(2, "0")}s`;
}

function formatFileSize(value) {
  const size = Number(value);
  if (!Number.isFinite(size) || size <= 0) {
    return "0 B";
  }

  const units = ["B", "KB", "MB", "GB", "TB"];
  const exponent = Math.min(Math.floor(Math.log(size) / Math.log(1024)), units.length - 1);
  const formatted = size / 1024 ** exponent;
  const digits = formatted >= 10 || exponent === 0 ? 0 : 1;
  return `${formatted.toFixed(digits)} ${units[exponent]}`;
}

async function copyText(text) {
  if (!text) return false;

  try {
    if (navigator.clipboard?.writeText) {
      await navigator.clipboard.writeText(text);
      return true;
    }
  } catch {
    // Fall back to document.execCommand below.
  }

  try {
    const textarea = document.createElement("textarea");
    textarea.value = text;
    textarea.setAttribute("readonly", "");
    textarea.style.position = "fixed";
    textarea.style.opacity = "0";
    textarea.style.pointerEvents = "none";
    document.body.appendChild(textarea);
    textarea.select();
    textarea.setSelectionRange(0, textarea.value.length);
    const copied = document.execCommand("copy");
    document.body.removeChild(textarea);
    return copied;
  } catch {
    return false;
  }
}

function groupScriptsByCategory(scripts) {
  return scripts.reduce((groups, script) => {
    const category = script.category || "Other";
    if (!groups[category]) {
      groups[category] = [];
    }
    groups[category].push(script);
    return groups;
  }, {});
}

function CategoryIcon({ category }) {
  const normalized = (category || "Other").toLowerCase();
  let path = "M12 3.5a8.5 8.5 0 1 0 8.5 8.5A8.51 8.51 0 0 0 12 3.5Zm0 4.25a1.25 1.25 0 1 1-1.25 1.25A1.25 1.25 0 0 1 12 7.75Zm2.25 8.5h-4.5v-1.5h.75v-2.25h-.75V11h3v3.75h1.5Z";

  if (normalized === "identity") {
    path = "M12 3.5c-2.9 0-5.25 2.35-5.25 5.25S9.1 14 12 14s5.25-2.35 5.25-5.25S14.9 3.5 12 3.5Zm0 12c-3.43 0-6.38 1.9-7.88 4.7l1.33.8c1.23-2.3 3.68-3.75 6.55-3.75s5.32 1.45 6.55 3.75l1.33-.8c-1.5-2.8-4.45-4.7-7.88-4.7Z";
  } else if (normalized === "exchange") {
    path = "M4 6.25A2.25 2.25 0 0 1 6.25 4h11.5A2.25 2.25 0 0 1 20 6.25v11.5A2.25 2.25 0 0 1 17.75 20H6.25A2.25 2.25 0 0 1 4 17.75ZM7 8v1.5h10V8Zm0 3.25v1.5h6.5v-1.5Zm0 3.25V16h8v-1.5Z";
  } else if (normalized === "security") {
    path = "M12 3.5 5.5 6v4.75c0 4.15 2.8 7.97 6.5 8.95 3.7-.98 6.5-4.8 6.5-8.95V6Zm0 4a2.5 2.5 0 0 1 2.5 2.5v.5h.5v4h-6v-4h.5V10A2.5 2.5 0 0 1 12 7.5Zm0 1.5A1 1 0 0 0 11 10v.5h2V10a1 1 0 0 0-1-1Z";
  } else if (normalized === "sharepoint") {
    path = "M7 5.25A3.75 3.75 0 1 0 10.75 9 3.75 3.75 0 0 0 7 5.25Zm10 2A2.75 2.75 0 1 0 19.75 10 2.75 2.75 0 0 0 17 7.25ZM8 14c-3.05 0-5.65 1.68-7 4.17L2.28 19C3.38 16.96 5.5 15.5 8 15.5c1.13 0 2.18.3 3.08.83l.73-1.31A7.97 7.97 0 0 0 8 14Zm9.25.5c-2.14 0-4.02 1.07-5.16 2.7l1.23.86c.87-1.25 2.31-2.06 3.93-2.06 1.28 0 2.45.5 3.33 1.32l1.02-1.1a6.25 6.25 0 0 0-4.35-1.72Z";
  } else if (normalized === "teams") {
    path = "M6 6.25A2.25 2.25 0 0 1 8.25 4h5.5A2.25 2.25 0 0 1 16 6.25v11.5A2.25 2.25 0 0 1 13.75 20h-5.5A2.25 2.25 0 0 1 6 17.75ZM9 8v1.5h1.75V16h1.5V9.5H14V8Zm9.25.5a1.75 1.75 0 1 0 0 3.5 1.75 1.75 0 0 0 0-3.5Zm-1.5 4.75a2.75 2.75 0 0 0-2.75 2.75V17h5.5v-1a2.75 2.75 0 0 0-2.75-2.75Z";
  } else if (normalized === "reporting") {
    path = "M5 19.25h14v1.5H3.5V5H5Zm2-2.5V10h1.75v6.75Zm4 0V7.5h1.75v9.25Zm4 0V12h1.75v4.75Z";
  } else if (normalized === "licensing") {
    path = "M12 3.75 5 6.5v5.25c0 4.1 2.72 7.87 7 8.75 4.28-.88 7-4.65 7-8.75V6.5Zm-1.25 5h2.5v2.5h2.5v2.5h-2.5v2.5h-2.5v-2.5h-2.5v-2.5h2.5Z";
  } else if (normalized === "incident response") {
    path = "M12 4 3.5 19.5h17ZM12 9a.75.75 0 0 1 .75.75v4.5a.75.75 0 0 1-1.5 0v-4.5A.75.75 0 0 1 12 9Zm0 8a1 1 0 1 1 0-2 1 1 0 0 1 0 2Z";
  } else if (normalized === "operations") {
    path = "m12 4 1.15 2.53 2.78.33-1.87 2 0.5 2.74L12 10.35 9.44 11.6l.5-2.74-1.87-2 2.78-.33Zm-6.5 8 1.15 2.53 2.78.33-1.87 2 .5 2.74L5.5 18.35 2.94 19.6l.5-2.74-1.87-2 2.78-.33Zm13 0 1.15 2.53 2.78.33-1.87 2 .5 2.74-2.56-1.25-2.56 1.25.5-2.74-1.87-2 2.78-.33Z";
  } else if (normalized === "collaboration") {
    path = "M8 6.25A2.25 2.25 0 1 0 10.25 8.5 2.25 2.25 0 0 0 8 6.25Zm8 0A2.25 2.25 0 1 0 18.25 8.5 2.25 2.25 0 0 0 16 6.25ZM4.75 17A3.25 3.25 0 0 1 8 13.75h.5A3.25 3.25 0 0 1 11.75 17v.75h-7Zm7.5.75V17a4.7 4.7 0 0 0-.77-2.56 3.23 3.23 0 0 1 1.77-.69H16A3.25 3.25 0 0 1 19.25 17v.75Z";
  } else if (normalized === "devices") {
    path = "M7.25 4h9.5A2.25 2.25 0 0 1 19 6.25v11.5A2.25 2.25 0 0 1 16.75 20h-9.5A2.25 2.25 0 0 1 5 17.75V6.25A2.25 2.25 0 0 1 7.25 4ZM11 16.5h2v-1.25h-2Z";
  }

  return (
    <span className="catalog-group-icon" aria-hidden="true">
      <svg viewBox="0 0 24 24" focusable="false">
        <path d={path} fill="currentColor" />
      </svg>
    </span>
  );
}

function getFavoriteScriptIds() {
  try {
    const raw = window.localStorage.getItem("m365-toolbox-favorites");
    return raw ? JSON.parse(raw) : [];
  } catch {
    return [];
  }
}

function getThemePreference() {
  try {
    return window.localStorage.getItem("m365-toolbox-theme") || "light";
  } catch {
    return "light";
  }
}

function Field({ field, value, onChange }) {
  if (field.type === "checkbox") {
    return (
      <label className="checkbox-field">
        <input
          type="checkbox"
          checked={Boolean(value)}
          onChange={(event) => onChange(field.id, event.target.checked)}
        />
        <span>{field.label}</span>
      </label>
    );
  }

  if (field.type === "textarea") {
    return (
      <label className="form-field">
        <span>{field.label}</span>
        <textarea
          rows={4}
          placeholder={field.placeholder || ""}
          value={value || ""}
          onChange={(event) => onChange(field.id, event.target.value)}
        />
        {field.helpText ? <small>{field.helpText}</small> : null}
      </label>
    );
  }

  if (field.type === "multiselect") {
    const selected = Array.isArray(value) ? value : [];
    return (
      <div className="form-field">
        <span>{field.label}</span>
        <div className="multiselect-grid">
          {field.options.map((option) => (
            <label key={option} className="checkbox-field">
              <input
                type="checkbox"
                checked={selected.includes(option)}
                onChange={(event) => {
                  const next = event.target.checked
                    ? [...selected, option]
                    : selected.filter((entry) => entry !== option);
                  onChange(field.id, next);
                }}
              />
              <span>{option}</span>
            </label>
          ))}
        </div>
      </div>
    );
  }

  return (
    <label className="form-field">
      <span>{field.label}</span>
      <input
        type={field.type === "number" ? "number" : field.type === "password" ? "password" : "text"}
        min={field.min}
        max={field.max}
        placeholder={field.placeholder || ""}
        value={value ?? ""}
        onChange={(event) => onChange(field.id, field.type === "number" ? Number(event.target.value) : event.target.value)}
      />
      {field.helpText ? <small>{field.helpText}</small> : null}
    </label>
  );
}

export function App() {
  const reportCardRef = useRef(null);
  const [sidebarWidth, setSidebarWidth] = useState(280);
  const [isResizingSidebar, setIsResizingSidebar] = useState(false);
  const [scripts, setScripts] = useState([]);
  const [selectedScript, setSelectedScript] = useState(null);
  const [formValues, setFormValues] = useState({});
  const [runs, setRuns] = useState([]);
  const [activeRun, setActiveRun] = useState(null);
  const [artifacts, setArtifacts] = useState([]);
  const [status, setStatus] = useState(null);
  const [error, setError] = useState("");
  const [success, setSuccess] = useState("");
  const [submitting, setSubmitting] = useState(false);
  const [canceling, setCanceling] = useState(false);
  const [runDetailsOpen, setRunDetailsOpen] = useState(true);
  const [recentRunsOpen, setRecentRunsOpen] = useState(true);
  const [devicePromptDismissed, setDevicePromptDismissed] = useState(false);
  const [expandedCategories, setExpandedCategories] = useState({});
  const [scriptSearch, setScriptSearch] = useState("");
  const [favoriteScriptIds, setFavoriteScriptIds] = useState(() => getFavoriteScriptIds());
  const [favoritesOnly, setFavoritesOnly] = useState(false);
  const [modeFilter, setModeFilter] = useState("all");
  const [theme, setTheme] = useState(() => getThemePreference());

  useEffect(() => {
    const load = async () => {
      const [scriptsResponse, runsResponse, statusResponse] = await Promise.all([
        fetch(`${apiBase}/scripts`),
        fetch(`${apiBase}/runs`),
        fetch(`${apiBase}/status`)
      ]);

      const scriptsData = await parseApiResponse(scriptsResponse);
      const runsData = await parseApiResponse(runsResponse);
      const statusData = await parseApiResponse(statusResponse);
      setScripts(scriptsData);
      setRuns(runsData);
      setStatus(statusData);
      setSelectedScript(null);
      setFormValues({});
      setExpandedCategories({});
    };

    load().catch((loadError) => setError(loadError.message));
  }, []);

  useEffect(() => {
    document.body.dataset.theme = theme;

    try {
      window.localStorage.setItem("m365-toolbox-theme", theme);
    } catch {
      // Ignore storage errors.
    }
  }, [theme]);

  useEffect(() => {
    if (!activeRun || !["running", "queued", "canceling"].includes(activeRun.status)) {
      return undefined;
    }

    const timer = window.setInterval(async () => {
      const response = await fetch(`${apiBase}/runs/${activeRun.id}`);
      const data = await parseApiResponse(response);
      setActiveRun(data);
      const runsResponse = await fetch(`${apiBase}/runs`);
      setRuns(await parseApiResponse(runsResponse));
    }, 2000);

    return () => window.clearInterval(timer);
  }, [activeRun]);

  useEffect(() => {
    if (!activeRun?.id) {
      setArtifacts([]);
      return undefined;
    }

    let cancelled = false;

    const loadArtifacts = async () => {
      try {
        const response = await fetch(`${apiBase}/runs/${activeRun.id}/artifacts`);
        const data = await parseApiResponse(response);
        if (!cancelled) {
          setArtifacts(data);
        }
      } catch {
        if (!cancelled) {
          setArtifacts([]);
        }
      }
    };

    loadArtifacts();

    return () => {
      cancelled = true;
    };
  }, [activeRun?.id, activeRun?.status]);

  useEffect(() => {
    if (activeRun?.status === "completed") {
      setSuccess(`Completed ${activeRun.scriptName} successfully.`);
    }
  }, [activeRun?.status, activeRun?.scriptName]);

  useEffect(() => {
    const hasHtmlArtifact = artifacts.some((artifact) => artifact.type === "html");
    if (activeRun?.id && hasHtmlArtifact) {
      setRunDetailsOpen(false);
      setRecentRunsOpen(false);
    }
  }, [activeRun?.id, artifacts]);

  useEffect(() => {
    const hasHtmlArtifact = artifacts.some((artifact) => artifact.type === "html");
    if (!activeRun?.id || !hasHtmlArtifact) {
      return;
    }

    const frame = window.requestAnimationFrame(() => {
      reportCardRef.current?.scrollIntoView({ behavior: "smooth", block: "start" });
      reportCardRef.current?.focus({ preventScroll: true });
    });

    return () => window.cancelAnimationFrame(frame);
  }, [activeRun?.id, artifacts]);

  useEffect(() => {
    if (activeRun?.status === "running") {
      setDevicePromptDismissed(false);
    }
  }, [activeRun?.id, activeRun?.status]);

  useEffect(() => {
    if (!isResizingSidebar) {
      return undefined;
    }

    const handlePointerMove = (event) => {
      const nextWidth = Math.min(Math.max(event.clientX, 220), 420);
      setSidebarWidth(nextWidth);
    };

    const handlePointerUp = () => {
      setIsResizingSidebar(false);
    };

    window.addEventListener("mousemove", handlePointerMove);
    window.addEventListener("mouseup", handlePointerUp);

    return () => {
      window.removeEventListener("mousemove", handlePointerMove);
      window.removeEventListener("mouseup", handlePointerUp);
    };
  }, [isResizingSidebar]);

  useEffect(() => {
    try {
      window.localStorage.setItem("m365-toolbox-favorites", JSON.stringify(favoriteScriptIds));
    } catch {
      // Ignore storage errors.
    }
  }, [favoriteScriptIds]);

  const handleScriptSelect = (script) => {
    setSelectedScript(script);
    setFormValues(normalizeDefaults(script.fields));
    setExpandedCategories({ [script.category || "Other"]: true });
    setError("");
    setSuccess("");
    setActiveRun(null);
    setArtifacts([]);
    setRunDetailsOpen(true);
    setRecentRunsOpen(true);
    setDevicePromptDismissed(false);
  };

  const handleOpenRun = (run) => {
    setArtifacts([]);
    setActiveRun(run ? { ...run } : null);
    setRunDetailsOpen(true);
    setRecentRunsOpen(true);
    setDevicePromptDismissed(false);
    const matchingScript = scripts.find((script) => script.id === run?.scriptId);
    if (matchingScript) {
      setSelectedScript(matchingScript);
    }
  };

  const handleChange = (fieldId, nextValue) => {
    setFormValues((current) => ({ ...current, [fieldId]: nextValue }));
  };

  const handleToggleFavorite = (scriptId) => {
    setFavoriteScriptIds((current) =>
      current.includes(scriptId)
        ? current.filter((id) => id !== scriptId)
        : [...current, scriptId]
    );
  };

  const handleSubmit = async (event) => {
    event.preventDefault();
    if (!selectedScript) return;

    const approvalConfirmed = !selectedScript.approvalRequired || window.confirm(
      `This script is marked as remediation and can change tenant state.\n\nDo you want to approve and launch "${selectedScript.name}"?`
    );

    if (!approvalConfirmed) {
      return;
    }

    setSubmitting(true);
    setError("");
    setSuccess("");

    try {
      const response = await fetch(`${apiBase}/scripts/${selectedScript.id}/run`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ ...formValues, approvalConfirmed })
      });

      const data = await parseApiResponse(response);
      if (!response.ok) {
        throw new Error(data.message || "Failed to start run.");
      }
      setActiveRun(data);
      setDevicePromptDismissed(false);
      setRunDetailsOpen(true);
      setRecentRunsOpen(true);
      setSuccess(data.status === "queued" ? "Run queued successfully." : "Run started successfully.");
      const runsResponse = await fetch(`${apiBase}/runs`);
      setRuns(await parseApiResponse(runsResponse));
    } catch (submitError) {
      setError(submitError.message);
    } finally {
      setSubmitting(false);
    }
  };

  const handleCancelRun = async () => {
    if (!activeRun?.id) return;

    setCanceling(true);
    setError("");
    setSuccess("");

    try {
      const response = await fetch(`${apiBase}/runs/${activeRun.id}/cancel`, {
        method: "POST"
      });
      const data = await parseApiResponse(response);
      if (!response.ok) {
        throw new Error(data.message || "Failed to cancel run.");
      }
      setActiveRun(data);
      setSuccess("Run cancellation requested.");
      const runsResponse = await fetch(`${apiBase}/runs`);
      setRuns(await parseApiResponse(runsResponse));
    } catch (cancelError) {
      setError(cancelError.message);
    } finally {
      setCanceling(false);
    }
  };

  const devicePrompt = extractDeviceCodePrompt([activeRun?.stdout, activeRun?.stderr].filter(Boolean).join("\n"));
  const showDevicePrompt = Boolean(devicePrompt?.code) && activeRun?.status === "running" && !devicePromptDismissed;
  const normalizedSearch = scriptSearch.trim().toLowerCase();
  const runCountsByScriptId = runs.reduce((acc, run) => {
    acc[run.scriptId] = (acc[run.scriptId] || 0) + 1;
    return acc;
  }, {});
  const favoriteScripts = scripts.filter((script) => favoriteScriptIds.includes(script.id));
  const recentFavoriteScripts = favoriteScripts.slice(0, 3);
  const mostUsedScripts = [...scripts]
    .filter((script) => runCountsByScriptId[script.id])
    .sort((a, b) => (runCountsByScriptId[b.id] || 0) - (runCountsByScriptId[a.id] || 0))
    .slice(0, 3);
  const hasHtmlArtifact = artifacts.some((artifact) => artifact.type === "html");
  const filteredScripts = scripts.filter((script) => {
    const matchesSearch = !normalizedSearch || [
      script.name,
      script.category,
      script.summary,
      script.description,
      script.id
      ]
      .filter(Boolean)
      .some((value) => value.toLowerCase().includes(normalizedSearch));
    const matchesFavorite = !favoritesOnly || favoriteScriptIds.includes(script.id);
    const matchesMode = modeFilter === "all" || script.mode === modeFilter;
    return matchesSearch && matchesFavorite && matchesMode;
  });
  const scriptGroups = groupScriptsByCategory(filteredScripts);
  const sortedCategories = Object.keys(scriptGroups).sort((a, b) => a.localeCompare(b));
  return (
    <div className="app-shell">
      {showDevicePrompt ? (
        <div className="modal-backdrop" role="presentation">
          <div className="modal-card" role="dialog" aria-modal="true" aria-labelledby="device-auth-title">
            <div className="card-header">
              <span className="card-title" id="device-auth-title">Microsoft Authentication Required</span>
              <button type="button" className="modal-close" onClick={() => setDevicePromptDismissed(true)}>
                Close
              </button>
            </div>
            <div className="card-body modal-body">
              <p className="empty-sub">Open the Microsoft device login page and enter this code to continue the script run.</p>
              <div className="device-code-row">
                <div className="device-code-box">{devicePrompt.code}</div>
                <button
                  type="button"
                  className="device-copy-btn"
                  onClick={async () => {
                    const copied = await copyText(devicePrompt.code);
                    setSuccess(copied ? "Device code copied." : "Unable to copy device code automatically.");
                  }}
                  aria-label="Copy device code"
                  title="Copy device code"
                >
                  <svg viewBox="0 0 24 24" aria-hidden="true">
                    <path d="M9 9.75A2.25 2.25 0 0 1 11.25 7.5h7.5A2.25 2.25 0 0 1 21 9.75v9A2.25 2.25 0 0 1 18.75 21h-7.5A2.25 2.25 0 0 1 9 18.75Zm-6-4.5A2.25 2.25 0 0 1 5.25 3h7.5A2.25 2.25 0 0 1 15 5.25V6.5h-1.5V5.25a.75.75 0 0 0-.75-.75h-7.5a.75.75 0 0 0-.75.75v9a.75.75 0 0 0 .75.75H6.5v1.5H5.25A2.25 2.25 0 0 1 3 14.25Z" fill="currentColor" />
                  </svg>
                </button>
              </div>
              <a className="filter-btn active-all" href={devicePrompt.url} target="_blank" rel="noreferrer">
                Open Microsoft Device Login
              </a>
            </div>
          </div>
        </div>
      ) : null}

      <header className="topbar">
        <div className="topbar-logo">M365 Toolbox</div>
        <div className="topbar-title">Web-based PowerShell operations for Microsoft 365</div>
        <div className="topbar-right">
          <button
            type="button"
            className={`theme-switch${theme === "dark" ? " dark" : ""}`}
            onClick={() => setTheme((current) => current === "dark" ? "light" : "dark")}
            aria-label={theme === "dark" ? "Switch to light mode" : "Switch to dark mode"}
            aria-pressed={theme === "dark"}
          >
            <span className="theme-switch-icon sun" aria-hidden="true">
              <svg viewBox="0 0 24 24">
                <path d="M12 6.5A5.5 5.5 0 1 0 17.5 12 5.51 5.51 0 0 0 12 6.5Zm0-3.5h1.5v2.25H12Zm0 15.75h1.5V21H12ZM3 11.25h2.25v1.5H3Zm15.75 0H21v1.5h-2.25ZM5.47 4.41l1.06-1.06 1.59 1.59-1.06 1.06Zm10.41 10.41 1.06-1.06 1.59 1.59-1.06 1.06ZM4.41 18.53l-1.06-1.06 1.59-1.59 1.06 1.06Zm10.41-10.41-1.06-1.06 1.59-1.59 1.06 1.06Z" fill="currentColor" />
              </svg>
            </span>
            <span className="theme-switch-track">
              <span className="theme-switch-thumb" />
            </span>
            <span className="theme-switch-icon moon" aria-hidden="true">
              <svg viewBox="0 0 24 24">
                <path d="M14.73 3.2a8.84 8.84 0 0 0 0 17.6 8.5 8.5 0 0 0 6.07-2.54 9.62 9.62 0 0 1-3.02.48A9.75 9.75 0 0 1 8.04 9a9.62 9.62 0 0 1 .48-3.02A8.5 8.5 0 0 0 5.98 12a8.84 8.84 0 0 0 8.75 8.8A8.84 8.84 0 0 0 14.73 3.2Z" fill="currentColor" />
              </svg>
            </span>
          </button>
          <div className="topbar-count">{scripts.length} script{scripts.length === 1 ? "" : "s"}</div>
          <div className="topbar-count">{runs.length} run{runs.length === 1 ? "" : "s"}</div>
        </div>
      </header>

      <div className="layout" style={{ "--sidebar-w": `${sidebarWidth}px` }}>
        <aside className="sidebar">
          <div className="sidebar-header">
            <div className="sidebar-label">Script Catalog</div>
            <label className="sidebar-search">
              <input
                type="text"
                placeholder="Search scripts..."
                value={scriptSearch}
                onChange={(event) => setScriptSearch(event.target.value)}
              />
            </label>
            <div className="sidebar-filter-group">
              <div className="sidebar-filter-label">Filters</div>
              <div className="chip-row">
                <button
                  type="button"
                  className={favoritesOnly ? "filter-chip active" : "filter-chip"}
                  onClick={() => setFavoritesOnly((current) => !current)}
                >
                  Favorites Only
                </button>
                <button type="button" className={modeFilter === "all" ? "filter-chip active" : "filter-chip"} onClick={() => setModeFilter("all")}>
                  All Modes
                </button>
                <button type="button" className={modeFilter === "read-only" ? "filter-chip active" : "filter-chip"} onClick={() => setModeFilter("read-only")}>
                  Read-only
                </button>
                <button type="button" className={modeFilter === "remediation" ? "filter-chip active" : "filter-chip"} onClick={() => setModeFilter("remediation")}>
                  Remediation
                </button>
              </div>
            </div>
          </div>
          <div className="tenant-list">
            {sortedCategories.length === 0 ? (
              <div className="empty-row">No scripts match your search.</div>
            ) : null}
            {sortedCategories.map((category) => {
              const isExpanded = normalizedSearch ? true : Boolean(expandedCategories[category]);
              return (
                <div key={category} className="catalog-group">
                  <button
                    type="button"
                    className={`catalog-group-header${isExpanded ? " expanded" : ""}`}
                    onClick={() =>
                      setExpandedCategories(isExpanded ? {} : { [category]: true })
                    }
                  >
                    <span className="catalog-group-title">
                      <CategoryIcon category={category} />
                      <span>{category}</span>
                    </span>
                    <span className="catalog-group-count">{scriptGroups[category].length}</span>
                    <span className="catalog-group-chevron">{isExpanded ? "▾" : "▸"}</span>
                  </button>

                  {isExpanded ? (
                    <div className="catalog-group-items">
                      {scriptGroups[category].map((script) => (
                        <div
                          key={script.id}
                          className={selectedScript?.id === script.id ? "tenant-item active" : "tenant-item"}
                          onClick={() => handleScriptSelect(script)}
                          role="button"
                          tabIndex={0}
                          onKeyDown={(event) => {
                            if (event.key === "Enter" || event.key === " ") {
                              handleScriptSelect(script);
                            }
                          }}
                        >
                          <div className="tenant-avatar">{script.name.slice(0, 2).toUpperCase()}</div>
                          <div className="tenant-info">
                            <div className="tenant-name">{script.name}</div>
                            <div className="tenant-meta">{script.summary}</div>
                            <div className="tenant-tags">
                              <span className={`mini-pill ${script.mode === "remediation" ? "badge-crit" : "badge-ok"}`}>{script.mode}</span>
                              <span className="mini-pill badge-neutral">{script.estimatedRuntimeMinutes} min</span>
                            </div>
                          </div>
                          <button
                            type="button"
                            className={favoriteScriptIds.includes(script.id) ? "favorite-btn active" : "favorite-btn"}
                            onClick={(event) => {
                              event.stopPropagation();
                              handleToggleFavorite(script.id);
                            }}
                          >
                            Fav
                          </button>
                        </div>
                      ))}
                    </div>
                  ) : null}
                </div>
              );
            })}
          </div>

          <div className="sidebar-footer">
            <div>{runs.length === 0 ? "No runs yet." : `${runs.length} tracked run${runs.length === 1 ? "" : "s"}`}</div>
            <div className="sidebar-footer-links">
              <a
                className="sidebar-repo-link"
                href="https://github.com/albercuba/M365-Toolbox"
                target="_blank"
                rel="noreferrer"
                aria-label="Open GitHub repository"
              >
                <svg viewBox="0 0 24 24" aria-hidden="true">
                  <path
                    d="M12 2C6.48 2 2 6.58 2 12.24c0 4.53 2.87 8.37 6.84 9.73.5.1.68-.22.68-.49 0-.24-.01-1.04-.01-1.89-2.78.62-3.37-1.21-3.37-1.21-.45-1.2-1.11-1.52-1.11-1.52-.91-.64.07-.63.07-.63 1 .07 1.53 1.06 1.53 1.06.9 1.58 2.35 1.12 2.92.86.09-.67.35-1.12.63-1.37-2.22-.26-4.56-1.15-4.56-5.13 0-1.13.39-2.05 1.03-2.77-.1-.26-.45-1.31.1-2.74 0 0 .84-.28 2.75 1.06A9.3 9.3 0 0 1 12 6.84c.85 0 1.71.12 2.51.35 1.91-1.34 2.75-1.06 2.75-1.06.55 1.43.2 2.48.1 2.74.64.72 1.03 1.64 1.03 2.77 0 3.99-2.34 4.87-4.57 5.12.36.32.68.95.68 1.92 0 1.39-.01 2.51-.01 2.85 0 .27.18.59.69.49A10.26 10.26 0 0 0 22 12.24C22 6.58 17.52 2 12 2Z"
                    fill="currentColor"
                  />
                </svg>
                <span>GitHub</span>
              </a>
              <a
                className="sidebar-repo-link"
                href="https://www.paypal.com/donate/?hosted_button_id=VBBHTH9XQ5CA2"
                target="_blank"
                rel="noreferrer"
                aria-label="Buy me a coffee with PayPal"
              >
                <svg viewBox="0 0 24 24" aria-hidden="true">
                  <path
                    d="M7.2 4h7.68c2.98 0 5.02 1.04 5.02 4.03 0 1.98-.75 3.48-2.18 4.38-1.07.68-2.49 1-4.22 1h-2.4L10.18 20H6.6l.94-5.35L9 6.3h3.95c1.42 0 2.4.16 2.94.56.47.35.68.88.68 1.62 0 .92-.31 1.64-.91 2.12-.62.5-1.56.74-2.84.74H11.4l-.33 1.9h1.71c2.46 0 4.39-.47 5.76-1.54 1.43-1.11 2.16-2.78 2.16-5 0-1.56-.53-2.78-1.56-3.57C18.13 2.38 16.66 2 14.7 2H8.77L7.2 4Z"
                    fill="currentColor"
                  />
                </svg>
                <span>Buy me a coffee</span>
              </a>
            </div>
          </div>
        </aside>
        <div
          className={`sidebar-resizer${isResizingSidebar ? " active" : ""}`}
          role="separator"
          aria-orientation="vertical"
          aria-label="Resize script catalog"
          onMouseDown={() => setIsResizingSidebar(true)}
        />

        <main className="main">
          {error ? (
            <div className="flash-wrap">
              <div className="flash flash-error">{error}</div>
            </div>
          ) : null}
          {success ? (
            <div className="flash-wrap">
              <div className="flash flash-success">{success}</div>
            </div>
          ) : null}

          {selectedScript ? (
            <>
              <div className="dash-topstrip">
                <div className="strip-item">
                  <div className="strip-label">Script</div>
                  <div className="strip-value">{selectedScript.name}</div>
                </div>
                <div className="strip-item">
                  <div className="strip-label">Mode</div>
                  <div className="strip-value">{selectedScript.mode}</div>
                </div>
                <div className="strip-item">
                  <div className="strip-label">Estimated Runtime</div>
                  <div className="strip-value">{selectedScript.estimatedRuntimeMinutes} min</div>
                </div>
                <div className="strip-item">
                  <div className="strip-label">Script Runs</div>
                  <div className="strip-value">{runCountsByScriptId[selectedScript.id] || 0}</div>
                </div>
              </div>

              <div className="dash-page">
                <div className="sections">
                  <div className="card">
                    <div className="card-header">
                      <span className="card-title">Script Details</span>
                      <span className={`card-badge ${selectedScript.mode === "remediation" ? "badge-crit" : "badge-ok"}`}>{selectedScript.mode}</span>
                      {selectedScript.approvalRequired ? <span className="card-badge badge-warn">approval required</span> : null}
                    </div>
                    <div className="card-body">
                      <div className="method-grid">
                        <div className="method-item method-item-selected">
                          <div className="method-info">
                            <div className="method-label">Summary</div>
                            <div className="method-count">{selectedScript.summary}</div>
                          </div>
                        </div>
                      </div>
                      <div className="empty-row" style={{ marginTop: "0.85rem" }}>{selectedScript.description}</div>
                    </div>
                  </div>

                  <div className="manage-workspace">
                    <div className="sections">
                      <div className="card">
                        <div className="card-header">
                          <span className="card-title">Run Script</span>
                          <span className={`card-badge ${selectedScript.mode === "remediation" ? "badge-crit" : "badge-ok"}`}>{selectedScript.mode}</span>
                        </div>
                        <div className="card-body">
                          {selectedScript.approvalRequired ? (
                            <div className="approval-banner">This remediation workflow requires an approval confirmation before launch.</div>
                          ) : null}
                          <form className="settings-row" onSubmit={handleSubmit}>
                            {selectedScript.fields.map((field) => (
                              <Field key={field.id} field={field} value={formValues[field.id]} onChange={handleChange} />
                            ))}
                            <button className="add-btn" type="submit" disabled={submitting}>
                              {submitting ? "Starting..." : selectedScript.approvalRequired ? "Approve and Run" : "Run in Toolbox"}
                            </button>
                          </form>
                        </div>
                      </div>

                      <div className={`card ${runDetailsOpen ? "" : "card-collapsed"}`}>
                        <button type="button" className="card-header card-header-button" onClick={() => setRunDetailsOpen((current) => !current)}>
                          <span className="card-title">Run Details</span>
                          <span className={`card-badge ${activeRun ? (activeRun.status === "completed" ? "badge-ok" : activeRun.status === "failed" || activeRun.status === "canceled" || activeRun.status === "interrupted" ? "badge-crit" : "badge-warn") : "badge-neutral"}`}>
                            {activeRun ? activeRun.status : "idle"}
                          </span>
                          <span className="card-chevron">{runDetailsOpen ? "▾" : "▸"}</span>
                        </button>
                        {runDetailsOpen ? (
                        <div className="card-body">
                          {!activeRun ? <div className="empty-row">Select a recent run or start a new one to inspect the output.</div> : null}
                          {activeRun ? (
                            <>
                              <div className="quick-summary-grid">
                                <div className="quick-summary-item">
                                  <div className="method-label">Requested</div>
                                  <div className="method-count">{formatDate(activeRun.requestedAt)}</div>
                                </div>
                                <div className="quick-summary-item">
                                  <div className="method-label">Started</div>
                                  <div className="method-count">{formatDate(activeRun.startedAt)}</div>
                                </div>
                                <div className="quick-summary-item">
                                  <div className="method-label">Finished</div>
                                  <div className="method-count">{formatDate(activeRun.finishedAt)}</div>
                                </div>
                                <div className="quick-summary-item">
                                  <div className="method-label">Duration</div>
                                  <div className="method-count">{formatDuration(activeRun.durationMs)}</div>
                                </div>
                                <div className="quick-summary-item">
                                  <div className="method-label">Current Step</div>
                                  <div className="method-count">{activeRun.currentStep || "Waiting for logs"}</div>
                                </div>
                                <div className="quick-summary-item">
                                  <div className="method-label">Exit Code</div>
                                  <div className="method-count">{activeRun.exitCode ?? "Pending"}</div>
                                </div>
                              </div>
                              {activeRun.errorSummary ? (
                                <div className="flash flash-error soft">{activeRun.errorSummary}</div>
                              ) : null}
                              <div className="manage-form-panel" style={{ marginTop: "1rem" }}>
                                <div className="panel-toolbar">
                                  <h4>Command</h4>
                                  <div className="run-actions">
                                  <button
                                    type="button"
                                    className="filter-btn active-all"
                                    onClick={async () => {
                                      const copied = await copyText(activeRun.command);
                                      setSuccess(copied ? "Command copied." : "Unable to copy command automatically.");
                                    }}
                                  >
                                    Copy Command
                                  </button>
                                    {["running", "queued", "canceling"].includes(activeRun.status) ? (
                                      <button type="button" className="filter-btn destructive" onClick={handleCancelRun} disabled={canceling}>
                                        {canceling ? "Canceling..." : "Cancel Run"}
                                      </button>
                                    ) : null}
                                  </div>
                                </div>
                                <pre className="manage-response">{activeRun.command}</pre>
                              </div>
                              <div className="manage-form-panel">
                                <h4>Structured Log</h4>
                                <div className="run-log-list">
                                  {activeRun.logs?.length ? activeRun.logs.slice(-12).reverse().map((entry) => (
                                    <div key={entry.id} className={`run-log-entry level-${entry.level}`}>
                                      <span className="run-log-time">{formatDate(entry.timestamp)}</span>
                                      <span className="run-log-message">{entry.message}</span>
                                    </div>
                                  )) : (
                                    <div className="empty-row compact">No structured log entries yet.</div>
                                  )}
                                </div>
                              </div>
                              <div className="manage-form-panel">
                                <h4>Stdout</h4>
                                <pre className="manage-response">{activeRun.stdout || "No stdout yet."}</pre>
                              </div>
                              <div className="manage-form-panel">
                                <h4>Stderr</h4>
                                <pre className="manage-response">{activeRun.stderr || "No stderr."}</pre>
                              </div>
                            </>
                          ) : null}
                        </div>
                        ) : null}
                      </div>

                      <div className="card">
                        <div className="card-header">
                          <span className="card-title">Artifacts</span>
                          <span className="card-badge badge-neutral">{artifacts.length}</span>
                        </div>
                        <div className="card-body">
                          {artifacts.length === 0 ? (
                            <div className="empty-row">Artifacts will appear here when a run exports HTML, CSV, XLSX, log, or text files.</div>
                          ) : (
                            <div className="artifact-list">
                              {artifacts.map((artifact) => (
                                <div key={artifact.id} className="artifact-item">
                                  <div>
                                    <div className="artifact-name">{artifact.name}</div>
                                    <div className="artifact-meta">
                                      {artifact.type.toUpperCase()} | {formatFileSize(artifact.size)} | {formatDate(artifact.createdAt)}
                                    </div>
                                  </div>
                                  <div className="artifact-actions">
                                    <a className="filter-btn active-all" href={`${apiBase}/runs/${activeRun?.id}/artifacts/${encodeURIComponent(artifact.id)}`}>
                                      Download
                                    </a>
                                    {artifact.type === "html" ? (
                                      <a className="filter-btn" href={`${apiBase}/runs/${activeRun?.id}/html`} target="_blank" rel="noreferrer">
                                        Preview
                                      </a>
                                    ) : null}
                                  </div>
                                </div>
                              ))}
                            </div>
                          )}
                        </div>
                      </div>

                      <div className={`card ${recentRunsOpen ? "" : "card-collapsed"}`}>
                        <button type="button" className="card-header card-header-button" onClick={() => setRecentRunsOpen((current) => !current)}>
                          <span className="card-title">Recent Runs</span>
                          <span className="card-badge badge-neutral">{runs.length}</span>
                          <span className="card-chevron">{recentRunsOpen ? "▾" : "▸"}</span>
                        </button>
                        {recentRunsOpen ? (
                        <div className="card-body">
                          {runs.length === 0 ? (
                            <div className="empty-row">No runs yet. Launch a report to start building persistent history.</div>
                          ) : (
                            <div className="table-scroll">
                              <table>
                                <thead>
                                  <tr>
                                    <th>Script</th>
                                    <th>Status</th>
                                    <th>Requested</th>
                                    <th>Action</th>
                                  </tr>
                                </thead>
                                <tbody>
                                  {runs.map((run) => (
                                    <tr key={run.id}>
                                      <td>{run.scriptName}</td>
                                      <td>
                                        <span className={`pill ${run.status === "completed" ? "badge-ok" : run.status === "failed" || run.status === "canceled" || run.status === "interrupted" ? "badge-crit" : "badge-warn"}`}>
                                          {run.status}
                                        </span>
                                      </td>
                                      <td>{formatDate(run.requestedAt || run.startedAt)}</td>
                                      <td className="table-actions">
                                        <button type="button" className="filter-btn active-all" onClick={() => handleOpenRun(run)}>
                                          Open
                                        </button>
                                      </td>
                                    </tr>
                                  ))}
                                </tbody>
                              </table>
                            </div>
                          )}
                        </div>
                        ) : null}
                      </div>

                      {activeRun?.id && hasHtmlArtifact ? (
                        <div className="card" ref={reportCardRef} tabIndex={-1}>
                          <div className="card-header">
                            <span className="card-title">HTML Report</span>
                            <span className="card-badge badge-ok">preview</span>
                            {activeRun?.id ? (
                              <a className="filter-btn active-all" href={`${apiBase}/runs/${activeRun.id}/html`} download={`m365-mfa-report-${activeRun.id}.html`}>
                                Download
                              </a>
                            ) : null}
                          </div>
                          <div className="card-body report-card-body">
                            <iframe title="HTML report preview" className="report-frame" src={`${apiBase}/runs/${activeRun.id}/html`} />
                          </div>
                        </div>
                      ) : null}
                    </div>
                  </div>
                </div>
              </div>
            </>
          ) : (
            <div className="dash-page">
              <div className="sections">
                <div className="card">
                  <div className="card-header">
                    <span className="card-title">Dashboard</span>
                    <span className="card-badge badge-neutral">{scripts.length} scripts</span>
                  </div>
                  <div className="card-body">
                    <div className="method-grid">
                      <div className="method-item method-item-selected">
                        <div className="method-info">
                          <div className="method-label">Catalog</div>
                          <div className="method-count">{scripts.length} scripts across {Object.keys(groupScriptsByCategory(scripts)).length} categories.</div>
                        </div>
                      </div>
                      <div className="method-item">
                        <div className="method-info">
                          <div className="method-label">Persistent Runs</div>
                          <div className="method-count">{runs.length} tracked runs survive backend restarts.</div>
                        </div>
                      </div>
                      <div className="method-item">
                        <div className="method-info">
                          <div className="method-label">Favorites</div>
                          <div className="method-count">{favoriteScriptIds.length} saved favorites ready for quick access.</div>
                        </div>
                      </div>
                      <div className="method-item">
                        <div className="method-info">
                          <div className="method-label">Backend Status</div>
                          <div className="method-count">
                            {status?.powerShell?.available ? "PowerShell ready" : "Status unavailable"} | {status?.modules?.ready ? "Modules ready" : "Module check pending"}
                          </div>
                        </div>
                      </div>
                    </div>
                  </div>
                </div>

                <div className="card">
                  <div className="card-header">
                    <span className="card-title">Shortcuts</span>
                    <button type="button" className="filter-btn" onClick={() => fetch(`${apiBase}/status`).then(parseApiResponse).then(setStatus).catch((loadError) => setError(loadError.message))}>
                      Refresh Status
                    </button>
                  </div>
                  <div className="card-body">
                    <div className="shortcut-grid">
                      <div className="shortcut-card">
                        <div className="method-label">Recent Favorites</div>
                        {recentFavoriteScripts.length ? recentFavoriteScripts.map((script) => (
                          <button key={script.id} type="button" className="shortcut-link" onClick={() => handleScriptSelect(script)}>
                            {script.name}
                          </button>
                        )) : <div className="empty-row compact">Mark a script as favorite to pin it here.</div>}
                      </div>
                      <div className="shortcut-card">
                        <div className="method-label">Most Used</div>
                        {mostUsedScripts.length ? mostUsedScripts.map((script) => (
                          <button key={script.id} type="button" className="shortcut-link" onClick={() => handleScriptSelect(script)}>
                            {script.name}
                          </button>
                        )) : <div className="empty-row compact">Most-used shortcuts appear after a few runs.</div>}
                      </div>
                      <div className="shortcut-card">
                        <div className="method-label">Health Snapshot</div>
                        <div className="status-list">
                          <div>Output path: {status?.paths?.outputWritable ? "writable" : "unavailable"}</div>
                          <div>Scripts mount: {status?.paths?.scriptsMounted ? "mounted" : "missing"}</div>
                          <div>PowerShell: {status?.powerShell?.available ? `v${status.powerShell.version}` : "not ready"}</div>
                          <div>Modules: {status?.modules?.ready ? "ready" : "check failed or pending"}</div>
                        </div>
                      </div>
                    </div>
                  </div>
                </div>

                <div className="card">
                  <div className="card-header">
                    <span className="card-title">Start Here</span>
                    <span className="card-badge badge-neutral">home</span>
                  </div>
                  <div className="card-body">
                    <div className="empty-row">Browse categories on the left, use filters to separate read-only and remediation scripts, and open any script to see prerequisites, examples, and runtime expectations before launch.</div>
                  </div>
                </div>

              </div>
            </div>
          )}
        </main>
      </div>
    </div>
  );
}

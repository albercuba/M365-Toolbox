import { useEffect, useState } from "react";

const apiBase = "/api";

function extractDeviceCodePrompt(stdout) {
  if (!stdout) return null;

  const match = stdout.match(
    /To sign in, use a web browser to open the page (https:\/\/[^\s]+) and enter the code ([A-Z0-9-]+) to authenticate\./i
  );

  if (!match) return null;

  return {
    url: match[1],
    code: match[2]
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
  const [scripts, setScripts] = useState([]);
  const [selectedScript, setSelectedScript] = useState(null);
  const [formValues, setFormValues] = useState({});
  const [runs, setRuns] = useState([]);
  const [activeRun, setActiveRun] = useState(null);
  const [activeRunHtml, setActiveRunHtml] = useState("");
  const [error, setError] = useState("");
  const [submitting, setSubmitting] = useState(false);
  const [runDetailsOpen, setRunDetailsOpen] = useState(true);
  const [recentRunsOpen, setRecentRunsOpen] = useState(true);
  const [devicePromptDismissed, setDevicePromptDismissed] = useState(false);

  useEffect(() => {
    const load = async () => {
      const [scriptsResponse, runsResponse] = await Promise.all([
        fetch(`${apiBase}/scripts`),
        fetch(`${apiBase}/runs`)
      ]);

      const scriptsData = await parseApiResponse(scriptsResponse);
      const runsData = await parseApiResponse(runsResponse);
      setScripts(scriptsData);
      setRuns(runsData);

      if (scriptsData.length > 0) {
        setSelectedScript(scriptsData[0]);
        setFormValues(normalizeDefaults(scriptsData[0].fields));
      }
    };

    load().catch((loadError) => setError(loadError.message));
  }, []);

  useEffect(() => {
    if (!activeRun || activeRun.status !== "running") {
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
    let cancelled = false;

    const loadRunHtml = async () => {
      if (!activeRun?.id || activeRun.status !== "completed" || !activeRun.artifacts?.htmlPath) {
        setActiveRunHtml("");
        return;
      }

      try {
        const response = await fetch(`${apiBase}/runs/${activeRun.id}/html`);
        if (!response.ok) {
          setActiveRunHtml("");
          return;
        }

        const html = await response.text();
        if (!cancelled) {
          setActiveRunHtml(html);
        }
      } catch {
        if (!cancelled) {
          setActiveRunHtml("");
        }
      }
    };

    loadRunHtml();

    return () => {
      cancelled = true;
    };
  }, [activeRun]);

  useEffect(() => {
    if (activeRun?.status === "completed") {
      setRunDetailsOpen(false);
      setRecentRunsOpen(false);
    }
  }, [activeRun?.status]);

  useEffect(() => {
    if (activeRun?.status === "running") {
      setDevicePromptDismissed(false);
    }
  }, [activeRun?.id, activeRun?.status]);

  const handleScriptSelect = (script) => {
    setSelectedScript(script);
    setFormValues(normalizeDefaults(script.fields));
    setError("");
    setActiveRun(null);
    setActiveRunHtml("");
    setRunDetailsOpen(true);
    setRecentRunsOpen(true);
    setDevicePromptDismissed(false);
  };

  const handleOpenRun = (run) => {
    setActiveRunHtml("");
    setActiveRun(run ? { ...run } : null);
    setRunDetailsOpen(true);
    setRecentRunsOpen(true);
    setDevicePromptDismissed(false);
  };

  const handleChange = (fieldId, nextValue) => {
    setFormValues((current) => ({ ...current, [fieldId]: nextValue }));
  };

  const handleSubmit = async (event) => {
    event.preventDefault();
    if (!selectedScript) return;

    setSubmitting(true);
    setError("");

    try {
      const response = await fetch(`${apiBase}/scripts/${selectedScript.id}/run`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(formValues)
      });

      const data = await parseApiResponse(response);
      if (!response.ok) {
        throw new Error(data.message || "Failed to start run.");
      }

      setActiveRunHtml("");
      setActiveRun(data);
      setDevicePromptDismissed(false);
      setRunDetailsOpen(true);
      setRecentRunsOpen(true);
      const runsResponse = await fetch(`${apiBase}/runs`);
      setRuns(await parseApiResponse(runsResponse));
    } catch (submitError) {
      setError(submitError.message);
    } finally {
      setSubmitting(false);
    }
  };

  const devicePrompt = extractDeviceCodePrompt(activeRun?.stdout);
  const showDevicePrompt = Boolean(devicePrompt) && activeRun?.status === "running" && !devicePromptDismissed;

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
              <div className="device-code-box">{devicePrompt.code}</div>
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
          <div className="topbar-count">{scripts.length} script{scripts.length === 1 ? "" : "s"}</div>
          <div className="topbar-count">{runs.length} run{runs.length === 1 ? "" : "s"}</div>
        </div>
      </header>

      <div className="layout">
        <aside className="sidebar">
          <div className="sidebar-header">
            <div className="sidebar-label">Script Catalog</div>
          </div>
          <div className="tenant-list">
            {scripts.map((script) => (
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
                  <div className="tenant-meta">{script.category}</div>
                </div>
              </div>
            ))}
          </div>

          <div className="sidebar-footer">
            {runs.length === 0 ? "No runs yet." : `${runs.length} tracked run${runs.length === 1 ? "" : "s"}`}
          </div>
        </aside>

        <main className="main">
          {error ? (
            <div className="flash-wrap">
              <div className="flash flash-error">{error}</div>
            </div>
          ) : null}

          {selectedScript ? (
            <>
              <div className="dash-topstrip">
                <div className="strip-item">
                  <div className="strip-label">Toolbox</div>
                  <div className="strip-value">M365 Toolbox</div>
                </div>
                <div className="strip-item">
                  <div className="strip-label">Script</div>
                  <div className="strip-value">{selectedScript.name}</div>
                </div>
                <div className="strip-item">
                  <div className="strip-label">Category</div>
                  <div className="strip-value">{selectedScript.category}</div>
                </div>
                <div className="strip-item">
                  <div className="strip-label">Recent Runs</div>
                  <div className="strip-value">{runs.length}</div>
                </div>
              </div>

              <div className="dash-page">
                <div className="sections">
                  <div className="card">
                    <div className="card-header">
                      <span className="card-title">Script Overview</span>
                      <span className="card-badge badge-neutral">{selectedScript.id}</span>
                    </div>
                    <div className="card-body">
                      <div className="method-grid">
                        <div className="method-item method-item-selected">
                          <div className="method-info">
                            <div className="method-label">Summary</div>
                            <div className="method-count">{selectedScript.summary}</div>
                          </div>
                        </div>
                        <div className="method-item">
                          <div className="method-info">
                            <div className="method-label">Growth Model</div>
                            <div className="method-count">Registry-based script catalog for future M365 tools</div>
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
                          <span className="card-badge badge-ok">{submitting ? "Starting" : "Ready"}</span>
                        </div>
                        <div className="card-body">
                          <form className="settings-row" onSubmit={handleSubmit}>
                            {selectedScript.fields.map((field) => (
                              <Field key={field.id} field={field} value={formValues[field.id]} onChange={handleChange} />
                            ))}
                            <button className="add-btn" type="submit" disabled={submitting}>
                              {submitting ? "Starting..." : "Run in Toolbox"}
                            </button>
                          </form>
                        </div>
                      </div>

                      <div className={`card ${runDetailsOpen ? "" : "card-collapsed"}`}>
                        <button type="button" className="card-header card-header-button" onClick={() => setRunDetailsOpen((current) => !current)}>
                          <span className="card-title">Run Details</span>
                          <span className={`card-badge ${activeRun ? (activeRun.status === "completed" ? "badge-ok" : activeRun.status === "failed" ? "badge-crit" : "badge-warn") : "badge-neutral"}`}>
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
                                  <div className="method-label">Started</div>
                                  <div className="method-count">{formatDate(activeRun.startedAt)}</div>
                                </div>
                                <div className="quick-summary-item">
                                  <div className="method-label">Finished</div>
                                  <div className="method-count">{formatDate(activeRun.finishedAt)}</div>
                                </div>
                                <div className="quick-summary-item">
                                  <div className="method-label">Exit Code</div>
                                  <div className="method-count">{activeRun.exitCode ?? "Running"}</div>
                                </div>
                              </div>
                              <div className="manage-form-panel" style={{ marginTop: "1rem" }}>
                                <h4>Command</h4>
                                <pre className="manage-response">{activeRun.command}</pre>
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

                      <div className={`card ${recentRunsOpen ? "" : "card-collapsed"}`}>
                        <button type="button" className="card-header card-header-button" onClick={() => setRecentRunsOpen((current) => !current)}>
                          <span className="card-title">Recent Runs</span>
                          <span className="card-badge badge-neutral">{runs.length}</span>
                          <span className="card-chevron">{recentRunsOpen ? "▾" : "▸"}</span>
                        </button>
                        {recentRunsOpen ? (
                        <div className="card-body">
                          {runs.length === 0 ? (
                            <div className="empty-row">No runs yet.</div>
                          ) : (
                            <div className="table-scroll">
                              <table>
                                <thead>
                                  <tr>
                                    <th>Script</th>
                                    <th>Status</th>
                                    <th>Started</th>
                                    <th>Action</th>
                                  </tr>
                                </thead>
                                <tbody>
                                  {runs.map((run) => (
                                    <tr key={run.id}>
                                      <td>{run.scriptName}</td>
                                      <td>
                                        <span className={`pill ${run.status === "completed" ? "badge-ok" : run.status === "failed" ? "badge-crit" : "badge-warn"}`}>
                                          {run.status}
                                        </span>
                                      </td>
                                      <td>{formatDate(run.startedAt)}</td>
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

                      {activeRunHtml ? (
                        <div className="card">
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
                            <iframe title="MFA HTML report preview" className="report-frame" srcDoc={activeRunHtml} />
                          </div>
                        </div>
                      ) : null}
                    </div>
                  </div>
                </div>
              </div>
            </>
          ) : (
            <section className="empty-state">
              <div className="empty-icon">M365</div>
              <div className="empty-title">Loading Toolbox...</div>
              <div className="empty-sub">{error || "Waiting for the script catalog from the backend."}</div>
            </section>
          )}
        </main>
      </div>
    </div>
  );
}

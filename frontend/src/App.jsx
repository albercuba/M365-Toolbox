import { useEffect, useState } from "react";

const apiBase = "/api";

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
        type={field.type === "number" ? "number" : "text"}
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
  const [error, setError] = useState("");
  const [submitting, setSubmitting] = useState(false);

  useEffect(() => {
    const load = async () => {
      const [scriptsResponse, runsResponse] = await Promise.all([
        fetch(`${apiBase}/scripts`),
        fetch(`${apiBase}/runs`)
      ]);

      const scriptsData = await scriptsResponse.json();
      const runsData = await runsResponse.json();
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
      const data = await response.json();
      setActiveRun(data);
      const runsResponse = await fetch(`${apiBase}/runs`);
      setRuns(await runsResponse.json());
    }, 2000);

    return () => window.clearInterval(timer);
  }, [activeRun]);

  const handleScriptSelect = (script) => {
    setSelectedScript(script);
    setFormValues(normalizeDefaults(script.fields));
    setError("");
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

      const data = await response.json();
      if (!response.ok) {
        throw new Error(data.message || "Failed to start run.");
      }

      setActiveRun(data);
      const runsResponse = await fetch(`${apiBase}/runs`);
      setRuns(await runsResponse.json());
    } catch (submitError) {
      setError(submitError.message);
    } finally {
      setSubmitting(false);
    }
  };

  return (
    <div className="app-shell">
      <aside className="sidebar">
        <div className="brand-block">
          <div className="brand-kicker">M365 Toolbox</div>
          <h1>M365 Toolbox</h1>
          <p>Web-based PowerShell operations for Microsoft 365</p>
        </div>

        <div className="panel">
          <h2>Script Catalog</h2>
          <div className="catalog-list">
            {scripts.map((script) => (
              <button
                key={script.id}
                type="button"
                className={selectedScript?.id === script.id ? "catalog-item active" : "catalog-item"}
                onClick={() => handleScriptSelect(script)}
              >
                <strong>{script.name}</strong>
                <span>{script.category}</span>
              </button>
            ))}
          </div>
        </div>

        <div className="panel">
          <h2>Recent Runs</h2>
          <div className="run-list">
            {runs.length === 0 ? <p>No runs yet.</p> : null}
            {runs.map((run) => (
              <button key={run.id} type="button" className="run-item" onClick={() => setActiveRun(run)}>
                <strong>{run.scriptName}</strong>
                <span>{run.status}</span>
                <small>{formatDate(run.startedAt)}</small>
              </button>
            ))}
          </div>
        </div>
      </aside>

      <main className="content">
        {selectedScript ? (
          <>
            <section className="hero-card">
              <div>
                <p className="hero-category">{selectedScript.category}</p>
                <h2>{selectedScript.name}</h2>
                <p>{selectedScript.summary}</p>
              </div>
              <div className="hero-note">
                <strong>Designed for growth</strong>
                <p>This toolbox uses a script registry so additional M365 runbooks can be added later without changing the UI shape.</p>
              </div>
            </section>

            <section className="grid">
              <form className="panel form-panel" onSubmit={handleSubmit}>
                <div className="panel-header">
                  <h2>Run Script</h2>
                  <span>{selectedScript.id}</span>
                </div>
                <p className="description">{selectedScript.description}</p>
                {selectedScript.fields.map((field) => (
                  <Field key={field.id} field={field} value={formValues[field.id]} onChange={handleChange} />
                ))}
                {error ? <div className="error-box">{error}</div> : null}
                <button className="primary-button" type="submit" disabled={submitting}>
                  {submitting ? "Starting..." : "Run in Toolbox"}
                </button>
              </form>

              <section className="panel result-panel">
                <div className="panel-header">
                  <h2>Run Details</h2>
                  <span>{activeRun ? activeRun.status : "idle"}</span>
                </div>
                {!activeRun ? <p>Select a run or start one to inspect the output.</p> : null}
                {activeRun ? (
                  <>
                    <div className="meta-grid">
                      <div>
                        <span>Started</span>
                        <strong>{formatDate(activeRun.startedAt)}</strong>
                      </div>
                      <div>
                        <span>Finished</span>
                        <strong>{formatDate(activeRun.finishedAt)}</strong>
                      </div>
                      <div>
                        <span>Exit Code</span>
                        <strong>{activeRun.exitCode ?? "Running"}</strong>
                      </div>
                    </div>
                    <div className="output-block">
                      <h3>Command</h3>
                      <pre>{activeRun.command}</pre>
                    </div>
                    <div className="output-block">
                      <h3>Stdout</h3>
                      <pre>{activeRun.stdout || "No stdout yet."}</pre>
                    </div>
                    <div className="output-block">
                      <h3>Stderr</h3>
                      <pre>{activeRun.stderr || "No stderr."}</pre>
                    </div>
                  </>
                ) : null}
              </section>
            </section>
          </>
        ) : (
          <section className="panel">
            <h2>Loading Toolbox...</h2>
            {error ? <div className="error-box">{error}</div> : <p>Waiting for the script catalog from the backend.</p>}
          </section>
        )}
      </main>
    </div>
  );
}

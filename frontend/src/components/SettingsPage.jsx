import { useEffect, useState } from "react";

const ROLE_OPTIONS = ["administrator", "privileged_user", "restricted_user"];

function roleLabel(role) {
  return role.split("_").map((part) => part.charAt(0).toUpperCase() + part.slice(1)).join(" ");
}

export function SettingsPage({
  apiBase,
  apiFetch,
  currentUser,
  isAdministrator,
  companies,
  companyDraft,
  companyImportInputRef,
  editingCompanyDraft,
  editingCompanyId,
  InfoTooltip,
  onAddCompany,
  onCancelEditCompany,
  onExportCompanies,
  onImportCompanies,
  onRemoveCompany,
  onSaveCompany,
  onStartEditCompany,
  setCompanyDraft,
  setEditingCompanyDraft,
  onSessionRefresh
}) {
  const [passwordDraft, setPasswordDraft] = useState({ currentPassword: "", newPassword: "" });
  const [authConfig, setAuthConfig] = useState(null);
  const [mappings, setMappings] = useState([]);
  const [mappingDraft, setMappingDraft] = useState({ groupName: "", groupId: "", assignedRole: "restricted_user" });
  const [settingsError, setSettingsError] = useState("");
  const [settingsSuccess, setSettingsSuccess] = useState("");

  useEffect(() => {
    if (!isAdministrator) {
      return;
    }

    const loadAuthSettings = async () => {
      try {
        const [configResponse, mappingsResponse] = await Promise.all([
          apiFetch(`${apiBase}/settings/auth/microsoft`),
          apiFetch(`${apiBase}/settings/auth/group-role-mappings`)
        ]);
        const configData = await configResponse.json();
        const mappingsData = await mappingsResponse.json();
        if (!configResponse.ok) throw new Error(configData.message || "Failed to load Microsoft configuration.");
        if (!mappingsResponse.ok) throw new Error(mappingsData.message || "Failed to load group mappings.");
        setAuthConfig(configData);
        setMappings(mappingsData);
      } catch (error) {
        setSettingsError(error.message);
      }
    };

    loadAuthSettings();
  }, [apiBase, apiFetch, isAdministrator]);

  const handleChangePassword = async (event) => {
    event.preventDefault();
    setSettingsError("");
    setSettingsSuccess("");
    try {
      const response = await apiFetch(`${apiBase}/auth/change-password`, {
        method: "POST",
        body: JSON.stringify(passwordDraft)
      });
      const data = await response.json();
      if (!response.ok) {
        throw new Error(data.message || "Failed to change password.");
      }
      setPasswordDraft({ currentPassword: "", newPassword: "" });
      await onSessionRefresh?.();
      setSettingsSuccess("Password changed.");
    } catch (error) {
      setSettingsError(error.message);
    }
  };

  const saveMicrosoftConfig = async (event) => {
    event.preventDefault();
    setSettingsError("");
    setSettingsSuccess("");
    try {
      const response = await apiFetch(`${apiBase}/settings/auth/microsoft`, {
        method: "PUT",
        body: JSON.stringify(authConfig)
      });
      const data = await response.json();
      if (!response.ok) {
        throw new Error(data.message || "Failed to save Microsoft configuration.");
      }
      setAuthConfig(data);
      setSettingsSuccess("Microsoft app registration configuration saved.");
    } catch (error) {
      setSettingsError(error.message);
    }
  };

  const addMapping = async (event) => {
    event.preventDefault();
    setSettingsError("");
    setSettingsSuccess("");
    try {
      const response = await apiFetch(`${apiBase}/settings/auth/group-role-mappings`, {
        method: "POST",
        body: JSON.stringify(mappingDraft)
      });
      const data = await response.json();
      if (!response.ok) {
        throw new Error(data.message || "Failed to add group mapping.");
      }
      setMappings((current) => [...current, data]);
      setMappingDraft({ groupName: "", groupId: "", assignedRole: "restricted_user" });
      setSettingsSuccess("Group mapping added.");
    } catch (error) {
      setSettingsError(error.message);
    }
  };

  const updateMapping = async (mapping, patch) => {
    setSettingsError("");
    setSettingsSuccess("");
    const next = { ...mapping, ...patch };
    try {
      const response = await apiFetch(`${apiBase}/settings/auth/group-role-mappings/${mapping.id}`, {
        method: "PUT",
        body: JSON.stringify(next)
      });
      const data = await response.json();
      if (!response.ok) {
        throw new Error(data.message || "Failed to update group mapping.");
      }
      setMappings((current) => current.map((entry) => entry.id === mapping.id ? data : entry));
    } catch (error) {
      setSettingsError(error.message);
    }
  };

  const deleteMapping = async (mappingId) => {
    setSettingsError("");
    setSettingsSuccess("");
    try {
      const response = await apiFetch(`${apiBase}/settings/auth/group-role-mappings/${mappingId}`, {
        method: "DELETE"
      });
      if (!response.ok) {
        const data = await response.json();
        throw new Error(data.message || "Failed to delete group mapping.");
      }
      setMappings((current) => current.filter((entry) => entry.id !== mappingId));
      setSettingsSuccess("Group mapping removed.");
    } catch (error) {
      setSettingsError(error.message);
    }
  };

  return (
    <div className="dash-page settings-page">
      <div className="sections">
        {settingsError ? <div className="flash flash-error soft">{settingsError}</div> : null}
        {settingsSuccess ? <div className="flash soft">{settingsSuccess}</div> : null}
        <div className="card">
          <div className="card-header">
            <span className="card-title">Account</span>
            <span className="card-badge badge-neutral">{roleLabel(currentUser.role)}</span>
          </div>
          <div className="card-body">
            {currentUser.mustChangePassword ? (
              <div className="approval-banner">This account is still using the default password. Change it before continuing operational work.</div>
            ) : null}
            <div className="quick-summary-grid">
              <div className="quick-summary-item">
                <div className="method-label">User</div>
                <div className="method-count">{currentUser.displayName || currentUser.username}</div>
              </div>
              <div className="quick-summary-item">
                <div className="method-label">Provider</div>
                <div className="method-count">{currentUser.authProvider}</div>
              </div>
            </div>
            {currentUser.authProvider === "local" ? (
              <form className="settings-row" onSubmit={handleChangePassword} style={{ marginTop: "1rem" }}>
                <label className="form-field">
                  <span>Current Password</span>
                  <input
                    type="password"
                    value={passwordDraft.currentPassword}
                    onChange={(event) => setPasswordDraft((current) => ({ ...current, currentPassword: event.target.value }))}
                    autoComplete="current-password"
                  />
                </label>
                <label className="form-field">
                  <span>New Password</span>
                  <input
                    type="password"
                    value={passwordDraft.newPassword}
                    onChange={(event) => setPasswordDraft((current) => ({ ...current, newPassword: event.target.value }))}
                    autoComplete="new-password"
                  />
                </label>
                <button type="submit" className="add-btn">Change Password</button>
              </form>
            ) : null}
          </div>
        </div>

        {isAdministrator ? (
        <div className="card">
          <div className="card-header">
            <span className="card-title">Companies</span>
            <span className="card-badge badge-neutral">{companies.length} companies</span>
            <div className="run-actions">
              <InfoTooltip label="Company CSV format">
                CSV format: Company Name,Tenant ID or Domain. Example: Contoso,contoso.onmicrosoft.com. Wrap company names with commas in quotes.
              </InfoTooltip>
              <input
                ref={companyImportInputRef}
                type="file"
                accept=".csv,text/csv"
                className="visually-hidden"
                onChange={onImportCompanies}
              />
              <button type="button" className="filter-btn" onClick={() => companyImportInputRef.current?.click()}>
                Import CSV
              </button>
              <button type="button" className="filter-btn" onClick={onExportCompanies} disabled={!companies.length}>
                Export CSV
              </button>
            </div>
          </div>
          <div className="card-body">
            <form className="company-form" onSubmit={onAddCompany}>
              <label className="form-field">
                <span>Company Name</span>
                <input
                  value={companyDraft.name}
                  onChange={(event) => setCompanyDraft((current) => ({ ...current, name: event.target.value }))}
                  placeholder="Contoso"
                />
              </label>
              <label className="form-field">
                <span>Tenant ID or Domain</span>
                <input
                  value={companyDraft.tenant}
                  onChange={(event) => setCompanyDraft((current) => ({ ...current, tenant: event.target.value }))}
                  placeholder="contoso.onmicrosoft.com"
                />
              </label>
              <button type="submit" className="add-btn">Add Company</button>
            </form>
            {companies.length ? (
              <div className="company-list">
                {companies.map((company) => {
                  const isEditing = editingCompanyId === company.id;

                  return (
                    <div key={company.id} className={`company-item${isEditing ? " editing" : ""}`}>
                      <div className="tenant-avatar">{company.name.slice(0, 2).toUpperCase()}</div>
                      {isEditing ? (
                        <div className="company-edit-grid">
                          <label className="form-field">
                            <span>Company Name</span>
                            <input
                              value={editingCompanyDraft.name}
                              onChange={(event) => setEditingCompanyDraft((current) => ({ ...current, name: event.target.value }))}
                            />
                          </label>
                          <label className="form-field">
                            <span>Tenant ID or Domain</span>
                            <input
                              value={editingCompanyDraft.tenant}
                              onChange={(event) => setEditingCompanyDraft((current) => ({ ...current, tenant: event.target.value }))}
                            />
                          </label>
                        </div>
                      ) : (
                        <div className="tenant-info">
                          <div className="tenant-name">{company.name}</div>
                          <div className="tenant-meta">{company.tenant}</div>
                        </div>
                      )}
                      <div className="company-actions">
                        {isEditing ? (
                          <>
                            <button type="button" className="filter-btn active-all" onClick={() => onSaveCompany(company.id)}>
                              Save
                            </button>
                            <button type="button" className="filter-btn" onClick={onCancelEditCompany}>
                              Cancel
                            </button>
                          </>
                        ) : (
                          <button type="button" className="filter-btn" onClick={() => onStartEditCompany(company)}>
                            Edit
                          </button>
                        )}
                        <button type="button" className="filter-btn destructive" onClick={() => onRemoveCompany(company.id)}>
                          Remove
                        </button>
                      </div>
                    </div>
                  );
                })}
              </div>
            ) : (
              <div className="empty-row">Add companies here, then type a company name, tenant ID, or domain in any script tenant field.</div>
            )}
          </div>
        </div>
        ) : (
          <div className="card">
            <div className="card-header">
              <span className="card-title">Settings</span>
              <span className="card-badge badge-neutral">read only</span>
            </div>
            <div className="card-body">
              <div className="empty-row">Administrator-only settings are hidden for this role.</div>
            </div>
          </div>
        )}

        {isAdministrator && authConfig ? (
          <div className="card">
            <div className="card-header">
              <span className="card-title">Microsoft App Registration Configuration</span>
              <span className={`card-badge ${authConfig.enabled ? "badge-ok" : "badge-neutral"}`}>{authConfig.enabled ? "enabled" : "disabled"}</span>
            </div>
            <div className="card-body">
              <form className="settings-row" onSubmit={saveMicrosoftConfig}>
                <label className="checkbox-field">
                  <input
                    type="checkbox"
                    checked={authConfig.enabled}
                    onChange={(event) => setAuthConfig((current) => ({ ...current, enabled: event.target.checked }))}
                  />
                  <span>Enable Microsoft login</span>
                </label>
                <label className="form-field">
                  <span>Tenant ID</span>
                  <input value={authConfig.tenantId} onChange={(event) => setAuthConfig((current) => ({ ...current, tenantId: event.target.value }))} />
                </label>
                <label className="form-field">
                  <span>Frontend Client ID</span>
                  <input value={authConfig.clientId} onChange={(event) => setAuthConfig((current) => ({ ...current, clientId: event.target.value }))} />
                </label>
                <label className="form-field">
                  <span>Backend API Audience</span>
                  <input value={authConfig.apiClientId} onChange={(event) => setAuthConfig((current) => ({ ...current, apiClientId: event.target.value }))} />
                </label>
                <label className="form-field">
                  <span>Authority URL</span>
                  <input value={authConfig.authorityUrl} onChange={(event) => setAuthConfig((current) => ({ ...current, authorityUrl: event.target.value }))} placeholder="https://login.microsoftonline.com/{tenantId}" />
                </label>
                <button type="submit" className="add-btn">Save Microsoft Configuration</button>
              </form>
              <div className="quick-summary-grid" style={{ marginTop: "1rem" }}>
                <div className="quick-summary-item">
                  <div className="method-label">Redirect URI</div>
                  <div className="method-count">{window.location.origin}</div>
                </div>
                <div className="quick-summary-item">
                  <div className="method-label">API Scope</div>
                  <div className="method-count">{authConfig.apiClientId ? `api://${authConfig.apiClientId}/access_as_user` : "Set API audience first"}</div>
                </div>
              </div>
              <div className="manage-form-panel">
                <h4>Entra Group Role Mappings</h4>
                <form className="settings-row" onSubmit={addMapping}>
                  <label className="form-field">
                    <span>Entra Group Name</span>
                    <input value={mappingDraft.groupName} onChange={(event) => setMappingDraft((current) => ({ ...current, groupName: event.target.value }))} />
                  </label>
                  <label className="form-field">
                    <span>Group Object ID</span>
                    <input value={mappingDraft.groupId} onChange={(event) => setMappingDraft((current) => ({ ...current, groupId: event.target.value }))} />
                  </label>
                  <label className="form-field">
                    <span>Assigned Role</span>
                    <select value={mappingDraft.assignedRole} onChange={(event) => setMappingDraft((current) => ({ ...current, assignedRole: event.target.value }))}>
                      {ROLE_OPTIONS.map((role) => <option key={role} value={role}>{role}</option>)}
                    </select>
                  </label>
                  <button type="submit" className="add-btn">Add Mapping</button>
                </form>
                <div className="company-list" style={{ marginTop: "1rem" }}>
                  {mappings.map((mapping) => (
                    <div key={mapping.id} className="company-item">
                      <div className="tenant-info">
                        <div className="tenant-name">{mapping.groupName}</div>
                        <div className="tenant-meta">{mapping.groupId || "Matching by claim group name"}</div>
                      </div>
                      <div className="company-actions">
                        <select value={mapping.assignedRole} onChange={(event) => updateMapping(mapping, { assignedRole: event.target.value })}>
                          {ROLE_OPTIONS.map((role) => <option key={role} value={role}>{role}</option>)}
                        </select>
                        <button type="button" className="filter-btn destructive" onClick={() => deleteMapping(mapping.id)}>Remove</button>
                      </div>
                    </div>
                  ))}
                  {!mappings.length ? <div className="empty-row">Add at least one Entra group mapping before enabling Microsoft login for operators.</div> : null}
                </div>
              </div>
            </div>
          </div>
        ) : null}
      </div>
    </div>
  );
}

import { useEffect, useState } from "react";

const ROLE_OPTIONS = ["administrator", "privileged_user", "restricted_user"];
const SETTINGS_SECTIONS = new Set(["companies", "microsoft", "users"]);

function roleLabel(role) {
  return role.split("_").map((part) => part.charAt(0).toUpperCase() + part.slice(1)).join(" ");
}

function emptyUserDraft() {
  return {
    username: "",
    displayName: "",
    email: "",
    password: "",
    role: "restricted_user",
    mustChangePassword: true
  };
}

function userToDraft(user) {
  return {
    username: user.username || "",
    displayName: user.displayName || "",
    email: user.email || "",
    password: "",
    role: user.role || "restricted_user",
    mustChangePassword: Boolean(user.mustChangePassword)
  };
}

export function SettingsPage({
  apiBase,
  apiFetch,
  currentUser,
  isAdministrator,
  activeSettingsSection = "companies",
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
  const selectedSection = SETTINGS_SECTIONS.has(activeSettingsSection) ? activeSettingsSection : "companies";
  const [authConfig, setAuthConfig] = useState(null);
  const [mappings, setMappings] = useState([]);
  const [mappingDraft, setMappingDraft] = useState({ groupName: "", groupId: "", assignedRole: "restricted_user" });
  const [users, setUsers] = useState([]);
  const [userDraft, setUserDraft] = useState(emptyUserDraft);
  const [editingUserId, setEditingUserId] = useState("");
  const [editingUserDraft, setEditingUserDraft] = useState(emptyUserDraft);
  const [settingsError, setSettingsError] = useState("");
  const [settingsSuccess, setSettingsSuccess] = useState("");

  useEffect(() => {
    if (!isAdministrator) {
      return;
    }

    const loadAuthSettings = async () => {
      try {
        const [configResponse, mappingsResponse, usersResponse] = await Promise.all([
          apiFetch(`${apiBase}/settings/auth/microsoft`),
          apiFetch(`${apiBase}/settings/auth/group-role-mappings`),
          apiFetch(`${apiBase}/settings/users`)
        ]);
        const configData = await configResponse.json();
        const mappingsData = await mappingsResponse.json();
        const usersData = await usersResponse.json();
        if (!configResponse.ok) throw new Error(configData.message || "Failed to load Microsoft configuration.");
        if (!mappingsResponse.ok) throw new Error(mappingsData.message || "Failed to load group mappings.");
        if (!usersResponse.ok) throw new Error(usersData.message || "Failed to load users.");
        setAuthConfig(configData);
        setMappings(mappingsData);
        setUsers(usersData);
      } catch (error) {
        setSettingsError(error.message);
      }
    };

    loadAuthSettings();
  }, [apiBase, apiFetch, isAdministrator]);



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

  const addUser = async (event) => {
    event.preventDefault();
    setSettingsError("");
    setSettingsSuccess("");
    try {
      const response = await apiFetch(`${apiBase}/settings/users`, {
        method: "POST",
        body: JSON.stringify(userDraft)
      });
      const data = await response.json();
      if (!response.ok) {
        throw new Error(data.message || "Failed to create user.");
      }
      setUsers((current) => [...current, data].sort((left, right) => left.username.localeCompare(right.username)));
      setUserDraft(emptyUserDraft());
      setSettingsSuccess("User created.");
    } catch (error) {
      setSettingsError(error.message);
    }
  };

  const startEditUser = (user) => {
    setEditingUserId(user.id);
    setEditingUserDraft(userToDraft(user));
  };

  const cancelEditUser = () => {
    setEditingUserId("");
    setEditingUserDraft(emptyUserDraft());
  };

  const saveUser = async (userId) => {
    setSettingsError("");
    setSettingsSuccess("");
    try {
      const response = await apiFetch(`${apiBase}/settings/users/${encodeURIComponent(userId)}`, {
        method: "PUT",
        body: JSON.stringify(editingUserDraft)
      });
      const data = await response.json();
      if (!response.ok) {
        throw new Error(data.message || "Failed to update user.");
      }
      setUsers((current) => current.map((entry) => entry.id === userId ? data : entry));
      if (currentUser.id === userId) {
        await onSessionRefresh?.();
      }
      cancelEditUser();
      setSettingsSuccess("User updated.");
    } catch (error) {
      setSettingsError(error.message);
    }
  };

  const removeUser = async (userId) => {
    setSettingsError("");
    setSettingsSuccess("");
    try {
      const response = await apiFetch(`${apiBase}/settings/users/${encodeURIComponent(userId)}`, {
        method: "DELETE"
      });
      if (!response.ok) {
        const data = await response.json();
        throw new Error(data.message || "Failed to remove user.");
      }
      setUsers((current) => current.filter((entry) => entry.id !== userId));
      if (editingUserId === userId) {
        cancelEditUser();
      }
      setSettingsSuccess("User removed.");
    } catch (error) {
      setSettingsError(error.message);
    }
  };



  const renderCompaniesSection = () => (
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
  );

  const renderMicrosoftSection = () => authConfig ? (
    <>
    <div className="card">
      <div className="card-header">
        <span className="card-title">Microsoft Integration</span>
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
    <details className="card integration-help-card">
      <summary className="card-header integration-help-summary">
        <span className="card-title">Microsoft Integration Setup Instructions</span>
        <span className="card-badge badge-neutral">setup guide</span>
      </summary>
      <div className="card-body integration-help-body">
        <div className="manage-form-panel">
          <h4>1. Create the backend API app registration</h4>
          <ol className="settings-instruction-list">
            <li>In Microsoft Entra admin center, create an app registration for the Toolbox backend API.</li>
            <li>Expose an API and set the Application ID URI to `api://&lt;backend-api-client-id&gt;`.</li>
            <li>Add a delegated scope named `access_as_user` for signed-in Toolbox operators.</li>
            <li>Copy the backend API application/client ID into `Backend API Audience` in Toolbox.</li>
          </ol>
        </div>
        <div className="manage-form-panel">
          <h4>2. Create the frontend SPA app registration</h4>
          <ol className="settings-instruction-list">
            <li>Create a second app registration for the Toolbox frontend SPA.</li>
            <li>Add a Single-page application redirect URI that matches the Toolbox browser origin shown above.</li>
            <li>Grant the frontend permission to call the backend API scope.</li>
            <li>Copy the frontend application/client ID into `Frontend Client ID` in Toolbox.</li>
          </ol>
        </div>
        <div className="manage-form-panel">
          <h4>3. Configure roles and enable sign-in</h4>
          <ol className="settings-instruction-list">
            <li>Copy your Entra tenant ID into `Tenant ID`.</li>
            <li>Add Entra group role mappings in Toolbox. Prefer group object IDs over display names.</li>
            <li>Assign operators to the Entra groups that map to `administrator`, `privileged_user`, or `restricted_user`.</li>
            <li>Enable Microsoft login, save the configuration, and test sign-in with a non-break-glass account.</li>
          </ol>
        </div>
      </div>
    </details>
    </>
  ) : null;

  const renderUsersSection = () => (
    <div className="card">
      <div className="card-header">
        <span className="card-title">Users</span>
        <span className="card-badge badge-neutral">{users.length} users</span>
      </div>
      <div className="card-body">
        <div className="manage-form-panel">
          <h4>Create Local User</h4>
          <form className="user-form" onSubmit={addUser}>
            <label className="form-field">
              <span>Username</span>
              <input value={userDraft.username} onChange={(event) => setUserDraft((current) => ({ ...current, username: event.target.value }))} autoComplete="off" />
            </label>
            <label className="form-field">
              <span>Display Name</span>
              <input value={userDraft.displayName} onChange={(event) => setUserDraft((current) => ({ ...current, displayName: event.target.value }))} />
            </label>
            <label className="form-field">
              <span>Email</span>
              <input type="email" value={userDraft.email} onChange={(event) => setUserDraft((current) => ({ ...current, email: event.target.value }))} />
            </label>
            <label className="form-field">
              <span>Initial Password</span>
              <input type="password" value={userDraft.password} onChange={(event) => setUserDraft((current) => ({ ...current, password: event.target.value }))} autoComplete="new-password" />
            </label>
            <label className="form-field">
              <span>Role</span>
              <select value={userDraft.role} onChange={(event) => setUserDraft((current) => ({ ...current, role: event.target.value }))}>
                {ROLE_OPTIONS.map((role) => <option key={role} value={role}>{roleLabel(role)}</option>)}
              </select>
            </label>
            <label className="checkbox-field user-form-checkbox">
              <input
                type="checkbox"
                checked={userDraft.mustChangePassword}
                onChange={(event) => setUserDraft((current) => ({ ...current, mustChangePassword: event.target.checked }))}
              />
              <span>Require password change on first login</span>
            </label>
            <button type="submit" className="add-btn">Create User</button>
          </form>
        </div>

        <div className="company-list">
          {users.map((user) => {
            const isEditing = editingUserId === user.id;
            const isLocal = user.authProvider === "local";
            return (
              <div key={user.id} className={`company-item user-item${isEditing ? " editing" : ""}`}>
                <div className="tenant-avatar">{(user.displayName || user.username).slice(0, 2).toUpperCase()}</div>
                {isEditing ? (
                  <div className="user-edit-grid">
                    <label className="form-field">
                      <span>Username</span>
                      <input
                        value={editingUserDraft.username}
                        disabled={!isLocal}
                        onChange={(event) => setEditingUserDraft((current) => ({ ...current, username: event.target.value }))}
                      />
                    </label>
                    <label className="form-field">
                      <span>Display Name</span>
                      <input value={editingUserDraft.displayName} onChange={(event) => setEditingUserDraft((current) => ({ ...current, displayName: event.target.value }))} />
                    </label>
                    <label className="form-field">
                      <span>Email</span>
                      <input type="email" value={editingUserDraft.email} onChange={(event) => setEditingUserDraft((current) => ({ ...current, email: event.target.value }))} />
                    </label>
                    <label className="form-field">
                      <span>Role</span>
                      <select disabled={!isLocal} value={editingUserDraft.role} onChange={(event) => setEditingUserDraft((current) => ({ ...current, role: event.target.value }))}>
                        {ROLE_OPTIONS.map((role) => <option key={role} value={role}>{roleLabel(role)}</option>)}
                      </select>
                    </label>
                    {isLocal ? (
                      <>
                        <label className="form-field">
                          <span>New Password</span>
                          <input type="password" value={editingUserDraft.password} onChange={(event) => setEditingUserDraft((current) => ({ ...current, password: event.target.value }))} placeholder="Leave blank to keep current" autoComplete="new-password" />
                        </label>
                        <label className="checkbox-field user-form-checkbox">
                          <input
                            type="checkbox"
                            checked={editingUserDraft.mustChangePassword}
                            onChange={(event) => setEditingUserDraft((current) => ({ ...current, mustChangePassword: event.target.checked }))}
                          />
                          <span>Require password change</span>
                        </label>
                      </>
                    ) : null}
                  </div>
                ) : (
                  <div className="tenant-info">
                    <div className="tenant-name">{user.displayName || user.username}</div>
                    <div className="tenant-meta">{user.username}{user.email ? ` • ${user.email}` : ""}</div>
                    <div className="tenant-tags">
                      <span className="mini-pill badge-neutral">{user.authProvider}</span>
                      <span className="mini-pill badge-ok">{roleLabel(user.role)}</span>
                      {user.mustChangePassword ? <span className="mini-pill badge-warn">password change required</span> : null}
                      {user.authProvider === "microsoft" ? <span className="mini-pill badge-neutral">role from Entra mapping</span> : null}
                    </div>
                  </div>
                )}
                <div className="company-actions">
                  {isEditing ? (
                    <>
                      <button type="button" className="filter-btn active-all" onClick={() => saveUser(user.id)}>Save</button>
                      <button type="button" className="filter-btn" onClick={cancelEditUser}>Cancel</button>
                    </>
                  ) : (
                    <button type="button" className="filter-btn" onClick={() => startEditUser(user)}>Edit</button>
                  )}
                  <button type="button" className="filter-btn destructive" disabled={user.id === currentUser.id} onClick={() => removeUser(user.id)}>Remove</button>
                </div>
              </div>
            );
          })}
          {!users.length ? <div className="empty-row">No users are configured yet.</div> : null}
        </div>
      </div>
    </div>
  );

  return (
    <div className="dash-page settings-page">
      <div className="sections">
        {settingsError ? <div className="flash flash-error soft">{settingsError}</div> : null}
        {settingsSuccess ? <div className="flash soft">{settingsSuccess}</div> : null}

        {isAdministrator ? (
          <>
            {selectedSection === "companies" ? renderCompaniesSection() : null}
            {selectedSection === "microsoft" ? renderMicrosoftSection() : null}
            {selectedSection === "users" ? renderUsersSection() : null}
          </>
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
      </div>
    </div>
  );
}

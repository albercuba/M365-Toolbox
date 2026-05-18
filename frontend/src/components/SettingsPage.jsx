import { Fragment, useEffect, useState } from "react";

const ROLE_OPTIONS = ["administrator", "privileged_user", "restricted_user"];
const SETTINGS_SECTIONS = new Set(["account", "companies", "microsoft", "users"]);

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
  const selectedSection = SETTINGS_SECTIONS.has(activeSettingsSection) ? activeSettingsSection : isAdministrator ? "companies" : "account";
  const [passwordDraft, setPasswordDraft] = useState({ currentPassword: "", newPassword: "" });
  const [authConfig, setAuthConfig] = useState(null);
  const [mappings, setMappings] = useState([]);
  const [mappingDraft, setMappingDraft] = useState({ groupName: "", groupId: "", assignedRole: "restricted_user" });
  const [users, setUsers] = useState([]);
  const [userDraft, setUserDraft] = useState(emptyUserDraft);
  const [userSearch, setUserSearch] = useState("");
  const [editingUserId, setEditingUserId] = useState("");
  const [editingUserDraft, setEditingUserDraft] = useState(emptyUserDraft);
  const [settingsError, setSettingsError] = useState("");
  const [settingsSuccess, setSettingsSuccess] = useState("");

  useEffect(() => {
    if (!settingsSuccess) {
      return undefined;
    }

    const timer = window.setTimeout(() => setSettingsSuccess(""), 5000);
    return () => window.clearTimeout(timer);
  }, [settingsSuccess]);

  useEffect(() => {
    if (!isAdministrator) {
      return;
    }

    const loadAuthSettings = async () => {
      const errors = [];

      const [configResult, mappingsResult, usersResult] = await Promise.allSettled([
        apiFetch(`${apiBase}/settings/auth/microsoft`),
        apiFetch(`${apiBase}/settings/auth/group-role-mappings`),
        apiFetch(`${apiBase}/settings/users`)
      ]);

      if (configResult.status === "fulfilled") {
        const configData = await configResult.value.json();
        if (configResult.value.ok) {
          setAuthConfig(configData);
        } else {
          errors.push(configData.message || "Failed to load Microsoft configuration.");
        }
      } else {
        errors.push(configResult.reason?.message || "Failed to load Microsoft configuration.");
      }

      if (mappingsResult.status === "fulfilled") {
        const mappingsData = await mappingsResult.value.json();
        if (mappingsResult.value.ok) {
          setMappings(mappingsData);
        } else {
          errors.push(mappingsData.message || "Failed to load group mappings.");
        }
      } else {
        errors.push(mappingsResult.reason?.message || "Failed to load group mappings.");
      }

      if (usersResult.status === "fulfilled") {
        const usersData = await usersResult.value.json();
        if (usersResult.value.ok) {
          setUsers(usersData);
        } else {
          errors.push(usersData.message || "Failed to load users.");
        }
      } else {
        errors.push(usersResult.reason?.message || "Failed to load users.");
      }

      if (errors.length) {
        setSettingsError(errors.join(" "));
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

  const renderAccountSection = () => (
    <div className="card">
      <div className="card-header">
        <span className="card-title">Account Security</span>
        <span className="card-badge badge-neutral">{roleLabel(currentUser.role)}</span>
      </div>
      <div className="card-body">
        {settingsError ? <div className="flash flash-error soft">{settingsError}</div> : null}
        {settingsSuccess ? <div className="flash flash-success soft">{settingsSuccess}</div> : null}
        {currentUser.authProvider === "local" ? (
          <form className="settings-row" onSubmit={handleChangePassword}>
            <div className="quick-summary-grid">
              <div className="quick-summary-item">
                <div className="method-label">User</div>
                <div className="method-count">{currentUser.displayName || currentUser.username}</div>
              </div>
              <div className="quick-summary-item">
                <div className="method-label">Authentication</div>
                <div className="method-count">{currentUser.authProvider}</div>
              </div>
            </div>
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
        ) : (
          <div className="empty-row">Password changes for Microsoft users are managed in Microsoft Entra.</div>
        )}
      </div>
    </div>
  );

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

  const renderMicrosoftSection = () => {
    if (!authConfig) {
      return (
        <div className="card">
          <div className="card-header">
            <span className="card-title">Microsoft Integration</span>
            <span className="card-badge badge-neutral">loading</span>
          </div>
          <div className="card-body">
            <div className="empty-row">Loading Microsoft integration settings...</div>
          </div>
        </div>
      );
    }

    return (
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
            <li>Go to Microsoft Entra admin center.</li>
            <li>Open App registrations.</li>
            <li>Create a new registration for the Toolbox backend API.</li>
            <li>Use single-tenant unless this deployment intentionally supports multiple tenants.</li>
            <li>Do not configure a redirect URI for this backend API registration.</li>
            <li>Copy the Application client ID. This value is used later in Toolbox as `Backend API Audience`.</li>
            <li>Copy the Directory tenant ID. This value is used later in Toolbox as `Tenant ID`.</li>
          </ol>
        </div>
        <div className="manage-form-panel">
          <h4>2. Expose the backend API scope</h4>
          <ol className="settings-instruction-list">
            <li>Open the backend API app registration.</li>
            <li>Go to Expose an API.</li>
            <li>Set the Application ID URI to `api://&lt;backend-api-client-id&gt;`.</li>
            <li>Add a delegated scope named `access_as_user`.</li>
            <li>Toolbox expects the frontend to request `api://&lt;backend-api-client-id&gt;/access_as_user`.</li>
            <li>Save the scope.</li>
          </ol>
        </div>
        <div className="manage-form-panel">
          <h4>3. Configure group claims</h4>
          <ol className="settings-instruction-list">
            <li>Open the backend API app registration.</li>
            <li>Go to Token configuration.</li>
            <li>Add a groups claim.</li>
            <li>Include the groups claim in access tokens.</li>
            <li>Prefer group object IDs instead of display names.</li>
            <li>Use security groups or groups assigned to the application where possible.</li>
            <li>Users in too many groups can trigger group overage claims. The Toolbox backend rejects tokens with group overage because it cannot safely map those tokens to a Toolbox role.</li>
          </ol>
        </div>
        <div className="manage-form-panel">
          <h4>4. Create the frontend SPA app registration</h4>
          <ol className="settings-instruction-list">
            <li>Create a second app registration for the Toolbox frontend.</li>
            <li>Use the same tenant as the backend API app registration.</li>
            <li>Copy the Application client ID. This value is used later in Toolbox as `Frontend Client ID`.</li>
            <li>Go to Authentication.</li>
            <li>Add a Single-page application platform.</li>
            <li>Add the redirect URI that exactly matches the Toolbox browser origin shown in this Microsoft Integration page.</li>
            <li>Example production origin: `https://toolbox.example.com`.</li>
            <li>Example local development origin: `http://localhost:5173`, if that is the local browser origin.</li>
            <li>Do not append `/api`, `/auth`, or another path unless the actual browser origin includes it.</li>
            <li>Do not create a client secret for the SPA registration. This integration uses SPA redirect login through MSAL.</li>
          </ol>
        </div>
        <div className="manage-form-panel">
          <h4>5. Grant frontend permission to call the backend API</h4>
          <ol className="settings-instruction-list">
            <li>Open the frontend SPA app registration.</li>
            <li>Go to API permissions.</li>
            <li>Add a permission.</li>
            <li>Choose My APIs.</li>
            <li>Select the Toolbox backend API app registration.</li>
            <li>Select the delegated permission `access_as_user`.</li>
            <li>Add the permission.</li>
            <li>Grant admin consent if the tenant requires it.</li>
          </ol>
        </div>
        <div className="manage-form-panel">
          <h4>6. Configure Toolbox</h4>
          <ol className="settings-instruction-list">
            <li>Sign in to Toolbox as a local administrator.</li>
            <li>Open Settings.</li>
            <li>Open Microsoft Integration.</li>
            <li>Enter `Tenant ID` using the Directory tenant ID.</li>
            <li>Enter `Frontend Client ID` using the frontend SPA application client ID.</li>
            <li>Enter `Backend API Audience` using the backend API application client ID.</li>
            <li>Leave `Authority URL` blank unless a custom authority is required.</li>
            <li>When `Authority URL` is blank, Toolbox uses `https://login.microsoftonline.com/{tenantId}`.</li>
            <li>Add at least one Entra Group Role Mapping before enabling Microsoft login.</li>
            <li>Prefer entering the group object ID.</li>
            <li>Map each group to one of `administrator`, `privileged_user`, or `restricted_user`.</li>
            <li>Make sure each Microsoft user is a member of at least one mapped Entra group.</li>
            <li>Enable Microsoft login.</li>
            <li>Save the Microsoft configuration.</li>
          </ol>
        </div>
        <div className="manage-form-panel">
          <h4>7. Test sign-in</h4>
          <ol className="settings-instruction-list">
            <li>Sign out of Toolbox.</li>
            <li>Click `Sign in with Microsoft`.</li>
            <li>Complete the Microsoft login flow.</li>
            <li>Confirm the user lands back in Toolbox.</li>
            <li>Confirm the assigned Toolbox role matches the Entra group mapping.</li>
            <li>Test first with a non-break-glass account.</li>
          </ol>
        </div>
        <div className="manage-form-panel">
          <h4>8. Troubleshooting</h4>
          <ol className="settings-instruction-list">
            <li>If the Microsoft button is disabled, verify Microsoft login is enabled and the required IDs are saved.</li>
            <li>If the app says Microsoft login is not fully configured, verify `Tenant ID`, `Frontend Client ID`, and `Backend API Audience`.</li>
            <li>If token validation fails, verify the backend API client ID matches `Backend API Audience`, the exposed scope is `access_as_user`, and the frontend app has permission to that scope.</li>
            <li>If login succeeds but access is denied, verify the user belongs to a mapped Entra group and the group object ID matches the mapping.</li>
            <li>If group overage appears, reduce emitted groups by assigning groups to the application or by using a smaller security group set.</li>
            <li>If redirect fails, verify the SPA redirect URI exactly matches `window.location.origin` for the deployed Toolbox URL.</li>
          </ol>
        </div>
      </div>
    </details>
    </>
    );
  };

  const normalizedUserSearch = userSearch.trim().toLowerCase();
  const visibleUsers = normalizedUserSearch
    ? users.filter((user) => [
      user.displayName,
      user.username,
      user.email,
      user.authProvider,
      roleLabel(user.role)
    ].filter(Boolean).some((value) => String(value).toLowerCase().includes(normalizedUserSearch)))
    : users;

  const renderUsersSection = () => (
    <div className="card">
      <div className="card-header">
        <span className="card-title">Users</span>
        <span className="card-badge badge-neutral">{users.length} users</span>
      </div>
      <div className="card-body">
        {settingsError ? <div className="flash flash-error soft">{settingsError}</div> : null}
        {settingsSuccess ? <div className="flash flash-success soft">{settingsSuccess}</div> : null}

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

        <div className="panel-toolbar users-table-toolbar">
          <label className="form-field users-search-field">
            <span>Search Users</span>
            <input value={userSearch} onChange={(event) => setUserSearch(event.target.value)} placeholder="Search name, username, email, authentication, or role" />
          </label>
        </div>

        {visibleUsers.length ? (
          <div className="table-scroll users-table-scroll">
            <table className="users-table">
              <thead>
                <tr>
                  <th>Display name</th>
                  <th>Username</th>
                  <th>Email</th>
                  <th>Authentication</th>
                  <th>Local role</th>
                  <th>Microsoft role</th>
                  <th>Password change required</th>
                  <th>Actions</th>
                </tr>
              </thead>
              <tbody>
                {visibleUsers.map((user) => {
                  const isEditing = editingUserId === user.id;
                  const isLocal = user.authProvider === "local";
                  const displayLabel = user.displayName || user.username;

                  return (
                    <Fragment key={user.id}>
                    <tr>
                      <td>
                        {isEditing ? (
                          <input className="table-input" value={editingUserDraft.displayName} onChange={(event) => setEditingUserDraft((current) => ({ ...current, displayName: event.target.value }))} />
                        ) : displayLabel}
                      </td>
                      <td>
                        {isEditing ? (
                          <input
                            className="table-input"
                            value={editingUserDraft.username}
                            disabled={!isLocal}
                            onChange={(event) => setEditingUserDraft((current) => ({ ...current, username: event.target.value }))}
                          />
                        ) : user.username}
                      </td>
                      <td>
                        {isEditing ? (
                          <input className="table-input" type="email" value={editingUserDraft.email} onChange={(event) => setEditingUserDraft((current) => ({ ...current, email: event.target.value }))} />
                        ) : user.email || "—"}
                      </td>
                      <td><span className="mini-pill badge-neutral">{user.authProvider}</span></td>
                      <td>
                        {isEditing && isLocal ? (
                          <select className="table-select" value={editingUserDraft.role} onChange={(event) => setEditingUserDraft((current) => ({ ...current, role: event.target.value }))}>
                            {ROLE_OPTIONS.map((role) => <option key={role} value={role}>{roleLabel(role)}</option>)}
                          </select>
                        ) : isLocal ? <span className="mini-pill badge-ok">{roleLabel(user.role)}</span> : "—"}
                      </td>
                      <td>{user.authProvider === "microsoft" ? <span className="mini-pill badge-ok">{roleLabel(user.role)}</span> : "—"}</td>
                      <td>
                        {isEditing && isLocal ? (
                          <label className="table-checkbox">
                            <input
                              type="checkbox"
                              checked={editingUserDraft.mustChangePassword}
                              onChange={(event) => setEditingUserDraft((current) => ({ ...current, mustChangePassword: event.target.checked }))}
                            />
                            <span>{editingUserDraft.mustChangePassword ? "Yes" : "No"}</span>
                          </label>
                        ) : user.mustChangePassword ? <span className="mini-pill badge-warn">Yes</span> : "No"}
                      </td>
                      <td>
                        <div className="table-actions icon-actions">
                          {isEditing ? (
                            <>
                              <button type="button" className="icon-action-btn success" onClick={() => saveUser(user.id)} aria-label={`Save ${displayLabel}`} title="Save">
                                <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M9.55 17.2 4.8 12.45l1.4-1.4 3.35 3.35 8.25-8.25 1.4 1.4Z" fill="currentColor" /></svg>
                              </button>
                              <button type="button" className="icon-action-btn" onClick={cancelEditUser} aria-label={`Cancel editing ${displayLabel}`} title="Cancel">
                                <svg viewBox="0 0 24 24" aria-hidden="true"><path d="m6.4 19-1.4-1.4 5.6-5.6L5 6.4 6.4 5l5.6 5.6L17.6 5 19 6.4 13.4 12l5.6 5.6-1.4 1.4-5.6-5.6Z" fill="currentColor" /></svg>
                              </button>
                            </>
                          ) : (
                            <button type="button" className="icon-action-btn" onClick={() => startEditUser(user)} aria-label={`Edit ${displayLabel}`} title="Edit">
                              <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M5 19h1.4l9.85-9.85-1.4-1.4L5 17.6Zm-2 2v-4.25L16.25 3.5a2.12 2.12 0 0 1 3 0l1.25 1.25a2.12 2.12 0 0 1 0 3L7.25 21Zm14.65-13.25 1.25-1.25-1.4-1.4-1.25 1.25Z" fill="currentColor" /></svg>
                            </button>
                          )}
                          <button type="button" className="icon-action-btn destructive" disabled={user.id === currentUser.id} onClick={() => removeUser(user.id)} aria-label={`Remove ${displayLabel}`} title="Remove">
                            <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M7 21q-.82 0-1.41-.59A1.92 1.92 0 0 1 5 19V7H4V5h5V4h6v1h5v2h-1v12q0 .82-.59 1.41A1.92 1.92 0 0 1 17 21Zm2-4h2V9H9Zm4 0h2V9h-2Z" fill="currentColor" /></svg>
                          </button>
                        </div>
                      </td>
                    </tr>
                    {isEditing && isLocal ? (
                      <tr className="user-edit-details-row">
                        <td colSpan={8}>
                          <div className="user-edit-details-panel">
                            <label className="form-field user-password-reset-field">
                              <span>New Password</span>
                              <input className="table-input" type="password" value={editingUserDraft.password} onChange={(event) => setEditingUserDraft((current) => ({ ...current, password: event.target.value }))} placeholder="Leave blank to keep current password" autoComplete="new-password" />
                            </label>
                            <div className="empty-row compact">Password reset is optional. Leave this field blank when you only want to update profile details or role.</div>
                          </div>
                        </td>
                      </tr>
                    ) : null}
                    </Fragment>
                  );
                })}
              </tbody>
            </table>
          </div>
        ) : <div className="empty-row">{users.length ? "No users match your search." : "No users are configured yet."}</div>}
      </div>
    </div>
  );

  return (
    <div className="dash-page settings-page">
      <div className="sections">
        {selectedSection !== "users" && selectedSection !== "account" && settingsError ? <div className="flash flash-error soft">{settingsError}</div> : null}
        {selectedSection !== "users" && selectedSection !== "account" && settingsSuccess ? <div className="flash flash-success soft">{settingsSuccess}</div> : null}

        {selectedSection === "account" ? renderAccountSection() : null}
        {isAdministrator ? (
          <>
            {selectedSection === "companies" ? renderCompaniesSection() : null}
            {selectedSection === "microsoft" ? renderMicrosoftSection() : null}
            {selectedSection === "users" ? renderUsersSection() : null}
          </>
        ) : selectedSection !== "account" ? (
          <div className="card">
            <div className="card-header">
              <span className="card-title">Settings</span>
              <span className="card-badge badge-neutral">read only</span>
            </div>
            <div className="card-body">
              <div className="empty-row">Administrator-only settings are hidden for this role.</div>
            </div>
          </div>
        ) : null}
      </div>
    </div>
  );
}

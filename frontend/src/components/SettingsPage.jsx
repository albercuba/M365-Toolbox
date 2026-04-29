export function SettingsPage({
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
  setEditingCompanyDraft
}) {
  return (
    <div className="dash-page settings-page">
      <div className="sections">
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
      </div>
    </div>
  );
}

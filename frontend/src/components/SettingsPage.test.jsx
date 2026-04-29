import { render, screen } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { expect, test, vi } from "vitest";
import { createRef } from "react";
import { SettingsPage } from "./SettingsPage.jsx";

function renderSettings(overrides = {}) {
  const props = {
    companies: [
      {
        id: "company-1",
        name: "Contoso",
        tenant: "contoso.onmicrosoft.com"
      }
    ],
    companyDraft: { name: "", tenant: "" },
    companyImportInputRef: createRef(),
    editingCompanyDraft: { name: "", tenant: "" },
    editingCompanyId: "",
    InfoTooltip: ({ children }) => <span>{children}</span>,
    onAddCompany: vi.fn((event) => event.preventDefault()),
    onCancelEditCompany: vi.fn(),
    onExportCompanies: vi.fn(),
    onImportCompanies: vi.fn(),
    onRemoveCompany: vi.fn(),
    onSaveCompany: vi.fn(),
    onStartEditCompany: vi.fn(),
    setCompanyDraft: vi.fn(),
    setEditingCompanyDraft: vi.fn(),
    ...overrides
  };

  render(<SettingsPage {...props} />);
  return props;
}

test("renders company settings and CSV actions", () => {
  renderSettings();

  expect(screen.getByText("Companies")).toBeInTheDocument();
  expect(screen.getByText("Contoso")).toBeInTheDocument();
  expect(screen.getByText("contoso.onmicrosoft.com")).toBeInTheDocument();
  expect(screen.getByRole("button", { name: "Import CSV" })).toBeInTheDocument();
  expect(screen.getByRole("button", { name: "Export CSV" })).toBeInTheDocument();
});

test("starts inline edit for a company", async () => {
  const user = userEvent.setup();
  const props = renderSettings();

  await user.click(screen.getByRole("button", { name: "Edit" }));

  expect(props.onStartEditCompany).toHaveBeenCalledWith(expect.objectContaining({ id: "company-1" }));
});

test("renders editing controls", () => {
  renderSettings({
    editingCompanyId: "company-1",
    editingCompanyDraft: {
      name: "Contoso",
      tenant: "contoso.onmicrosoft.com"
    }
  });

  expect(screen.getByRole("button", { name: "Save" })).toBeInTheDocument();
  expect(screen.getByRole("button", { name: "Cancel" })).toBeInTheDocument();
});

import test from "node:test";
import assert from "node:assert/strict";
import { buildExecutionPlan } from "./scriptExecution.js";

test("compromised account actions are passed as a PowerShell array", () => {
  const plan = buildExecutionPlan(
    {
      id: "m365-compromised-account-remediation",
      scriptRelativePath: "M365-CompromisedAccountRemediation.ps1"
    },
    {
      userPrincipalName: "user@example.com",
      actions: [
        "ReviewMfaMethods",
        "ReviewInboxRules",
        "ReviewMailboxForwarding",
        "ReviewMailboxDelegates",
        "ReviewRecentSignIns",
        "ExportAuditLog"
      ]
    }
  );

  const actionsIndex = plan.commandArgs.indexOf("-Actions");
  assert.notEqual(actionsIndex, -1);
  assert.deepEqual(plan.commandArgs.slice(actionsIndex + 1, actionsIndex + 7), [
    "ReviewMfaMethods",
    "ReviewInboxRules",
    "ReviewMailboxForwarding",
    "ReviewMailboxDelegates",
    "ReviewRecentSignIns",
    "ExportAuditLog"
  ]);
  assert.equal(plan.commandArgs[actionsIndex + 1].includes(","), false);
});

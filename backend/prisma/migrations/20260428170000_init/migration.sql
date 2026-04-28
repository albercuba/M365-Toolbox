CREATE TABLE "runs" (
  "id" UUID PRIMARY KEY,
  "script_id" TEXT NOT NULL,
  "script_name" TEXT,
  "mode" TEXT,
  "status" TEXT NOT NULL,
  "tenant_id" UUID,
  "tenant_hint" TEXT,
  "requested_by" TEXT,
  "parameters" JSONB,
  "parameters_redacted" JSONB,
  "result" JSONB,
  "summary" TEXT,
  "exit_code" INTEGER,
  "started_at" TIMESTAMP(3),
  "finished_at" TIMESTAMP(3),
  "created_at" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
  "updated_at" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
  "queued_at" TIMESTAMP(3),
  "last_activity_at" TIMESTAMP(3),
  "cancel_requested_at" TIMESTAMP(3),
  "current_step" TEXT,
  "error_summary" TEXT,
  "duration_ms" BIGINT,
  "command" TEXT,
  "command_args" JSONB,
  "script_path" TEXT
);

CREATE TABLE "run_logs" (
  "id" UUID PRIMARY KEY,
  "run_id" UUID NOT NULL REFERENCES "runs"("id") ON DELETE CASCADE,
  "stream" TEXT NOT NULL,
  "level" TEXT,
  "message" TEXT NOT NULL,
  "created_at" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE "run_artifacts" (
  "id" UUID PRIMARY KEY,
  "run_id" UUID NOT NULL REFERENCES "runs"("id") ON DELETE CASCADE,
  "type" TEXT NOT NULL,
  "filename" TEXT NOT NULL,
  "path" TEXT NOT NULL,
  "size_bytes" INTEGER,
  "created_at" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
  CONSTRAINT "run_artifacts_run_id_path_key" UNIQUE ("run_id", "path")
);

CREATE TABLE "approvals" (
  "id" UUID PRIMARY KEY,
  "run_id" UUID NOT NULL UNIQUE REFERENCES "runs"("id") ON DELETE CASCADE,
  "status" TEXT NOT NULL,
  "requested_by" TEXT,
  "approved_by" TEXT,
  "reason" TEXT,
  "created_at" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
  "updated_at" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE INDEX "runs_status_created_at_idx" ON "runs" ("status", "created_at" DESC);
CREATE INDEX "runs_script_id_created_at_idx" ON "runs" ("script_id", "created_at" DESC);
CREATE INDEX "runs_tenant_id_created_at_idx" ON "runs" ("tenant_id", "created_at" DESC);
CREATE INDEX "runs_tenant_hint_created_at_idx" ON "runs" ("tenant_hint", "created_at" DESC);
CREATE INDEX "runs_requested_by_created_at_idx" ON "runs" ("requested_by", "created_at" DESC);
CREATE INDEX "run_logs_run_id_created_at_idx" ON "run_logs" ("run_id", "created_at");
CREATE INDEX "run_artifacts_run_id_created_at_idx" ON "run_artifacts" ("run_id", "created_at");
CREATE INDEX "approvals_status_created_at_idx" ON "approvals" ("status", "created_at");

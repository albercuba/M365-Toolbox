CREATE TABLE "companies" (
  "id" UUID PRIMARY KEY,
  "name" TEXT NOT NULL,
  "tenant" TEXT NOT NULL,
  "created_at" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
  "updated_at" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE UNIQUE INDEX "companies_name_key" ON "companies" ("name");
CREATE UNIQUE INDEX "companies_tenant_key" ON "companies" ("tenant");
CREATE INDEX "companies_name_idx" ON "companies" ("name");

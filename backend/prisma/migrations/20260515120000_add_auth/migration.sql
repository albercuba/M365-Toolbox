CREATE TYPE "AuthProvider" AS ENUM ('local', 'microsoft');
CREATE TYPE "UserRole" AS ENUM ('administrator', 'privileged_user', 'restricted_user');

CREATE TABLE "users" (
    "id" UUID NOT NULL,
    "username" TEXT NOT NULL,
    "password_hash" TEXT,
    "display_name" TEXT,
    "email" TEXT,
    "auth_provider" "AuthProvider" NOT NULL,
    "entra_object_id" TEXT,
    "role" "UserRole" NOT NULL,
    "must_change_password" BOOLEAN NOT NULL DEFAULT false,
    "created_at" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updated_at" TIMESTAMP(3) NOT NULL,
    "last_login_at" TIMESTAMP(3),

    CONSTRAINT "users_pkey" PRIMARY KEY ("id")
);

CREATE TABLE "microsoft_auth_config" (
    "id" TEXT NOT NULL DEFAULT 'singleton',
    "tenant_id" TEXT,
    "client_id" TEXT,
    "api_client_id" TEXT,
    "authority_url" TEXT,
    "enabled" BOOLEAN NOT NULL DEFAULT false,
    "created_at" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updated_at" TIMESTAMP(3) NOT NULL,

    CONSTRAINT "microsoft_auth_config_pkey" PRIMARY KEY ("id")
);

CREATE TABLE "entra_group_role_mappings" (
    "id" UUID NOT NULL,
    "group_name" TEXT NOT NULL,
    "group_id" TEXT,
    "assigned_role" "UserRole" NOT NULL,
    "created_at" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updated_at" TIMESTAMP(3) NOT NULL,

    CONSTRAINT "entra_group_role_mappings_pkey" PRIMARY KEY ("id")
);

ALTER TABLE "runs" ADD COLUMN "created_by_user_id" UUID;
ALTER TABLE "runs" ADD COLUMN "created_by_username" TEXT;
ALTER TABLE "runs" ADD COLUMN "created_by_display_name" TEXT;
ALTER TABLE "runs" ADD COLUMN "created_by_auth_provider" "AuthProvider";

CREATE UNIQUE INDEX "users_username_key" ON "users"("username");
CREATE UNIQUE INDEX "users_email_key" ON "users"("email");
CREATE UNIQUE INDEX "users_entra_object_id_key" ON "users"("entra_object_id");
CREATE INDEX "users_auth_provider_idx" ON "users"("auth_provider");
CREATE INDEX "users_role_idx" ON "users"("role");

CREATE UNIQUE INDEX "entra_group_role_mappings_group_id_key" ON "entra_group_role_mappings"("group_id");
CREATE INDEX "entra_group_role_mappings_group_name_idx" ON "entra_group_role_mappings"("group_name");
CREATE INDEX "entra_group_role_mappings_assigned_role_idx" ON "entra_group_role_mappings"("assigned_role");
CREATE INDEX "runs_created_by_user_id_created_at_idx" ON "runs"("created_by_user_id", "created_at" DESC);

ALTER TABLE "runs" ADD CONSTRAINT "runs_created_by_user_id_fkey" FOREIGN KEY ("created_by_user_id") REFERENCES "users"("id") ON DELETE SET NULL ON UPDATE CASCADE;

import fs from "node:fs";
import path from "node:path";
import { spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";

const currentDir = path.dirname(fileURLToPath(import.meta.url));
const backendRoot = path.resolve(currentDir, "..");
const schemaPath = path.join(backendRoot, "prisma", "schema.prisma");

if (process.env.SKIP_PRISMA_GENERATE === "1") {
  process.exit(0);
}

if (!fs.existsSync(schemaPath)) {
  process.exit(0);
}

const npxCommand = process.platform === "win32" ? "npx.cmd" : "npx";
const result = spawnSync(
  npxCommand,
  ["prisma", "generate", "--schema", "./prisma/schema.prisma"],
  {
    cwd: backendRoot,
    stdio: "inherit",
    env: process.env
  }
);

if (result.error) {
  throw result.error;
}

process.exit(result.status ?? 0);

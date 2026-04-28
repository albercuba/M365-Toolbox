import fs from "node:fs";
import path from "node:path";
import { spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";

const currentDir = path.dirname(fileURLToPath(import.meta.url));
const backendRoot = path.resolve(currentDir, "..");
const schemaPath = path.join(backendRoot, "prisma", "schema.prisma");
const configPath = path.join(backendRoot, "prisma.config.ts");

if (process.env.SKIP_PRISMA_GENERATE === "1" || !fs.existsSync(schemaPath) || !fs.existsSync(configPath)) {
  process.exit(0);
}

const prismaCliPathCandidates = [
  path.resolve(backendRoot, "..", "node_modules", "prisma", "build", "index.js"),
  path.resolve(backendRoot, "node_modules", "prisma", "build", "index.js")
];
const prismaCliPath = prismaCliPathCandidates.find((candidate) => fs.existsSync(candidate));

if (!prismaCliPath) {
  throw new Error("Unable to locate the Prisma CLI entrypoint for client generation.");
}

const result = spawnSync(
  process.execPath,
  [prismaCliPath, "generate", "--config", "./prisma.config.ts", "--schema", "./prisma/schema.prisma"],
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

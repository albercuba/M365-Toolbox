import { spawnSync } from "node:child_process";

const npxCommand = process.platform === "win32" ? "npx.cmd" : "npx";
const result = spawnSync(
  npxCommand,
  ["prisma", "migrate", "deploy", "--config", "./prisma.config.ts", "--schema", "./prisma/schema.prisma"],
  {
    stdio: "inherit",
    env: process.env
  }
);

if (result.error) {
  throw result.error;
}

process.exit(result.status ?? 0);

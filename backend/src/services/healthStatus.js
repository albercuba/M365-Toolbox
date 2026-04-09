import fs from "node:fs";
import path from "node:path";
import { execFile } from "node:child_process";

const OUTPUT_DIR = process.env.OUTPUT_DIR || path.resolve(process.cwd(), "../output");
const TOOLBOX_SCRIPT_MOUNT_ROOT = process.env.TOOLBOX_SCRIPT_MOUNT_ROOT || path.resolve(process.cwd(), "../scripts");

let lastStatus = null;
let lastStatusAt = 0;

function execPwsh(args) {
  return new Promise((resolve, reject) => {
    execFile("pwsh", args, { timeout: 10000 }, (error, stdout, stderr) => {
      if (error) {
        reject(new Error(stderr?.trim() || error.message));
        return;
      }

      resolve((stdout || "").trim());
    });
  });
}

function getModuleNames() {
  return execPwsh([
    "-NoLogo",
    "-NoProfile",
    "-Command",
    "(Get-Module -ListAvailable Microsoft.Graph.Authentication,ExchangeOnlineManagement,ImportExcel | Select-Object -ExpandProperty Name) -join ','"
  ]);
}

export async function getSystemStatus() {
  if (lastStatus && Date.now() - lastStatusAt < 15000) {
    return lastStatus;
  }

  const outputWritable = (() => {
    try {
      fs.mkdirSync(OUTPUT_DIR, { recursive: true });
      fs.accessSync(OUTPUT_DIR, fs.constants.W_OK);
      return true;
    } catch {
      return false;
    }
  })();

  const scriptsMounted = fs.existsSync(TOOLBOX_SCRIPT_MOUNT_ROOT);

  let powerShell = { available: false, version: null, error: null };
  let modules = { ready: false, installed: [], error: null };

  try {
    const version = await execPwsh(["-NoLogo", "-NoProfile", "-Command", "$PSVersionTable.PSVersion.ToString()"]);
    powerShell = { available: true, version, error: null };
  } catch (error) {
    powerShell = { available: false, version: null, error: error.message };
  }

  if (powerShell.available) {
    try {
      const installed = (await getModuleNames()).split(",").filter(Boolean);
      modules = {
        ready: installed.includes("Microsoft.Graph.Authentication") && installed.includes("ExchangeOnlineManagement"),
        installed,
        error: null
      };
    } catch (error) {
      modules = { ready: false, installed: [], error: error.message };
    }
  }

  lastStatus = {
    checkedAt: new Date().toISOString(),
    backend: { status: "ok" },
    paths: {
      outputDir: OUTPUT_DIR,
      outputWritable,
      scriptMountRoot: TOOLBOX_SCRIPT_MOUNT_ROOT,
      scriptsMounted
    },
    powerShell,
    modules
  };
  lastStatusAt = Date.now();
  return lastStatus;
}

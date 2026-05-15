import { ensureDefaultAdmin } from "../services/auth.js";
import { ensureDatabaseReady } from "../services/db.js";

await ensureDatabaseReady();
await ensureDefaultAdmin();
console.log("Default local administrator is present.");

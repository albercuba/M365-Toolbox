import { PrismaClient } from "@prisma/client";

const globalScope = globalThis;

export const prisma =
  globalScope.__m365ToolboxPrisma ||
  new PrismaClient({
    log: process.env.NODE_ENV === "development" ? ["error", "warn"] : ["error"]
  });

if (!globalScope.__m365ToolboxPrisma) {
  globalScope.__m365ToolboxPrisma = prisma;
}

export async function ensureDatabaseReady() {
  try {
    await prisma.$connect();
    await prisma.$queryRawUnsafe("SELECT 1");
  } catch (error) {
    throw new Error(`Unable to connect to PostgreSQL. Check DATABASE_URL and database availability. ${error.message}`);
  }

  try {
    await prisma.$queryRawUnsafe('SELECT 1 FROM "runs" LIMIT 1');
  } catch (error) {
    throw new Error(`PostgreSQL is reachable, but the schema is not ready. Run 'npm run db:migrate'. ${error.message}`);
  }
}

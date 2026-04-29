import { v4 as uuidv4 } from "uuid";
import { prisma } from "./db.js";
import { createError } from "./validation.js";

function normalizeCompany(company) {
  return {
    id: company.id,
    name: company.name,
    tenant: company.tenant,
    createdAt: company.createdAt instanceof Date ? company.createdAt.toISOString() : new Date(company.createdAt).toISOString(),
    updatedAt: company.updatedAt instanceof Date ? company.updatedAt.toISOString() : new Date(company.updatedAt).toISOString()
  };
}

function normalizeCompanyInput(input = {}) {
  const name = String(input.name || "").trim();
  const tenant = String(input.tenant || "").trim();

  if (!name || !tenant) {
    throw createError("Company name and tenant ID or domain are required.");
  }

  return { name, tenant };
}

function isUniqueConstraintError(error) {
  return error?.code === "P2002";
}

export async function listCompanies() {
  const companies = await prisma.company.findMany({
    orderBy: { name: "asc" }
  });
  return companies.map(normalizeCompany);
}

export async function createCompany(input) {
  const data = normalizeCompanyInput(input);

  try {
    const company = await prisma.company.create({
      data: {
        id: uuidv4(),
        ...data
      }
    });
    return normalizeCompany(company);
  } catch (error) {
    if (isUniqueConstraintError(error)) {
      throw createError("That company name or tenant value already exists.", 409);
    }
    throw error;
  }
}

export async function updateCompany(id, input) {
  const data = normalizeCompanyInput(input);

  try {
    const company = await prisma.company.update({
      where: { id },
      data
    });
    return normalizeCompany(company);
  } catch (error) {
    if (isUniqueConstraintError(error)) {
      throw createError("That company name or tenant value already exists.", 409);
    }
    if (error?.code === "P2025") {
      throw createError("Company not found.", 404);
    }
    throw error;
  }
}

export async function deleteCompany(id) {
  try {
    await prisma.company.delete({ where: { id } });
  } catch (error) {
    if (error?.code === "P2025") {
      throw createError("Company not found.", 404);
    }
    throw error;
  }
}

export async function replaceCompanies(inputs = []) {
  const normalized = inputs.map(normalizeCompanyInput);
  const seen = new Set();

  for (const company of normalized) {
    const nameKey = company.name.toLowerCase();
    const tenantKey = company.tenant.toLowerCase();
    if (seen.has(nameKey) || seen.has(tenantKey)) {
      throw createError("Imported companies contain duplicate names or tenant values.");
    }
    seen.add(nameKey);
    seen.add(tenantKey);
  }

  await prisma.$transaction([
    prisma.company.deleteMany({}),
    ...normalized.map((company) =>
      prisma.company.create({
        data: {
          id: uuidv4(),
          ...company
        }
      })
    )
  ]);

  return listCompanies();
}

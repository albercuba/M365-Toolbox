import { Router } from "express";
import { requireRole } from "../middleware/auth.js";
import {
  createGroupRoleMapping,
  createLocalUser,
  deleteGroupRoleMapping,
  deleteUser,
  getMicrosoftAuthConfig,
  listGroupRoleMappings,
  listUsers,
  updateGroupRoleMapping,
  updateMicrosoftAuthConfig,
  updateUser
} from "../services/auth.js";

export const authSettingsRouter = Router();

authSettingsRouter.use("/settings", requireRole("administrator"));

authSettingsRouter.get("/settings/auth/microsoft", async (_req, res, next) => {
  try {
    res.json(await getMicrosoftAuthConfig());
  } catch (error) {
    next(error);
  }
});

authSettingsRouter.put("/settings/auth/microsoft", async (req, res) => {
  try {
    res.json(await updateMicrosoftAuthConfig(req.body));
  } catch (error) {
    res.status(error.statusCode || 400).json({ message: error.message });
  }
});

authSettingsRouter.get("/settings/auth/group-role-mappings", async (_req, res, next) => {
  try {
    res.json(await listGroupRoleMappings());
  } catch (error) {
    next(error);
  }
});

authSettingsRouter.post("/settings/auth/group-role-mappings", async (req, res) => {
  try {
    res.status(201).json(await createGroupRoleMapping(req.body));
  } catch (error) {
    res.status(error.statusCode || 400).json({ message: error.message });
  }
});

authSettingsRouter.put("/settings/auth/group-role-mappings/:id", async (req, res) => {
  try {
    res.json(await updateGroupRoleMapping(req.params.id, req.body));
  } catch (error) {
    res.status(error.statusCode || 400).json({ message: error.message });
  }
});

authSettingsRouter.delete("/settings/auth/group-role-mappings/:id", async (req, res) => {
  try {
    await deleteGroupRoleMapping(req.params.id);
    res.status(204).end();
  } catch (error) {
    res.status(error.statusCode || 400).json({ message: error.message });
  }
});

authSettingsRouter.get("/settings/users", async (_req, res, next) => {
  try {
    res.json(await listUsers());
  } catch (error) {
    next(error);
  }
});

authSettingsRouter.post("/settings/users", async (req, res) => {
  try {
    res.status(201).json(await createLocalUser(req.body));
  } catch (error) {
    res.status(error.statusCode || 400).json({ message: error.message });
  }
});

authSettingsRouter.put("/settings/users/:id", async (req, res) => {
  try {
    res.json(await updateUser(req.params.id, req.body));
  } catch (error) {
    res.status(error.statusCode || 400).json({ message: error.message });
  }
});

authSettingsRouter.delete("/settings/users/:id", async (req, res) => {
  try {
    await deleteUser(req.params.id, req.user.id);
    res.status(204).end();
  } catch (error) {
    res.status(error.statusCode || 400).json({ message: error.message });
  }
});

import { Router } from "express";
import {
  changeLocalPassword,
  clearSessionCookie,
  getPublicMicrosoftAuthConfig,
  loginLocal,
  loginMicrosoft,
  setSessionCookie
} from "../services/auth.js";
import { requireAuth } from "../middleware/auth.js";

export const authRouter = Router();

authRouter.get("/auth/config", async (_req, res, next) => {
  try {
    res.json(await getPublicMicrosoftAuthConfig());
  } catch (error) {
    next(error);
  }
});

authRouter.post("/auth/login", async (req, res) => {
  try {
    const user = await loginLocal(req.body?.username, req.body?.password);
    setSessionCookie(res, user);
    res.json({ user });
  } catch (error) {
    res.status(error.statusCode || 400).json({ message: error.message });
  }
});

authRouter.post("/auth/microsoft", async (req, res) => {
  try {
    const user = await loginMicrosoft(req.body?.token);
    setSessionCookie(res, user);
    res.json({ user });
  } catch (error) {
    res.status(error.statusCode || 400).json({ message: error.message });
  }
});

authRouter.get("/auth/me", requireAuth, (req, res) => {
  res.json({ user: req.user });
});

authRouter.post("/auth/logout", (_req, res) => {
  clearSessionCookie(res);
  res.status(204).end();
});

authRouter.post("/auth/change-password", requireAuth, async (req, res) => {
  try {
    const user = await changeLocalPassword(req.user.id, req.body?.currentPassword, req.body?.newPassword);
    setSessionCookie(res, user);
    res.json({ user });
  } catch (error) {
    res.status(error.statusCode || 400).json({ message: error.message });
  }
});

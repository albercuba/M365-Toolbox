import { Router } from "express";
import {
  completeMicrosoftLogin,
  createMicrosoftLoginUrl,
  getSessionInfo,
  logoutSession,
  startAnonymousSession
} from "../services/auth.js";

export const authRouter = Router();

authRouter.get("/session", (req, res) => {
  res.json(getSessionInfo(req));
});

authRouter.post("/anonymous", (req, res) => {
  try {
    res.json(startAnonymousSession(req, res));
  } catch (error) {
    res.status(error.statusCode || 400).json({ message: error.message });
  }
});

authRouter.post("/logout", (req, res) => {
  logoutSession(req, res);
  res.json({ ok: true });
});

authRouter.get("/login", (req, res) => {
  try {
    res.json({ url: createMicrosoftLoginUrl(req) });
  } catch (error) {
    res.status(error.statusCode || 400).json({ message: error.message });
  }
});

authRouter.get("/callback", async (req, res, next) => {
  try {
    const redirectUrl = await completeMicrosoftLogin(req, res);
    res.redirect(302, redirectUrl);
  } catch (error) {
    next(error);
  }
});

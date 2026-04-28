import crypto from "node:crypto";
import { ARTIFACT_TOKEN_SECRET, ARTIFACT_TOKEN_TTL_SECONDS } from "../config/runtime.js";

function base64UrlEncode(value) {
  return Buffer.from(value).toString("base64url");
}

function base64UrlDecode(value) {
  return Buffer.from(value, "base64url").toString("utf8");
}

function signPayload(payload) {
  return crypto.createHmac("sha256", ARTIFACT_TOKEN_SECRET).update(payload).digest("base64url");
}

export function issueArtifactToken({ runId, artifactId = null, kind = "download" }) {
  const payload = JSON.stringify({
    runId,
    artifactId,
    kind,
    exp: Math.floor(Date.now() / 1000) + ARTIFACT_TOKEN_TTL_SECONDS
  });
  return `${base64UrlEncode(payload)}.${signPayload(payload)}`;
}

export function verifyArtifactToken(token, expected) {
  if (!token || typeof token !== "string" || !token.includes(".")) {
    return false;
  }

  try {
    const [encodedPayload, signature] = token.split(".", 2);
    const payload = base64UrlDecode(encodedPayload);
    if (signPayload(payload) !== signature) {
      return false;
    }

    const parsed = JSON.parse(payload);
    if (!parsed.exp || parsed.exp < Math.floor(Date.now() / 1000)) {
      return false;
    }

    return (
      parsed.runId === expected.runId &&
      parsed.kind === expected.kind &&
      (expected.artifactId === undefined || parsed.artifactId === expected.artifactId)
    );
  } catch {
    return false;
  }
}

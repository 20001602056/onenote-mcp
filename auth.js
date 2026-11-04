// Microsoft sign-in via device code flow.
// Tokens are kept in the macOS Keychain, or in a 0600 file on other platforms.
import { execFileSync } from "node:child_process";
import { chmodSync, mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { homedir } from "node:os";
import { dirname, join } from "node:path";
import { pathToFileURL } from "node:url";

// Microsoft's public "Microsoft Graph Command Line Tools" client — no app registration needed.
const CLIENT_ID = process.env.ONENOTE_CLIENT_ID || "14d82eec-204b-4c2f-b7e8-296a70dab67e";
// "consumers" = personal Microsoft accounts. Use "organizations", "common" or a tenant id for work/school accounts.
const TENANT = process.env.ONENOTE_TENANT || "consumers";
const AUTHORITY = `https://login.microsoftonline.com/${TENANT}/oauth2/v2.0`;
const SCOPE = "Notes.ReadWrite User.Read offline_access";
const KC_SERVICE = "onenote-mcp";
const KC_ACCOUNT = "graph-token";
const USE_KEYCHAIN = process.platform === "darwin" && !process.env.ONENOTE_TOKEN_FILE;
const TOKEN_FILE =
  process.env.ONENOTE_TOKEN_FILE || join(process.env.XDG_CONFIG_HOME || join(homedir(), ".config"), "onenote-mcp", "tokens.json");
const STORE_NAME = USE_KEYCHAIN ? `macOS Keychain (service: ${KC_SERVICE})` : TOKEN_FILE;

function loadTokens() {
  try {
    const raw = USE_KEYCHAIN
      ? execFileSync("security", ["find-generic-password", "-s", KC_SERVICE, "-a", KC_ACCOUNT, "-w"], {
          stdio: ["ignore", "pipe", "ignore"],
        })
      : readFileSync(TOKEN_FILE);
    return JSON.parse(raw.toString());
  } catch {
    return null;
  }
}

function saveTokens(t) {
  const data = JSON.stringify({
    access_token: t.access_token,
    refresh_token: t.refresh_token,
    expires_at: Date.now() + (t.expires_in - 60) * 1000,
  });
  if (USE_KEYCHAIN) {
    execFileSync("security", ["add-generic-password", "-U", "-s", KC_SERVICE, "-a", KC_ACCOUNT, "-w", data]);
    return;
  }
  mkdirSync(dirname(TOKEN_FILE), { recursive: true, mode: 0o700 });
  writeFileSync(TOKEN_FILE, data, { mode: 0o600 });
  chmodSync(TOKEN_FILE, 0o600);
}

async function tokenRequest(params) {
  const res = await fetch(`${AUTHORITY}/token`, {
    method: "POST",
    body: new URLSearchParams({ client_id: CLIENT_ID, scope: SCOPE, ...params }),
  });
  return res.json();
}

export async function getAccessToken() {
  const t = loadTokens();
  if (!t) throw new Error("Not signed in to OneNote. Run `npm run login` in the onenote-mcp folder.");
  if (Date.now() < t.expires_at) return t.access_token;
  const r = await tokenRequest({ grant_type: "refresh_token", refresh_token: t.refresh_token });
  if (!r.access_token) throw new Error(`Token refresh failed (${r.error}). Run \`npm run login\` in the onenote-mcp folder.`);
  saveTokens(r);
  return r.access_token;
}

async function login() {
  const res = await fetch(`${AUTHORITY}/devicecode`, {
    method: "POST",
    body: new URLSearchParams({ client_id: CLIENT_ID, scope: SCOPE }),
  });
  const dc = await res.json();
  if (!dc.device_code) throw new Error(dc.error_description || dc.error || "Could not start sign-in");
  console.log(dc.message);
  const deadline = Date.now() + dc.expires_in * 1000;
  while (Date.now() < deadline) {
    await new Promise((r) => setTimeout(r, dc.interval * 1000));
    const r = await tokenRequest({ grant_type: "urn:ietf:params:oauth:grant-type:device_code", device_code: dc.device_code });
    if (r.access_token) {
      saveTokens(r);
      console.log(`Signed in. Tokens saved to ${STORE_NAME}.`);
      return;
    }
    if (r.error !== "authorization_pending" && r.error !== "slow_down") throw new Error(r.error_description || r.error);
  }
  throw new Error("Sign-in code expired. Run login again.");
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href && process.argv[2] === "login") {
  login().catch((e) => {
    console.error(e.message);
    process.exit(1);
  });
}

// Settings, all from environment variables (see README "Configuration").
import { homedir } from "node:os";
import { join, resolve } from "node:path";

export const env = (name, fallback = "") => (process.env[name] ?? "").trim() || fallback;
export const flag = (name, fallback = false) => {
  const v = env(name).toLowerCase();
  return v ? ["1", "true", "yes"].includes(v) : fallback;
};
export const list = (name) =>
  env(name)
    .split(",")
    .map((s) => s.trim())
    .filter(Boolean);
export const expandHome = (p) => resolve(p.replace(/^~(?=$|[\\/])/, homedir()));

export const READ_ONLY = flag("ONENOTE_READ_ONLY");
export const ALLOW_DELETE = flag("ONENOTE_ALLOW_DELETE");
export const DEFAULT_NOTEBOOK = env("ONENOTE_DEFAULT_NOTEBOOK");
export const DEFAULT_SECTION = env("ONENOTE_DEFAULT_SECTION");
export const DOWNLOAD_DIR = expandHome(env("ONENOTE_DOWNLOAD_DIR", join(homedir(), "Downloads")));
export const EXPORT_DIR = expandHome(env("ONENOTE_EXPORT_DIR", DOWNLOAD_DIR));
export const CHROME_PATH = env("ONENOTE_CHROME_PATH");

// Versioning: local snapshots of pages, kept for RETENTION_DAYS.
export const VERSIONING = flag("ONENOTE_VERSIONING", true);
export const VERSION_DIR = expandHome(
  env("ONENOTE_VERSION_DIR", join(process.env.XDG_DATA_HOME || join(homedir(), ".local", "share"), "onenote-mcp", "versions"))
);
export const RETENTION_DAYS = Math.max(1, Number(env("ONENOTE_VERSION_RETENTION_DAYS", "30")) || 30);

export const AUDIT_LOG = env("ONENOTE_AUDIT_LOG") && expandHome(env("ONENOTE_AUDIT_LOG"));

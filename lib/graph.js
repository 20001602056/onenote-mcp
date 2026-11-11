// Microsoft Graph helpers (signed in as you, see auth.js).
import { getAccessToken } from "../auth.js";

export const GRAPH = "https://graph.microsoft.com/v1.0/me/onenote";
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

// Retries throttled (429) and briefly unavailable (503/504) requests, honouring Retry-After.
export async function graph(path, { method = "GET", body, contentType, raw = false, binary = false, full = false } = {}) {
  const url = path.startsWith("http") ? path : GRAPH + path;
  for (let attempt = 0; ; attempt++) {
    const headers = { Authorization: `Bearer ${await getAccessToken()}` };
    if (contentType) headers["Content-Type"] = contentType;
    const res = await fetch(url, { method, headers, body });
    if ([429, 503, 504].includes(res.status) && attempt < 4) {
      await res.arrayBuffer().catch(() => {});
      await sleep(Math.min(Number(res.headers.get("Retry-After")) * 1000 || 2 ** attempt * 1000, 30000));
      continue;
    }
    if (!res.ok) throw new Error(`Graph ${method} ${path} -> ${res.status}: ${await res.text()}`);
    if (full) {
      const textBody = await res.text();
      return { status: res.status, location: res.headers.get("Operation-Location"), body: textBody ? JSON.parse(textBody) : null };
    }
    if (res.status === 204) return null;
    if (binary) return Buffer.from(await res.arrayBuffer());
    return raw ? res.text() : res.json();
  }
}

export async function getAll(path, limit = Infinity) {
  const items = [];
  let next = path;
  while (next && items.length < limit) {
    const r = await graph(next);
    items.push(...r.value);
    next = r["@odata.nextLink"];
  }
  return items.slice(0, limit);
}

export const notFound = (e) => /-> 404:/.test(e.message);
export const notImplemented = (e) => /-> 501:/.test(e.message);

export const q = (s) => s.replace(/'/g, "''");
export const esc = (s) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
export const same = (a, b) => (a || "").toLowerCase() === (b || "").toLowerCase();

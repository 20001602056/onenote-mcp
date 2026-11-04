#!/usr/bin/env node
// OneNote MCP server backed by Microsoft Graph (signed in as you, see auth.js).
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { getAccessToken } from "./auth.js";

const GRAPH = "https://graph.microsoft.com/v1.0/me/onenote";

async function graph(path, { method = "GET", body, contentType, raw = false } = {}) {
  const url = path.startsWith("http") ? path : GRAPH + path;
  const headers = { Authorization: `Bearer ${await getAccessToken()}` };
  if (contentType) headers["Content-Type"] = contentType;
  const res = await fetch(url, { method, headers, body });
  if (!res.ok) throw new Error(`Graph ${method} ${path} -> ${res.status}: ${await res.text()}`);
  if (res.status === 204) return null;
  return raw ? res.text() : res.json();
}

async function getAll(path) {
  const items = [];
  let next = path;
  while (next) {
    const r = await graph(next);
    items.push(...r.value);
    next = r["@odata.nextLink"];
  }
  return items;
}

// Guardrail: sections that must never be read, listed, searched or written.
// Matched by section name (case-insensitive) or section id, so renaming the section doesn't bypass it.
// Set with ONENOTE_BLOCKED_SECTIONS="Name1,Name2,<sectionId>". Nothing is blocked by default.
const BLOCKED_SECTIONS = (process.env.ONENOTE_BLOCKED_SECTIONS ?? "")
  .split(",")
  .map((s) => s.trim().toLowerCase())
  .filter(Boolean);

const isBlocked = (section) =>
  BLOCKED_SECTIONS.includes((section.displayName || "").toLowerCase()) || BLOCKED_SECTIONS.includes((section.id || "").toLowerCase());

const BLOCKED_MSG = "Access denied by guardrail: this section is blocked.";

async function assertSectionAllowed(sectionId) {
  const s = await graph(`/sections/${sectionId}?$select=id,displayName`);
  if (isBlocked(s)) throw new Error(BLOCKED_MSG);
  return s;
}

async function assertPageAllowed(pageId) {
  const p = await graph(`/pages/${pageId}?$select=id&$expand=parentSection($select=id,displayName)`);
  if (!p.parentSection || isBlocked(p.parentSection)) throw new Error(BLOCKED_MSG);
}

async function resolveSection({ sectionId, notebook, section }) {
  if (sectionId) return (await assertSectionAllowed(sectionId)).id;
  if (!notebook || !section) throw new Error("Provide sectionId, or both notebook and section");
  return (await findSection(notebook, section)).id;
}

const esc = (s) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

async function findSection(notebookName, sectionName) {
  const notebooks = await getAll(`/notebooks?$filter=displayName eq '${notebookName.replace(/'/g, "''")}'`);
  if (!notebooks.length) throw new Error(`Notebook "${notebookName}" not found`);
  const sections = (await getAll(`/notebooks/${notebooks[0].id}/sections`)).filter((x) => !isBlocked(x));
  if (BLOCKED_SECTIONS.includes(sectionName.toLowerCase())) throw new Error(BLOCKED_MSG);
  const s = sections.find((x) => x.displayName.toLowerCase() === sectionName.toLowerCase());
  if (!s) throw new Error(`Section "${sectionName}" not found. Available: ${sections.map((x) => x.displayName).join(", ")}`);
  return s;
}

const text = (v) => ({ content: [{ type: "text", text: typeof v === "string" ? v : JSON.stringify(v, null, 2) }] });

const server = new McpServer({ name: "onenote", version: "1.0.0" });

server.tool("list_notebooks", "List all OneNote notebooks with their sections.", {}, async () => {
  const notebooks = await getAll("/notebooks?$expand=sections($select=id,displayName)&$select=id,displayName");
  return text(notebooks.map((n) => ({ notebook: n.displayName, id: n.id, sections: n.sections.filter((s) => !isBlocked(s)).map((s) => ({ name: s.displayName, id: s.id })) })));
});

server.tool(
  "create_section",
  "Create a new section in a notebook (or return the existing one with that name).",
  { notebook: z.string(), section: z.string() },
  async ({ notebook, section }) => {
    if (BLOCKED_SECTIONS.includes(section.toLowerCase())) throw new Error(BLOCKED_MSG);
    const notebooks = await getAll(`/notebooks?$filter=displayName eq '${notebook.replace(/'/g, "''")}'`);
    if (!notebooks.length) throw new Error(`Notebook "${notebook}" not found`);
    const existing = (await getAll(`/notebooks/${notebooks[0].id}/sections`)).find((x) => x.displayName.toLowerCase() === section.toLowerCase());
    if (existing) return text({ section: existing.displayName, id: existing.id, created: false });
    const s = await graph(`/notebooks/${notebooks[0].id}/sections`, {
      method: "POST",
      contentType: "application/json",
      body: JSON.stringify({ displayName: section }),
    });
    return text({ section: s.displayName, id: s.id, created: true });
  }
);

server.tool(
  "list_pages",
  "List pages in a section, newest first. Identify the section by sectionId, or by notebook + section names.",
  { sectionId: z.string().optional(), notebook: z.string().optional(), section: z.string().optional() },
  async ({ sectionId, notebook, section }) => {
    const id = await resolveSection({ sectionId, notebook, section });
    const pages = await getAll(`/sections/${id}/pages?$select=id,title,createdDateTime,lastModifiedDateTime&$orderby=lastModifiedDateTime desc&$top=100`);
    return text(pages.map((p) => ({ title: p.title, id: p.id, created: p.createdDateTime, modified: p.lastModifiedDateTime })));
  }
);

server.tool("read_page", "Read a page's content as HTML.", { pageId: z.string() }, async ({ pageId }) => {
  await assertPageAllowed(pageId);
  return text(await graph(`/pages/${pageId}/content`, { raw: true }));
});

server.tool(
  "search_pages",
  "Search page titles (case-insensitive substring) in a section, identified by sectionId or notebook + section names.",
  { query: z.string(), sectionId: z.string().optional(), notebook: z.string().optional(), section: z.string().optional() },
  async ({ query, sectionId, notebook, section }) => {
    const id = await resolveSection({ sectionId, notebook, section });
    const pages = await getAll(`/sections/${id}/pages?$select=id,title,lastModifiedDateTime&$top=100`);
    const q = query.toLowerCase();
    return text(pages.filter((p) => (p.title || "").toLowerCase().includes(q)).map((p) => ({ title: p.title, id: p.id, modified: p.lastModifiedDateTime })));
  }
);

server.tool(
  "create_page",
  "Create a page in a section. bodyHtml is simple HTML (p, h1-h3, ul/ol/li, pre, code, b, i, table, a). Identify the section by sectionId or notebook + section names.",
  { title: z.string(), bodyHtml: z.string(), sectionId: z.string().optional(), notebook: z.string().optional(), section: z.string().optional() },
  async ({ title, bodyHtml, sectionId, notebook, section }) => {
    const id = await resolveSection({ sectionId, notebook, section });
    const html = `<!DOCTYPE html><html><head><title>${esc(title)}</title><meta name="created" content="${new Date().toISOString()}" /></head><body>${bodyHtml}</body></html>`;
    const page = await graph(`/sections/${id}/pages`, { method: "POST", body: html, contentType: "application/xhtml+xml" });
    return text({ created: page.title, id: page.id, link: page.links?.oneNoteWebUrl?.href });
  }
);

server.tool(
  "append_to_page",
  "Append HTML content to the end of an existing page.",
  { pageId: z.string(), html: z.string() },
  async ({ pageId, html }) => {
    await assertPageAllowed(pageId);
    await graph(`/pages/${pageId}/content`, {
      method: "PATCH",
      contentType: "application/json",
      body: JSON.stringify([{ target: "body", action: "append", content: html }]),
    });
    return text("Appended.");
  }
);

await server.connect(new StdioServerTransport());

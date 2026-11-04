#!/usr/bin/env node
// OneNote MCP server backed by Microsoft Graph (signed in as you, see auth.js).
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { readFile, writeFile, mkdir } from "node:fs/promises";
import { basename, dirname, extname, resolve } from "node:path";
import { homedir } from "node:os";
import { randomUUID } from "node:crypto";
import { z } from "zod";
import { getAccessToken } from "./auth.js";

const GRAPH = "https://graph.microsoft.com/v1.0/me/onenote";

// ---- Settings (environment variables) ----
const env = (name, fallback = "") => (process.env[name] ?? fallback).trim();
const flag = (name) => ["1", "true", "yes"].includes(env(name).toLowerCase());
const READ_ONLY = flag("ONENOTE_READ_ONLY");
const ALLOW_DELETE = flag("ONENOTE_ALLOW_DELETE");
const DEFAULT_NOTEBOOK = env("ONENOTE_DEFAULT_NOTEBOOK");
const DEFAULT_SECTION = env("ONENOTE_DEFAULT_SECTION");
const DOWNLOAD_DIR = env("ONENOTE_DOWNLOAD_DIR", resolve(homedir(), "Downloads"));

// ---- Graph helpers ----
async function graph(path, { method = "GET", body, contentType, raw = false, binary = false, full = false } = {}) {
  const url = path.startsWith("http") ? path : GRAPH + path;
  const headers = { Authorization: `Bearer ${await getAccessToken()}` };
  if (contentType) headers["Content-Type"] = contentType;
  const res = await fetch(url, { method, headers, body });
  if (!res.ok) throw new Error(`Graph ${method} ${path} -> ${res.status}: ${await res.text()}`);
  if (full) {
    const textBody = await res.text();
    return { status: res.status, location: res.headers.get("Operation-Location"), body: textBody ? JSON.parse(textBody) : null };
  }
  if (res.status === 204) return null;
  if (binary) return Buffer.from(await res.arrayBuffer());
  return raw ? res.text() : res.json();
}

async function getAll(path, limit = Infinity) {
  const items = [];
  let next = path;
  while (next && items.length < limit) {
    const r = await graph(next);
    items.push(...r.value);
    next = r["@odata.nextLink"];
  }
  return items.slice(0, limit);
}

const q = (s) => s.replace(/'/g, "''");
const esc = (s) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
const text = (v) => ({ content: [{ type: "text", text: typeof v === "string" ? v : JSON.stringify(v, null, 2) }] });
const same = (a, b) => (a || "").toLowerCase() === (b || "").toLowerCase();

// ---- Guardrail ----
// Sections that must never be read, listed, searched, copied or written.
// Matched by section name (case-insensitive) or section id, so renaming the section doesn't bypass it.
// Set with ONENOTE_BLOCKED_SECTIONS="Name1,Name2,<sectionId>". Nothing is blocked by default.
const BLOCKED_SECTIONS = env("ONENOTE_BLOCKED_SECTIONS")
  .split(",")
  .map((s) => s.trim().toLowerCase())
  .filter(Boolean);

const isBlocked = (section) =>
  BLOCKED_SECTIONS.includes((section.displayName || "").toLowerCase()) || BLOCKED_SECTIONS.includes((section.id || "").toLowerCase());
const isBlockedName = (name) => BLOCKED_SECTIONS.includes((name || "").toLowerCase());

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

// ---- Lookups ----
// Every section in every notebook, including sections inside section groups.
async function allSections() {
  return getAll(
    "/sections?$select=id,displayName,isDefault,createdDateTime,lastModifiedDateTime" +
      "&$expand=parentNotebook($select=id,displayName),parentSectionGroup($select=id,displayName)"
  );
}

async function allSectionGroups() {
  return getAll("/sectionGroups?$select=id,displayName&$expand=parentNotebook($select=id,displayName),parentSectionGroup($select=id,displayName)");
}

async function findNotebook(name) {
  const nbs = await getAll(`/notebooks?$filter=displayName eq '${q(name)}'&$select=id,displayName`);
  if (!nbs.length) throw new Error(`Notebook "${name}" not found`);
  return nbs[0];
}

async function findSectionGroup(notebookName, groupName) {
  const groups = (await allSectionGroups()).filter((g) => same(g.parentNotebook?.displayName, notebookName) && same(g.displayName, groupName));
  if (!groups.length) throw new Error(`Section group "${groupName}" not found in notebook "${notebookName}"`);
  return groups[0];
}

async function findSection(notebookName, sectionName, sectionGroupName) {
  if (isBlockedName(sectionName)) throw new Error(BLOCKED_MSG);
  const inNotebook = (await allSections()).filter((s) => same(s.parentNotebook?.displayName, notebookName) && !isBlocked(s));
  if (!inNotebook.length) throw new Error(`Notebook "${notebookName}" not found or has no sections`);
  const matches = inNotebook.filter(
    (s) => same(s.displayName, sectionName) && (!sectionGroupName || same(s.parentSectionGroup?.displayName, sectionGroupName))
  );
  if (!matches.length) {
    throw new Error(`Section "${sectionName}" not found. Available: ${inNotebook.map((s) => s.displayName).join(", ")}`);
  }
  // Prefer a top-level section when the same name exists inside a section group too.
  return matches.find((s) => !s.parentSectionGroup) || matches[0];
}

async function resolveSection({ sectionId, notebook, section, sectionGroup }) {
  if (sectionId) return (await assertSectionAllowed(sectionId)).id;
  notebook ||= DEFAULT_NOTEBOOK;
  section ||= DEFAULT_SECTION;
  if (!notebook || !section) {
    throw new Error("Provide sectionId, or notebook + section (or set ONENOTE_DEFAULT_NOTEBOOK / ONENOTE_DEFAULT_SECTION)");
  }
  return (await findSection(notebook, section, sectionGroup)).id;
}

const pageSummary = (p) => ({
  title: p.title,
  id: p.id,
  created: p.createdDateTime,
  modified: p.lastModifiedDateTime,
  ...(p.level !== undefined && { level: p.level, order: p.order }),
  ...(p.parentSection && { section: p.parentSection.displayName, sectionId: p.parentSection.id }),
  ...(p.links && { webUrl: p.links.oneNoteWebUrl?.href, clientUrl: p.links.oneNoteClientUrl?.href }),
});

// ---- Attachments (images and files embedded in pages) ----
const MIME = {
  ".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".gif": "image/gif", ".bmp": "image/bmp",
  ".tif": "image/tiff", ".tiff": "image/tiff", ".pdf": "application/pdf", ".txt": "text/plain", ".csv": "text/csv",
  ".json": "application/json", ".zip": "application/zip", ".doc": "application/msword",
  ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  ".xls": "application/vnd.ms-excel", ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  ".ppt": "application/vnd.ms-powerpoint", ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
};

const attachmentSchema = z
  .array(
    z.object({
      path: z.string().describe("Local file path to upload"),
      partName: z.string().regex(/^[A-Za-z0-9_-]+$/).optional().describe("Name to reference it in HTML as name:<partName>. Defaults to file1, file2, ..."),
      contentType: z.string().optional().describe("MIME type. Guessed from the extension if omitted"),
    })
  )
  .optional()
  .describe(
    "Local images/files to embed. Reference them in HTML as <img src=\"name:file1\"/> or " +
      "<object data=\"name:file1\" data-attachment=\"report.pdf\" type=\"application/pdf\"/>. Unreferenced parts are added at the end automatically."
  );

// Hand-built multipart/form-data body (Graph rejects some FormData encodings).
function multipart(parts) {
  const boundary = `onenote-mcp-${randomUUID()}`;
  const chunks = [];
  for (const p of parts) {
    const disposition = `form-data; name="${p.name}"${p.filename ? `; filename="${p.filename.replace(/"/g, "")}"` : ""}`;
    chunks.push(Buffer.from(`--${boundary}\r\nContent-Disposition: ${disposition}\r\nContent-Type: ${p.type}\r\n\r\n`));
    chunks.push(Buffer.isBuffer(p.data) ? p.data : Buffer.from(p.data), Buffer.from("\r\n"));
  }
  chunks.push(Buffer.from(`--${boundary}--\r\n`));
  return { body: Buffer.concat(chunks), contentType: `multipart/form-data; boundary=${boundary}` };
}

// Multipart body for page create ("Presentation" HTML part) or update ("Commands" JSON part), plus local files.
async function buildMultipart(main, attachments, mainType) {
  let content = main;
  const extra = [];
  const parts = [];
  for (const [i, a] of attachments.entries()) {
    const part = a.partName || `file${i + 1}`;
    const file = resolve(a.path.replace(/^~(?=$|\/)/, homedir()));
    const type = a.contentType || MIME[extname(file).toLowerCase()] || "application/octet-stream";
    parts.push({ name: part, data: await readFile(file), type, filename: basename(file) });
    if (!content.includes(`name:${part}`)) {
      extra.push(
        type.startsWith("image/")
          ? `<img src="name:${part}" alt="${esc(basename(file))}" />`
          : `<object data-attachment="${esc(basename(file))}" data="name:${part}" type="${type}" />`
      );
    }
  }
  if (extra.length) {
    if (mainType === "application/json") {
      const ops = JSON.parse(content);
      ops.push({ target: "body", action: "append", content: extra.join("") });
      content = JSON.stringify(ops);
    } else {
      content = content.includes("</body>") ? content.replace("</body>", `${extra.join("")}</body>`) : content + extra.join("");
    }
  }
  const mainName = mainType === "application/json" ? "Commands" : "Presentation";
  return multipart([{ name: mainName, type: mainType, data: content }, ...parts]);
}

async function patchPage(pageId, ops, attachments) {
  const json = JSON.stringify(ops);
  const req = attachments?.length ? await buildMultipart(json, attachments, "application/json") : { body: json, contentType: "application/json" };
  await graph(`/pages/${pageId}/content`, { method: "PATCH", ...req });
}

// Graph can only "replace" some elements (e.g. <p>) through OneNote's generated id, not a custom data-id.
// Swap "#my-data-id" targets for the generated id when needed.
async function resolveReplaceTargets(pageId, ops) {
  const wanted = ops.filter((o) => o.action === "replace" && o.target.startsWith("#"));
  if (!wanted.length) return ops;
  const html = await graph(`/pages/${pageId}/content?includeIDs=true`, { raw: true });
  const tags = [...html.matchAll(/<[a-z0-9]+\b[^>]*>/gi)].map((m) => m[0]);
  const attr = (tag, name) => tag.match(new RegExp(`\\s${name}="([^"]*)"`))?.[1];
  return ops.map((o) => {
    if (!wanted.includes(o)) return o;
    const key = o.target.slice(1);
    const tag = tags.find((t) => attr(t, "data-id") === key);
    const generated = tag && attr(tag, "id");
    return generated ? { ...o, target: generated } : o;
  });
}

// Copy a page by reading it, downloading its images/files and recreating it (personal accounts
// don't support Graph's copyToSection).
async function copyPageManually(pageId, destSectionId) {
  let html = await graph(`/pages/${pageId}/content`, { raw: true });
  const types = new Map();
  for (const [tag] of html.matchAll(/<(?:img|object)\b[^>]*>/gi)) {
    const url = tag.match(/\s(?:src|data)="([^"]*\/resources\/[^"]*)"/)?.[1];
    const type = tag.match(/\s(?:data-src-type|type)="([^"]*)"/)?.[1];
    if (url && type) types.set(url, type);
  }
  const urls = [...new Set([...html.matchAll(/https:\/\/graph\.microsoft\.com\/[^"]*?\/resources\/[^"/]+\/(?:\$value|content)/g)].map((m) => m[0]))];
  const parts = [];
  for (const [i, url] of urls.entries()) {
    const resourceId = url.match(/resources\/([^"/]+)\//)[1];
    const name = `res${i + 1}`;
    parts.push({ name, type: types.get(url) || "application/octet-stream", data: await graph(`/resources/${encodeURIComponent(resourceId)}/content`, { binary: true }) });
    html = html.split(url).join(`name:${name}`);
  }
  html = html.replace(/\sdata-fullres-src(?:-type)?="[^"]*"/g, "");
  const req = multipart([{ name: "Presentation", type: "text/html", data: html }, ...parts]);
  const page = await graph(`/sections/${destSectionId}/pages`, { method: "POST", ...req });
  return { id: page.id, title: page.title, resourcesCopied: parts.length, webUrl: page.links?.oneNoteWebUrl?.href };
}

const notImplemented = (e) => /-> 501:/.test(e.message);

// ---- Server ----
const server = new McpServer({ name: "onenote", version: "2.0.0" });
const RO = { readOnlyHint: true, openWorldHint: true };
const WRITE = { readOnlyHint: false, destructiveHint: false, openWorldHint: true };
const DESTRUCTIVE = { readOnlyHint: false, destructiveHint: true, openWorldHint: true };

function tool(name, description, inputSchema, annotations, handler) {
  if (READ_ONLY && !annotations.readOnlyHint) return;
  server.registerTool(name, { description, inputSchema, annotations }, handler);
}

const sectionTarget = {
  sectionId: z.string().optional().describe("Section id (takes precedence over names)"),
  notebook: z.string().optional().describe("Notebook name"),
  section: z.string().optional().describe("Section name"),
  sectionGroup: z.string().optional().describe("Section group name, to pick a section inside a group"),
};

// --- Notebooks ---
tool(
  "list_notebooks",
  "List all notebooks with their full structure: section groups (nested) and sections, with ids and links.",
  {},
  RO,
  async () => {
    const [notebooks, groups, sections] = await Promise.all([
      getAll("/notebooks?$select=id,displayName,isDefault,isShared,userRole,createdDateTime,lastModifiedDateTime,links"),
      allSectionGroups(),
      allSections(),
    ]);
    const visible = sections.filter((s) => !isBlocked(s));
    const sectionsOf = (pred) => visible.filter(pred).map((s) => ({ name: s.displayName, id: s.id, ...(s.isDefault && { isDefault: true }) }));
    const groupTree = (pred) =>
      groups.filter(pred).map((g) => ({
        sectionGroup: g.displayName,
        id: g.id,
        sections: sectionsOf((s) => s.parentSectionGroup?.id === g.id),
        sectionGroups: groupTree((c) => c.parentSectionGroup?.id === g.id),
      }));
    return text(
      notebooks.map((n) => ({
        notebook: n.displayName,
        id: n.id,
        isDefault: n.isDefault,
        isShared: n.isShared,
        role: n.userRole,
        modified: n.lastModifiedDateTime,
        webUrl: n.links?.oneNoteWebUrl?.href,
        sections: sectionsOf((s) => s.parentNotebook?.id === n.id && !s.parentSectionGroup),
        sectionGroups: groupTree((g) => g.parentNotebook?.id === n.id && !g.parentSectionGroup),
      }))
    );
  }
);

tool(
  "get_recent_notebooks",
  "List recently opened notebooks with last-access time and links to open them in OneNote.",
  {},
  RO,
  async () => {
    const r = await graph("/notebooks/getRecentNotebooks(includePersonalNotebooks=true)");
    return text(
      r.value.map((n) => ({
        notebook: n.displayName,
        lastAccessed: n.lastAccessedTime,
        source: n.sourceService,
        webUrl: n.links?.oneNoteWebUrl?.href,
        clientUrl: n.links?.oneNoteClientUrl?.href,
      }))
    );
  }
);

tool("create_notebook", "Create a new notebook (or return the existing one with that name).", { name: z.string() }, WRITE, async ({ name }) => {
  const existing = await getAll(`/notebooks?$filter=displayName eq '${q(name)}'&$select=id,displayName`);
  if (existing.length) return text({ notebook: existing[0].displayName, id: existing[0].id, created: false });
  const n = await graph("/notebooks", { method: "POST", contentType: "application/json", body: JSON.stringify({ displayName: name }) });
  return text({ notebook: n.displayName, id: n.id, created: true, webUrl: n.links?.oneNoteWebUrl?.href });
});

// --- Section groups ---
tool(
  "create_section_group",
  "Create a section group in a notebook, or nested inside another section group (or return the existing one).",
  {
    notebook: z.string(),
    name: z.string(),
    parentSectionGroup: z.string().optional().describe("Name of an existing section group to nest inside"),
  },
  WRITE,
  async ({ notebook, name, parentSectionGroup }) => {
    const parentPath = parentSectionGroup
      ? `/sectionGroups/${(await findSectionGroup(notebook, parentSectionGroup)).id}`
      : `/notebooks/${(await findNotebook(notebook)).id}`;
    const existing = (await getAll(`${parentPath}/sectionGroups?$select=id,displayName`)).find((g) => same(g.displayName, name));
    if (existing) return text({ sectionGroup: existing.displayName, id: existing.id, created: false });
    const g = await graph(`${parentPath}/sectionGroups`, { method: "POST", contentType: "application/json", body: JSON.stringify({ displayName: name }) });
    return text({ sectionGroup: g.displayName, id: g.id, created: true });
  }
);

// --- Sections ---
tool(
  "list_sections",
  "List sections (including those inside section groups), optionally only in one notebook.",
  { notebook: z.string().optional() },
  RO,
  async ({ notebook }) => {
    const sections = (await allSections()).filter((s) => !isBlocked(s) && (!notebook || same(s.parentNotebook?.displayName, notebook)));
    return text(
      sections.map((s) => ({
        section: s.displayName,
        id: s.id,
        notebook: s.parentNotebook?.displayName,
        sectionGroup: s.parentSectionGroup?.displayName,
        isDefault: s.isDefault,
        modified: s.lastModifiedDateTime,
      }))
    );
  }
);

tool(
  "create_section",
  "Create a section in a notebook or inside a section group (or return the existing one with that name).",
  { notebook: z.string(), section: z.string(), sectionGroup: z.string().optional().describe("Create inside this section group") },
  WRITE,
  async ({ notebook, section, sectionGroup }) => {
    if (isBlockedName(section)) throw new Error(BLOCKED_MSG);
    const parentPath = sectionGroup
      ? `/sectionGroups/${(await findSectionGroup(notebook, sectionGroup)).id}`
      : `/notebooks/${(await findNotebook(notebook)).id}`;
    const existing = (await getAll(`${parentPath}/sections?$select=id,displayName`)).find((x) => same(x.displayName, section));
    if (existing) return text({ section: existing.displayName, id: existing.id, created: false });
    const s = await graph(`${parentPath}/sections`, { method: "POST", contentType: "application/json", body: JSON.stringify({ displayName: section }) });
    return text({ section: s.displayName, id: s.id, created: true });
  }
);

tool(
  "copy_section",
  "Copy a section (with all its pages) into a notebook or section group. Work/school accounts copy in the background " +
    "(check with get_operation); personal accounts copy page by page immediately (slower for big sections).",
  {
    ...sectionTarget,
    toNotebook: z.string().describe("Destination notebook name"),
    toSectionGroup: z.string().optional().describe("Destination section group name (inside toNotebook)"),
    renameAs: z.string().optional().describe("Name for the copy"),
  },
  WRITE,
  async ({ toNotebook, toSectionGroup, renameAs, ...src }) => {
    const id = await resolveSection(src);
    if (renameAs && isBlockedName(renameAs)) throw new Error(BLOCKED_MSG);
    const dest = toSectionGroup ? await findSectionGroup(toNotebook, toSectionGroup) : await findNotebook(toNotebook);
    const action = toSectionGroup ? "copyToSectionGroup" : "copyToNotebook";
    try {
      const r = await graph(`/sections/${id}/${action}`, {
        method: "POST",
        contentType: "application/json",
        body: JSON.stringify({ id: dest.id, ...(renameAs && { renameAs }) }),
        full: true,
      });
      return text({ started: true, operationId: r.body?.id, status: r.body?.status, statusUrl: r.location });
    } catch (e) {
      if (!notImplemented(e)) throw e;
    }
    // Manual fallback: new section in the destination, then copy each page (oldest first).
    const source = await graph(`/sections/${id}?$select=id,displayName`);
    const name = renameAs || source.displayName;
    const parentPath = toSectionGroup ? `/sectionGroups/${dest.id}` : `/notebooks/${dest.id}`;
    if ((await getAll(`${parentPath}/sections?$select=id,displayName`)).some((s) => same(s.displayName, name))) {
      throw new Error(`A section named "${name}" already exists there. Pass renameAs to copy under a new name.`);
    }
    const target = await graph(`${parentPath}/sections`, { method: "POST", contentType: "application/json", body: JSON.stringify({ displayName: name }) });
    const pages = await getAll(`/sections/${id}/pages?$select=id,title&$orderby=createdDateTime asc&$top=100`);
    const copied = [];
    const failedPages = [];
    for (const p of pages) {
      try {
        copied.push((await copyPageManually(p.id, target.id)).title);
      } catch (err) {
        failedPages.push({ title: p.title, error: err.message.slice(0, 200) });
      }
    }
    return text({ copied: true, method: "manual", section: name, sectionId: target.id, pagesCopied: copied.length, failedPages });
  }
);

// --- Pages ---
tool(
  "list_pages",
  "List pages in a section with their indent level and order, newest first by default.",
  {
    ...sectionTarget,
    orderBy: z.enum(["lastModifiedDateTime desc", "lastModifiedDateTime asc", "createdDateTime desc", "createdDateTime asc", "title asc", "title desc"]).optional(),
    limit: z.number().int().min(1).max(1000).optional().describe("Max pages to return (default 100)"),
  },
  RO,
  async ({ orderBy = "lastModifiedDateTime desc", limit = 100, ...target }) => {
    const id = await resolveSection(target);
    const pages = await getAll(
      `/sections/${id}/pages?$select=id,title,createdDateTime,lastModifiedDateTime,level,order&$orderby=${encodeURIComponent(orderBy)}&$top=${Math.min(limit, 100)}`,
      limit
    );
    return text(pages.map(pageSummary));
  }
);

tool(
  "search_pages",
  "Search page titles (case-insensitive) in one section, or across every section of every notebook when no section is given.",
  { query: z.string(), ...sectionTarget, limit: z.number().int().min(1).max(500).optional().describe("Max results (default 50)") },
  RO,
  async ({ query, limit = 50, ...target }) => {
    const filter = `$filter=contains(tolower(title),'${q(query.toLowerCase())}')`;
    const select = "$select=id,title,createdDateTime,lastModifiedDateTime&$expand=parentSection($select=id,displayName)";
    let sections;
    if (target.sectionId || target.section) {
      sections = [{ id: await resolveSection(target) }];
    } else {
      sections = (await allSections()).filter((s) => !isBlocked(s) && (!target.notebook || same(s.parentNotebook?.displayName, target.notebook)));
    }
    const results = [];
    for (const s of sections) {
      if (results.length >= limit) break;
      const pages = await getAll(`/sections/${s.id}/pages?${filter}&${select}&$top=100`, limit - results.length);
      results.push(...pages.filter((p) => !p.parentSection || !isBlocked(p.parentSection)));
    }
    return text(results.map(pageSummary));
  }
);

tool(
  "get_page",
  "Get a page's details: title, section, notebook, level/order, links and a short text preview (no full content).",
  { pageId: z.string() },
  RO,
  async ({ pageId }) => {
    await assertPageAllowed(pageId);
    const [p, preview] = await Promise.all([
      graph(
        `/pages/${pageId}?$select=id,title,createdDateTime,lastModifiedDateTime,level,order,createdByAppId,links` +
          "&$expand=parentSection($select=id,displayName),parentNotebook($select=id,displayName)"
      ),
      graph(`/pages/${pageId}/preview`).catch(() => null),
    ]);
    return text({ ...pageSummary(p), notebook: p.parentNotebook?.displayName, createdByApp: p.createdByAppId, preview: preview?.previewText });
  }
);

tool(
  "read_page",
  "Read a page's content as HTML. Set includeIds=true to get element ids (data-id/id) to target with update_page. " +
    "Tags show as data-tag (e.g. to-do, to-do:completed, important, question). Images/files show as resource URLs; save them with download_resource.",
  { pageId: z.string(), includeIds: z.boolean().optional() },
  RO,
  async ({ pageId, includeIds }) => {
    await assertPageAllowed(pageId);
    return text(await graph(`/pages/${pageId}/content${includeIds ? "?includeIDs=true" : ""}`, { raw: true }));
  }
);

tool(
  "create_page",
  "Create a page. bodyHtml supports p, h1-h6, ul/ol/li, table, pre, code, b, i, u, a, img (src can be a public URL or name:<part>), " +
    "and tags via data-tag, e.g. <p data-tag=\"to-do\">Task</p>, data-tag=\"to-do:completed\", \"important\", \"question\", \"idea\", \"critical\", \"contact\", \"remember-for-later\". " +
    "Local images/files go in attachments. Use data-id=\"...\" on elements you want to update later.",
  { title: z.string(), bodyHtml: z.string(), ...sectionTarget, attachments: attachmentSchema },
  WRITE,
  async ({ title, bodyHtml, attachments, ...target }) => {
    const id = await resolveSection(target);
    const html = `<!DOCTYPE html><html><head><title>${esc(title)}</title><meta name="created" content="${new Date().toISOString()}" /></head><body>${bodyHtml}</body></html>`;
    const req = attachments?.length ? await buildMultipart(html, attachments, "text/html") : { body: html, contentType: "application/xhtml+xml" };
    const page = await graph(`/sections/${id}/pages`, { method: "POST", ...req });
    return text({ created: page.title, id: page.id, webUrl: page.links?.oneNoteWebUrl?.href, clientUrl: page.links?.oneNoteClientUrl?.href });
  }
);

tool(
  "update_page",
  "Edit a page in place. Each change targets 'body', 'title', or an element by '#<data-id>' or its generated id (from read_page with includeIds=true). " +
    "Actions: append (end of target), prepend (start of target), insert (before/after target), replace (swap target; use target 'title' + replace " +
    "to rename the page). Example: tick a to-do with {target:'#task1', action:'replace', content:'<p data-tag=\"to-do:completed\" data-id=\"task1\">Done</p>'}. " +
    "All changes in one call succeed or fail together.",
  {
    pageId: z.string(),
    changes: z
      .array(
        z.object({
          target: z.string().describe("'body', 'title', or '#<element id>'"),
          action: z.enum(["append", "prepend", "insert", "replace"]),
          position: z.enum(["before", "after"]).optional().describe("For insert (default after) or append to a list/table"),
          content: z.string().describe("HTML (or plain text for a title replace)"),
        })
      )
      .min(1),
    attachments: attachmentSchema,
  },
  DESTRUCTIVE,
  async ({ pageId, changes, attachments }) => {
    await assertPageAllowed(pageId);
    const ops = changes.map((c) => ({ target: c.target, action: c.action, ...(c.position && { position: c.position }), content: c.content }));
    await patchPage(pageId, await resolveReplaceTargets(pageId, ops), attachments);
    return text(`Applied ${ops.length} change(s).`);
  }
);

tool(
  "append_to_page",
  "Append HTML content to the end of an existing page.",
  { pageId: z.string(), html: z.string(), attachments: attachmentSchema },
  WRITE,
  async ({ pageId, html, attachments }) => {
    await assertPageAllowed(pageId);
    await patchPage(pageId, [{ target: "body", action: "append", content: html }], attachments);
    return text("Appended.");
  }
);

tool(
  "copy_page",
  "Copy a page (with its images and files) into another section. To 'move' a page: copy, then delete the original. " +
    "Work/school accounts copy in the background (check with get_operation); personal accounts copy immediately and return the new page id.",
  {
    pageId: z.string(),
    toSectionId: z.string().optional(),
    toNotebook: z.string().optional(),
    toSection: z.string().optional(),
    toSectionGroup: z.string().optional(),
  },
  WRITE,
  async ({ pageId, toSectionId, toNotebook, toSection, toSectionGroup }) => {
    await assertPageAllowed(pageId);
    const dest = await resolveSection({ sectionId: toSectionId, notebook: toNotebook, section: toSection, sectionGroup: toSectionGroup });
    try {
      const r = await graph(`/pages/${pageId}/copyToSection`, {
        method: "POST",
        contentType: "application/json",
        body: JSON.stringify({ id: dest }),
        full: true,
      });
      return text({ started: true, operationId: r.body?.id, status: r.body?.status, statusUrl: r.location });
    } catch (e) {
      if (!notImplemented(e)) throw e;
      return text({ copied: true, method: "manual", ...(await copyPageManually(pageId, dest)) });
    }
  }
);

tool(
  "get_operation",
  "Check the status of a background copy (copy_page / copy_section). Returns status and, when finished, the new item's location.",
  { operationId: z.string() },
  RO,
  async ({ operationId }) => {
    const o = await graph(`/operations/${encodeURIComponent(operationId)}`);
    return text({ status: o.status, percentComplete: o.percentComplete, resourceId: o.resourceId, resourceLocation: o.resourceLocation, error: o.error });
  }
);

tool(
  "download_resource",
  "Save an image or attached file from a page to disk. Get resource ids/URLs from read_page (…/resources/<id>/$value or /content).",
  {
    pageId: z.string().describe("The page that contains the resource"),
    resource: z.string().describe("Resource id, or the resource URL from read_page"),
    outputPath: z.string().optional().describe("File path to write. Defaults to ONENOTE_DOWNLOAD_DIR/<resource id>"),
  },
  WRITE,
  async ({ pageId, resource, outputPath }) => {
    await assertPageAllowed(pageId);
    const resourceId = resource.match(/resources\/([^/?]+)/)?.[1] || resource;
    const content = await graph(`/pages/${pageId}/content`, { raw: true });
    if (!content.includes(resourceId)) throw new Error("That resource is not on this page.");
    const data = await graph(`/resources/${encodeURIComponent(resourceId)}/content`, { binary: true });
    const out = resolve((outputPath || resolve(DOWNLOAD_DIR, resourceId)).replace(/^~(?=$|\/)/, homedir()));
    await mkdir(dirname(out), { recursive: true });
    await writeFile(out, data);
    return text({ saved: out, bytes: data.length });
  }
);

if (ALLOW_DELETE) {
  tool("delete_page", "Permanently delete a page. Cannot be undone through the API.", { pageId: z.string() }, DESTRUCTIVE, async ({ pageId }) => {
    await assertPageAllowed(pageId);
    await graph(`/pages/${pageId}`, { method: "DELETE" });
    return text("Deleted.");
  });
}

tool(
  "server_settings",
  "Show this server's active settings (read-only mode, delete permission, defaults, blocked section count).",
  {},
  RO,
  async () =>
    text({
      readOnly: READ_ONLY,
      deleteEnabled: ALLOW_DELETE && !READ_ONLY,
      defaultNotebook: DEFAULT_NOTEBOOK || null,
      defaultSection: DEFAULT_SECTION || null,
      downloadDir: DOWNLOAD_DIR,
      blockedSections: BLOCKED_SECTIONS.length,
    })
);

await server.connect(new StdioServerTransport());

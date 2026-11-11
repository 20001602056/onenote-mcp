// OneNote lookups, page resources (images/files) and page create/update helpers.
import { readFile } from "node:fs/promises";
import { basename, extname } from "node:path";
import { randomUUID } from "node:crypto";
import { DEFAULT_NOTEBOOK, DEFAULT_SECTION, expandHome } from "./config.js";
import { esc, getAll, graph, q, same } from "./graph.js";
import { BLOCKED_MSG, assertSectionAllowed, checkSection, describeSection, isSectionBlocked, sectionGroups } from "./guardrail.js";

// ---- Lookups ----
// Every section in every notebook, including sections inside section groups.
export async function allSections() {
  return getAll(
    "/sections?$select=id,displayName,isDefault,createdDateTime,lastModifiedDateTime" +
      "&$expand=parentNotebook($select=id,displayName),parentSectionGroup($select=id,displayName)"
  );
}

// Sections the guardrail allows, each with .info = { id, name, notebook, path }.
export async function visibleSections(notebook) {
  const out = [];
  for (const s of await allSections()) {
    if (notebook && !same(s.parentNotebook?.displayName, notebook)) continue;
    const info = await describeSection(s);
    if (!isSectionBlocked(info)) out.push({ ...s, info });
  }
  return out;
}

export const allSectionGroups = async () => (await sectionGroups()).groups;

export async function findNotebook(name) {
  const nbs = await getAll(`/notebooks?$filter=displayName eq '${q(name)}'&$select=id,displayName`);
  if (!nbs.length) throw new Error(`Notebook "${name}" not found`);
  return nbs[0];
}

export async function findSectionGroup(notebookName, groupName) {
  const groups = (await allSectionGroups()).filter((g) => same(g.parentNotebook?.displayName, notebookName) && same(g.displayName, groupName));
  if (!groups.length) throw new Error(`Section group "${groupName}" not found in notebook "${notebookName}"`);
  return groups[0];
}

export async function findSection(notebookName, sectionName, sectionGroupName) {
  // Refuse blocked names before looking anything up, so a blocked section can't be probed for existence.
  if (isSectionBlocked({ name: sectionName, path: null })) throw new Error(BLOCKED_MSG);
  const inNotebook = await visibleSections(notebookName);
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

// Returns the section's guardrail info ({ id, name, notebook, path }). Pass { write: true } before changing it.
export async function resolveSection({ sectionId, notebook, section, sectionGroup }, opts = {}) {
  if (sectionId) return assertSectionAllowed(sectionId, opts);
  notebook ||= DEFAULT_NOTEBOOK;
  section ||= DEFAULT_SECTION;
  if (!notebook || !section) {
    throw new Error("Provide sectionId, or notebook + section (or set ONENOTE_DEFAULT_NOTEBOOK / ONENOTE_DEFAULT_SECTION)");
  }
  return checkSection((await findSection(notebook, section, sectionGroup)).info, opts);
}

export const pageSummary = (p) => ({
  title: p.title,
  id: p.id,
  created: p.createdDateTime,
  modified: p.lastModifiedDateTime,
  ...(p.level !== undefined && { level: p.level, order: p.order }),
  ...(p.parentSection && { section: p.parentSection.displayName, sectionId: p.parentSection.id }),
  ...(p.links && { webUrl: p.links.oneNoteWebUrl?.href, clientUrl: p.links.oneNoteClientUrl?.href }),
});

// ---- Attachments (images and files embedded in pages) ----
export const MIME = {
  ".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".gif": "image/gif", ".bmp": "image/bmp",
  ".tif": "image/tiff", ".tiff": "image/tiff", ".pdf": "application/pdf", ".txt": "text/plain", ".csv": "text/csv",
  ".json": "application/json", ".zip": "application/zip", ".doc": "application/msword",
  ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  ".xls": "application/vnd.ms-excel", ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  ".ppt": "application/vnd.ms-powerpoint", ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
};
export const extForType = (type) => Object.keys(MIME).find((ext) => MIME[ext] === type) || "";

// Hand-built multipart/form-data body (Graph rejects some FormData encodings).
export function multipart(parts) {
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
export async function buildMultipart(main, attachments, mainType) {
  let content = main;
  const extra = [];
  const parts = [];
  for (const [i, a] of attachments.entries()) {
    const part = a.partName || `file${i + 1}`;
    const file = expandHome(a.path);
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

export async function patchPage(pageId, ops, attachments) {
  const json = JSON.stringify(ops);
  const req = attachments?.length ? await buildMultipart(json, attachments, "application/json") : { body: json, contentType: "application/json" };
  await graph(`/pages/${pageId}/content`, { method: "PATCH", ...req });
}

// Graph can only "replace" some elements (e.g. <p>) through OneNote's generated id, not a custom data-id.
// Swap "#my-data-id" targets for the generated id when needed.
export async function resolveReplaceTargets(pageId, ops) {
  const wanted = ops.filter((o) => o.action === "replace" && o.target.startsWith("#"));
  if (!wanted.length) return ops;
  const html = await graph(`/pages/${pageId}/content?includeIDs=true`, { raw: true });
  const tags = [...html.matchAll(/<[a-z0-9]+\b[^>]*>/gi)].map((m) => m[0]);
  return ops.map((o) => {
    if (!wanted.includes(o)) return o;
    const key = o.target.slice(1);
    const tag = tags.find((t) => attr(t, "data-id") === key);
    const generated = tag && attr(tag, "id");
    return generated ? { ...o, target: generated } : o;
  });
}

// ---- Page resources ----
const attr = (tag, name) => tag.match(new RegExp(`\\s${name}="([^"]*)"`))?.[1];
// Resource URLs contain users('email'), so they are matched up to the closing double quote only.
const RESOURCE_URL = /https:\/\/graph\.microsoft\.com\/[^"]*?\/resources\/[^"/]+\/(?:\$value|content)/g;

// Full-resolution image copies point at Graph too; drop them and keep the rendered image.
export const stripFullres = (html) => html.replace(/\sdata-fullres-src(?:-type)?="[^"]*"/g, "");

// Images and files on a page: [{ url, id, type, name }] (name is the file name of attachments).
export function pageResources(html) {
  const info = new Map();
  for (const [tag] of html.matchAll(/<(?:img|object)\b[^>]*>/gi)) {
    const url = tag.match(/\s(?:src|data)="([^"]*\/resources\/[^"]*)"/)?.[1];
    if (url) info.set(url, { type: attr(tag, "data-src-type") || attr(tag, "type"), name: attr(tag, "data-attachment") });
  }
  const urls = [...new Set([...html.matchAll(RESOURCE_URL)].map((m) => m[0]))];
  return urls.map((url) => ({ url, id: url.match(/resources\/([^"/]+)\//)[1], ...info.get(url) }));
}

export const fetchResource = (resourceId) => graph(`/resources/${encodeURIComponent(resourceId)}/content`, { binary: true });

export function setTitle(html, title) {
  const tag = `<title>${esc(title)}</title>`;
  if (/<title>[\s\S]*?<\/title>/i.test(html)) return html.replace(/<title>[\s\S]*?<\/title>/i, tag);
  return html.replace(/<head\b[^>]*>/i, (m) => m + tag);
}

// Create a page from full page HTML. resources: [{ url, type, data }] whose URLs are swapped for uploaded parts.
export async function createPageFromHtml(sectionId, html, resources = []) {
  const parts = [];
  for (const [i, r] of resources.entries()) {
    const name = `res${i + 1}`;
    parts.push({ name, type: r.type || "application/octet-stream", data: r.data });
    html = html.split(r.url).join(`name:${name}`);
  }
  const req = multipart([{ name: "Presentation", type: "text/html", data: html }, ...parts]);
  const page = await graph(`/sections/${sectionId}/pages`, { method: "POST", ...req });
  return { id: page.id, title: page.title, resourcesCopied: parts.length, webUrl: page.links?.oneNoteWebUrl?.href };
}

// Copy a page by reading it, downloading its images/files and recreating it (personal accounts
// don't support Graph's copyToSection).
export async function copyPageManually(pageId, destSectionId) {
  const html = stripFullres(await graph(`/pages/${pageId}/content`, { raw: true }));
  const resources = [];
  for (const r of pageResources(html)) resources.push({ ...r, data: await fetchResource(r.id) });
  return createPageFromHtml(destSectionId, html, resources);
}

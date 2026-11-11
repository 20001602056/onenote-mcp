#!/usr/bin/env node
// OneNote MCP server backed by Microsoft Graph (signed in as you, see auth.js).
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { writeFile, mkdir, stat } from "node:fs/promises";
import { basename, dirname, join, resolve } from "node:path";
import { z } from "zod";
import {
  ALLOW_DELETE, DEFAULT_NOTEBOOK, DEFAULT_SECTION, DOWNLOAD_DIR, EXPORT_DIR, READ_ONLY, RETENTION_DAYS, VERSIONING, VERSION_DIR, AUDIT_LOG,
  expandHome,
} from "./lib/config.js";
import { esc, getAll, graph, notFound, notImplemented, q, same } from "./lib/graph.js";
import {
  NEED_PATHS, RULES, assertPageAllowed, audit, checkNewPageTitle, checkNewSection, checkStored, containerPath,
  forgetSectionGroups, isPageBlocked,
} from "./lib/guardrail.js";
import {
  allSectionGroups, buildMultipart, copyPageManually, createPageFromHtml, findNotebook, findSectionGroup, pageSummary, patchPage,
  resolveReplaceTargets, resolveSection, setTitle, visibleSections,
} from "./lib/onenote.js";
import { EXT, FORMATS, convert, findBrowser, localizeResources, renderCombined, renderPage, safeName, uniquePath, writeOutput } from "./lib/export.js";
import { listVersionedPages, listVersions, prune, readVersion, snapshotPage, snapshotSections, versionResources, versionSummary } from "./lib/versions.js";

const text = (v) => ({ content: [{ type: "text", text: typeof v === "string" ? v : JSON.stringify(v, null, 2) }] });

// ---- Server ----
const server = new McpServer({ name: "onenote", version: "2.1.0" });
const RO = { readOnlyHint: true, openWorldHint: true };
const WRITE = { readOnlyHint: false, destructiveHint: false, openWorldHint: true };
const DESTRUCTIVE = { readOnlyHint: false, destructiveHint: true, openWorldHint: true };

// Identifying arguments recorded in the audit log (never page content).
const AUDIT_KEYS = ["pageId", "sectionId", "notebook", "section", "sectionGroup", "name", "title", "toSectionId", "toNotebook", "toSection", "toSectionGroup", "renameAs", "versionId", "format"];

function tool(name, description, inputSchema, annotations, handler) {
  if (READ_ONLY && !annotations.readOnlyHint) return;
  server.registerTool(name, { description, inputSchema, annotations }, async (args, extra) => {
    const logged = Object.fromEntries(AUDIT_KEYS.filter((k) => typeof args?.[k] === "string").map((k) => [k, args[k]]));
    try {
      const result = await handler(args, extra);
      if (!annotations.readOnlyHint) await audit("call", { tool: name, ...logged });
      return result;
    } catch (e) {
      if (/guardrail/.test(e.message)) await audit("denied", { tool: name, ...logged, error: e.message });
      throw e;
    }
  });
}

const sectionTarget = {
  sectionId: z.string().optional().describe("Section id (takes precedence over names)"),
  notebook: z.string().optional().describe("Notebook name"),
  section: z.string().optional().describe("Section name"),
  sectionGroup: z.string().optional().describe("Section group name, to pick a section inside a group"),
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

const readFormat = z.enum(["html", "markdown", "text"]).optional().describe("html (default), markdown or text");

// Snapshot a page before the server changes it. Returns a warning string if the snapshot failed.
async function snapshotBefore(page, reason, { required = false } = {}) {
  if (!VERSIONING) return null;
  try {
    await snapshotPage(page, reason);
    return null;
  } catch (e) {
    if (required) throw new Error(`Refusing to continue: could not save a version first (${e.message.slice(0, 200)})`);
    return `Warning: could not save a version before this change (${e.message.slice(0, 200)})`;
  }
}

const withWarning = (result, warning) => (warning ? (typeof result === "string" ? `${result}\n${warning}` : { ...result, warning }) : result);

// --- Notebooks ---
tool(
  "list_notebooks",
  "List all notebooks with their full structure: section groups (nested) and sections, with ids and links.",
  {},
  RO,
  async () => {
    const [notebooks, groups, visible] = await Promise.all([
      getAll("/notebooks?$select=id,displayName,isDefault,isShared,userRole,createdDateTime,lastModifiedDateTime,links"),
      allSectionGroups(),
      visibleSections(),
    ]);
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
    forgetSectionGroups();
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
    const sections = await visibleSections(notebook);
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
    checkNewSection(section, null);
    const group = sectionGroup ? await findSectionGroup(notebook, sectionGroup) : null;
    const parentPath = group ? `/sectionGroups/${group.id}` : `/notebooks/${(await findNotebook(notebook)).id}`;
    checkNewSection(section, await containerPath(group?.parentNotebook?.displayName ?? notebook, group));
    const existing = (await getAll(`${parentPath}/sections?$select=id,displayName`)).find((x) => same(x.displayName, section));
    if (existing) return text({ section: existing.displayName, id: existing.id, created: false });
    const s = await graph(`${parentPath}/sections`, { method: "POST", contentType: "application/json", body: JSON.stringify({ displayName: section }) });
    return text({ section: s.displayName, id: s.id, created: true });
  }
);

tool(
  "copy_section",
  "Copy a section (with all its pages) into a notebook or section group. Work/school accounts copy in the background " +
    "(check with get_operation); personal accounts copy page by page immediately (slower for big sections). Pages blocked by the guardrail are left out.",
  {
    ...sectionTarget,
    toNotebook: z.string().describe("Destination notebook name"),
    toSectionGroup: z.string().optional().describe("Destination section group name (inside toNotebook)"),
    renameAs: z.string().optional().describe("Name for the copy"),
  },
  WRITE,
  async ({ toNotebook, toSectionGroup, renameAs, ...src }) => {
    const source = await resolveSection(src);
    const name = renameAs || source.name;
    const dest = toSectionGroup ? await findSectionGroup(toNotebook, toSectionGroup) : await findNotebook(toNotebook);
    checkNewSection(name, await containerPath(toSectionGroup ? dest.parentNotebook?.displayName : dest.displayName, toSectionGroup ? dest : null), { write: true });
    const pages = await getAll(`/sections/${source.id}/pages?$select=id,title&$orderby=createdDateTime asc&$top=100`);
    const allowed = pages.filter((p) => !isPageBlocked(p, source));
    // Graph's own copy takes every page, so only use it when no page in the section is blocked.
    if (allowed.length === pages.length) {
      const action = toSectionGroup ? "copyToSectionGroup" : "copyToNotebook";
      try {
        const r = await graph(`/sections/${source.id}/${action}`, {
          method: "POST",
          contentType: "application/json",
          body: JSON.stringify({ id: dest.id, ...(renameAs && { renameAs }) }),
          full: true,
        });
        return text({ started: true, operationId: r.body?.id, status: r.body?.status, statusUrl: r.location });
      } catch (e) {
        if (!notImplemented(e)) throw e;
      }
    }
    // Manual copy: new section in the destination, then copy each allowed page (oldest first).
    const parentPath = toSectionGroup ? `/sectionGroups/${dest.id}` : `/notebooks/${dest.id}`;
    if ((await getAll(`${parentPath}/sections?$select=id,displayName`)).some((s) => same(s.displayName, name))) {
      throw new Error(`A section named "${name}" already exists there. Pass renameAs to copy under a new name.`);
    }
    const target = await graph(`${parentPath}/sections`, { method: "POST", contentType: "application/json", body: JSON.stringify({ displayName: name }) });
    const copied = [];
    const failedPages = [];
    for (const p of allowed) {
      try {
        copied.push((await copyPageManually(p.id, target.id)).title);
      } catch (err) {
        failedPages.push({ title: p.title, error: err.message.slice(0, 200) });
      }
    }
    return text({
      copied: true,
      method: "manual",
      section: name,
      sectionId: target.id,
      pagesCopied: copied.length,
      failedPages,
      ...(allowed.length < pages.length && { skippedByGuardrail: pages.length - allowed.length }),
    });
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
    const section = await resolveSection(target);
    const pages = await getAll(
      `/sections/${section.id}/pages?$select=id,title,createdDateTime,lastModifiedDateTime,level,order&$orderby=${encodeURIComponent(orderBy)}&$top=${Math.min(limit, 100)}`,
      limit
    );
    return text(pages.filter((p) => !isPageBlocked(p, section)).map(pageSummary));
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
    const sections = target.sectionId || target.section ? [{ info: await resolveSection(target) }] : await visibleSections(target.notebook);
    const results = [];
    for (const s of sections) {
      if (results.length >= limit) break;
      const pages = await getAll(`/sections/${s.info.id}/pages?${filter}&${select}&$top=100`, limit - results.length);
      results.push(...pages.filter((p) => !isPageBlocked(p, s.info)));
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
    const versions = VERSIONING ? (await listVersions(pageId)).length : undefined;
    return text({ ...pageSummary(p), notebook: p.parentNotebook?.displayName, createdByApp: p.createdByAppId, preview: preview?.previewText, versions });
  }
);

tool(
  "read_page",
  "Read a page's content. format=html (default) returns OneNote HTML; set includeIds=true to get element ids (data-id/id) to target with update_page. " +
    "Tags show as data-tag (e.g. to-do, to-do:completed, important, question). Images/files show as resource URLs; save them with download_resource. " +
    "format=markdown or text returns a lighter, readable version.",
  { pageId: z.string(), includeIds: z.boolean().optional(), format: readFormat },
  RO,
  async ({ pageId, includeIds, format = "html" }) => {
    await assertPageAllowed(pageId);
    const html = await graph(`/pages/${pageId}/content${includeIds && format === "html" ? "?includeIDs=true" : ""}`, { raw: true });
    return text(convert(html, format));
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
    const section = await resolveSection(target, { write: true });
    checkNewPageTitle(title, section);
    const html = `<!DOCTYPE html><html><head><title>${esc(title)}</title><meta name="created" content="${new Date().toISOString()}" /></head><body>${bodyHtml}</body></html>`;
    const req = attachments?.length ? await buildMultipart(html, attachments, "text/html") : { body: html, contentType: "application/xhtml+xml" };
    const page = await graph(`/sections/${section.id}/pages`, { method: "POST", ...req });
    return text({ created: page.title, id: page.id, webUrl: page.links?.oneNoteWebUrl?.href, clientUrl: page.links?.oneNoteClientUrl?.href });
  }
);

tool(
  "update_page",
  "Edit a page in place. Each change targets 'body', 'title', or an element by '#<data-id>' or its generated id (from read_page with includeIds=true). " +
    "Actions: append (end of target), prepend (start of target), insert (before/after target), replace (swap target; use target 'title' + replace " +
    "to rename the page). Example: tick a to-do with {target:'#task1', action:'replace', content:'<p data-tag=\"to-do:completed\" data-id=\"task1\">Done</p>'}. " +
    "All changes in one call succeed or fail together. A version of the page is saved first (see list_page_versions).",
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
    const page = await assertPageAllowed(pageId, { write: true });
    for (const c of changes) if (c.target === "title" && c.action === "replace") checkNewPageTitle(c.content, page.section);
    const warning = await snapshotBefore(page, "before update_page");
    const ops = changes.map((c) => ({ target: c.target, action: c.action, ...(c.position && { position: c.position }), content: c.content }));
    await patchPage(pageId, await resolveReplaceTargets(pageId, ops), attachments);
    return text(withWarning(`Applied ${ops.length} change(s).`, warning));
  }
);

tool(
  "append_to_page",
  "Append HTML content to the end of an existing page. A version of the page is saved first.",
  { pageId: z.string(), html: z.string(), attachments: attachmentSchema },
  WRITE,
  async ({ pageId, html, attachments }) => {
    const page = await assertPageAllowed(pageId, { write: true });
    const warning = await snapshotBefore(page, "before append_to_page");
    await patchPage(pageId, [{ target: "body", action: "append", content: html }], attachments);
    return text(withWarning("Appended.", warning));
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
    const page = await assertPageAllowed(pageId);
    const dest = await resolveSection({ sectionId: toSectionId, notebook: toNotebook, section: toSection, sectionGroup: toSectionGroup }, { write: true });
    checkNewPageTitle(page.title, dest);
    try {
      const r = await graph(`/pages/${pageId}/copyToSection`, {
        method: "POST",
        contentType: "application/json",
        body: JSON.stringify({ id: dest.id }),
        full: true,
      });
      return text({ started: true, operationId: r.body?.id, status: r.body?.status, statusUrl: r.location });
    } catch (e) {
      if (!notImplemented(e)) throw e;
      return text({ copied: true, method: "manual", ...(await copyPageManually(pageId, dest.id)) });
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
    const out = expandHome(outputPath || resolve(DOWNLOAD_DIR, resourceId));
    await mkdir(dirname(out), { recursive: true });
    await writeFile(out, data);
    return text({ saved: out, bytes: data.length });
  }
);

if (ALLOW_DELETE) {
  tool(
    "delete_page",
    "Permanently delete a page from OneNote. A version is saved first (when versioning is on), so it can be brought back with restore_page_version " +
      "within the retention period.",
    { pageId: z.string() },
    DESTRUCTIVE,
    async ({ pageId }) => {
      const page = await assertPageAllowed(pageId, { write: true });
      await snapshotBefore(page, "before delete_page", { required: true });
      await graph(`/pages/${pageId}`, { method: "DELETE" });
      return text(VERSIONING ? "Deleted. A version was saved; restore it with restore_page_version." : "Deleted.");
    }
  );
}

// --- Export ---
const formatSchema = z.enum(FORMATS).describe("markdown, html, text, json or pdf (pdf needs Chrome, Edge, Chromium or Brave installed)");
const resourcesSchema = z
  .enum(["download", "inline", "none"])
  .optional()
  .describe("Images/files: download (default; saved in a folder next to the export), inline (embedded as data: URIs), or none. text ignores this; pdf always embeds.");

const resourceMode = (format, resources = "download") => (format === "pdf" ? "inline" : format === "text" ? "none" : resources);

// Export metadata for a page the guardrail already allowed.
const exportMeta = (page) => ({
  id: page.id,
  title: page.title,
  created: page.created,
  modified: page.modified,
  notebook: page.section.notebook,
  section: page.section.name,
  path: page.section.path ? `${page.section.path}/${page.title}` : undefined,
  webUrl: page.webUrl,
});

tool(
  "export_page",
  "Export a page to a local file as Markdown, HTML, plain text, JSON or PDF, with its images and attached files.",
  {
    pageId: z.string(),
    format: formatSchema,
    resources: resourcesSchema,
    outputPath: z.string().optional().describe("File (or existing folder) to write. Defaults to ONENOTE_EXPORT_DIR/<page title>.<ext>, never overwriting"),
  },
  WRITE,
  async ({ pageId, format, resources, outputPath }) => {
    const page = await assertPageAllowed(pageId);
    const ext = EXT[format];
    let out;
    if (outputPath) {
      out = expandHome(outputPath);
      const isDir = /[\\/]$/.test(outputPath) || (await stat(out).then((s) => s.isDirectory(), () => false));
      if (isDir) out = await uniquePath(join(out, safeName(page.title) + ext));
    } else {
      out = await uniquePath(join(EXPORT_DIR, safeName(page.title) + ext));
    }
    const assetsDir = out.slice(0, out.length - ext.length) + "_files";
    const mode = resourceMode(format, resources);
    const raw = await graph(`/pages/${pageId}/content`, { raw: true });
    const { html, count } = await localizeResources(raw, mode, { assetsDir, assetsHref: basename(assetsDir) });
    const bytes = await writeOutput(out, renderPage(exportMeta(page), html, format), format);
    return text({ saved: out, format, bytes, resources: count, ...(mode === "download" && count && { resourcesDir: assetsDir }) });
  }
);

tool(
  "export_section",
  "Export every page in a section to local files as Markdown, HTML, plain text, JSON or PDF. By default each page gets its own file in a folder " +
    "named after the section (with an index for markdown/html); combine=true writes one file with all pages. Pages blocked by the guardrail are left out.",
  {
    ...sectionTarget,
    format: formatSchema,
    resources: resourcesSchema,
    combine: z.boolean().optional().describe("One file for the whole section instead of one per page"),
    outputDir: z.string().optional().describe("Folder to create the export in. Defaults to ONENOTE_EXPORT_DIR"),
  },
  WRITE,
  async ({ format, resources, combine = false, outputDir, ...target }) => {
    const section = await resolveSection(target);
    const listed = await getAll(
      `/sections/${section.id}/pages?$select=id,title,createdDateTime,lastModifiedDateTime,order,links&$orderby=createdDateTime asc&$top=100`
    );
    const pages = listed.filter((p) => !isPageBlocked(p, section)).sort((a, b) => (a.order ?? 0) - (b.order ?? 0));
    const base = outputDir ? expandHome(outputDir) : EXPORT_DIR;
    const ext = EXT[format];
    const mode = resourceMode(format, resources);
    const meta = (p) =>
      exportMeta({ id: p.id, title: p.title, created: p.createdDateTime, modified: p.lastModifiedDateTime, webUrl: p.links?.oneNoteWebUrl?.href, section });
    const failedPages = [];
    let resourceCount = 0;

    if (combine) {
      const out = await uniquePath(join(base, safeName(section.name) + ext));
      const assetsDir = out.slice(0, out.length - ext.length) + "_files";
      const rendered = [];
      for (const p of pages) {
        try {
          const raw = await graph(`/pages/${p.id}/content`, { raw: true });
          const { html, count } = await localizeResources(raw, mode, { assetsDir, assetsHref: basename(assetsDir) });
          resourceCount += count;
          rendered.push({ page: meta(p), html });
        } catch (e) {
          failedPages.push({ title: p.title, error: e.message.slice(0, 200) });
        }
      }
      const bytes = await writeOutput(out, renderCombined(section, rendered, format), format);
      return text({ saved: out, format, pages: rendered.length, bytes, resources: resourceCount, failedPages });
    }

    const dir = await uniquePath(join(base, safeName(section.name)));
    const used = new Set();
    const files = [];
    for (const p of pages) {
      try {
        const out = await uniquePath(join(dir, safeName(p.title) + ext), used);
        const raw = await graph(`/pages/${p.id}/content`, { raw: true });
        const { html, count } = await localizeResources(raw, mode, { assetsDir: join(dir, "assets"), assetsHref: "assets" });
        resourceCount += count;
        await writeOutput(out, renderPage(meta(p), html, format), format);
        files.push({ title: p.title || "Untitled", file: basename(out) });
      } catch (e) {
        failedPages.push({ title: p.title, error: e.message.slice(0, 200) });
      }
    }
    if (format === "markdown" || format === "html") {
      const link = (f) => encodeURI(f.file);
      const index =
        format === "markdown"
          ? `# ${section.name}\n\n${files.map((f) => `- [${f.title}](${link(f)})`).join("\n")}\n`
          : `<!DOCTYPE html>\n<html lang="en"><head><meta charset="utf-8"><title>${esc(section.name)}</title></head><body>` +
            `<h1>${esc(section.name)}</h1><ol>${files.map((f) => `<li><a href="${link(f)}">${esc(f.title)}</a></li>`).join("")}</ol></body></html>\n`;
      await writeOutput(join(dir, `index${ext}`), index, format);
    }
    return text({ saved: dir, format, pages: files.length, resources: resourceCount, failedPages });
  }
);

// --- Versions ---
// Guardrail for a page's stored versions: check the live page, or the stored location if it was deleted.
async function assertVersionsAllowed(pageId) {
  try {
    return await assertPageAllowed(pageId);
  } catch (e) {
    if (!notFound(e)) throw e;
  }
  const [latest] = await listVersions(pageId);
  if (!latest) throw new Error("No versions saved for this page.");
  checkStored({ id: pageId, title: latest.title }, { id: latest.sectionId, name: latest.section, path: latest.sectionPath ?? null });
  return null;
}

if (VERSIONING) {
  tool(
    "snapshot_pages",
    `Save a version of every page that changed since its last version, in one section, one notebook, or every section (no arguments). ` +
      `Versions are kept for ${RETENTION_DAYS} days. Pages edited through this server are also saved automatically before each change.`,
    { ...sectionTarget, force: z.boolean().optional().describe("Re-check every page's content, not just pages whose modified time changed") },
    WRITE,
    async ({ force, ...target }) => {
      const sections =
        target.sectionId || target.section ? [{ info: await resolveSection(target) }] : await visibleSections(target.notebook);
      const complete = !target.sectionId && !target.section && !target.notebook;
      return text(await snapshotSections(sections, { force, complete, reason: "manual" }));
    }
  );

  tool(
    "list_page_versions",
    `List saved versions of a page (newest first). Without pageId, list every page that has versions, including deleted pages. ` +
      `Versions are kept for ${RETENTION_DAYS} days.`,
    { pageId: z.string().optional() },
    RO,
    async ({ pageId }) => {
      if (pageId) {
        await assertVersionsAllowed(pageId);
        return text((await listVersions(pageId)).map(versionSummary));
      }
      const pages = (await listVersionedPages()).filter((p) => {
        try {
          checkStored({ id: p.pageId, title: p.title }, { id: p.sectionId, name: p.section, path: p.sectionPath ?? null });
          return true;
        } catch {
          return false;
        }
      });
      return text(pages);
    }
  );

  tool(
    "read_page_version",
    "Read a saved version of a page as html (default), markdown or text.",
    { pageId: z.string(), versionId: z.string(), format: readFormat },
    RO,
    async ({ pageId, versionId, format = "html" }) => {
      await assertVersionsAllowed(pageId);
      const { meta, html } = await readVersion(pageId, versionId);
      return text(`<!-- version ${meta.versionId} of "${meta.title}", captured ${meta.capturedAt} (${meta.reason}) -->\n${convert(html, format)}`);
    }
  );

  tool(
    "restore_page_version",
    "Restore a saved version as a new page (with its images and files), in its original section or another one. " +
      "The current page is left as it is; delete it afterwards if you want the restored copy to replace it.",
    {
      pageId: z.string(),
      versionId: z.string(),
      title: z.string().optional().describe('Title for the restored page. Defaults to "<original title> (restored)"'),
      toSectionId: z.string().optional(),
      toNotebook: z.string().optional(),
      toSection: z.string().optional(),
      toSectionGroup: z.string().optional(),
    },
    WRITE,
    async ({ pageId, versionId, title, toSectionId, toNotebook, toSection, toSectionGroup }) => {
      await assertVersionsAllowed(pageId);
      const { meta, html } = await readVersion(pageId, versionId);
      const dest = toSectionId || toSection
        ? await resolveSection({ sectionId: toSectionId, notebook: toNotebook, section: toSection, sectionGroup: toSectionGroup }, { write: true })
        : await resolveSection({ sectionId: meta.sectionId }, { write: true });
      const newTitle = title || `${meta.title || "Untitled"} (restored)`;
      checkNewPageTitle(newTitle, dest);
      const page = await createPageFromHtml(dest.id, setTitle(html, newTitle), await versionResources(pageId, meta));
      return text({ restored: true, from: meta.versionId, ...page, section: dest.name });
    }
  );
}

tool(
  "server_settings",
  "Show this server's active settings (read-only mode, delete permission, defaults, guardrail rule counts, versioning, export).",
  {},
  RO,
  async () =>
    text({
      readOnly: READ_ONLY,
      deleteEnabled: ALLOW_DELETE && !READ_ONLY,
      defaultNotebook: DEFAULT_NOTEBOOK || null,
      defaultSection: DEFAULT_SECTION || null,
      downloadDir: DOWNLOAD_DIR,
      exportDir: EXPORT_DIR,
      pdfBrowser: findBrowser(),
      guardrail: {
        blockedSections: RULES.blockedSections.length,
        blockedPages: RULES.blockedPages.length,
        protectedSections: RULES.protectedSections.length,
        protectedPages: RULES.protectedPages.length,
        pathRules: NEED_PATHS,
        auditLog: AUDIT_LOG || null,
      },
      versioning: VERSIONING ? { retentionDays: RETENTION_DAYS, versionDir: VERSION_DIR } : false,
    })
);

// Age out old versions in the background; never blocks startup.
if (VERSIONING) prune().catch(() => {});

await server.connect(new StdioServerTransport());

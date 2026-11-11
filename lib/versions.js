// Page versioning: local snapshots of page HTML plus its images/files, kept for RETENTION_DAYS.
// Graph has no version history for OneNote pages, so this server keeps its own:
//   - before every update/append/delete made through the server, and
//   - for every page in every section on demand (snapshot_pages tool, `npm run snapshot`, schedule it).
// Layout: VERSION_DIR/<page id>/<version id>.html + .json, and resources/<resource id> shared by its versions.
import { createHash } from "node:crypto";
import { mkdir, readdir, readFile, rm, stat, writeFile } from "node:fs/promises";
import { join } from "node:path";
import { RETENTION_DAYS, VERSION_DIR } from "./config.js";
import { getAll, graph } from "./graph.js";
import { isPageBlocked } from "./guardrail.js";
import { fetchResource, pageResources, stripFullres } from "./onenote.js";

const DAY = 86_400_000;
const safe = (s) => s.replace(/[^A-Za-z0-9!._-]/g, "_");
const dirOf = (pageId) => join(VERSION_DIR, safe(pageId));
const exists = (p) => stat(p).then(() => true, () => false);
const writePrivate = (file, data) => writeFile(file, data, { mode: 0o600 });
const VERSION_ID = /^\d{4}-\d{2}-\d{2}T\d{2}-\d{2}-\d{2}-\d{3}Z$/;

// Newest first.
export async function listVersions(pageId) {
  const dir = dirOf(pageId);
  let files;
  try {
    files = await readdir(dir);
  } catch {
    return [];
  }
  const metas = [];
  for (const f of files.filter((f) => f.endsWith(".json"))) {
    try {
      metas.push(JSON.parse(await readFile(join(dir, f), "utf8")));
    } catch {
      // A half-written snapshot; ignore it.
    }
  }
  return metas.sort((a, b) => b.versionId.localeCompare(a.versionId));
}

export const versionSummary = (m) => ({
  versionId: m.versionId,
  captured: m.capturedAt,
  reason: m.reason,
  title: m.title,
  pageModified: m.pageModified,
  bytes: m.bytes,
  resources: m.resources.length,
});

// Every page with stored versions, including pages since deleted from OneNote.
export async function listVersionedPages() {
  let dirs;
  try {
    dirs = await readdir(VERSION_DIR, { withFileTypes: true });
  } catch {
    return [];
  }
  const pages = [];
  for (const d of dirs.filter((d) => d.isDirectory())) {
    const versions = await listVersions(d.name);
    if (!versions.length) continue;
    const latest = versions[0];
    pages.push({
      pageId: latest.pageId,
      title: latest.title,
      notebook: latest.notebook,
      section: latest.section,
      sectionId: latest.sectionId,
      sectionPath: latest.sectionPath,
      versions: versions.length,
      latest: latest.capturedAt,
      oldest: versions.at(-1).capturedAt,
    });
  }
  return pages.sort((a, b) => b.latest.localeCompare(a.latest));
}

// page: { id, title, created, modified, section: { id, name, notebook, path } }
export async function snapshotPage(page, reason) {
  const html = stripFullres(await graph(`/pages/${page.id}/content`, { raw: true }));
  const sha256 = createHash("sha256").update(html).digest("hex");
  const [latest] = await listVersions(page.id);
  if (latest?.sha256 === sha256) {
    // Same content: remember the newer modified time so scheduled runs stop re-downloading the page.
    if (page.modified && latest.pageModified !== page.modified) {
      await writePrivate(join(dirOf(page.id), `${latest.versionId}.json`), JSON.stringify({ ...latest, pageModified: page.modified }, null, 2));
    }
    return { saved: false, unchanged: true, versionId: latest.versionId };
  }

  const dir = dirOf(page.id);
  await mkdir(join(dir, "resources"), { recursive: true, mode: 0o700 });
  const resources = [];
  for (const r of pageResources(html)) {
    const file = safe(r.id);
    const path = join(dir, "resources", file);
    if (!(await exists(path))) await writePrivate(path, await fetchResource(r.id));
    resources.push({ id: r.id, url: r.url, type: r.type, name: r.name, file });
  }

  let now = new Date();
  let versionId = now.toISOString().replace(/[:.]/g, "-");
  while (await exists(join(dir, `${versionId}.json`))) {
    now = new Date(now.getTime() + 1);
    versionId = now.toISOString().replace(/[:.]/g, "-");
  }
  const meta = {
    versionId,
    pageId: page.id,
    title: page.title,
    created: page.created,
    pageModified: page.modified,
    capturedAt: now.toISOString(),
    reason,
    notebook: page.section.notebook,
    section: page.section.name,
    sectionId: page.section.id,
    sectionPath: page.section.path,
    sha256,
    bytes: Buffer.byteLength(html),
    resources,
  };
  await writePrivate(join(dir, `${versionId}.html`), html);
  await writePrivate(join(dir, `${versionId}.json`), JSON.stringify(meta, null, 2));
  return { saved: true, ...versionSummary(meta) };
}

export async function readVersion(pageId, versionId) {
  if (!VERSION_ID.test(versionId)) throw new Error(`Invalid versionId "${versionId}"`);
  const dir = dirOf(pageId);
  try {
    const meta = JSON.parse(await readFile(join(dir, `${versionId}.json`), "utf8"));
    return { meta, html: await readFile(join(dir, `${versionId}.html`), "utf8") };
  } catch {
    throw new Error(`Version ${versionId} not found for this page. List them with list_page_versions.`);
  }
}

// The version's images/files as [{ url, type, data }] for createPageFromHtml / export.
export async function versionResources(pageId, meta) {
  const dir = dirOf(pageId);
  return Promise.all(meta.resources.map(async (r) => ({ ...r, data: await readFile(join(dir, "resources", r.file)) })));
}

// Delete versions older than RETENTION_DAYS. The newest version of each page is kept as its baseline
// (it matches the page as it is now), unless `alive` is given and the page is no longer in it.
export async function prune({ alive } = {}) {
  const cutoff = new Date(Date.now() - RETENTION_DAYS * DAY).toISOString();
  let dirs;
  try {
    dirs = await readdir(VERSION_DIR, { withFileTypes: true });
  } catch {
    return 0;
  }
  let removed = 0;
  for (const d of dirs.filter((d) => d.isDirectory())) {
    const dir = join(VERSION_DIR, d.name);
    const versions = await listVersions(d.name);
    const keepBaseline = versions.length && (!alive || alive.has(versions[0].pageId));
    const keep = versions.filter((v, i) => v.capturedAt >= cutoff || (i === 0 && keepBaseline));
    for (const v of versions.filter((v) => !keep.includes(v))) {
      await rm(join(dir, `${v.versionId}.json`), { force: true });
      await rm(join(dir, `${v.versionId}.html`), { force: true });
      removed++;
    }
    if (!keep.length) {
      await rm(dir, { recursive: true, force: true });
      continue;
    }
    const used = new Set(keep.flatMap((v) => v.resources.map((r) => r.file)));
    for (const f of await readdir(join(dir, "resources")).catch(() => [])) {
      if (!used.has(f)) await rm(join(dir, "resources", f), { force: true });
    }
  }
  return removed;
}

// Snapshot every changed page in the given sections (from visibleSections). Pages whose
// lastModifiedDateTime matches their newest version are skipped without downloading them.
// complete=true means these are all sections, so versions of pages that no longer exist age out normally.
export async function snapshotSections(sections, { force = false, complete = false, reason = "scheduled" } = {}) {
  const result = { sections: sections.length, pagesChecked: 0, saved: 0, unchanged: 0, failed: [] };
  const alive = new Set();
  let listedAll = true;
  for (const s of sections) {
    let pages;
    try {
      pages = await getAll(`/sections/${s.info.id}/pages?$select=id,title,createdDateTime,lastModifiedDateTime&$top=100`);
    } catch (e) {
      listedAll = false;
      result.failed.push({ section: s.info.name, error: e.message.slice(0, 200) });
      continue;
    }
    for (const p of pages) {
      if (isPageBlocked(p, s.info)) continue;
      alive.add(p.id);
      result.pagesChecked++;
      const [latest] = await listVersions(p.id);
      if (!force && latest && latest.pageModified === p.lastModifiedDateTime) {
        result.unchanged++;
        continue;
      }
      try {
        const page = { id: p.id, title: p.title, created: p.createdDateTime, modified: p.lastModifiedDateTime, section: s.info };
        (await snapshotPage(page, reason)).saved ? result.saved++ : result.unchanged++;
      } catch (e) {
        result.failed.push({ page: p.title, pageId: p.id, error: e.message.slice(0, 200) });
      }
    }
  }
  result.pruned = await prune(complete && listedAll ? { alive } : {});
  result.retentionDays = RETENTION_DAYS;
  result.versionDir = VERSION_DIR;
  return result;
}

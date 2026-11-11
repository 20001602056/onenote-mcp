// Guardrail: sections and pages that tools must never touch (blocked) or never change (protected).
// Rules are comma-separated environment variables. Each entry is one of:
//   - an id: a section id (section rules) or page id (page rules)
//   - a name: a section name or page title, case-insensitive, where * matches anything
//   - a path, when it contains "/": "Notebook/Section group/Section" for sections, or ".../Section/Page title"
//     for pages. * matches inside one level and ** matches any number of levels, so "My Notebook/Private/**"
//     covers everything inside the Private section group.
import { appendFile } from "node:fs/promises";
import { AUDIT_LOG, list } from "./config.js";
import { getAll, graph } from "./graph.js";

function compile(raw) {
  const p = raw.toLowerCase();
  const isPath = p.includes("/");
  const body = p
    .split(/(\*\*|\*)/)
    .map((t) => (t === "**" ? ".*" : t === "*" ? (isPath ? "[^/]*" : ".*") : t.replace(/[.+?^${}()|[\]\\]/g, "\\$&")))
    .join("");
  return { isPath, re: new RegExp(`^${body}$`) };
}

const rules = (name) => list(name).map(compile);
export const RULES = {
  blockedSections: rules("ONENOTE_BLOCKED_SECTIONS"),
  blockedPages: rules("ONENOTE_BLOCKED_PAGES"),
  protectedSections: rules("ONENOTE_PROTECTED_SECTIONS"),
  protectedPages: rules("ONENOTE_PROTECTED_PAGES"),
};
// Path rules need each section's full path (notebook + section groups), which costs extra lookups.
export const NEED_PATHS = Object.values(RULES).some((rs) => rs.some((r) => r.isPath));

export const BLOCKED_MSG = "Access denied by guardrail: this section is blocked.";
export const BLOCKED_PAGE_MSG = "Access denied by guardrail: this page is blocked.";
export const PROTECTED_MSG = "Write denied by guardrail: this section is read-only.";
export const PROTECTED_PAGE_MSG = "Write denied by guardrail: this page is read-only.";

// item: { id, name, path }. A path rule never matches an item whose path is unknown, so callers that
// can't work out a path while NEED_PATHS is set must refuse instead (see checkStored).
const matches = (ruleList, item) =>
  ruleList.some((r) =>
    r.isPath
      ? item.path != null && r.re.test(item.path.toLowerCase())
      : r.re.test((item.id || "").toLowerCase()) || r.re.test((item.name || "").toLowerCase())
  );

// ---- Section paths ----
let groupCache = null;
export async function sectionGroups() {
  if (!groupCache || Date.now() - groupCache.at > 60_000) {
    const groups = await getAll(
      "/sectionGroups?$select=id,displayName&$expand=parentNotebook($select=id,displayName),parentSectionGroup($select=id,displayName)"
    );
    groupCache = { at: Date.now(), groups, byId: new Map(groups.map((g) => [g.id, g])) };
  }
  return groupCache;
}
export const forgetSectionGroups = () => {
  groupCache = null;
};

// "Notebook/Group/Subgroup" for a notebook name and an (optional) section group.
export async function containerPath(notebookName, group) {
  const names = [];
  let refreshed = false;
  while (group && names.length < 50) {
    let full = (await sectionGroups()).byId.get(group.id);
    if (!full && !refreshed) {
      forgetSectionGroups();
      refreshed = true;
      full = (await sectionGroups()).byId.get(group.id);
    }
    names.unshift(full?.displayName ?? group.displayName);
    group = full?.parentSectionGroup;
  }
  return [notebookName ?? "", ...names].join("/");
}

// Turn a Graph section (with parentNotebook / parentSectionGroup expanded) into { id, name, notebook, path }.
export async function describeSection(s) {
  const path = s.parentNotebook ? `${await containerPath(s.parentNotebook.displayName, s.parentSectionGroup)}/${s.displayName}` : null;
  return { id: s.id, name: s.displayName, notebook: s.parentNotebook?.displayName, path };
}

// ---- Checks ----
const pageItem = (p, section) => ({ id: p.id, name: p.title, path: section.path != null ? `${section.path}/${p.title ?? ""}` : null });

export const isSectionBlocked = (section) => matches(RULES.blockedSections, section);
export const isPageBlocked = (p, section) => isSectionBlocked(section) || matches(RULES.blockedPages, pageItem(p, section));

export function checkSection(section, { write = false } = {}) {
  if (isSectionBlocked(section)) throw new Error(BLOCKED_MSG);
  if (write && matches(RULES.protectedSections, section)) throw new Error(PROTECTED_MSG);
  return section;
}

export function checkPage(p, section, { write = false } = {}) {
  checkSection(section, { write });
  if (matches(RULES.blockedPages, pageItem(p, section))) throw new Error(BLOCKED_PAGE_MSG);
  if (write && matches(RULES.protectedPages, pageItem(p, section))) throw new Error(PROTECTED_PAGE_MSG);
}

// A section about to be created (or copied) under parentPath ("Notebook" or "Notebook/Group", null if unknown).
export function checkNewSection(name, parentPath, { write = false } = {}) {
  checkSection({ id: "", name, path: parentPath != null ? `${parentPath}/${name}` : null }, { write });
}

// A page about to be created in, renamed in, copied or restored into a section.
export const checkNewPageTitle = (title, section) => checkPage({ id: "", title }, section, { write: true });

// Pages and sections known only from a stored snapshot (the page may be gone).
export function checkStored(p, section) {
  if (NEED_PATHS && section.path == null) throw new Error(BLOCKED_PAGE_MSG);
  checkPage(p, section);
}

const SECTION_SELECT = "$select=id,displayName&$expand=parentNotebook($select=id,displayName),parentSectionGroup($select=id,displayName)";

export async function assertSectionAllowed(sectionId, opts) {
  const s = await graph(`/sections/${sectionId}?${SECTION_SELECT}`);
  return checkSection(await describeSection(s), opts);
}

// Returns { id, title, created, modified, webUrl, section } for a page the guardrail allows.
export async function assertPageAllowed(pageId, opts = {}) {
  const p = await graph(
    `/pages/${pageId}?$select=id,title,createdDateTime,lastModifiedDateTime,links` +
      "&$expand=parentSection($select=id,displayName),parentNotebook($select=id,displayName)"
  );
  if (!p.parentSection) throw new Error(BLOCKED_PAGE_MSG);
  const section = NEED_PATHS
    ? await describeSection(await graph(`/sections/${p.parentSection.id}?${SECTION_SELECT}`))
    : { id: p.parentSection.id, name: p.parentSection.displayName, notebook: p.parentNotebook?.displayName, path: null };
  checkPage(p, section, opts);
  return { id: p.id, title: p.title, created: p.createdDateTime, modified: p.lastModifiedDateTime, webUrl: p.links?.oneNoteWebUrl?.href, section };
}

// ---- Audit log (optional, ONENOTE_AUDIT_LOG) ----
export async function audit(event, details) {
  if (!AUDIT_LOG) return;
  const line = JSON.stringify({ at: new Date().toISOString(), event, ...details });
  await appendFile(AUDIT_LOG, `${line}\n`, { mode: 0o600 }).catch(() => {});
}

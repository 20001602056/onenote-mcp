#!/usr/bin/env node
// Save a version of every changed page (all sections, or --notebook / --section), then prune old versions.
// Meant to run on a schedule (cron, launchd, Task Scheduler); see README "Versioning".
//   node snapshot.js [--notebook "My Notebook"] [--section "Quick Notes"] [--force]
import { parseArgs } from "node:util";
import { VERSIONING } from "./lib/config.js";
import { resolveSection, visibleSections } from "./lib/onenote.js";
import { snapshotSections } from "./lib/versions.js";

const { values } = parseArgs({
  options: { notebook: { type: "string" }, section: { type: "string" }, force: { type: "boolean", default: false } },
});

if (!VERSIONING) {
  console.error("Versioning is turned off (ONENOTE_VERSIONING=false).");
  process.exit(1);
}

try {
  const sections = values.section
    ? [{ info: await resolveSection({ notebook: values.notebook, section: values.section }) }]
    : await visibleSections(values.notebook);
  const complete = !values.section && !values.notebook;
  const r = await snapshotSections(sections, { force: values.force, complete, reason: "scheduled" });
  console.log(
    `${new Date().toISOString()} sections=${r.sections} pages=${r.pagesChecked} saved=${r.saved} unchanged=${r.unchanged} ` +
      `pruned=${r.pruned} failed=${r.failed.length}`
  );
  for (const f of r.failed) console.log(`  failed: ${f.page ?? f.section}: ${f.error}`);
  process.exit(r.failed.length ? 2 : 0);
} catch (e) {
  console.error(e.message);
  process.exit(1);
}

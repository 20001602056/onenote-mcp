// End-to-end test of every tool against ONENOTE_TEST_NOTEBOOK / ONENOTE_TEST_SECTION.
// Creates a test page (plus a copy and a restored version), edits, exports and versions it, then deletes
// them when ONENOTE_ALLOW_DELETE=true. Versions and exports go to a temp folder, not your real ones.
import { mkdtemp, writeFile, stat, readFile, readdir } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { NOTEBOOK, SECTION, connect, requireTarget } from "./config.mjs";

requireTarget();
const dir = await mkdtemp(join(tmpdir(), "onenote-mcp-test-"));
process.env.ONENOTE_VERSION_DIR = join(dir, "versions");
process.env.ONENOTE_EXPORT_DIR = join(dir, "export");
const c = await connect("feature-test");
let failed = 0;
async function call(name, args = {}, { allowError = false } = {}) {
  const r = await c.callTool({ name, arguments: args });
  const out = r.content[0].text;
  const ok = allowError || !r.isError;
  if (!ok) failed++;
  console.log(`${ok ? "PASS" : "FAIL"}  ${name}${ok ? "" : `\n      ${out.slice(0, 300)}`}`);
  return r.isError ? null : out;
}
const json = (s) => (s ? JSON.parse(s) : null);
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

const tools = (await c.listTools()).tools.map((t) => t.name);
console.log(`tools: ${tools.join(", ")}\n`);

// Read-only tools
await call("server_settings");
const nbs = json(await call("list_notebooks"));
await call("get_recent_notebooks");
await call("list_sections", { notebook: NOTEBOOK });
await call("list_pages", { notebook: NOTEBOOK, section: SECTION, limit: 5, orderBy: "title asc" });

// Local attachments: a 1x1 PNG and a small text file
const png = join(dir, "dot.png");
const txt = join(dir, "notes.txt");
await writeFile(png, Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8DwHwAFBQIAX8jx0gAAAABJRU5ErkJggg==", "base64"));
await writeFile(txt, "attachment from onenote-mcp feature test\n");

const stamp = new Date().toISOString().replace(/[:.]/g, "-");
const title = `onenote-mcp feature test ${stamp}`;
const created = json(
  await call("create_page", {
    notebook: NOTEBOOK,
    section: SECTION,
    title,
    bodyHtml:
      '<h2>Feature test</h2><p data-tag="to-do" data-id="task1">Tick me</p><p data-tag="important">Important line</p>' +
      '<p data-id="status">status: created</p><img src="name:pic" alt="dot" />',
    attachments: [{ path: png, partName: "pic" }, { path: txt }],
  })
);
const pageId = created?.id;
if (!pageId) {
  console.log("\nCannot continue without a page.");
  process.exit(1);
}

await sleep(3000);
await call("search_pages", { query: "onenote-mcp feature test", notebook: NOTEBOOK, section: SECTION });
await call("get_page", { pageId });

await call("update_page", {
  pageId,
  changes: [
    { target: "title", action: "replace", content: `${title} (renamed)` },
    { target: "#task1", action: "replace", content: '<p data-tag="to-do:completed" data-id="task1">Ticked</p>' },
    { target: "#status", action: "insert", position: "after", content: "<p>inserted after status</p>" },
  ],
});
await call("append_to_page", { pageId, html: "<p>appended with attachment</p>", attachments: [{ path: txt, partName: "again" }] });

await sleep(3000);
const html = await call("read_page", { pageId, includeIds: true });
const checks = {
  "title renamed": html?.includes("(renamed)"),
  "to-do ticked": html?.includes("to-do:completed"),
  "insert applied": html?.includes("inserted after status"),
  "append applied": html?.includes("appended with attachment"),
  "image embedded": /<img[^>]+resources\//.test(html || ""),
};
for (const [k, v] of Object.entries(checks)) {
  if (!v) failed++;
  console.log(`${v ? "PASS" : "FAIL"}  content: ${k}`);
}

const resourceUrl = (html || "").match(/https:\/\/graph\.microsoft\.com\/[^"]+?\/resources\/[^"/]+\/\$value/)?.[0];
if (resourceUrl) {
  const saved = json(await call("download_resource", { pageId, resource: resourceUrl, outputPath: join(dir, "downloaded.png") }));
  if (saved) console.log(`      saved ${(await stat(saved.saved)).size} bytes`);
} else console.log("SKIP  download_resource (no resource URL found)");

// Copy (background operation)
let copiedId, restoredId;
const op = json(await call("copy_page", { pageId, toNotebook: NOTEBOOK, toSection: SECTION }));
if (op?.copied) {
  copiedId = op.id;
  console.log(`${copiedId ? "PASS" : (failed++, "FAIL")}  copy completed (${op.method}, ${op.resourcesCopied} resource(s))`);
} else if (op?.operationId) {
  for (let i = 0; i < 20; i++) {
    await sleep(2000);
    const st = json(await call("get_operation", { operationId: op.operationId }, { allowError: true }));
    if (st?.status === "Completed" || st?.status === "completed") {
      copiedId = st.resourceId;
      break;
    }
    if (/fail/i.test(st?.status || "")) break;
  }
  console.log(`${copiedId ? "PASS" : (failed++, "FAIL")}  copy completed`);
}

// read_page in other formats
const md = await call("read_page", { pageId, format: "markdown" });
console.log(`${md?.includes("- [x] Ticked") ? "PASS" : (failed++, "FAIL")}  read_page markdown shows the ticked to-do`);

// Export: every format for the page, then the section as one combined Markdown file
for (const format of ["markdown", "html", "text", "json", "pdf"]) {
  const r = json(await call("export_page", { pageId, format }, { allowError: format === "pdf" }));
  if (!r) {
    console.log("      (pdf needs a local Chrome/Edge; skipped)");
    continue;
  }
  const body = format === "pdf" ? "" : await readFile(r.saved, "utf8");
  const ok = r.bytes > 0 && (format === "pdf" || body.includes("inserted after status"));
  if (!ok) failed++;
  console.log(`${ok ? "PASS" : "FAIL"}  export ${format}: ${r.bytes} bytes, ${r.resources} resource(s)`);
}
const sec = json(await call("export_section", { notebook: NOTEBOOK, section: SECTION, format: "markdown", combine: true, resources: "none" }));
if (sec) console.log(`${sec.pages > 0 ? "PASS" : (failed++, "FAIL")}  export_section combined: ${sec.pages} page(s)`);

// Versions: update_page and append_to_page each saved one first
if (tools.includes("list_page_versions")) {
  const versions = json(await call("list_page_versions", { pageId })) || [];
  console.log(`${versions.length >= 2 ? "PASS" : (failed++, "FAIL")}  ${versions.length} version(s) saved before edits`);
  const oldest = versions.at(-1);
  if (oldest) {
    const v = await call("read_page_version", { pageId, versionId: oldest.versionId, format: "text" });
    console.log(`${v?.includes("Tick me") && !v.includes("Ticked") ? "PASS" : (failed++, "FAIL")}  oldest version has the original content`);
    const restored = json(await call("restore_page_version", { pageId, versionId: oldest.versionId, title: `${title} (restored)` }));
    restoredId = restored?.id;
    console.log(`${restored?.resourcesCopied >= 1 ? "PASS" : (failed++, "FAIL")}  restored with ${restored?.resourcesCopied} resource(s)`);
  }
  const snap = json(await call("snapshot_pages", { notebook: NOTEBOOK, section: SECTION }));
  if (snap) console.log(`${snap.failed.length ? (failed++, "FAIL") : "PASS"}  snapshot_pages: ${snap.pagesChecked} checked, ${snap.saved} saved`);
}

// Guardrail sanity: blocked names must be refused when configured
const blockedName = (process.env.ONENOTE_BLOCKED_SECTIONS || "").split(",").map((s) => s.trim()).find((s) => s && !s.includes("!"));
if (blockedName) {
  const r = await c.callTool({ name: "create_section", arguments: { notebook: NOTEBOOK, section: blockedName } });
  const ok = r.isError && r.content[0].text.includes("guardrail");
  if (!ok) failed++;
  console.log(`${ok ? "PASS" : "FAIL"}  create_section refuses blocked name`);
}

// Cleanup
if (tools.includes("delete_page")) {
  await call("delete_page", { pageId });
  if (copiedId) await call("delete_page", { pageId: copiedId });
  if (restoredId) await call("delete_page", { pageId: restoredId });
  const left = await readdir(join(dir, "versions")).catch(() => []);
  console.log(`${left.length ? "PASS" : (failed++, "FAIL")}  delete_page saved a version first (${left.length} page folder(s))`);
} else {
  console.log(`\nNOTE  delete_page is disabled; remove the test page(s) manually: "${title} (renamed)"`);
}

await c.close();
console.log(failed ? `\n${failed} check(s) failed` : "\nAll feature checks passed");
process.exit(failed ? 1 : 0);

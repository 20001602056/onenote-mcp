// Verifies the blocked section is refused by every tool and that an allowed section still works.
// Needs ONENOTE_BLOCKED_SECTIONS set to "<SectionName>,<sectionId>", plus ONENOTE_TEST_NOTEBOOK and
// ONENOTE_TEST_SECTION (an allowed section in the same notebook).
import { getAccessToken } from "../auth.js";
import { NOTEBOOK, SECTION, connect, requireTarget } from "./config.mjs";

requireTarget();
const blocked = (process.env.ONENOTE_BLOCKED_SECTIONS || "").split(",").map((s) => s.trim()).filter(Boolean);
const BLOCKED_NAME = blocked.find((s) => !s.includes("!"));
const BLOCKED_ID = blocked.find((s) => s.includes("!"));
if (!BLOCKED_NAME || !BLOCKED_ID) {
  console.error('Set ONENOTE_BLOCKED_SECTIONS="<SectionName>,<sectionId>" (get the id from list_notebooks before blocking it).');
  process.exit(1);
}

const c = await connect("guardrail-test");

// Fetch only the id (no title or content) of one blocked page, to probe page-level tools.
const r = await fetch(`https://graph.microsoft.com/v1.0/me/onenote/sections/${BLOCKED_ID}/pages?$select=id&$top=1`, {
  headers: { Authorization: `Bearer ${await getAccessToken()}` },
});
const blockedPageId = (await r.json()).value?.[0]?.id;

let failed = 0;
async function expect(label, name, args, shouldBlock) {
  const res = await c.callTool({ name, arguments: args });
  const out = res.content[0].text;
  const isBlocked = res.isError && out.includes("guardrail");
  const ok = shouldBlock ? isBlocked : !res.isError;
  if (!ok) failed++;
  console.log(`${ok ? "PASS" : "FAIL"}  ${label}${ok ? "" : `\n      ${out.slice(0, 200)}`}`);
  return out;
}

const nb = await expect("list_notebooks works", "list_notebooks", {}, false);
const hidden = !nb.toLowerCase().includes(`"${BLOCKED_NAME.toLowerCase()}"`) && !nb.includes(BLOCKED_ID);
console.log(`${hidden ? "PASS" : (failed++, "FAIL")}  list_notebooks hides ${BLOCKED_NAME}`);
await expect("list_pages by name blocked", "list_pages", { notebook: NOTEBOOK, section: BLOCKED_NAME }, true);
await expect("list_pages by name (other case) blocked", "list_pages", { notebook: NOTEBOOK, section: BLOCKED_NAME.toUpperCase() }, true);
await expect("list_pages by id blocked", "list_pages", { sectionId: BLOCKED_ID }, true);
await expect("search_pages blocked", "search_pages", { query: "", sectionId: BLOCKED_ID }, true);
await expect("create_page blocked", "create_page", { title: "x", bodyHtml: "<p>x</p>", sectionId: BLOCKED_ID }, true);
await expect("export_section blocked", "export_section", { sectionId: BLOCKED_ID, format: "markdown" }, true);
await expect("snapshot_pages blocked", "snapshot_pages", { sectionId: BLOCKED_ID }, true);
await expect("create_section with blocked name refused", "create_section", { notebook: NOTEBOOK, section: BLOCKED_NAME.toLowerCase() }, true);
if (blockedPageId) {
  await expect("read_page blocked", "read_page", { pageId: blockedPageId }, true);
  await expect("append_to_page blocked", "append_to_page", { pageId: blockedPageId, html: "<p>x</p>" }, true);
  await expect("read_page markdown blocked", "read_page", { pageId: blockedPageId, format: "markdown" }, true);
  await expect("export_page blocked", "export_page", { pageId: blockedPageId, format: "text" }, true);
  await expect("list_page_versions blocked", "list_page_versions", { pageId: blockedPageId }, true);
  await expect("copy_page from blocked section refused", "copy_page", { pageId: blockedPageId, toNotebook: NOTEBOOK, toSection: SECTION }, true);
} else console.log(`SKIP  page-level checks (${BLOCKED_NAME} has no pages)`);
await expect(`${SECTION} list_pages still works`, "list_pages", { notebook: NOTEBOOK, section: SECTION }, false);

await c.close();
console.log(failed ? `\n${failed} check(s) failed` : "\nAll guardrail checks passed");
process.exit(failed ? 1 : 0);

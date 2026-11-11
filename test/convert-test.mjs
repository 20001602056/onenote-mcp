// Offline checks (no Microsoft sign-in needed): HTML conversion for export, and guardrail rule matching.
process.env.ONENOTE_BLOCKED_SECTIONS = "Credentials,0-ABC!123,My Notebook/Private/**";
process.env.ONENOTE_BLOCKED_PAGES = "Passwords*,0-PAGE!9";
process.env.ONENOTE_PROTECTED_SECTIONS = "Archive";
process.env.ONENOTE_PROTECTED_PAGES = "_Global,Work/*/Runbook";

const { toMarkdown, toText, unposition, safeName, decode } = await import("../lib/export.js");
const g = await import("../lib/guardrail.js");

let failed = 0;
function check(label, ok, detail = "") {
  if (!ok) failed++;
  console.log(`${ok ? "PASS" : "FAIL"}  ${label}${ok ? "" : `\n      ${detail}`}`);
}
const throws = (fn, needle) => {
  try {
    fn();
    return false;
  } catch (e) {
    return e.message.includes(needle);
  }
};

// ---- Conversion ----
const page = `<html><head><title>T &amp; x</title></head><body data-absolute-enabled="true" style="font-family:Calibri">
  <div style="position:absolute;left:48px;top:120px;width:576px">
    <h2 style="color:#2e75b5">Overview</h2>
    <p data-tag="to-do">Open task</p>
    <p data-tag="to-do:completed">Done task</p>
    <ul><li>Item with <a href="https://example.com">link</a></li><li>Run <span style="font-family:consolas">npm test</span></li></ul>
    <table><tr><td>A</td><td>B | C</td></tr><tr><td>1</td><td>2</td></tr></table>
    <img src="assets/x.png" alt="pic" />
    <object data-attachment="notes.txt" type="text/plain" data="assets/notes.txt"></object>
    <p>Q&#39;s &lt;tag&gt;</p>
  </div></body></html>`;

const md = toMarkdown(page);
check("md heading", md.includes("## Overview"), md);
check("md open to-do", md.includes("- [ ] Open task"), md);
check("md done to-do", md.includes("- [x] Done task"), md);
check("md link", md.includes("[link](https://example.com)"), md);
check("md inline code", md.includes("`npm test`"), md);
check("md table header", /\| A \| B \\\| C \|\n\| --- \| --- \|\n\| 1 \| 2 \|/.test(md), md);
check("md image", md.includes("![pic](assets/x.png)"), md);
check("md attachment", md.includes("[Attachment: notes.txt](assets/notes.txt)"), md);
check("md keeps content after a self-closing attachment", toMarkdown('<p>a</p><object data-attachment="f.txt" data="x/f.txt" type="text/plain" /><p>after</p>').includes("after"));
check("md entities", md.includes("Q's <tag>") || md.includes("Q's \\<tag\\>"), md);

const txt = toText(page);
check("text to-dos", txt.includes("[ ] Open task") && txt.includes("[x] Done task"), txt);
check("text list", txt.includes("- Item with link"), txt);
check("text table", txt.includes("A\tB | C"), txt);
check("text attachment", txt.includes("[attachment: notes.txt]"), txt);
check("text entities", txt.includes("Q's <tag>"), txt);
check("text no stray indentation", !/\n[ \t]+\S/.test(txt), JSON.stringify(txt));

check("unposition", unposition('<div style="position:absolute;left:48px;top:120px;width:576px">') === '<div style="width:576px">');
check("safeName", safeName('a/b:c*?"<>|  d') === "a-b-c------ d" && safeName("") === "Untitled");
check("decode", decode("&#x1F600;&amp;&nbsp;") === "😀& ");

// ---- Guardrail rules ----
const sec = (name, path, id = "s1") => ({ id, name, path });
check("blocked section by name (any case)", g.isSectionBlocked(sec("CREDENTIALS", null)));
check("blocked section by id", g.isSectionBlocked(sec("Renamed", null, "0-abc!123")));
check("blocked section by path **", g.isSectionBlocked(sec("Bank", "My Notebook/Private/Deep/Bank")));
check("other section allowed", !g.isSectionBlocked(sec("Work", "My Notebook/Work")));
check("path rule needs a path", !g.isSectionBlocked(sec("Bank", null)));

const work = sec("Claude", "My Notebook/Claude");
check("blocked page by title wildcard", g.isPageBlocked({ id: "p1", title: "passwords 2026" }, work));
check("blocked page by id", g.isPageBlocked({ id: "0-PAGE!9", title: "x" }, work));
check("page in blocked section blocked", g.isPageBlocked({ id: "p1", title: "x" }, sec("Credentials", null)));
check("normal page allowed", !g.isPageBlocked({ id: "p1", title: "Notes" }, work));

check("protected section: read ok", (g.checkSection(sec("Archive", null)), true));
check("protected section: write refused", throws(() => g.checkSection(sec("Archive", null), { write: true }), "read-only"));
check("protected page: read ok", (g.checkPage({ id: "p", title: "_Global" }, work), true));
check("protected page: write refused", throws(() => g.checkPage({ id: "p", title: "_global" }, work, { write: true }), "read-only"));
check(
  "protected page by path (* = one level)",
  throws(() => g.checkPage({ id: "p", title: "Runbook" }, sec("Ops", "Work/Ops"), { write: true }), "read-only") &&
    !throws(() => g.checkPage({ id: "p", title: "Runbook" }, sec("Ops", "Work/Team/Ops"), { write: true }), "read-only")
);
check("new page with blocked title refused", throws(() => g.checkNewPageTitle("Passwords", work), "guardrail"));
check("new section with blocked name refused", throws(() => g.checkNewSection("credentials", "My Notebook"), "guardrail"));
check("new section inside blocked group refused", throws(() => g.checkNewSection("New", "My Notebook/Private"), "guardrail"));
check("stored version without path refused when path rules exist", throws(() => g.checkStored({ id: "p", title: "x" }, sec("Work", null)), "guardrail"));

console.log(failed ? `\n${failed} check(s) failed` : "\nAll conversion and guardrail checks passed");
process.exit(failed ? 1 : 0);

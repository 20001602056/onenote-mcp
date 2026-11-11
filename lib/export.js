// Page conversion (HTML -> Markdown / text) and export to files.
import { execFile } from "node:child_process";
import { existsSync } from "node:fs";
import { mkdir, mkdtemp, rm, stat, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { basename, delimiter, dirname, extname, join } from "node:path";
import { pathToFileURL } from "node:url";
import { promisify } from "node:util";
import TurndownService from "turndown";
import { CHROME_PATH } from "./config.js";
import { esc } from "./graph.js";
import { extForType, fetchResource, pageResources, stripFullres } from "./onenote.js";

export const FORMATS = ["markdown", "html", "text", "json", "pdf"];
export const EXT = { markdown: ".md", html: ".html", text: ".txt", json: ".json", pdf: ".pdf" };

// ---- Conversion ----
const ENTITIES = { amp: "&", lt: "<", gt: ">", quot: '"', apos: "'", nbsp: " " };
export const decode = (s) =>
  s.replace(/&(#x[0-9a-f]+|#\d+|[a-z]+);/gi, (m, e) => {
    if (e[0] !== "#") return ENTITIES[e.toLowerCase()] ?? m;
    const code = e[1].toLowerCase() === "x" ? parseInt(e.slice(2), 16) : Number(e.slice(1));
    return code <= 0x10ffff ? String.fromCodePoint(code) : m;
  });

export const titleOf = (html) => decode(html.match(/<title>([\s\S]*?)<\/title>/i)?.[1] ?? "").trim();
export const bodyOf = (html) => html.match(/<body\b[^>]*>([\s\S]*)<\/body>/i)?.[1] ?? html;

// OneNote writes attachments as self-closing <object ... />, which HTML parsers (and browsers) don't close,
// so everything after one would become hidden fallback content. Turn each into a plain link instead.
export const linkAttachments = (html) =>
  html.replace(/<object\b([^>]*?)\/?>(?:\s*<\/object>)?/gi, (m, attrs) => {
    const get = (name) => attrs.match(new RegExp(`\\s${name}="([^"]*)"`))?.[1];
    const label = `Attachment: ${get("data-attachment") || "file"}`;
    const href = get("data");
    return `<p class="attachment">${href ? `<a href="${href}">${label}</a>` : label}</p>`;
  });

// OneNote positions each outline absolutely; drop that so exported pages flow top to bottom.
export const unposition = (html) =>
  html.replace(/\sstyle="([^"]*)"/gi, (m, css) => {
    const kept = css
      .split(";")
      .filter((d) => d.trim() && !/^\s*(position|left|top)\s*:/i.test(d))
      .join(";");
    return kept ? ` style="${kept}"` : "";
  });

const tagOf = (n) => (n.getAttribute && n.getAttribute("data-tag")) || "";
const turndown = new TurndownService({ headingStyle: "atx", bulletListMarker: "-", codeBlockStyle: "fenced", emDelimiter: "*" });
turndown.remove(["script", "style", "title", "meta"]);
turndown.addRule("todo", {
  filter: (n) => n.nodeName === "P" && tagOf(n).includes("to-do"),
  replacement: (content, n) => `\n- [${tagOf(n).includes("to-do:completed") ? "x" : " "}] ${content.trim()}\n`,
});
// OneNote writes inline code as a Consolas span.
turndown.addRule("monospace", {
  filter: (n) => n.nodeName === "SPAN" && /font-family:\s*'?(consolas|courier)/i.test(n.getAttribute("style") || ""),
  replacement: (_c, n) => {
    const code = n.textContent;
    if (!code.trim()) return code;
    const fence = code.includes("`") ? "``" : "`";
    return `${fence}${code}${fence}`;
  },
});
// OneNote tables have no header row, so the first row becomes the Markdown header.
turndown.addRule("table", {
  filter: "table",
  replacement: (_c, table) => {
    const rows = [];
    const walk = (node) => {
      for (const child of node.childNodes) {
        if (child.nodeName === "TR") rows.push(child);
        else if (["THEAD", "TBODY", "TFOOT"].includes(child.nodeName)) walk(child);
      }
    };
    walk(table);
    const cells = rows.map((tr) =>
      [...tr.childNodes]
        .filter((c) => c.nodeName === "TD" || c.nodeName === "TH")
        .map((c) => turndown.turndown(c.innerHTML).replace(/\n+/g, "<br>").replace(/\|/g, "\\|").trim())
    );
    if (!cells.length) return "";
    const width = Math.max(...cells.map((r) => r.length));
    const line = (r) => `| ${Array.from({ length: width }, (_, i) => r[i] ?? "").join(" | ")} |`;
    return `\n\n${[line(cells[0]), line(Array(width).fill("---")), ...cells.slice(1).map(line)].join("\n")}\n\n`;
  },
});

export const toMarkdown = (html) => `${turndown.turndown(linkAttachments(bodyOf(html))).trim()}\n`;

export function toText(html) {
  const s = bodyOf(html)
    .replace(/(<pre\b[\s\S]*?<\/pre>)|\s+/gi, (m, pre) => pre || " ")
    .replace(/<(script|style)\b[\s\S]*?<\/\1>/gi, "")
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<li\b[^>]*>/gi, "\n- ")
    .replace(/<p\b[^>]*\sdata-tag="([^"]*)"[^>]*>/gi, (m, tag) =>
      tag.includes("to-do:completed") ? "\n[x] " : tag.includes("to-do") ? "\n[ ] " : "\n"
    )
    .replace(/<\/(td|th)>/gi, "\t")
    .replace(/<\/?(p|div|h[1-6]|tr|table|ul|ol|pre|blockquote)\b[^>]*>/gi, "\n")
    .replace(/<img\b[^>]*\salt="([^"]*)"[^>]*>/gi, "[image: $1]")
    .replace(/<object\b[^>]*\sdata-attachment="([^"]*)"[^>]*>/gi, "[attachment: $1]")
    .replace(/<[^>]+>/g, "");
  return `${decode(s)
    .replace(/\n /g, "\n")
    .replace(/[ \t]+\n/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim()}\n`;
}

// Page HTML in the format read_page / read_page_version return.
export const convert = (html, format) => (format === "markdown" ? toMarkdown(html) : format === "text" ? toText(html) : html);

// ---- Files ----
export const safeName = (s, fallback = "Untitled") =>
  (s || "")
    .replace(/[/\\:*?"<>|\x00-\x1f]/g, "-")
    .replace(/\s+/g, " ")
    .trim()
    .replace(/^\.+/, "")
    .slice(0, 120)
    .trim() || fallback;

const exists = (p) => stat(p).then(() => true, () => false);

// p, or "name (2).ext", "name (3).ext", ... when p (or a name already used in this export) exists.
export async function uniquePath(p, used = new Set()) {
  const ext = extname(p);
  const stem = p.slice(0, p.length - ext.length);
  for (let i = 1; ; i++) {
    const candidate = i === 1 ? p : `${stem} (${i})${ext}`;
    if (!used.has(candidate.toLowerCase()) && !(await exists(candidate))) {
      used.add(candidate.toLowerCase());
      return candidate;
    }
  }
}

// Swap Graph resource URLs for files in assetsDir ("download"), data: URIs ("inline"), or drop them ("none").
export async function localizeResources(html, mode, { assetsDir, assetsHref, load = fetchResource } = {}) {
  html = stripFullres(html);
  const resources = pageResources(html);
  if (mode === "none") {
    return { html: html.replace(/<img\b[^>]*\/resources\/[^>]*>/gi, "").replace(/\sdata="[^"]*\/resources\/[^"]*"/gi, ""), count: 0 };
  }
  for (const r of resources) {
    const data = await load(r.id);
    let ref;
    if (mode === "inline") {
      ref = `data:${r.type || "application/octet-stream"};base64,${data.toString("base64")}`;
    } else {
      const file = safeName(r.id) + (r.name ? extname(r.name) : extForType(r.type));
      await mkdir(assetsDir, { recursive: true });
      await writeFile(join(assetsDir, file), data);
      ref = encodeURI(`${assetsHref}/${file}`).replace(/\(/g, "%28").replace(/\)/g, "%29");
    }
    html = html.split(r.url).join(ref);
  }
  return { html, count: resources.length };
}

// ---- Rendering ----
const exportBody = (html) => linkAttachments(unposition(bodyOf(html)));

// page: { id, title, created, modified, notebook, section, path, webUrl }
const metaLine = (p) => [p.path || [p.notebook, p.section].filter(Boolean).join("/"), p.modified && `modified ${p.modified}`].filter(Boolean).join(" · ");

const frontMatter = (fields) =>
  `---\n${Object.entries(fields)
    .filter(([, v]) => v != null && v !== "")
    .map(([k, v]) => `${k}: ${JSON.stringify(v)}`)
    .join("\n")}\n---\n\n`;

const pageFields = (p) => ({ title: p.title, id: p.id, notebook: p.notebook, section: p.section, created: p.created, modified: p.modified, url: p.webUrl });

const STYLE = `body{font-family:Calibri,"Segoe UI",Helvetica,Arial,sans-serif;font-size:11pt;line-height:1.45;max-width:820px;margin:2rem auto;padding:0 1rem;color:#222}
img{max-width:100%;height:auto}table{border-collapse:collapse;margin:.5rem 0}td,th{border:1px solid #bbb;padding:4px 8px;vertical-align:top}
.meta{color:#666;font-size:9pt;margin-top:-.5rem}article+article{border-top:1px solid #ddd;margin-top:2rem;padding-top:1rem}
[data-tag*="to-do"]::before{content:"\\2610  "}[data-tag*="to-do:completed"]::before{content:"\\2611  "}
@media print{body{margin:0;max-width:none}article{break-before:page}article:first-of-type{break-before:auto}}`;

const htmlDoc = (title, inner) =>
  `<!DOCTYPE html>\n<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">` +
  `<title>${esc(title)}</title><style>${STYLE}</style></head><body>\n${inner}\n</body></html>\n`;

const pageArticle = (p, html, heading = "h1") =>
  `<article><${heading}>${esc(p.title || "Untitled")}</${heading}><p class="meta">${esc(metaLine(p))}</p>\n${exportBody(html)}</article>`;

// One page -> file content (pdf is rendered as html first).
export function renderPage(p, html, format) {
  switch (format) {
    case "markdown":
      return `${frontMatter(pageFields(p))}# ${p.title || "Untitled"}\n\n${toMarkdown(html)}`;
    case "text":
      return `${p.title || "Untitled"}\n${"=".repeat(Math.min((p.title || "Untitled").length, 80))}\n${metaLine(p)}\n\n${toText(html)}`;
    case "json":
      return `${JSON.stringify({ ...pageFields(p), path: p.path, html: exportBody(html), markdown: toMarkdown(html) }, null, 2)}\n`;
    default:
      return htmlDoc(p.title || "Untitled", pageArticle(p, html));
  }
}

// Several pages -> one file. section: { name, notebook, path }
export function renderCombined(section, pages, format) {
  const title = section.name;
  switch (format) {
    case "markdown":
      return (
        frontMatter({ section: section.name, notebook: section.notebook, pages: pages.length, exported: new Date().toISOString() }) +
        pages.map(({ page, html }) => `# ${page.title || "Untitled"}\n\n${toMarkdown(html)}`).join("\n---\n\n")
      );
    case "text":
      return pages.map(({ page, html }) => renderPage(page, html, "text")).join(`\n${"-".repeat(40)}\n\n`);
    case "json":
      return `${JSON.stringify(
        {
          section: section.name,
          notebook: section.notebook,
          path: section.path,
          exported: new Date().toISOString(),
          pages: pages.map(({ page, html }) => ({ ...pageFields(page), html: exportBody(html), markdown: toMarkdown(html) })),
        },
        null,
        2
      )}\n`;
    default: {
      const toc = pages.map(({ page }, i) => `<li><a href="#p${i + 1}">${esc(page.title || "Untitled")}</a></li>`).join("");
      const body = pages.map(({ page, html }, i) => pageArticle(page, html, "h2").replace("<article>", `<article id="p${i + 1}">`)).join("\n");
      return htmlDoc(title, `<h1>${esc(title)}</h1><p class="meta">${esc(section.path || "")}</p><ol>${toc}</ol>\n${body}`);
    }
  }
}

// ---- PDF (printed by a local Chrome / Edge / Chromium / Brave in headless mode) ----
export function findBrowser() {
  if (CHROME_PATH) return existsSync(CHROME_PATH) ? CHROME_PATH : null;
  let candidates;
  if (process.platform === "darwin") {
    candidates = ["Google Chrome", "Microsoft Edge", "Chromium", "Brave Browser"].map((app) => `/Applications/${app}.app/Contents/MacOS/${app}`);
  } else if (process.platform === "win32") {
    const roots = [process.env.PROGRAMFILES, process.env["PROGRAMFILES(X86)"], process.env.LOCALAPPDATA].filter(Boolean);
    candidates = roots.flatMap((r) => [join(r, "Google", "Chrome", "Application", "chrome.exe"), join(r, "Microsoft", "Edge", "Application", "msedge.exe")]);
  } else {
    const dirs = (process.env.PATH || "").split(delimiter).filter(Boolean);
    candidates = ["google-chrome", "google-chrome-stable", "chromium", "chromium-browser", "microsoft-edge", "brave-browser"].flatMap((b) =>
      dirs.map((d) => join(d, b))
    );
  }
  return candidates.find((c) => existsSync(c)) || null;
}

export async function htmlToPdf(html, outFile) {
  const browser = findBrowser();
  if (!browser) throw new Error("PDF export needs Chrome, Edge, Chromium or Brave. Install one, or set ONENOTE_CHROME_PATH to its executable.");
  const dir = await mkdtemp(join(tmpdir(), "onenote-pdf-"));
  try {
    const src = join(dir, "page.html");
    await writeFile(src, html);
    await mkdir(dirname(outFile), { recursive: true });
    await promisify(execFile)(
      browser,
      ["--headless=new", "--disable-gpu", "--no-pdf-header-footer", `--user-data-dir=${join(dir, "profile")}`, `--print-to-pdf=${outFile}`, pathToFileURL(src).href],
      { timeout: 120_000 }
    );
    if (!(await exists(outFile))) throw new Error(`${basename(browser)} did not produce a PDF`);
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
}

// Write rendered content (or print it to PDF) and return the byte count.
export async function writeOutput(file, content, format) {
  await mkdir(dirname(file), { recursive: true });
  if (format === "pdf") await htmlToPdf(content, file);
  else await writeFile(file, content);
  return (await stat(file)).size;
}

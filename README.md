# onenote-mcp

A [Model Context Protocol](https://modelcontextprotocol.io) server that lets Claude (Claude Code, Claude Desktop, or any MCP client) read and write your Microsoft OneNote notebooks through Microsoft Graph.

- **No Azure app registration.** It signs in as you with Microsoft's public "Microsoft Graph Command Line Tools" client, using device-code login.
- **Tokens stay on your machine:** in the macOS Keychain, or in a `0600` file on Linux and Windows. They refresh automatically.
- **Guardrail:** block sections, section groups or pages from every tool, or make them read-only. Rules match ids, names (with wildcards) or paths.
- **Export:** pages and whole sections to Markdown, HTML, plain text, JSON or PDF, with their images and files.
- **Versioning:** a snapshot of each page is saved before every change and on a schedule, kept for 30 days, and can be restored with images and files.
- **Safety switches:** read-only mode, page deletion that stays off unless you turn it on, and an optional audit log.

## Tools

Each tool tells your MCP client whether it only reads, writes, or can destroy content, so clients can ask before risky actions.

**Notebooks and structure**

| Tool | What it does |
|---|---|
| `list_notebooks` | Full tree of every notebook: nested section groups and sections, with ids, default/shared flags and links |
| `get_recent_notebooks` | Recently opened notebooks, with last-access time and links |
| `create_notebook` | Create a notebook (or return the existing one) |
| `create_section_group` | Create a section group in a notebook or nested inside another group |
| `list_sections` | Every section, including those inside section groups, optionally for one notebook |
| `create_section` | Create a section in a notebook or a section group |
| `copy_section` | Copy a section and all its pages into another notebook or section group |

**Pages**

| Tool | What it does |
|---|---|
| `list_pages` | Pages in a section, with indent level and order; choose sort order and limit |
| `search_pages` | Search page titles in one section, one notebook, or everywhere |
| `get_page` | Page details: section, notebook, links and a short text preview |
| `read_page` | Full page as HTML (optionally with element ids for precise edits), Markdown or plain text |
| `create_page` | New page from HTML, with tags, web images and local images/files |
| `update_page` | Edit in place: rename, replace/insert/append/prepend at any element, tick to-dos, add attachments (saves a version first) |
| `append_to_page` | Append HTML (and attachments) to the end of a page (saves a version first) |
| `copy_page` | Copy a page with its images and files to another section (copy + delete = move) |
| `delete_page` | Delete a page (only when `ONENOTE_ALLOW_DELETE=true`; saves a version first and refuses if it can't) |
| `download_resource` | Save an image or attached file from a page to disk |
| `get_operation` | Check a background copy (work/school accounts) |
| `server_settings` | Show the server's active settings |

**Export**

| Tool | What it does |
|---|---|
| `export_page` | Save a page as Markdown, HTML, text, JSON or PDF, with its images and files |
| `export_section` | Save every page in a section, one file per page (plus an index) or one combined file |

**Versions**

| Tool | What it does |
|---|---|
| `snapshot_pages` | Save a version of every changed page in a section, a notebook, or everywhere |
| `list_page_versions` | A page's saved versions, or every page with versions (including deleted pages) |
| `read_page_version` | Read a saved version as HTML, Markdown or text |
| `restore_page_version` | Recreate a saved version, with its images and files, as a new page |

Sections are identified by `sectionId`, or by `notebook` + `section` names (+ `sectionGroup` for a section inside a group). With `ONENOTE_DEFAULT_NOTEBOOK` and `ONENOTE_DEFAULT_SECTION` set, you can leave them out.

### Page content tips

- **Tags:** `<p data-tag="to-do">Buy milk</p>`. Other tags: `to-do:completed`, `important`, `question`, `idea`, `critical`, `contact`, `remember-for-later`, `definition`, `highlight`, and more.
- **Ticking a to-do:** give the line a `data-id` when you create it, then `update_page` with `{target: "#my-id", action: "replace", content: "<p data-tag=\"to-do:completed\" data-id=\"my-id\">Done</p>"}`.
- **Images from the web:** `<img src="https://..." />`. OneNote downloads and embeds them.
- **Local images and files:** pass `attachments: [{ path: "~/Pictures/chart.png" }]`. Reference them as `name:file1`, or leave them unreferenced and they're added at the end.
- **Renaming a page:** `update_page` with `{target: "title", action: "replace", content: "New title"}`.

### What Microsoft Graph doesn't support

These OneNote features aren't available through Graph, so no OneNote MCP server can provide them (page version history is one of them, which is why this server keeps its own; see [Versioning](#versioning)):

- OneNote **app settings** (themes, sync options, proofing, etc.)
- **Renaming or deleting** notebooks, section groups or sections (pages can be renamed and deleted)
- Setting a page's **indent level or order**, section colors, or page templates
- Ink/drawings as editable strokes, audio recording, password-protected sections

On **personal Microsoft accounts**, Graph's built-in copy isn't available. `copy_page` and `copy_section` copy content by hand instead (read → download images/files → recreate). The result is the same content, but the copy has a new created date.

## Exporting pages and sections

`export_page` and `export_section` write files to `ONENOTE_EXPORT_DIR` (default `~/Downloads`), or to the path you give. Without an explicit file path they never overwrite: a second export becomes `Name (2).md`.

| Format | Output |
|---|---|
| `markdown` | YAML front matter (title, id, notebook, section, dates, link), then the page. To-dos become `- [ ]` / `- [x]`, tables become Markdown tables, attachments become links |
| `html` | A standalone page that opens in any browser (OneNote's absolute positioning removed, to-do boxes shown) |
| `text` | Plain text with `[ ]` / `[x]` to-dos and tab-separated tables |
| `json` | Page metadata plus its HTML and Markdown, for scripts |
| `pdf` | Printed by a local Chrome, Edge, Chromium or Brave in headless mode (set `ONENOTE_CHROME_PATH` if it isn't found) |

Images and attached files (`resources`):

- `download` (default): saved in `<name>_files/` next to the export (or `assets/` in a section folder), with relative links.
- `inline`: embedded as `data:` URIs, so one file holds everything. PDF always does this.
- `none`: left out. Attachments still show their file name.

`export_section` writes `<Section>/` with one file per page in section order and an `index.md` / `index.html`. With `combine: true` it writes a single `<Section>.<ext>` instead. Pages blocked by the guardrail are left out.

## Versioning

Microsoft Graph has no version history for OneNote pages, so the server keeps its own:

- **Before every change** made through the server (`update_page`, `append_to_page`, `delete_page`), it saves the page as it was. `delete_page` refuses to delete if that snapshot fails.
- **For every page in every section**, on demand with `snapshot_pages`, or on a schedule with `npm run snapshot`. Only pages whose modified time changed since their last version are downloaded, and identical content is never stored twice.

Each version stores the page HTML plus its images and files, so `restore_page_version` brings the page back complete, as a new page in its original section (or one you choose). The current page is left alone; delete it afterwards if the restored copy should replace it. Deleted pages can be restored too: find them with `list_page_versions` without a `pageId`.

**Retention:** versions older than `ONENOTE_VERSION_RETENTION_DAYS` (default 30) are deleted on each snapshot run and when the server starts. The newest version of a page that still exists is always kept, because it matches the page as it is now. Versions of deleted pages age out fully after 30 days.

**Where:** `~/.local/share/onenote-mcp/versions` (or `$XDG_DATA_HOME/onenote-mcp/versions`, or `ONENOTE_VERSION_DIR`). Folders are `0700` and files `0600`, but the content is not encrypted, so treat this folder like the notebook itself. Blocked sections and pages are never snapshotted.

### Scheduling snapshots

`npm run snapshot` snapshots every section the guardrail allows, prints a one-line summary and exits (`--notebook`, `--section` and `--force` narrow or widen it). It reads the same environment variables as the server, so pass your guardrail settings to it too.

**macOS (launchd)**, hourly. Save as `~/Library/LaunchAgents/com.onenote-mcp.snapshot.plist`, then run `launchctl load ~/Library/LaunchAgents/com.onenote-mcp.snapshot.plist`:

```xml
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0"><dict>
  <key>Label</key><string>com.onenote-mcp.snapshot</string>
  <key>ProgramArguments</key><array>
    <string>/usr/local/bin/node</string><string>/absolute/path/to/onenote-mcp/snapshot.js</string>
  </array>
  <key>EnvironmentVariables</key><dict>
    <key>ONENOTE_BLOCKED_SECTIONS</key><string>Credentials,0-ABC123...!12345</string>
  </dict>
  <key>StartInterval</key><integer>3600</integer>
  <key>StandardOutPath</key><string>/tmp/onenote-mcp-snapshot.log</string>
  <key>StandardErrorPath</key><string>/tmp/onenote-mcp-snapshot.log</string>
</dict></plist>
```

**Linux / macOS (cron)**, hourly:

```
0 * * * * ONENOTE_BLOCKED_SECTIONS="Credentials,0-ABC123...!12345" /usr/bin/node /absolute/path/to/onenote-mcp/snapshot.js >> ~/.onenote-mcp-snapshot.log 2>&1
```

**Windows:** create a Task Scheduler task that runs `node C:\path\to\onenote-mcp\snapshot.js` hourly.

The first run downloads every page; later runs only fetch pages that changed.

## Requirements

- Node.js 18 or newer
- A Microsoft account with OneNote (personal accounts work out of the box; see [Work or school accounts](#work-or-school-accounts))

## Installation

### 1. Get the code

```bash
git clone https://github.com/20001602056/onenote-mcp.git
cd onenote-mcp
npm install
```

### 2. Sign in to Microsoft

```bash
npm run login
```

The command prints a code and a URL (https://microsoft.com/devicelogin). Open the URL, enter the code, sign in, and approve access to OneNote. When the terminal says **Signed in**, you're done.

Where the tokens are stored:

| Platform | Location |
|---|---|
| macOS | Keychain, service `onenote-mcp` |
| Linux / Windows | `~/.config/onenote-mcp/tokens.json` (or `$XDG_CONFIG_HOME/onenote-mcp/tokens.json`), readable only by you |

Run `npm run login` again whenever the session expires.

### 3. Register the server with your MCP client

Use the absolute path to `server.js` in the folder you cloned. Print it with:

```bash
echo "$(pwd)/server.js"
```

**Claude Code**

```bash
claude mcp add --scope user onenote -- node /absolute/path/to/onenote-mcp/server.js
```

**Claude Desktop**: add this to `claude_desktop_config.json` (macOS: `~/Library/Application Support/Claude/`, Windows: `%APPDATA%\Claude\`), then restart Claude Desktop:

```json
{
  "mcpServers": {
    "onenote": {
      "command": "node",
      "args": ["/absolute/path/to/onenote-mcp/server.js"]
    }
  }
}
```

### 4. Check it works

Ask Claude: *"List my OneNote notebooks."* You can also run the read test against a section you own:

```bash
ONENOTE_TEST_NOTEBOOK="My Notebook" ONENOTE_TEST_SECTION="Quick Notes" npm run test:read
```

## Configuration

All settings are environment variables. Set them in your MCP client's server config (`env`), or with `claude mcp add -e NAME=value ...` in Claude Code.

| Variable | Default | Purpose |
|---|---|---|
| `ONENOTE_BLOCKED_SECTIONS` | *(empty)* | Sections every tool must refuse (ids, names or paths; see [Guardrail](#guardrail-blocking-and-protecting-sections-and-pages)) |
| `ONENOTE_BLOCKED_PAGES` | *(empty)* | Pages every tool must refuse (page ids, titles or paths) |
| `ONENOTE_PROTECTED_SECTIONS` | *(empty)* | Sections that can be read but not changed |
| `ONENOTE_PROTECTED_PAGES` | *(empty)* | Pages that can be read but not changed |
| `ONENOTE_AUDIT_LOG` | *(none)* | File to append a JSON line to for every change and every guardrail denial |
| `ONENOTE_READ_ONLY` | `false` | `true` hides every tool that changes OneNote or writes files |
| `ONENOTE_ALLOW_DELETE` | `false` | `true` enables `delete_page` (ignored in read-only mode) |
| `ONENOTE_DEFAULT_NOTEBOOK` | *(none)* | Notebook used when a tool call names no notebook |
| `ONENOTE_DEFAULT_SECTION` | *(none)* | Section used when a tool call names no section |
| `ONENOTE_DOWNLOAD_DIR` | `~/Downloads` | Where `download_resource` saves files when no path is given |
| `ONENOTE_EXPORT_DIR` | `ONENOTE_DOWNLOAD_DIR` | Where `export_page` / `export_section` write when no path is given |
| `ONENOTE_CHROME_PATH` | auto-detected | Chrome/Edge/Chromium/Brave executable for PDF export |
| `ONENOTE_VERSIONING` | `true` | `false` turns off snapshots and hides the version tools |
| `ONENOTE_VERSION_RETENTION_DAYS` | `30` | How long versions are kept |
| `ONENOTE_VERSION_DIR` | `~/.local/share/onenote-mcp/versions` | Where versions are stored |
| `ONENOTE_TENANT` | `consumers` | Sign-in authority: `consumers` (personal), `organizations`, `common`, or a tenant id |
| `ONENOTE_CLIENT_ID` | Graph CLI public client | Use your own app registration instead |
| `ONENOTE_TOKEN_FILE` | see above | Store tokens in this file (also on macOS, instead of the Keychain) |

### Guardrail: blocking and protecting sections and pages

Four rule lists, each a comma-separated environment variable:

| Variable | Effect |
|---|---|
| `ONENOTE_BLOCKED_SECTIONS` | Hidden from every tool: not listed, searched, read, exported, snapshotted, copied or written |
| `ONENOTE_BLOCKED_PAGES` | Same, for individual pages |
| `ONENOTE_PROTECTED_SECTIONS` | Readable, but no page can be created, edited, deleted, copied or restored into them |
| `ONENOTE_PROTECTED_PAGES` | Readable, but can't be edited, renamed or deleted |

Each entry is one of:

- **An id**: a section id (`0-ABC123...!12345`) or page id. Ids survive renames, so use them for anything sensitive.
- **A name**: section name or page title, case-insensitive. `*` matches anything: `Passwords*`, `*secret*`.
- **A path**, when the entry contains `/`: `Notebook/Section group/Section`, and for pages `Notebook/Section/Page title`. `*` matches within one level and `**` across levels. `My Notebook/Private/**` covers everything in the Private section group, however deeply nested.

Example: block the Credentials section by name and id, plus any page whose title starts with "Passwords", and make the `_Global` page read-only:

```bash
claude mcp add --scope user onenote \
  -e ONENOTE_BLOCKED_SECTIONS="Credentials,0-ABC123...!12345" \
  -e ONENOTE_BLOCKED_PAGES="Passwords*" \
  -e ONENOTE_PROTECTED_PAGES="_Global" \
  -- node /absolute/path/to/onenote-mcp/server.js
```

Get ids from `list_notebooks` (sections) or `list_pages` (pages) *before* you block them.

How it's enforced:

- Every page tool checks the page's section and the page itself before doing anything, and blocked items return "Access denied by guardrail". Writes to protected items return "Write denied by guardrail".
- Lists, searches, `export_section`, `copy_section` and snapshots leave blocked pages out.
- New names are checked too: you can't create a section, or create, rename, copy or restore a page, under a name or path that a rule blocks or protects.
- Stored versions follow the same rules. A deleted page's versions are checked against where it lived.
- With `ONENOTE_AUDIT_LOG=~/onenote-audit.log`, every change and every denial is appended as a JSON line (tool, ids and names only, never page content).

> The guardrail lives in this server (and `snapshot.js`). Your Microsoft token itself can still reach every notebook, so keep it private.

### Work or school accounts

Set `ONENOTE_TENANT=organizations` (or your tenant id) and run `npm run login` again. Some organizations require an admin to approve the Graph Command Line Tools app. If yours does, ask your admin, or register your own app and set `ONENOTE_CLIENT_ID`.

## Tests

The tests talk to your real OneNote, so point them at a notebook and section you own:

```bash
export ONENOTE_TEST_NOTEBOOK="My Notebook"
export ONENOTE_TEST_SECTION="Quick Notes"

npm run test:convert     # offline: Markdown/text conversion and guardrail rules (no sign-in needed)
npm run test:read        # lists tools, notebooks and pages in the section
npm run test:write       # creates a "Test from onenote-mcp" page (delete it afterwards)

# full test of every page tool: create, tags, attachments, edits, download, copy, export in every
# format, versions and restore. Versions and exports go to a temp folder.
# With delete enabled it removes the pages it created; otherwise delete them by hand.
ONENOTE_ALLOW_DELETE=true npm run test:features

# guardrail test: also needs a blocked section name + id in the same notebook
ONENOTE_BLOCKED_SECTIONS="Credentials,0-ABC123...!12345" npm run test:guardrail
```

## Revoke access

- Delete the tokens: on macOS, remove the `onenote-mcp` item in Keychain Access; elsewhere, delete `~/.config/onenote-mcp/tokens.json`.
- Remove the app's consent at https://account.live.com/consent/Manage (personal accounts) or https://myapps.microsoft.com (work/school).

## License

[MIT](LICENSE)

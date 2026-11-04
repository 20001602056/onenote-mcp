# onenote-mcp

A [Model Context Protocol](https://modelcontextprotocol.io) server that lets Claude (Claude Code, Claude Desktop, or any MCP client) read and write your Microsoft OneNote notebooks through Microsoft Graph.

- **No Azure app registration.** It signs in as you with Microsoft's public "Microsoft Graph Command Line Tools" client, using device-code login.
- **Tokens stay on your machine:** in the macOS Keychain, or in a `0600` file on Linux and Windows. They refresh automatically.
- **Optional guardrail:** block sections (for example, one holding passwords) from every tool.
- **Safety switches:** read-only mode, and page deletion that stays off unless you turn it on.

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
| `read_page` | Full page HTML; optionally with element ids for precise edits |
| `create_page` | New page from HTML, with tags, web images and local images/files |
| `update_page` | Edit in place: rename, replace/insert/append/prepend at any element, tick to-dos, add attachments |
| `append_to_page` | Append HTML (and attachments) to the end of a page |
| `copy_page` | Copy a page with its images and files to another section (copy + delete = move) |
| `delete_page` | Delete a page (only when `ONENOTE_ALLOW_DELETE=true`) |
| `download_resource` | Save an image or attached file from a page to disk |
| `get_operation` | Check a background copy (work/school accounts) |
| `server_settings` | Show the server's active settings |

Sections are identified by `sectionId`, or by `notebook` + `section` names (+ `sectionGroup` for a section inside a group). With `ONENOTE_DEFAULT_NOTEBOOK` and `ONENOTE_DEFAULT_SECTION` set, you can leave them out.

### Page content tips

- **Tags:** `<p data-tag="to-do">Buy milk</p>`. Other tags: `to-do:completed`, `important`, `question`, `idea`, `critical`, `contact`, `remember-for-later`, `definition`, `highlight`, and more.
- **Ticking a to-do:** give the line a `data-id` when you create it, then `update_page` with `{target: "#my-id", action: "replace", content: "<p data-tag=\"to-do:completed\" data-id=\"my-id\">Done</p>"}`.
- **Images from the web:** `<img src="https://..." />`. OneNote downloads and embeds them.
- **Local images and files:** pass `attachments: [{ path: "~/Pictures/chart.png" }]`. Reference them as `name:file1`, or leave them unreferenced and they're added at the end.
- **Renaming a page:** `update_page` with `{target: "title", action: "replace", content: "New title"}`.

### What Microsoft Graph doesn't support

These OneNote features aren't available through Graph, so no OneNote MCP server can provide them:

- OneNote **app settings** (themes, sync options, proofing, etc.)
- **Renaming or deleting** notebooks, section groups or sections (pages can be renamed and deleted)
- Setting a page's **indent level or order**, section colors, or page templates
- Ink/drawings as editable strokes, audio recording, password-protected sections

On **personal Microsoft accounts**, Graph's built-in copy isn't available. `copy_page` and `copy_section` copy content by hand instead (read → download images/files → recreate). The result is the same content, but the copy has a new created date.

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
| `ONENOTE_BLOCKED_SECTIONS` | *(empty)* | Comma-separated section names and/or ids that every tool must refuse |
| `ONENOTE_READ_ONLY` | `false` | `true` hides every tool that changes OneNote or writes files |
| `ONENOTE_ALLOW_DELETE` | `false` | `true` enables `delete_page` (ignored in read-only mode) |
| `ONENOTE_DEFAULT_NOTEBOOK` | *(none)* | Notebook used when a tool call names no notebook |
| `ONENOTE_DEFAULT_SECTION` | *(none)* | Section used when a tool call names no section |
| `ONENOTE_DOWNLOAD_DIR` | `~/Downloads` | Where `download_resource` saves files when no path is given |
| `ONENOTE_TENANT` | `consumers` | Sign-in authority: `consumers` (personal), `organizations`, `common`, or a tenant id |
| `ONENOTE_CLIENT_ID` | Graph CLI public client | Use your own app registration instead |
| `ONENOTE_TOKEN_FILE` | see above | Store tokens in this file (also on macOS, instead of the Keychain) |

### Guardrail: blocking sections

To hide a section (for example `Credentials`) from Claude:

1. Ask Claude to run `list_notebooks` and note the section's id. It looks like `0-ABC123...!12345`.
2. Set both the name and the id, so renaming the section doesn't get around the block:

   ```bash
   claude mcp add --scope user onenote \
     -e ONENOTE_BLOCKED_SECTIONS="Credentials,0-ABC123...!12345" \
     -- node /absolute/path/to/onenote-mcp/server.js
   ```

Blocked sections are hidden from `list_notebooks`, and listing, searching, reading, creating or appending in them returns "Access denied by guardrail". Page tools check each page's parent section first.

> The guardrail lives in this server. Your Microsoft token itself can still reach every notebook, so keep it private.

### Work or school accounts

Set `ONENOTE_TENANT=organizations` (or your tenant id) and run `npm run login` again. Some organizations require an admin to approve the Graph Command Line Tools app. If yours does, ask your admin, or register your own app and set `ONENOTE_CLIENT_ID`.

## Tests

The tests talk to your real OneNote, so point them at a notebook and section you own:

```bash
export ONENOTE_TEST_NOTEBOOK="My Notebook"
export ONENOTE_TEST_SECTION="Quick Notes"

npm run test:read        # lists tools, notebooks and pages in the section
npm run test:write       # creates a "Test from onenote-mcp" page (delete it afterwards)

# full test of every page tool: create, tags, attachments, edits, download, copy.
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

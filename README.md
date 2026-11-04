# onenote-mcp

A [Model Context Protocol](https://modelcontextprotocol.io) server that lets Claude (Claude Code, Claude Desktop, or any MCP client) read and write your Microsoft OneNote notebooks through Microsoft Graph.

- **No Azure app registration.** It signs in as you with Microsoft's public "Microsoft Graph Command Line Tools" client, using device-code login.
- **Tokens stay on your machine:** in the macOS Keychain, or in a `0600` file on Linux and Windows. They refresh automatically.
- **Optional guardrail:** block sections (for example, one holding passwords) from every tool.

## Tools

| Tool | What it does |
|---|---|
| `list_notebooks` | List notebooks and their sections |
| `create_section` | Create a section in a notebook (or return the existing one) |
| `list_pages` | List pages in a section, newest first |
| `search_pages` | Search page titles in a section |
| `read_page` | Read a page's content as HTML |
| `create_page` | Create a page from simple HTML |
| `append_to_page` | Append HTML to the end of a page |

Sections are identified by `sectionId`, or by `notebook` + `section` names.

## Requirements

- Node.js 18 or newer
- A Microsoft account with OneNote (personal accounts work out of the box; see [Work or school accounts](#work-or-school-accounts))

## Installation

### 1. Get the code

```bash
git clone https://github.com/<your-user>/onenote-mcp.git
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

# guardrail test: also needs a blocked section name + id in the same notebook
ONENOTE_BLOCKED_SECTIONS="Credentials,0-ABC123...!12345" npm run test:guardrail
```

## Revoke access

- Delete the tokens: on macOS, remove the `onenote-mcp` item in Keychain Access; elsewhere, delete `~/.config/onenote-mcp/tokens.json`.
- Remove the app's consent at https://account.live.com/consent/Manage (personal accounts) or https://myapps.microsoft.com (work/school).

# OneNote MCP Server

A stdio [Model Context Protocol](https://modelcontextprotocol.io/) server for Microsoft OneNote. It uses Microsoft Graph delegated permissions and Microsoft device-code authentication.

## Requirements

- Node.js 20 or newer
- A Microsoft account with OneNote access
- An MCP client such as Codex

## Install

### Option 1: Add from the desktop GUI

1. Open the ChatGPT desktop app, select **Plugins**, and open the marketplace
   settings next to **Installed**.

   ![Open the Plugins marketplace settings](https://raw.githubusercontent.com/seanGSISG/onenote-mcp/main/docs/assets/gui-install/01-open-marketplace-dialog.png)

2. Select **Add plugin marketplace** and enter:

   - **Source:** `https://github.com/seanGSISG/onenote-mcp.git`
   - **Git ref:** `main`
   - **Sparse paths:** leave blank

   ![Enter the OneNote marketplace repository](https://raw.githubusercontent.com/seanGSISG/onenote-mcp/main/docs/assets/gui-install/02-enter-repository.png)

3. Select **Add marketplace**, return to **Plugins**, and select the
   **Personal** tab. If the new marketplace does not appear, completely restart
   the desktop app and return to **Plugins**.
4. Select **Sean's OneNote Plugin** and choose **Install** beside
   **Microsoft OneNote**.

   ![Install Microsoft OneNote from the new marketplace](https://raw.githubusercontent.com/seanGSISG/onenote-mcp/main/docs/assets/gui-install/03-install-plugin.png)

5. Restart the desktop app again, open a new task, and ask:
   `Authenticate with OneNote.`

**Watch or download the full 26-second walkthrough:**
[H.264 MP4](https://raw.githubusercontent.com/seanGSISG/onenote-mcp/main/docs/assets/gui-install/add-marketplace.mp4) ·
[WebM](https://raw.githubusercontent.com/seanGSISG/onenote-mcp/main/docs/assets/gui-install/add-marketplace.webm)

> GitHub recommends H.264 for the broadest browser compatibility. The original
> WebM recording is included as an alternative.

### Option 2: Let Codex configure it

Paste this into a Codex desktop task—no clone is required:

> Read https://raw.githubusercontent.com/seanGSISG/onenote-mcp/main/AGENTS.md and follow it to install the complete OneNote plugin in the Codex desktop app.

Codex will add this repository as a plugin marketplace, install the complete
OneNote plugin, verify its MCP server, and explain when a desktop-app restart or
new task is required. Users do not need to clone the repository or edit MCP JSON
or TOML files.

### Option 3: Install for development

```powershell
git clone https://github.com/seanGSISG/onenote-mcp.git
cd onenote-mcp
npm install
npm test
```

The MCP SDK is installed from npm. A separate SDK checkout and build are no longer required.

## Advanced: MCP-only configuration

Add this to `~/.codex/config.toml`, replacing the path with the absolute path to `onenote-mcp.mjs`:

```toml
[mcp_servers.onenote]
type = "stdio"
command = "node"
args = ["C:/absolute/path/to/onenote-mcp.mjs"]
enabled = true
startup_timeout_sec = 30
```

Restart Codex after changing MCP configuration.

## Authentication

1. Call `authenticationStatus`.
2. If unauthenticated, call `authenticate`.
3. Open the returned Microsoft device-login URL and enter the returned code. Use `force: true` to replace a stale or unwanted login.
4. Call `authenticationStatus` again to confirm completion.

Authentication starts in the background, so the `authenticate` tool returns the device code immediately. The delegated access token is stored at `~/.onenote-mcp/access-token.json`. Set `ONENOTE_TOKEN_PATH` to override that location, or supply `GRAPH_ACCESS_TOKEN` in the server environment. `ONENOTE_CLIENT_ID` and `ONENOTE_TENANT_ID` override the default public client and `common` tenant.

The server requests the delegated `Notes.ReadWrite` and `User.Read` scopes.

## Tools

| Tool | Purpose |
| --- | --- |
| `authenticationStatus` | Check authentication state and expiration |
| `authenticate` | Start device-code authentication |
| `listNotebooks` | List up to `maxResults` notebooks |
| `getNotebook` | Get one notebook by exact ID |
| `listSections` | List up to `maxResults` sections, optionally in a notebook |
| `listPages` | List up to `maxResults` pages, optionally in a section |
| `getPage` | Get page metadata and optionally its HTML |
| `createPage` | Create a titled XHTML page in an exact section |
| `searchPages` | Search page titles |

All resource-specific operations use explicit IDs. The server never silently chooses the first notebook, section, or page.

## Development

```powershell
npm run check
npm test
npm audit
```

The test suite covers token migration and expiration, asynchronous device-code flow, Graph request construction and errors, XHTML title escaping, and a real MCP stdio handshake.

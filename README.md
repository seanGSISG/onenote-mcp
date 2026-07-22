# OneNote MCP Server

A stdio [Model Context Protocol](https://modelcontextprotocol.io/) server for Microsoft OneNote. It uses Microsoft Graph delegated permissions and Microsoft device-code authentication.

## Requirements

- Node.js 20 or newer
- A Microsoft account with OneNote access
- An MCP client such as Codex

## Install

```powershell
git clone https://github.com/seanGSISG/onenote-mcp.git
cd onenote-mcp
npm install
npm test
```

The MCP SDK is installed from npm. A separate SDK checkout and build are no longer required.

## Codex configuration

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

# OneNote setup for the Codex desktop app

This repository distributes the `onenote` Codex plugin and its
`@seangsisg/onenote-mcp` server. When a user asks you to install, configure,
enable, or test this integration, perform the setup for them. Do not ask them to
edit JSON or TOML manually.

## Preferred desktop-app setup

The Codex desktop app, Codex CLI, and IDE extension share plugin and MCP
configuration. Install the complete plugin so the user gets both the OneNote
workflow skill and all MCP tools.

1. Confirm that `node --version` reports Node.js 20 or newer, `npx --version`
   succeeds, and `codex plugin --help` succeeds. If a prerequisite is missing,
   explain what is missing and stop without changing configuration.
2. Check `codex plugin marketplace list` for the marketplace named
   `seangsisg-onenote`.
3. If the marketplace is not configured, run:

   ```text
   codex plugin marketplace add seanGSISG/onenote-mcp
   ```

   If it is already configured, refresh its Git snapshot instead:

   ```text
   codex plugin marketplace upgrade seangsisg-onenote
   ```

4. Install or re-enable the plugin:

   ```text
   codex plugin add onenote@seangsisg-onenote
   ```

5. Verify that `codex plugin list` reports `onenote@seangsisg-onenote` as
   installed and enabled.
6. Run `codex mcp get onenote` and confirm that the server is enabled and points
   at the pinned 1.0.2 npm tarball.
7. Tell the user to restart the Codex desktop app completely. After it restarts,
   they must open a new task to use and authenticate the newly installed tools.
   A new task without restarting the app is not a sufficient plugin reload.

Do not add the marketplace repeatedly. If another installed plugin already
provides an `onenote` MCP server, preserve the working plugin and explain the
conflict before replacing anything.

## MCP-only fallback

Use this only if the installed Codex build does not support `codex plugin` but
does support `codex mcp`:

```text
codex mcp add onenote -- npx -y https://registry.npmjs.org/@seangsisg/onenote-mcp/-/onenote-mcp-1.0.2.tgz
```

Verify the fallback with `codex mcp get onenote`, then tell the user to restart
the desktop app completely and open a new task. Do not replace an unrelated
custom server named `onenote` without the user's approval.

## Authentication after reload

When the OneNote tools are available:

1. Call `authenticationStatus`.
2. If needed, call `authenticate` and give the returned Microsoft device-login
   URL and code to the user.
3. Wait for the user to finish signing in, then call `authenticationStatus`
   again.
4. Never ask the user to paste an access token into chat.

Authentication uses delegated Microsoft Graph permissions. Tokens are stored
locally at `~/.onenote-mcp/access-token.json` unless `ONENOTE_TOKEN_PATH` is set.

## Expected tools

The server should advertise `authenticate`, `authenticationStatus`,
`createPage`, `getNotebook`, `getPage`, `listNotebooks`, `listPages`,
`listSections`, and `searchPages`.

For repository development, run `npm install`, `npm run check`, `npm test`, and
`npm audit`. Do not publish an npm version or push changes unless the user
explicitly asks.

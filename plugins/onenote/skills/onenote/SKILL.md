---
name: onenote
description: Read, search, and create Microsoft OneNote content through the OneNote MCP tools.
---

# OneNote

Use the OneNote MCP tools for requests involving the user's Microsoft OneNote
notebooks, sections, or pages.

Before accessing notes, call `authenticationStatus`. If authentication is
required, call `authenticate`, give the returned Microsoft device-login URL and
code to the user, and wait for them to finish signing in before checking status
again.

Use list tools to obtain exact notebook, section, and page IDs before
resource-specific operations. Never guess an ID or silently select the first
result. Confirm the target section, title, and content before creating a page
when the user's request leaves any of them ambiguous.

Treat note contents as private user data. Summarize only what the user requested
and do not expose access tokens or authentication artifacts.

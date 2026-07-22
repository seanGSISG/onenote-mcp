import { McpServer } from '@modelcontextprotocol/server';
import path from 'node:path';
import os from 'node:os';
import { z } from 'zod/v4';
import { AuthManager } from './auth-manager.mjs';
import { OneNoteClient } from './onenote-client.mjs';

const DEFAULT_TOKEN_PATH = process.env.ONENOTE_TOKEN_PATH
  || path.join(os.homedir(), '.onenote-mcp', 'access-token.json');
const MaxResults = z.number().int().min(1).max(100).default(50);

export function createOneNoteServer({ authManager, oneNoteClient } = {}) {
  const auth = authManager || new AuthManager({ tokenFilePath: DEFAULT_TOKEN_PATH });
  const client = oneNoteClient || new OneNoteClient({ authManager: auth });
  const server = new McpServer(
    { name: 'onenote', version: '1.0.0' },
    {
      instructions: 'Call authenticationStatus before accessing OneNote. If needed, call authenticate and give the returned device-login URL and code to the user. Use list tools to discover resource IDs before get or create operations.',
    },
  );

  const ready = auth.initialize();
  const tool = (name, config, handler) => {
    server.registerTool(name, config, async input => {
      try {
        await ready;
        const result = await handler(input);
        return textResult(result);
      } catch (error) {
        console.error(`${name} failed:`, error);
        return {
          isError: true,
          content: [{ type: 'text', text: error instanceof Error ? error.message : String(error) }],
        };
      }
    });
  };

  tool('authenticationStatus', {
    description: 'Check whether OneNote is authenticated and whether device-code authentication is pending.',
    inputSchema: z.object({}),
    annotations: { readOnlyHint: true, idempotentHint: true },
  }, () => auth.getStatus());

  tool('authenticate', {
    description: 'Start Microsoft device-code authentication. Immediately returns the login URL and user code; authentication completes in the background.',
    inputSchema: z.object({
      force: z.boolean().default(false).describe('Start a new sign-in even if a stored token appears valid'),
    }),
    annotations: { readOnlyHint: true },
  }, input => auth.beginAuthentication(input));

  tool('listNotebooks', {
    description: 'List OneNote notebooks available to the signed-in user.',
    inputSchema: z.object({ maxResults: MaxResults.optional() }),
    annotations: { readOnlyHint: true, idempotentHint: true },
  }, input => client.listNotebooks(input));

  tool('getNotebook', {
    description: 'Get a OneNote notebook by its exact resource ID.',
    inputSchema: z.object({ notebookId: z.string().min(1) }),
    annotations: { readOnlyHint: true, idempotentHint: true },
  }, ({ notebookId }) => client.getNotebook(notebookId));

  tool('listSections', {
    description: 'List OneNote sections, optionally limited to one notebook.',
    inputSchema: z.object({ notebookId: z.string().min(1).optional(), maxResults: MaxResults.optional() }),
    annotations: { readOnlyHint: true, idempotentHint: true },
  }, input => client.listSections(input));

  tool('listPages', {
    description: 'List OneNote page metadata, optionally limited to one section.',
    inputSchema: z.object({ sectionId: z.string().min(1).optional(), maxResults: MaxResults.optional() }),
    annotations: { readOnlyHint: true, idempotentHint: true },
  }, input => client.listPages(input));

  tool('getPage', {
    description: 'Get a OneNote page by its exact resource ID, including HTML content by default.',
    inputSchema: z.object({ pageId: z.string().min(1), includeContent: z.boolean().default(true) }),
    annotations: { readOnlyHint: true, idempotentHint: true },
  }, ({ pageId, includeContent }) => client.getPage(pageId, includeContent));

  tool('createPage', {
    description: 'Create a OneNote page in an exact section using XHTML content.',
    inputSchema: z.object({
      sectionId: z.string().min(1),
      title: z.string().min(1).max(255),
      contentHtml: z.string().min(1).describe('HTML fragment for the page body'),
    }),
    annotations: { destructiveHint: false, idempotentHint: false },
  }, input => client.createPage(input));

  tool('searchPages', {
    description: 'Search page titles across OneNote. Returns page metadata; use getPage to read a result.',
    inputSchema: z.object({ query: z.string().min(1), maxResults: MaxResults.optional() }),
    annotations: { readOnlyHint: true, idempotentHint: true },
  }, input => client.searchPages(input));

  return server;
}

function textResult(value) {
  return {
    content: [{ type: 'text', text: typeof value === 'string' ? value : JSON.stringify(value, null, 2) }],
  };
}

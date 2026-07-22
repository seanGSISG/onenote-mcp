#!/usr/bin/env node

import 'dotenv/config';
import { StdioServerTransport } from '@modelcontextprotocol/server/stdio';
import { createOneNoteServer } from './src/server.mjs';

const server = createOneNoteServer();

async function shutdown() {
  await server.close();
  process.exit(0);
}

process.once('SIGINT', shutdown);
process.once('SIGTERM', shutdown);

try {
  await server.connect(new StdioServerTransport());
  console.error('OneNote MCP server started.');
} catch (error) {
  console.error('Failed to start OneNote MCP server:', error);
  process.exit(1);
}

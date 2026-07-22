import assert from 'node:assert/strict';
import test from 'node:test';
import { Client } from '@modelcontextprotocol/client';
import { StdioClientTransport } from '@modelcontextprotocol/client/stdio';

test('stdio server completes an MCP handshake and advertises validated tools', async t => {
  const client = new Client({ name: 'onenote-test-client', version: '1.0.0' });
  const transport = new StdioClientTransport({
    command: process.execPath,
    args: ['onenote-mcp.mjs'],
    cwd: process.cwd(),
    stderr: 'pipe',
  });
  t.after(() => client.close());
  await client.connect(transport);

  const { tools } = await client.listTools();
  assert.deepEqual(
    tools.map(tool => tool.name).sort(),
    [
      'authenticate',
      'authenticationStatus',
      'createPage',
      'getNotebook',
      'getPage',
      'listNotebooks',
      'listPages',
      'listSections',
      'searchPages',
    ].sort(),
  );
  const getPage = tools.find(tool => tool.name === 'getPage');
  assert.deepEqual(getPage.inputSchema.required, ['pageId']);

  const status = await client.callTool({ name: 'authenticationStatus', arguments: {} });
  assert.equal(status.isError, undefined);
  assert.match(status.content[0].text, /"authenticated"/);
});

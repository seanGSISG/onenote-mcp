import assert from 'node:assert/strict';
import test from 'node:test';
import { buildPageHtml, GraphRequestError, OneNoteClient } from '../src/onenote-client.mjs';

const authManager = { getAccessToken: async () => 'secret-token' };

test('buildPageHtml escapes the title and preserves the body fragment', () => {
  const html = buildPageHtml('R&D <plan>', '<p>Trusted body</p>');
  assert.match(html, /<title>R&amp;D &lt;plan&gt;<\/title>/);
  assert.match(html, /<body><p>Trusted body<\/p><\/body>/);
});

test('listPages targets the requested section and caps the result count', async () => {
  let request;
  const client = new OneNoteClient({
    authManager,
    fetchImpl: async (url, options) => {
      request = { url, options };
      return new Response(JSON.stringify({ value: [{ id: 'page-1' }] }), {
        status: 200,
        headers: { 'Content-Type': 'application/json' },
      });
    },
  });
  const pages = await client.listPages({ sectionId: 'section/id', maxResults: 12 });
  assert.deepEqual(pages, [{ id: 'page-1' }]);
  assert.equal(request.url, 'https://graph.microsoft.com/v1.0/me/onenote/sections/section%2Fid/pages?$top=12');
  assert.equal(request.options.headers.Authorization, 'Bearer secret-token');
});

test('Graph errors include the status and a bounded response body', async () => {
  const client = new OneNoteClient({
    authManager,
    fetchImpl: async () => new Response('permission denied', { status: 403, statusText: 'Forbidden' }),
  });
  await assert.rejects(() => client.listNotebooks(), error => {
    assert.ok(error instanceof GraphRequestError);
    assert.equal(error.status, 403);
    assert.match(error.message, /permission denied/);
    return true;
  });
});

test('searchPages filters in Microsoft Graph before applying the result limit', async () => {
  let requestedUrl;
  const client = new OneNoteClient({
    authManager,
    fetchImpl: async url => {
      requestedUrl = url;
      return new Response(JSON.stringify({ value: [] }), { status: 200 });
    },
  });
  await client.searchPages({ query: "Sam's plan", maxResults: 7 });
  const url = new URL(requestedUrl);
  assert.equal(url.searchParams.get('$filter'), "contains(title,'Sam''s plan')");
  assert.equal(url.searchParams.get('$top'), '7');
});

test('a 401 invalidates the cached token so authentication can restart', async () => {
  let invalidated = false;
  const client = new OneNoteClient({
    authManager: {
      getAccessToken: async () => 'expired',
      invalidateToken: async () => { invalidated = true; },
    },
    fetchImpl: async () => new Response('expired', { status: 401, statusText: 'Unauthorized' }),
  });
  await assert.rejects(() => client.listNotebooks(), GraphRequestError);
  assert.equal(invalidated, true);
});

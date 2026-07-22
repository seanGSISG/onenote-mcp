import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import test from 'node:test';

const files = {
  agents: new URL('../AGENTS.md', import.meta.url),
  package: new URL('../package.json', import.meta.url),
  pluginConfig: new URL('../plugins/onenote/.mcp.json', import.meta.url),
  pluginManifest: new URL('../plugins/onenote/.codex-plugin/plugin.json', import.meta.url),
};

test('desktop onboarding pins match the package and plugin release', async () => {
  const packageManifest = JSON.parse(await readFile(files.package, 'utf8'));
  const pluginManifest = JSON.parse(await readFile(files.pluginManifest, 'utf8'));
  const expectedVersion = packageManifest.version;
  const expectedTarball = `https://registry.npmjs.org/@seangsisg/onenote-mcp/-/onenote-mcp-${expectedVersion}.tgz`;

  assert.equal(pluginManifest.version, expectedVersion);

  for (const name of ['agents', 'pluginConfig']) {
    const content = await readFile(files[name], 'utf8');
    assert.match(content, new RegExp(escapeRegExp(expectedTarball)), `${name} must pin ${expectedTarball}`);
  }
});

test('AGENTS.md uses the marketplace manifest name for detection and install', async () => {
  const agents = await readFile(files.agents, 'utf8');
  assert.match(agents, /marketplace named\s+`seangsisg-onenote`/);
  assert.match(agents, /onenote@seangsisg-onenote/);
  assert.match(agents, /marketplace upgrade seangsisg-onenote/);
  assert.match(agents, /restart the Codex desktop app completely/);
});

function escapeRegExp(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

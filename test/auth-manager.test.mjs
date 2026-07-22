import assert from 'node:assert/strict';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { AuthManager, AuthenticationRequiredError } from '../src/auth-manager.mjs';

async function fixture(t) {
  const directory = await mkdtemp(path.join(os.tmpdir(), 'onenote-auth-'));
  t.after(() => rm(directory, { recursive: true, force: true }));
  return path.join(directory, 'token.json');
}

test('derives expiry when loading a legacy raw JWT token file', async t => {
  const tokenFilePath = await fixture(t);
  const token = jwtWithExpiry(Date.now() + 3_600_000);
  await writeFile(tokenFilePath, token);
  const auth = new AuthManager({ tokenFilePath, environmentToken: null });
  await auth.initialize();
  assert.equal(await auth.getAccessToken(), token);
  assert.equal(auth.getStatus().authenticated, true);
});

test('treats an opaque legacy token with no expiry as stale', async t => {
  const tokenFilePath = await fixture(t);
  await writeFile(tokenFilePath, 'opaque-legacy-token');
  const auth = new AuthManager({ tokenFilePath, environmentToken: null });
  await auth.initialize();
  await assert.rejects(() => auth.getAccessToken(), AuthenticationRequiredError);
});

test('rejects expired persisted tokens', async t => {
  const tokenFilePath = await fixture(t);
  await writeFile(tokenFilePath, JSON.stringify({ token: 'expired', expiresOnTimestamp: Date.now() - 1 }));
  const auth = new AuthManager({ tokenFilePath, environmentToken: null });
  await auth.initialize();
  await assert.rejects(() => auth.getAccessToken(), AuthenticationRequiredError);
});

test('returns the device code before authentication completes and persists the result', async t => {
  const tokenFilePath = await fixture(t);
  let completeAuthentication;
  const credentialFactory = ({ userPromptCallback }) => ({
    getToken: () => new Promise(resolve => {
      completeAuthentication = resolve;
      userPromptCallback({
        message: 'Open the device login page and enter ABC-123.',
        verificationUri: 'https://microsoft.com/devicelogin',
        userCode: 'ABC-123',
      });
    }),
  });
  const auth = new AuthManager({ tokenFilePath, environmentToken: null, credentialFactory });
  await auth.initialize();

  const prompt = await auth.beginAuthentication();
  assert.equal(prompt.status, 'pending');
  assert.equal(prompt.userCode, 'ABC-123');
  assert.equal(auth.getStatus().authenticationPending, true);

  const completion = auth.authentication.completion;
  completeAuthentication({ token: 'fresh-token', expiresOnTimestamp: Date.now() + 3_600_000 });
  await completion;
  assert.equal(await auth.getAccessToken(), 'fresh-token');
  const saved = JSON.parse(await readFile(tokenFilePath, 'utf8'));
  assert.equal(saved.token, 'fresh-token');
});

test('forced authentication replaces a token that appears valid', async t => {
  const tokenFilePath = await fixture(t);
  await writeFile(tokenFilePath, jwtWithExpiry(Date.now() + 3_600_000));
  const credentialFactory = ({ userPromptCallback }) => ({
    getToken: () => new Promise(() => {
      userPromptCallback({ message: 'Sign in again.', verificationUri: 'https://example.test', userCode: 'NEW' });
    }),
  });
  const auth = new AuthManager({ tokenFilePath, environmentToken: null, credentialFactory });
  await auth.initialize();
  const prompt = await auth.beginAuthentication({ force: true });
  assert.equal(prompt.userCode, 'NEW');
  assert.equal(auth.getStatus().authenticated, false);
});

function jwtWithExpiry(timestamp) {
  const encode = value => Buffer.from(JSON.stringify(value)).toString('base64url');
  return `${encode({ alg: 'none' })}.${encode({ exp: Math.floor(timestamp / 1000) })}.signature`;
}

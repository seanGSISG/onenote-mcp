import { DeviceCodeCredential } from '@azure/identity';
import { chmod, mkdir, readFile, rename, unlink, writeFile } from 'node:fs/promises';
import path from 'node:path';

const DEFAULT_CLIENT_ID = '14d82eec-204b-4c2f-b7e8-296a70dab67e';
const DEFAULT_SCOPES = ['Notes.ReadWrite', 'User.Read'];
const EXPIRY_SKEW_MS = 60_000;

export class AuthenticationRequiredError extends Error {
  constructor(message = 'OneNote is not authenticated. Call authenticate first.') {
    super(message);
    this.name = 'AuthenticationRequiredError';
  }
}

export class AuthManager {
  constructor({
    tokenFilePath,
    environmentToken = process.env.GRAPH_ACCESS_TOKEN,
    clientId = process.env.ONENOTE_CLIENT_ID || DEFAULT_CLIENT_ID,
    tenantId = process.env.ONENOTE_TENANT_ID || 'common',
    scopes = DEFAULT_SCOPES,
    credentialFactory = options => new DeviceCodeCredential(options),
  }) {
    this.tokenFilePath = tokenFilePath;
    this.environmentToken = environmentToken?.trim() || null;
    this.clientId = clientId;
    this.tenantId = tenantId;
    this.scopes = scopes;
    this.credentialFactory = credentialFactory;
    this.tokenRecord = null;
    this.credential = null;
    this.authentication = null;
  }

  async initialize() {
    if (this.environmentToken) {
      this.tokenRecord = { token: this.environmentToken, source: 'environment' };
      return;
    }

    try {
      const raw = (await readFile(this.tokenFilePath, 'utf8')).trim();
      if (!raw) return;
      try {
        const parsed = JSON.parse(raw);
        this.tokenRecord = normalizeTokenRecord(typeof parsed === 'string' ? { token: parsed } : parsed);
      } catch {
        this.tokenRecord = normalizeTokenRecord({ token: raw });
      }
    } catch (error) {
      if (error.code !== 'ENOENT') throw error;
    }
  }

  getStatus() {
    const expiresOn = this.tokenRecord?.expiresOnTimestamp
      ? new Date(this.tokenRecord.expiresOnTimestamp).toISOString()
      : null;
    return {
      authenticated: Boolean(this.tokenRecord?.token) && !this.#isExpired(this.tokenRecord),
      authenticationPending: Boolean(this.authentication),
      source: this.tokenRecord?.source || (this.tokenRecord?.token ? 'token-file' : null),
      expiresOn,
    };
  }

  async beginAuthentication({ force = false } = {}) {
    if (!force && this.getStatus().authenticated) {
      return { status: 'authenticated', message: 'OneNote is already authenticated.' };
    }

    if (this.authentication) return this.authentication.prompt;
    if (force) await this.invalidateToken();

    let resolvePrompt;
    let rejectPrompt;
    const prompt = new Promise((resolve, reject) => {
      resolvePrompt = resolve;
      rejectPrompt = reject;
    });

    this.credential = this.credentialFactory({
      clientId: this.clientId,
      tenantId: this.tenantId,
      userPromptCallback: info => {
        resolvePrompt({
          status: 'pending',
          message: info.message,
          verificationUri: info.verificationUri,
          userCode: info.userCode,
        });
      },
    });

    const completion = this.credential.getToken(this.scopes)
      .then(async token => {
        await this.#persistToken(token.token, token.expiresOnTimestamp);
        return token;
      })
      .catch(error => {
        rejectPrompt(error);
        throw error;
      })
      .finally(() => {
        this.authentication = null;
      });

    // Prevent a rejection after the prompt has been returned from becoming unhandled.
    completion.catch(error => console.error('OneNote authentication failed:', error.message));
    this.authentication = { prompt, completion };
    return prompt;
  }

  async saveAccessToken(token, expiresOnTimestamp) {
    const normalized = token?.trim();
    if (!normalized) throw new TypeError('Access token must not be empty.');
    this.credential = null;
    await this.#persistToken(normalized, expiresOnTimestamp);
  }

  async invalidateToken() {
    this.tokenRecord = null;
    this.credential = null;
    try {
      await unlink(this.tokenFilePath);
    } catch (error) {
      if (error.code !== 'ENOENT') throw error;
    }
  }

  async getAccessToken() {
    if (this.credential && this.tokenRecord?.token && this.#isExpired(this.tokenRecord)) {
      const refreshed = await this.credential.getToken(this.scopes);
      if (refreshed.token !== this.tokenRecord.token) {
        await this.#persistToken(refreshed.token, refreshed.expiresOnTimestamp);
      }
    }

    if (!this.tokenRecord?.token || this.#isExpired(this.tokenRecord)) {
      throw new AuthenticationRequiredError();
    }
    return this.tokenRecord.token;
  }

  #isExpired(record) {
    return Number.isFinite(record?.expiresOnTimestamp)
      && record.expiresOnTimestamp <= Date.now() + EXPIRY_SKEW_MS;
  }

  async #persistToken(token, expiresOnTimestamp) {
    this.tokenRecord = { token, expiresOnTimestamp, source: 'token-file' };
    const temporaryPath = `${this.tokenFilePath}.tmp`;
    await mkdir(path.dirname(this.tokenFilePath), { recursive: true });
    await writeFile(temporaryPath, JSON.stringify(this.tokenRecord), { encoding: 'utf8', mode: 0o600 });
    await rename(temporaryPath, this.tokenFilePath);
    try {
      await chmod(this.tokenFilePath, 0o600);
    } catch (error) {
      if (process.platform !== 'win32') throw error;
    }
  }
}

function normalizeTokenRecord(record) {
  if (!record || typeof record.token !== 'string' || !record.token.trim()) return null;
  const token = record.token.trim();
  const expiresOnTimestamp = record.expiresOnTimestamp || jwtExpiry(token) || 0;
  return { token, expiresOnTimestamp, source: record.source || 'token-file' };
}

function jwtExpiry(token) {
  try {
    const payload = token.split('.')[1];
    if (!payload) return null;
    const parsed = JSON.parse(Buffer.from(payload, 'base64url').toString('utf8'));
    return Number.isFinite(parsed.exp) ? parsed.exp * 1000 : null;
  } catch {
    return null;
  }
}

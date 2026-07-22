const GRAPH_ROOT = 'https://graph.microsoft.com/v1.0/me/onenote';

export class GraphRequestError extends Error {
  constructor(response, body) {
    super(`Microsoft Graph returned ${response.status} ${response.statusText}${body ? `: ${body}` : ''}`);
    this.name = 'GraphRequestError';
    this.status = response.status;
  }
}

export class OneNoteClient {
  constructor({ authManager, fetchImpl = globalThis.fetch }) {
    this.authManager = authManager;
    this.fetch = fetchImpl;
  }

  async listNotebooks({ maxResults = 50 } = {}) {
    return this.#collection('/notebooks', maxResults);
  }

  async getNotebook(notebookId) {
    return this.#request(`/notebooks/${encodeURIComponent(notebookId)}`);
  }

  async listSections({ notebookId, maxResults = 50 } = {}) {
    const path = notebookId
      ? `/notebooks/${encodeURIComponent(notebookId)}/sections`
      : '/sections';
    return this.#collection(path, maxResults);
  }

  async listPages({ sectionId, maxResults = 50 } = {}) {
    const path = sectionId
      ? `/sections/${encodeURIComponent(sectionId)}/pages`
      : '/pages';
    return this.#collection(path, maxResults);
  }

  async getPage(pageId, includeContent = true) {
    const encodedId = encodeURIComponent(pageId);
    const page = await this.#request(`/pages/${encodedId}`);
    if (includeContent) page.content = await this.#request(`/pages/${encodedId}/content`, { responseType: 'text' });
    return page;
  }

  async createPage({ sectionId, title, contentHtml }) {
    const document = buildPageHtml(title, contentHtml);
    return this.#request(`/sections/${encodeURIComponent(sectionId)}/pages`, {
      method: 'POST',
      body: document,
      headers: { 'Content-Type': 'application/xhtml+xml' },
    });
  }

  async searchPages({ query, maxResults = 50 }) {
    const escapedQuery = query.trim().replaceAll("'", "''");
    const parameters = new URLSearchParams({ '$filter': `contains(title,'${escapedQuery}')` });
    return this.#collection(`/pages?${parameters}`, maxResults);
  }

  async #collection(path, maxResults) {
    const separator = path.includes('?') ? '&' : '?';
    const response = await this.#request(`${path}${separator}$top=${maxResults}`);
    return response.value || [];
  }

  async #request(path, { method = 'GET', body, headers = {}, responseType = 'json' } = {}) {
    const token = await this.authManager.getAccessToken();
    const response = await this.fetch(`${GRAPH_ROOT}${path}`, {
      method,
      body,
      headers: { Authorization: `Bearer ${token}`, ...headers },
    });
    if (!response.ok) {
      const errorBody = (await response.text()).slice(0, 1000);
      if (response.status === 401) await this.authManager.invalidateToken?.();
      throw new GraphRequestError(response, errorBody);
    }
    if (responseType === 'text') return response.text();
    if (response.status === 204) return null;
    return response.json();
  }
}

export function buildPageHtml(title, contentHtml) {
  return `<!DOCTYPE html><html><head><title>${escapeHtml(title)}</title></head><body>${contentHtml}</body></html>`;
}

function escapeHtml(value) {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');
}

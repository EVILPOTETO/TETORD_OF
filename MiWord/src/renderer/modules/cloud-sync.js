import { AppModule } from './app-module.js';

const DEFAULT_BASE_URL = 'http://127.0.0.1:8787';
const TOKEN_KEY = 'tetord.cloud.jwt';
const MAX_DOC_BYTES = 2 * 1024 * 1024;

/** Optional REST synchronization for .tetord documents. Local/offline mode never depends on this module. */
export class CloudSyncModule extends AppModule {
  constructor(ctx = {}) {
    super(ctx);
    const config = window.TETORD_CLOUD || {};
    this.baseUrl = String(config.baseUrl || DEFAULT_BASE_URL).replace(/\/+$/, '');
    this.token = String(config.token || localStorage.getItem(TOKEN_KEY) || '');
    this.enabled = config.enabled !== false;
    this.status = 'offline';
  }

  setToken(token, persist = true) {
    this.token = String(token || '');
    if (persist) {
      if (this.token) localStorage.setItem(TOKEN_KEY, this.token);
      else localStorage.removeItem(TOKEN_KEY);
    }
    return this.token;
  }

  clearToken() { this.setToken('', true); this.status = 'offline'; }

  async request(path, options = {}) {
    if (!this.enabled) throw new Error('Cloud Sync está desactivado.');
    const headers = { 'content-type': 'application/json', ...(options.headers || {}) };
    if (this.token) headers.authorization = `Bearer ${this.token}`;
    const response = await fetch(`${this.baseUrl}${path}`, { ...options, headers });
    let body = null; try { body = await response.json(); } catch (_) {}
    if (!response.ok) throw new Error(body?.error || `HTTP ${response.status}`);
    return body;
  }

  async register(username, password, nick = username) {
    const body = await this.request('/auth/register', { method: 'POST', body: JSON.stringify({ username, password, nick }) });
    this.setToken(body.token); this.status = 'online'; return body.user;
  }

  async login(username, password) {
    const body = await this.request('/auth/login', { method: 'POST', body: JSON.stringify({ username, password }) });
    this.setToken(body.token); this.status = 'online'; return body.user;
  }

  async me() { const body = await this.request('/auth/me'); this.status = 'online'; return body.user; }

  async refreshToken() { const body = await this.request('/auth/refresh', { method: 'POST' }); this.setToken(body.token); this.status = 'online'; return body.user; }
  logout() { this.clearToken(); }

  async listRemoteDocuments() { const body = await this.request('/sync/documents'); return body.documents || []; }

  async getRemoteDocument(name) { return this.request(`/sync/documents/${encodeURIComponent(name)}`); }

  async uploadDocument(name, document, updatedAt = Date.now()) {
    const encoded = JSON.stringify(document);
    if (new Blob([encoded]).size > MAX_DOC_BYTES) throw new Error('El documento supera el límite de 2 MB.');
    return this.request(`/sync/documents/${encodeURIComponent(name)}`, { method: 'PUT', body: JSON.stringify({ document, updatedAt }) });
  }

  async syncDocument(relativePath, strategy = 'timestamp') {
    if (!this.storage?.desktop) throw new Error('Cloud Sync de archivos requiere TETORD Desktop para leer el documento local.');
    const local = await this.storage.readJSON(relativePath, null);
    if (!local?.schema || !/^tetord\.(writer|calc|slides)$/.test(local.schema)) throw new Error('No es un documento .tetord válido.');
    const stat = await this.storage.stat(relativePath);
    const localUpdatedAt = Number(stat?.mtimeMs) || Date.now();
    const name = String(relativePath).split('/').pop();
    let remote = null;
    try { remote = await this.getRemoteDocument(name); } catch (error) { if (!/404/.test(error.message)) throw error; }
    if (!remote) return this.uploadDocument(name, local, localUpdatedAt);
    const remoteUpdatedAt = Number(remote.updatedAt) || 0;
    if (Math.abs(remoteUpdatedAt - localUpdatedAt) < 1000 || remoteUpdatedAt === localUpdatedAt) return { action: 'unchanged', updatedAt: remoteUpdatedAt };
    if (strategy === 'remote-wins' || (strategy === 'timestamp' && remoteUpdatedAt > localUpdatedAt)) {
      const result = await this.storage.writeJSON(relativePath, remote.document);
      if (!result?.ok) throw new Error(result?.error || 'No se pudo restaurar la versión remota.');
      return { action: 'downloaded', updatedAt: remoteUpdatedAt };
    }
    try {
      const result = await this.uploadDocument(name, local, localUpdatedAt);
      return { action: 'uploaded', updatedAt: result.updatedAt };
    } catch (error) {
      if (/409/.test(error.message)) return { action: 'conflict', error: error.message };
      throw error;
    }
  }
}

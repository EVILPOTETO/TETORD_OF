/** TETORD disk storage facade. Electron uses the secure preload IPC API; browser mode falls back to localStorage. */
export const storage = {
  get desktop() { return Boolean(window.TETORD_DESKTOP?.storage); },
  async readJSON(path, fallback = null) {
    if (this.desktop) {
      const r = await window.TETORD_DESKTOP.storage.readJSON(path);
      return r?.ok ? r.value : fallback;
    }
    try { const raw = localStorage.getItem(`tetord.disk.${path}`); return raw == null ? fallback : JSON.parse(raw); } catch (_) { return fallback; }
  },
  async writeJSON(path, value) {
    if (this.desktop) return window.TETORD_DESKTOP.storage.writeJSON(path, value);
    try { localStorage.setItem(`tetord.disk.${path}`, JSON.stringify(value)); return { ok: true, path }; } catch (error) { return { ok: false, error: String(error) }; }
  },
  async remove(path) {
    if (this.desktop) return window.TETORD_DESKTOP.storage.remove(path);
    try { localStorage.removeItem(`tetord.disk.${path}`); return { ok: true }; } catch (error) { return { ok: false, error: String(error) }; }
  },
  async mkdir(path) {
    if (this.desktop) return window.TETORD_DESKTOP.storage.mkdir(path);
    return { ok: true, path };
  },
  async writeBytes(path, base64, mime = 'application/octet-stream') {
    if (this.desktop) return window.TETORD_DESKTOP.storage.writeBytes(path, base64, mime);
    try { localStorage.setItem(`tetord.asset.${path}`, JSON.stringify({ base64, mime })); return { ok: true, path }; } catch (error) { return { ok: false, error: String(error) }; }
  },
  async readDataURL(path) {
    if (!path) return '';
    if (this.desktop) {
      const r = await window.TETORD_DESKTOP.storage.readDataURL(path);
      return r?.ok ? r.dataUrl : '';
    }
    try { const v = JSON.parse(localStorage.getItem(`tetord.asset.${path}`) || 'null'); return v ? `data:${v.mime};base64,${v.base64}` : ''; } catch (_) { return ''; }
  },
  async list(path = '', recursive = true) {
    if (this.desktop) return window.TETORD_DESKTOP.storage.list(path, recursive);
    return { ok: true, entries: [] };
  },
  async stat(path) {
    if (this.desktop) return window.TETORD_DESKTOP.storage.stat(path);
    return { ok: false, missing: true };
  },
  async rename(from, to) {
    if (this.desktop) return window.TETORD_DESKTOP.storage.rename(from, to);
    return { ok: false, error: 'Renombrar requiere la versión de escritorio.' };
  },
  async copy(from, to) {
    if (this.desktop) return window.TETORD_DESKTOP.storage.copy(from, to);
    return { ok: false, error: 'Duplicar requiere la versión de escritorio.' };
  },
  async chooseFile(options = {}) {
    if (this.desktop) return window.TETORD_DESKTOP.storage.chooseFile(options);
    return { canceled: true };
  }
};

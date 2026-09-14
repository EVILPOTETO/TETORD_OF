import { AppModule } from './app-module.js';

const ROOT = '';
const DOC_EXT = '.tetord';
const KIND_META = { writer: 'Writer', calc: 'Calc', slides: 'Slides' };
const ICON = { writer: 'W', calc: 'C', slides: 'S', folder: '▰' };

/**
 * Native Files service. The renderer never receives arbitrary filesystem paths;
 * it works with relative TETORD paths exposed by the preload IPC boundary.
 */
export class FilesModule extends AppModule {
  constructor(ctx) {
    super(ctx);
    this.entries = [];
    this.currentFolder = '';
    this.search = '';
    this.filter = 'all';
    this.sort = 'updated';
    this.favorites = new Set();
    this.loaded = false;
  }

  async init() {
    await this.storage.mkdir(ROOT);
    const settings = await this.storage.readJSON('files/settings.json', { favorites: [] });
    this.favorites = new Set(Array.isArray(settings?.favorites) ? settings.favorites : []);
    await this.refresh();
    return this;
  }

  async refresh() {
    const result = await this.storage.list(ROOT, true);
    this.entries = result?.ok ? result.entries : [];
    const docs = this.entries.filter(e => e.type === 'file' && /\.tetord$/i.test(e.name));
    await Promise.all(docs.map(async e => { const r = await this.storage.readJSON(e.relativePath, null); e.moduleType = r?.schema === 'tetord.writer' ? 'writer' : r?.schema === 'tetord.calc' ? 'calc' : r?.schema === 'tetord.slides' ? 'slides' : null; e.documentId = r?.id || null; }));
    this.loaded = true;
    return this.entries;
  }

  async saveSettings() {
    return this.storage.writeJSON('files/settings.json', { version: 1, favorites: [...this.favorites] });
  }

  normalizeName(name, type = 'writer') {
    const clean = String(name || 'Sin título').replace(/[\\/:*?"<>|]+/g, '-').trim().slice(0, 120) || 'Sin título';
    return clean.toLowerCase().endsWith(DOC_EXT) ? clean : `${clean}${DOC_EXT}`;
  }

  typeFromPath(relativePath) {
    const p = String(relativePath || '').replace(/\\/g, '/');
    const first = p.split('/')[0].toLowerCase();
    if (KIND_META[first]) return first;
    return 'folder';
  }

  records() {
    return this.entries
      .filter(e => e.type === 'file' && /\.tetord$/i.test(e.name))
      .map(e => {
        const type = e.moduleType || this.typeFromPath(e.relativePath);
        return {
          id: e.relativePath,
          sourceId: e.relativePath,
          path: e.relativePath,
          name: e.name.replace(/\.tetord$/i, ''),
          fileName: e.name,
          type,
          label: KIND_META[type] || 'Archivo',
          format: 'TETORD',
          size: Number(e.size) || 0,
          updatedAt: e.mtimeMs ? new Date(e.mtimeMs).toISOString() : null,
          favorite: this.favorites.has(e.relativePath),
          action: () => this.openRecord({ path: e.relativePath, type, name: e.name.replace(/\.tetord$/i, '') })
        };
      });
  }

  folderRecords() {
    const prefix = this.currentFolder ? `${this.currentFolder}/` : '';
    return this.entries
      .filter(e => e.type === 'directory' && e.relativePath.startsWith(prefix))
      .map(e => {
        const rest = e.relativePath.slice(prefix.length);
        if (!rest || rest.includes('/')) return null;
        return { id: e.relativePath, path: e.relativePath, name: e.name, type: 'folder', isFolder: true, label: 'Carpeta', size: 0, updatedAt: e.mtimeMs ? new Date(e.mtimeMs).toISOString() : null, favorite: false };
      }).filter(Boolean);
  }

  visibleRecords() {
    const prefix = this.currentFolder ? `${this.currentFolder}/` : '';
    let rows = this.records().filter(r => {
      const rel = r.path.startsWith(prefix) ? r.path.slice(prefix.length) : '';
      return rel && !rel.includes('/');
    });
    const folders = this.folderRecords();
    const q = this.search.trim().toLocaleLowerCase();
    rows = rows.filter(r => (!q || r.name.toLocaleLowerCase().includes(q)) && (this.filter === 'all' || this.filter === r.type || (this.filter === 'favorites' && r.favorite)));
    let out = this.filter === 'folder' ? folders : this.filter === 'all' ? [...folders, ...rows] : rows;
    if (q) out = out.filter(r => r.name.toLocaleLowerCase().includes(q));
    if (this.sort === 'name') out.sort((a,b) => a.name.localeCompare(b.name, 'es'));
    else if (this.sort === 'type') out.sort((a,b) => String(a.type).localeCompare(String(b.type)) || a.name.localeCompare(b.name, 'es'));
    else out.sort((a,b) => String(b.updatedAt || '').localeCompare(String(a.updatedAt || '')));
    return out;
  }

  async createFolder(name = null) {
    const value = name ?? window.prompt('Nombre de la carpeta', 'Nueva carpeta');
    if (!value?.trim()) return false;
    const clean = value.trim().replace(/[\\/:*?"<>|]+/g, '-').slice(0, 60);
    const path = this.currentFolder ? `${this.currentFolder}/${clean}` : clean;
    const result = await this.storage.mkdir(path);
    if (!result?.ok) this.announce(result.error || 'No se pudo crear la carpeta.');
    await this.refresh(); this.render();
    return !!result?.ok;
  }

  async createDocument(type, name = null) {
    if (!KIND_META[type]) return false;
    const base = name || ({ writer: 'Documento sin título', calc: 'Libro de cálculo', slides: 'Presentación' }[type]);
    const id = `${Date.now().toString(36)}-${crypto.randomUUID().slice(0, 8)}`;
    const folder = this.currentFolder ? `${this.currentFolder}/` : `${type}/`;
    const path = `${folder}${this.normalizeName(base, type)}`;
    const payload = type === 'writer'
      ? { schema: 'tetord.writer', version: 1, id, name: base, content: '<p><br></p>', fileFormat: 'html', headerText: '', footerText: '', showPageNumber: false, differentFirstPage: false, styles: {} }
      : type === 'calc'
        ? { schema: 'tetord.calc', version: 1, id, name: base, rows: 20, cols: 12, data: {}, styles: {}, selected: 'A1', sheetIndex: 0, sheets: [{ name: 'Hoja 1', rows: 20, cols: 12, data: {}, styles: {} }] }
        : { schema: 'tetord.slides', version: 1, id, name: base, selected: 0, items: [{ title: base, body: '', theme: 'miku', background: '#dceeff', imagePath: '' }] };
    const result = await this.storage.writeJSON(path, payload);
    if (!result?.ok) { this.announce(result.error || 'No se pudo crear el documento.'); return false; }
    await this.refresh();
    await this.openRecord({ path, type, name: base });
    return true;
  }

  async openRecord(record) {
    const result = await this.storage.readJSON(record.path, null);
    if (!result?.value) return false;
    if (record.type === 'writer') return this.office?.modules?.writer?.openNativeRecord?.(record, result.value) ?? false;
    if (record.type === 'calc') return this.office?.modules?.calc?.openNativeRecord?.(record, result.value) ?? false;
    if (record.type === 'slides') return this.office?.modules?.slides?.openNativeRecord?.(record, result.value) ?? false;
    return false;
  }

  async renameRecord(record, requestedName = null) {
    const value = requestedName ?? window.prompt('Nuevo nombre', record.name);
    if (!value?.trim()) return false;
    const next = this.normalizeName(value);
    const dir = String(record.path).includes('/') ? String(record.path).slice(0, String(record.path).lastIndexOf('/')) : '';
    const target = dir ? `${dir}/${next}` : next;
    const result = await this.storage.rename(record.path, target);
    if (result?.ok) {
      const disk = await this.storage.readJSON(target, null);
      if (disk?.value && typeof disk.value === 'object') { disk.value.name = next.replace(/\.tetord$/i, ''); await this.storage.writeJSON(target, disk.value); }
      if (record.type === 'writer') { const d = window.MiWord?.LocalDocuments?.get?.(record.path); if (d) { d.name = next.replace(/\.tetord$/i, ''); d.filePath = target; window.MiWord.LocalDocuments.upsert(d); } }
    } else this.announce(result.error || 'No se pudo renombrar.');
    await this.refresh(); this.render();
    return !!result?.ok;
  }

  async duplicateRecord(record) {
    const dir = record.path.includes('/') ? record.path.slice(0, record.path.lastIndexOf('/')) : '';
    const targetName = `${record.name} (copia)${DOC_EXT}`;
    const target = dir ? `${dir}/${targetName}` : targetName;
    const source = await this.storage.readJSON(record.path, null);
    if (!source?.value) return false;
    const value = structuredClone(source.value);
    value.id = crypto.randomUUID(); value.name = `${record.name} (copia)`; value.updatedAt = new Date().toISOString();
    const result = await this.storage.writeJSON(target, value);
    if (!result?.ok) this.announce(result.error || 'No se pudo duplicar.');
    await this.refresh(); this.render();
    return !!result?.ok;
  }

  async deleteRecord(record) {
    if (!window.confirm(`¿Eliminar “${record.name}”? Esta acción no se puede deshacer.`)) return false;
    const result = await this.storage.remove(record.path);
    this.favorites.delete(record.path); await this.saveSettings();
    if (!result?.ok) this.announce(result.error || 'No se pudo eliminar.');
    await this.refresh(); this.render();
    return !!result?.ok;
  }

  async toggleFavorite(record) {
    if (this.favorites.has(record.path)) this.favorites.delete(record.path); else this.favorites.add(record.path);
    await this.saveSettings(); await this.refresh(); this.render();
  }

  async renameFolder(folder, requestedName = null) {
    const value = requestedName ?? window.prompt('Nuevo nombre', folder.name);
    if (!value?.trim()) return false;
    const dir = folder.path.includes('/') ? folder.path.slice(0, folder.path.lastIndexOf('/')) : '';
    const target = dir ? `${dir}/${value.trim().slice(0,60)}` : value.trim().slice(0,60);
    const result = await this.storage.rename(folder.path, target);
    await this.refresh(); this.render();
    return !!result?.ok;
  }

  async deleteFolder(folder) {
    if (folder.path === '') return false;
    if (!window.confirm(`¿Eliminar la carpeta “${folder.name}” y su contenido?`)) return false;
    const result = await this.storage.remove(folder.path);
    await this.refresh(); this.render();
    return !!result?.ok;
  }

  render() {
    const box = this.byId('officeFilesList'); if (!box) return;
    this.clear(box);
    const rows = this.visibleRecords();
    if (!rows.length) { box.appendChild(this.el('p', { className: 'office-empty office-empty--large', text: 'No se encontraron elementos aquí.' })); return; }
    rows.forEach((item, index) => {
      const row = this.el('div', { className: `office-file-row office-file-row--${item.type}${item.favorite ? ' is-favorite' : ''}`, attrs: { 'data-file-index': index, tabindex: '0' } });
      const main = this.button('', 'office-file-mainbutton');
      main.append(this.el('span', { className: 'office-file-icon', text: ICON[item.type] || '?' }), this.el('span', { className: 'office-file-main', children: [this.el('strong', { text: item.name }), this.el('small', { text: `${item.label}${item.size ? ` · ${this.formatSize(item.size)}` : ''}` })] }));
      main.addEventListener('click', () => { if (item.isFolder) { this.currentFolder = item.path; if (this.office) this.office.fileFolder = this.currentFolder; this.render(); } else this.openRecord(item); });
      row.appendChild(main);
      row.appendChild(this.el('span', { className: 'office-file-date', text: item.updatedAt ? new Date(item.updatedAt).toLocaleString('es-MX') : '' }));
      if (!item.isFolder) {
        const fav = this.button(item.favorite ? '★' : '☆', 'office-file-action', { title: 'Favorito', 'aria-label': 'Favorito' });
        fav.addEventListener('click', e => { e.stopPropagation(); this.toggleFavorite(item); });
        row.appendChild(fav);
      }
      const more = this.button('⋯', 'office-file-action', { title: 'Más acciones', 'aria-label': 'Más acciones' });
      more.addEventListener('click', e => { e.stopPropagation(); this.fileMenu(item); });
      row.appendChild(more); box.appendChild(row);
    });
    this.renderBreadcrumbs();
  }

  async fileMenu(item) {
    const action = window.prompt('Acciones: renombrar / duplicar / eliminar / favorito', 'renombrar')?.toLowerCase();
    if (action === 'renombrar') return item.isFolder ? this.renameFolder(item) : this.renameRecord(item);
    if (action === 'duplicar' && !item.isFolder) return this.duplicateRecord(item);
    if (action === 'eliminar') return item.isFolder ? this.deleteFolder(item) : this.deleteRecord(item);
    if (action === 'favorito' && !item.isFolder) return this.toggleFavorite(item);
    return false;
  }

  renderBreadcrumbs() {
    const box = this.byId('officeFilesBreadcrumbs'); if (!box) return;
    this.clear(box);
    const root = this.button('Este dispositivo'); root.addEventListener('click', () => { this.currentFolder = ''; if (this.office) this.office.fileFolder = ''; this.render(); }); box.appendChild(root);
    const parts = this.currentFolder ? this.currentFolder.split('/') : [];
    let path = '';
    parts.forEach((part, i) => { path = path ? `${path}/${part}` : part; box.appendChild(this.el('span', { text: '›' })); const b = this.button(part); const p = path; b.addEventListener('click', () => { this.currentFolder = p; if (this.office) this.office.fileFolder = p; this.render(); }); box.appendChild(b); });
  }

  formatSize(bytes) { const n = Number(bytes) || 0; if (n < 1024) return `${n} B`; if (n < 1024 * 1024) return `${(n / 1024).toFixed(1)} KB`; return `${(n / 1024 / 1024).toFixed(2)} MB`; }
}

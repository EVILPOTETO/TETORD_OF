import { AppModule } from './app-module.js';

const DEFAULT_SLIDES = {
  version: 2,
  selected: 0,
  items: [{ title: 'Título de la presentación', body: 'Escribe el contenido de esta diapositiva.', theme: 'miku', background: '#dceeff', imagePath: '' }]
};

function extFromMime(mime = '') {
  const map = { 'image/png': 'png', 'image/jpeg': 'jpg', 'image/webp': 'webp', 'image/gif': 'gif', 'image/svg+xml': 'svg' };
  return map[mime] || 'bin';
}

export class SlidesModule extends AppModule {
  constructor(ctx) {
    super(ctx);
    this.state = structuredClone(DEFAULT_SLIDES);
    this.storagePath = 'slides/Presentación.tetord';
    this.bound = false;
  }

  normalize(value) {
    const v = value && typeof value === 'object' ? value : {};
    const items = Array.isArray(v.items) && v.items.length ? v.items : DEFAULT_SLIDES.items;
    return {
      version: 2,
      selected: Math.max(0, Math.min(items.length - 1, Number(v.selected) || 0)),
      items: items.map(s => ({
        title: String(s?.title ?? 'Nueva diapositiva'),
        body: String(s?.body ?? ''),
        theme: String(s?.theme ?? 'miku'),
        background: String(s?.background ?? '#dceeff'),
        imagePath: String(s?.imagePath ?? '')
      }))
    };
  }

  async openNativeRecord(record, value) { this.storagePath = record.path; this.state = this.normalize(value); this.office.slides = this.state; this.office.open('slides'); await this.renderSlides(); return true; }

  async init() {
    let disk = await this.storage.readJSON(this.storagePath, null);
    if (!disk) disk = await this.migrateLegacyLocalStorage();
    this.state = this.normalize(disk || DEFAULT_SLIDES);
    await this.persist();
    this.bind();
    await this.renderSlides();
  }

  async migrateLegacyLocalStorage() {
    try {
      const raw = localStorage.getItem('tetord.office.slides');
      if (!raw) return null;
      const legacy = JSON.parse(raw);
      const state = this.normalize(legacy);
      for (let i = 0; i < state.items.length; i++) {
        const image = legacy?.items?.[i]?.image;
        if (typeof image === 'string' && image.startsWith('data:')) {
          const match = image.match(/^data:([^;]+);base64,(.*)$/s);
          if (match) {
            const path = `slides/assets/slide-${Date.now().toString(36)}-${i}.${extFromMime(match[1])}`;
            await this.storage.writeBytes(path, match[2], match[1]);
            state.items[i].imagePath = path;
          }
        }
      }
      await this.storage.writeJSON(this.storagePath, state);
      localStorage.removeItem('tetord.office.slides');
      return state;
    } catch (_) { return null; }
  }

  async persist() {
    const result = await this.storage.writeJSON(this.storagePath, this.state);
    if (!result?.ok) this.announce('No se pudo guardar la presentación.');
    if (this.office?.touchOfficeMeta) this.office.touchOfficeMeta('slides');
    return result;
  }

  blank(title = 'Nueva diapositiva') { return { title, body: '', theme: 'miku', background: '#dceeff', imagePath: '' }; }

  bind() {
    if (this.bound) return;
    this.bound = true;
    this.on('slideNewBtn', 'click', async () => { this.state = { version: 2, selected: 0, items: [this.blank('Nueva presentación')] }; await this.persist(); await this.renderSlides(); });
    this.on('slideAddBtn', 'click', async () => { const at = this.state.selected + 1; this.state.items.splice(at, 0, this.blank()); this.state.selected = at; await this.persist(); await this.renderSlides(); });
    this.on('slideDuplicateBtn', 'click', async () => { const copy = structuredClone(this.state.items[this.state.selected]); this.state.items.splice(this.state.selected + 1, 0, copy); this.state.selected++; await this.persist(); await this.renderSlides(); });
    this.on('slideDeleteBtn', 'click', async () => { if (this.state.items.length === 1) return; const [removed] = this.state.items.splice(this.state.selected, 1); if (removed?.imagePath) await this.storage.remove(removed.imagePath); this.state.selected = Math.max(0, this.state.selected - 1); await this.persist(); await this.renderSlides(); });
    this.on('slideTitleInput', 'input', async e => { this.state.items[this.state.selected].title = e.target.value; await this.persist(); await this.renderSlides(false); });
    this.on('slideBodyInput', 'input', async e => { this.state.items[this.state.selected].body = e.target.value; await this.persist(); await this.renderSlides(false); });
    this.on('slideThemeSelect', 'change', async e => { this.state.items[this.state.selected].theme = e.target.value; await this.persist(); await this.renderSlides(false); });
    this.on('slideBgColor', 'input', async e => { this.state.items[this.state.selected].background = e.target.value; await this.persist(); await this.renderSlides(false); });
    this.on('slideImageBtn', 'click', () => this.byId('slideImageInput')?.click());
    this.on('slideImageInput', 'change', e => this.handleImage(e));
    this.on('slidePresentBtn', 'click', () => this.present());
    this.on('slidePresentCloseBtn', 'click', () => { const o = this.byId('slidePresentOverlay'); if (o) o.hidden = true; });
    this.on('slidePrevBtn', 'click', () => this.slideNavigate(-1));
    this.on('slideNextBtn', 'click', () => this.slideNavigate(1));
    document.addEventListener('keydown', e => {
      const o = this.byId('slidePresentOverlay'); if (o?.hidden) return;
      if (e.key === 'ArrowRight' || e.key === ' ') { e.preventDefault(); this.slideNavigate(1); }
      else if (e.key === 'ArrowLeft') { e.preventDefault(); this.slideNavigate(-1); }
      else if (e.key === 'Escape') o.hidden = true;
    });
  }

  async handleImage(event) {
    const file = event?.target?.files?.[0]; if (!file) return;
    const buffer = await file.arrayBuffer();
    const bytes = new Uint8Array(buffer);
    let binary = ''; const chunk = 0x8000;
    for (let i = 0; i < bytes.length; i += chunk) binary += String.fromCharCode(...bytes.subarray(i, i + chunk));
    const base64 = btoa(binary);
    const path = `slides/assets/${crypto.randomUUID()}.${extFromMime(file.type)}`;
    const result = await this.storage.writeBytes(path, base64, file.type || 'application/octet-stream');
    if (!result?.ok) return this.announce('No se pudo guardar la imagen.');
    const old = this.state.items[this.state.selected].imagePath;
    this.state.items[this.state.selected].imagePath = path;
    if (old) await this.storage.remove(old);
    await this.persist(); await this.renderSlides(false);
    event.target.value = '';
  }

  async renderSlides(edit = true) {
    const box = this.byId('slideThumbs'); if (!box) return;
    this.clear(box);
    this.state.items.forEach((s, i) => {
      const b = this.button('', `slide-thumb${i === this.state.selected ? ' is-active' : ''}`, { 'aria-label': `Diapositiva ${i + 1}` });
      b.append(this.el('span', { text: i + 1 }), this.el('strong', { text: s.title || 'Sin título' }), this.el('small', { text: (s.body || '').slice(0, 60) }));
      b.addEventListener('click', async () => { this.state.selected = i; await this.renderSlides(); });
      box.appendChild(b);
    });
    const s = this.state.items[this.state.selected] || this.state.items[0];
    if (edit) {
      const title = this.byId('slideTitleInput'), body = this.byId('slideBodyInput'), theme = this.byId('slideThemeSelect'), bg = this.byId('slideBgColor');
      if (title) title.value = s.title; if (body) body.value = s.body; if (theme) theme.value = s.theme; if (bg) bg.value = s.background;
    }
    const canvas = this.byId('slideCanvas');
    if (canvas) { canvas.className = `slide-canvas theme-${s.theme}`; canvas.style.background = s.background; }
    const img = this.byId('slideImagePreview');
    if (img) { img.hidden = !s.imagePath; img.alt = s.title || 'Imagen de diapositiva'; img.src = s.imagePath ? await this.storage.readDataURL(s.imagePath) : ''; }
    const status = this.byId('slidesStatus'); if (status) status.textContent = `Presentación · ${this.state.items.length} diapositiva${this.state.items.length === 1 ? '' : 's'}`;
  }

  async renderPresent() {
    const s = this.state.items[this.state.selected]; const c = this.byId('slidePresentCard'); if (!c || !s) return;
    this.clear(c); c.style.background = s.background || '';
    c.append(this.el('h1', { text: s.title || 'Sin título' }));
    const body = this.el('div'); String(s.body || '').split('\n').forEach((line, i) => { if (i) body.appendChild(document.createElement('br')); body.appendChild(document.createTextNode(line)); });
    c.appendChild(body);
    if (s.imagePath) { const img = this.el('img', { attrs: { alt: '', decoding: 'async' } }); img.src = await this.storage.readDataURL(s.imagePath); c.appendChild(img); }
  }

  async present() { await this.renderPresent(); const o = this.byId('slidePresentOverlay'); if (o) o.hidden = false; }
  async slideNavigate(delta) { this.state.selected = Math.max(0, Math.min(this.state.items.length - 1, this.state.selected + delta)); await this.renderSlides(false); await this.renderPresent(); }
}

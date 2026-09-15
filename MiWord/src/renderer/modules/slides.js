import { AppModule } from './app-module.js';

const DEFAULT_SLIDES = {
  version: 3,
  selected: 0,
  items: [{
    title: 'Título de la presentación', body: 'Escribe el contenido de esta diapositiva.',
    theme: 'miku', background: '#dceeff', imagePath: '',
    objects: [
      { id: 'title', type: 'text', text: 'Título de la presentación', x: 8, y: 9, w: 84, h: 14, fontSize: 34, bold: true, color: '#19465d' },
      { id: 'body', type: 'text', text: 'Escribe el contenido de esta diapositiva.', x: 10, y: 28, w: 80, h: 28, fontSize: 21, bold: false, color: '#28576b' }
    ]
  }]
};

function extFromMime(mime = '') {
  const map = { 'image/png': 'png', 'image/jpeg': 'jpg', 'image/webp': 'webp', 'image/gif': 'gif', 'image/svg+xml': 'svg' };
  return map[mime] || 'bin';
}

const clone = value => structuredClone(value);
const uid = () => `obj-${crypto.randomUUID()}`;

export class SlidesModule extends AppModule {
  constructor(ctx) {
    super(ctx);
    this.state = clone(DEFAULT_SLIDES);
    this.storagePath = 'slides/Presentación.tetord';
    this.bound = false;
    this.drag = null;
    this.history = [];
    this.historyIndex = -1;
    this.clipboardObject = null;
    this._historyRestoring = false;
    this.homeVisible = true;
  }

  normalizeObject(o, fallback = {}) {
    const type = ['text', 'shape', 'image'].includes(o?.type) ? o.type : 'text';
    return {
      id: String(o?.id || uid()), type,
      text: String(o?.text ?? fallback.text ?? ''),
      x: Number.isFinite(Number(o?.x)) ? Number(o.x) : (fallback.x ?? 10),
      y: Number.isFinite(Number(o?.y)) ? Number(o.y) : (fallback.y ?? 10),
      w: Number.isFinite(Number(o?.w)) ? Number(o.w) : (fallback.w ?? 80),
      h: Number.isFinite(Number(o?.h)) ? Number(o.h) : (fallback.h ?? 20),
      fontSize: Math.max(10, Number(o?.fontSize) || fallback.fontSize || 24),
      bold: Boolean(o?.bold ?? fallback.bold),
      italic: Boolean(o?.italic ?? fallback.italic),
      underline: Boolean(o?.underline ?? fallback.underline),
      align: String(o?.align || fallback.align || 'left'),
      fontFamily: String(o?.fontFamily || fallback.fontFamily || 'Aptos'),
      zIndex: Number.isFinite(Number(o?.zIndex)) ? Number(o.zIndex) : (fallback.zIndex ?? 0),
      shapeKind: String(o?.shapeKind || fallback.shapeKind || 'rect'),
      color: String(o?.color || fallback.color || '#19465d'),
      fill: String(o?.fill || fallback.fill || '#78b9df'),
      radius: Math.max(0, Number(o?.radius) || 12),
      src: String(o?.src || fallback.src || '')
    };
  }

  legacyObjects(s) {
    const objects = [];
    if (s?.title) objects.push(this.normalizeObject({ id: 'legacy-title', type: 'text', text: s.title, x: 8, y: 9, w: 84, h: 14, fontSize: 34, bold: true, color: s.theme === 'dark' ? '#f4f8ff' : '#19465d' }));
    if (s?.body) objects.push(this.normalizeObject({ id: 'legacy-body', type: 'text', text: s.body, x: 10, y: 28, w: 80, h: 30, fontSize: 21, color: s.theme === 'dark' ? '#f4f8ff' : '#28576b' }));
    if (s?.imagePath) objects.push(this.normalizeObject({ id: 'legacy-image', type: 'image', src: s.imagePath, x: 15, y: 60, w: 70, h: 30 }));
    return objects;
  }

  normalize(value) {
    const v = value && typeof value === 'object' ? value : {};
    const rawItems = Array.isArray(v.items) && v.items.length ? v.items : DEFAULT_SLIDES.items;
    const items = rawItems.map((s, index) => {
      const objects = Array.isArray(s?.objects) && s.objects.length ? s.objects.map(o => this.normalizeObject(o)) : this.legacyObjects(s);
      return {
        title: String(s?.title ?? objects.find(o => o.type === 'text')?.text ?? `Diapositiva ${index + 1}`),
        body: String(s?.body ?? ''), theme: String(s?.theme ?? 'miku'), background: String(s?.background ?? '#dceeff'),
        imagePath: String(s?.imagePath ?? ''), objects
      };
    });
    return { version: 3, selected: Math.max(0, Math.min(items.length - 1, Number(v.selected) || 0)), items };
  }

  async openNativeRecord(record, value) { this.storagePath = record.path; this.state = this.normalize(value); this.office.slides = this.state; this.office.open('slides'); await this.renderSlides(); return true; }

  showHome() {
    this.homeVisible = true;
    const home = this.byId('slidesHome');
    if (home) home.hidden = false;
    const editor = this.byId('slidesEditorShell');
    if (editor) editor.hidden = true;
    this.renderHome();
  }

  showEditor() {
    this.homeVisible = false;
    const home = this.byId('slidesHome');
    if (home) home.hidden = true;
    const editor = this.byId('slidesEditorShell');
    if (editor) editor.hidden = false;
  }

  renderHome() {
    const box = this.byId('slidesHomeRecent');
    if (!box) return;
    while (box.firstChild) box.removeChild(box.firstChild);
    const rows = (this.office.fileRecords?.() || []).filter(f => f.type === 'slides').sort((a,b) => String(b.updatedAt||'').localeCompare(String(a.updatedAt||''))).slice(0,8);
    if (!rows.length) {
      const empty = document.createElement('div'); empty.className='slides-home-empty';
      const strong=document.createElement('strong'); strong.textContent='Aún no hay presentaciones recientes';
      const span=document.createElement('span'); span.textContent='Crea tu primera presentación para verla aquí.';
      empty.append(strong,span); box.appendChild(empty); return;
    }
    rows.forEach(f => {
      const b=document.createElement('button'); b.type='button'; b.className='slides-recent-card';
      const preview=document.createElement('span'); preview.className='slides-recent-preview'; preview.textContent='S';
      const main=document.createElement('span'); main.className='slides-recent-main';
      const strong=document.createElement('strong'); strong.textContent=String(f.name||'Presentación');
      const small=document.createElement('small'); small.textContent=String(f.updatedAt ? new Date(f.updatedAt).toLocaleString('es-MX') : 'TETORD Slides');
      main.append(strong,small); b.append(preview,main); b.addEventListener('click',()=>this.office.openFile?.(f)); box.appendChild(b);
    });
  }

  async createFromTemplate(template='blank') {
    const titles={blank:'Nueva presentación',title:'Nueva presentación',bigstat:'Resultados 2026',cards3:'Presentación ejecutiva',timeline:'Plan de trabajo',compare:'Comparación'};
    this.storagePath=`slides/${Date.now().toString(36)}-${crypto.randomUUID().slice(0,6)}.tetord`;
    this.state={version:3,selected:0,items:[this.blank(titles[template]||'Nueva presentación')]};
    this.state.items[0].title=titles[template]||'Nueva presentación';
    this.state.items[0].body='';
    this.commit();
    if(template!=='blank' && template!=='title') this.applyLayout(template);
    await this.persist();
    this.office.slides=this.state;
    this.office.touchOfficeMeta?.('slides');
    await this.renderSlides();
  }

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
            const imageObj = state.items[i].objects.find(o => o.type === 'image');
            if (imageObj) imageObj.src = path;
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

  blank(title = 'Nueva diapositiva') {
    return { title, body: '', theme: 'miku', background: '#dceeff', imagePath: '', objects: [this.normalizeObject({ type: 'text', text: title, x: 8, y: 9, w: 84, h: 14, fontSize: 34, bold: true, color: '#19465d' })] };
  }

  current() { return this.state.items[this.state.selected] || this.state.items[0]; }
  selectedObject() { return this.current()?.objects?.find(o => o.id === this._selectedObjectId) || null; }

  selectObject(id) {
    this._selectedObjectId = id;
    this.renderObjects();
    this.renderInspector();
  }

  addObject(object) {
    const s = this.current();
    const o = this.normalizeObject({ id: uid(), ...object });
    s.objects.push(o);
    this._selectedObjectId = o.id;
    this.persist();
    this.renderObjects();
    this.renderInspector();
  }

  snapshotHistory() {
    if (this._historyRestoring) return;
    const snap = JSON.stringify(this.state);
    if (this.history[this.historyIndex] === snap) return;
    this.history = this.history.slice(0, this.historyIndex + 1);
    this.history.push(snap);
    if (this.history.length > 80) this.history.shift();
    this.historyIndex = this.history.length - 1;
  }

  async restoreHistory(index) {
    if (index < 0 || index >= this.history.length) return;
    this._historyRestoring = true;
    this.state = this.normalize(JSON.parse(this.history[index]));
    this._selectedObjectId = null;
    this.historyIndex = index;
    this._historyRestoring = false;
    await this.persist();
    await this.renderSlides();
  }

  async undo() { await this.restoreHistory(this.historyIndex - 1); }
  async redo() { await this.restoreHistory(this.historyIndex + 1); }

  commit() {
    this.snapshotHistory();
    this.on('slidesHomeNew','click',()=>this.createFromTemplate('blank'));
    this.on('slidesHomeBack','click',()=>this.office.open('home'));
    this.on('slidesHomeRefresh','click',()=>this.renderHome());
    document.querySelectorAll('[data-slides-template]').forEach(btn=>btn.addEventListener('click',()=>this.createFromTemplate(btn.dataset.slidesTemplate||'blank')));
    this.persist();
  }

  applyTextFormat(prop, value) {
    const o = this.selectedObject();
    if (!o || o.type !== 'text') return;
    o[prop] = value;
    this.commit();
    this.renderObjects();
    this.renderInspector();
    this.updateRibbonState();
  }

  setLayer(direction) {
    const s = this.current(), o = this.selectedObject();
    if (!s || !o) return;
    const ordered = [...s.objects].sort((a,b) => (a.zIndex||0) - (b.zIndex||0));
    const idx = ordered.findIndex(x => x.id === o.id);
    if (idx < 0) return;
    if (direction === 'front') o.zIndex = (Math.max(...ordered.map(x=>x.zIndex||0), 0) + 1);
    else if (direction === 'back') o.zIndex = (Math.min(...ordered.map(x=>x.zIndex||0), 0) - 1);
    else if (direction === 'forward') { const n = ordered[idx + 1]; if (n) { const z=o.zIndex||0; o.zIndex=n.zIndex||0; n.zIndex=z; } }
    else if (direction === 'backward') { const n = ordered[idx - 1]; if (n) { const z=o.zIndex||0; o.zIndex=n.zIndex||0; n.zIndex=z; } }
    this.commit(); this.renderObjects();
  }

  async copyObject() {
    const o = this.selectedObject(); if (!o) return;
    this.clipboardObject = clone(o);
    try { await navigator.clipboard?.writeText(JSON.stringify(o)); } catch (_) {}
  }

  pasteObject() {
    if (!this.clipboardObject) return;
    this.addObject({ ...clone(this.clipboardObject), id: uid(), x: Math.min(88, this.clipboardObject.x + 4), y: Math.min(88, this.clipboardObject.y + 4) });
  }

  updateRibbonState() {
    const o = this.selectedObject();
    const set = (id, prop, value) => { const e=this.byId(id); if(e) e[prop]=value; };
    set('pptFontFamily','value',o?.fontFamily || 'Aptos');
    set('pptFontSize','value',String(Math.round(o?.fontSize || 24)));
    set('pptFontColor','value',o?.color || '#19465d');
    ['pptBoldBtn','pptItalicBtn','pptUnderlineBtn'].forEach((id,i)=>set(id,'className',`ppt-ribbon-toggle${o?.type==='text' && [o.bold,o.italic,o.underline][i]?' is-on':''}`));
    ['pptAlignLeftBtn','pptAlignCenterBtn','pptAlignRightBtn'].forEach((id,i)=>set(id,'className',`ppt-ribbon-toggle${o?.type==='text' && ['left','center','right'][i]===(o.align||'left')?' is-on':''}`));
  }

  bind() {
    if (this.bound) return;
    this.bound = true;
    this.snapshotHistory();
    const click = (id, fn) => this.on(id, 'click', fn);
    document.querySelectorAll('[data-ppt-tab]').forEach(tab => tab.addEventListener('click', () => {
      document.querySelectorAll('[data-ppt-tab]').forEach(t=>t.classList.toggle('is-active',t===tab));
      const key=tab.dataset.pptTab;
      document.querySelectorAll('[class*="ppt-ribbon-tab-"]').forEach(el=>{
        const active=el.classList.contains(`ppt-ribbon-tab-${key}`);
        el.hidden=!active;
      });
    }));
    click('slideNewBtn', async () => { this.state = { version: 3, selected: 0, items: [this.blank('Nueva presentación')] }; this._selectedObjectId = null; this.commit(); await this.renderSlides(); });
    click('slideAddQuickBtn', async () => { const at=this.state.selected+1; this.state.items.splice(at,0,this.blank()); this.state.selected=at; this._selectedObjectId=null; this.commit(); await this.renderSlides(); });
    click('slideAddBtn', async () => { const at=this.state.selected+1; this.state.items.splice(at,0,this.blank()); this.state.selected=at; this._selectedObjectId=null; this.commit(); await this.renderSlides(); });
    click('slideDuplicateBtn', async () => { const copy=clone(this.current()); copy.objects=copy.objects.map(o=>({...o,id:uid()})); this.state.items.splice(this.state.selected+1,0,copy); this.state.selected++; this._selectedObjectId=null; this.commit(); await this.renderSlides(); });
    click('slideDeleteBtn', async () => { if(this.state.items.length===1)return; this.state.items.splice(this.state.selected,1); this.state.selected=Math.max(0,this.state.selected-1); this._selectedObjectId=null; this.commit(); await this.renderSlides(); });
    click('pptInsertTextBtn', () => this.addObject({type:'text',text:'Texto nuevo',x:15,y:45,w:70,h:12,fontSize:24}));
    click('pptInsertShapeBtn', () => this.addObject({type:'shape',x:35,y:45,w:30,h:20,fill:'#78b9df',radius:10,shapeKind:'rect'})); click('slideAddTextBtn', () => this.addObject({type:'text',text:'Texto nuevo',x:15,y:45,w:70,h:12,fontSize:24})); click('slideAddShapeBtn', () => this.addObject({type:'shape',x:35,y:45,w:30,h:20,fill:'#78b9df',radius:10,shapeKind:'rect'}));
    click('pptCircleBtn', () => this.addObject({type:'shape',x:35,y:42,w:22,h:22,fill:'#78b9df',radius:50,shapeKind:'ellipse'}));
    click('pptLineBtn', () => this.addObject({type:'shape',x:20,y:48,w:60,h:1.2,fill:'#4e89b4',radius:0,shapeKind:'line'}));
    click('pptArrowBtn', () => this.addObject({type:'shape',x:25,y:45,w:50,h:4,fill:'#4e89b4',radius:2,shapeKind:'arrow'}));
    click('slideDuplicateObjectBtn', () => { const o=this.selectedObject(); if(!o)return; this.addObject({...clone(o),id:uid(),x:Math.min(88,o.x+3),y:Math.min(88,o.y+3)}); });
    click('pptThumbAddBtn', async () => { const at=this.state.selected+1; this.state.items.splice(at,0,this.blank()); this.state.selected=at; this._selectedObjectId=null; this.commit(); await this.renderSlides(); });
    click('slideDeleteObjectBtn', async () => { const s=this.current(); if(!this._selectedObjectId)return; s.objects=s.objects.filter(o=>o.id!==this._selectedObjectId); this._selectedObjectId=null; this.commit(); this.renderObjects(); this.renderInspector(); });
    click('pptUndoBtn',()=>this.undo()); click('pptRedoBtn',()=>this.redo());
    click('pptCopyBtn',()=>this.copyObject()); click('pptPasteBtn',()=>this.pasteObject());
    click('pptFrontBtn',()=>this.setLayer('front')); click('pptBackBtn',()=>this.setLayer('back')); click('pptForwardBtn',()=>this.setLayer('forward')); click('pptBackwardBtn',()=>this.setLayer('backward'));
    click('pptBoldBtn',()=>this.applyTextFormat('bold',!this.selectedObject()?.bold)); click('pptItalicBtn',()=>this.applyTextFormat('italic',!this.selectedObject()?.italic)); click('pptUnderlineBtn',()=>this.applyTextFormat('underline',!this.selectedObject()?.underline));
    click('pptAlignLeftBtn',()=>this.applyTextFormat('align','left')); click('pptAlignCenterBtn',()=>this.applyTextFormat('align','center')); click('pptAlignRightBtn',()=>this.applyTextFormat('align','right'));
    click('pptZoomInBtn',()=>this.setZoom(1.1)); click('pptZoomOutBtn',()=>this.setZoom(.9)); click('pptFitBtn',()=>this.setZoom(1));
    this.on('pptFontFamily','change',e=>this.applyTextFormat('fontFamily',e.target.value));
    this.on('pptFontSize','change',e=>this.applyTextFormat('fontSize',Math.max(10,Number(e.target.value)||24)));
    this.on('pptFontColor','input',e=>this.applyTextFormat('color',e.target.value));
    click('pptBulletBtn',()=>{ const o=this.selectedObject(); if(!o||o.type!=='text')return; o.text=o.text.split('\n').map(line=>line.trim().startsWith('•')?line:`• ${line}`).join('\n'); this.commit(); this.renderObjects(); this.renderInspector(); });
    click('pptPresentStartBtn',async()=>{this.state.selected=0;this._selectedObjectId=null;await this.renderSlides(false);await this.present();});
    click('pptShareBtn',()=>this.announce('Compartir estará disponible al conectar una cuenta TETORD.'));
    click('pptCommentsBtn',()=>this.announce('Comentarios: selecciona una diapositiva y usa Revisar para añadir comentarios.'));
    click('pptCommentsRibbonBtn',()=>this.announce('Comentario nuevo preparado para la diapositiva actual.'));
    click('pptSpellBtn',()=>this.announce('Revisión ortográfica básica: usa el corrector de tu sistema para texto libre.'));
    this.on('pptTransitionSelect','change',e=>{this.current().transition=e.target.value;this.commit();});
    this.on('pptTransitionDuration','input',e=>{this.current().transitionDuration=Number(e.target.value)||.5;this.commit();});
    this.on('pptAnimationSelect','change',e=>{this._pendingAnimation=e.target.value;});
    click('pptAnimationApplyBtn',()=>{const o=this.selectedObject();if(!o)return;const a=this._pendingAnimation||'Sin animación';o.animation=a;this.commit();this.renderObjects();});
    this.on('pptCanvasZoomOut','click',()=>this.setZoom(.9)); this.on('pptCanvasZoomIn','click',()=>this.setZoom(1.1)); this.on('pptCanvasFit','click',()=>this.setZoom(1));
    this.on('pptNotesInput','input',e=>{this.current().notes=e.target.value;this.commit();});
    this.on('slideThemeSelect','change',async e=>{this.current().theme=e.target.value;this.commit();await this.renderSlides(false);});
    this.on('slideApplyLayoutBtn','click',()=>this.applyLayout(this.byId('slideLayoutSelect')?.value||'title'));
    this.on('slideBgColor','input',async e=>{this.current().background=e.target.value;this.commit();await this.renderSlides(false);});
    click('slideImageBtn',()=>this.byId('slideImageInput')?.click()); this.on('slideImageInput','change',e=>this.handleImage(e));
    click('slidePresentBtn',()=>this.present()); click('slidePresentCloseBtn',()=>{const o=this.byId('slidePresentOverlay');if(o)o.hidden=true;});
    click('slidePrevBtn',()=>this.slideNavigate(-1)); click('slideNextBtn',()=>this.slideNavigate(1));
    for(const id of ['slideObjX','slideObjY','slideObjW','slideObjH','slideObjFont','slideObjColor','slideObjFill']) this.on(id,'input',e=>this.updateSelectedFromInspector(e));
    this.on('slideObjBold','change',e=>{const o=this.selectedObject();if(o){o.bold=e.target.checked;this.commit();this.renderObjects();this.updateRibbonState();}});
    this.on('slideObjText','input',e=>{const o=this.selectedObject();if(o){o.text=e.target.value;this.commit();this.renderObjects();}});
    this.on('slideObjItalic','change',e=>this.applyTextFormat('italic',e.target.checked)); this.on('slideObjUnderline','change',e=>this.applyTextFormat('underline',e.target.checked));
    this.on('slideObjAlign','change',e=>this.applyTextFormat('align',e.target.value)); this.on('slideObjFontFamily','change',e=>this.applyTextFormat('fontFamily',e.target.value));
    document.addEventListener('keydown',e=>{
      const target=e.target; const typing=target?.matches?.('input,textarea,select,[contenteditable="true"]');
      if((e.ctrlKey||e.metaKey)&&e.key.toLowerCase()==='z'&&!typing){e.preventDefault();e.shiftKey?this.redo():this.undo();return;}
      if((e.ctrlKey||e.metaKey)&&e.key.toLowerCase()==='y'&&!typing){e.preventDefault();this.redo();return;}
      if((e.ctrlKey||e.metaKey)&&e.key.toLowerCase()==='c'&&!typing){e.preventDefault();this.copyObject();return;}
      if((e.ctrlKey||e.metaKey)&&e.key.toLowerCase()==='v'&&!typing){e.preventDefault();this.pasteObject();return;}
      if((e.key==='Delete'||e.key==='Backspace')&&!typing&&this.selectedObject()){e.preventDefault();this.byId('slideDeleteObjectBtn')?.click();return;}
      const overlay=this.byId('slidePresentOverlay'); if(overlay&&!overlay.hidden){if(e.key==='ArrowRight'||e.key===' '){e.preventDefault();this.slideNavigate(1)}else if(e.key==='ArrowLeft'){e.preventDefault();this.slideNavigate(-1)}else if(e.key==='Escape')overlay.hidden=true;}
    });
    this.updateRibbonState();
  }


  setZoom(scale) {
    const canvas=this.byId('slideCanvas'); if(!canvas)return;
    const current=Number(canvas.dataset.zoom||1); const next=Math.max(.55,Math.min(1.45,current*scale));
    canvas.dataset.zoom=String(next); canvas.style.transform=`scale(${next})`; canvas.style.transformOrigin='center center';
  }

  applyLayout(layout) {
    const s = this.current();
    if (!s) return;
    const color = s.theme === 'dark' ? '#f4f8ff' : '#19465d';
    const accent = s.theme === 'dark' ? '#8fc8ef' : '#5d9ed0';
    const titleText = s.title || s.objects?.find(o => o.type === 'text')?.text || 'Título de la presentación';
    const text = s.body || 'Añade aquí la idea principal.';
    const mk = (obj) => this.normalizeObject({ id: uid(), color, ...obj });
    const layouts = {
      title: [
        mk({ type:'text', text:titleText, x:8, y:10, w:84, h:14, fontSize:34, bold:true }),
        mk({ type:'text', text, x:10, y:29, w:78, h:28, fontSize:21 })
      ],
      bigstat: [
        mk({ type:'text', text:titleText, x:8, y:9, w:84, h:12, fontSize:28, bold:true }),
        mk({ type:'text', text:'75%', x:10, y:28, w:42, h:30, fontSize:58, bold:true, color:accent }),
        mk({ type:'text', text, x:52, y:32, w:36, h:24, fontSize:18 })
      ],
      cards3: [
        mk({ type:'text', text:titleText, x:7, y:8, w:86, h:12, fontSize:30, bold:true }),
        mk({ type:'shape', x:7, y:27, w:27, h:49, fill:'#d7e8fb', radius:12 }),
        mk({ type:'shape', x:36.5, y:27, w:27, h:49, fill:'#e5f1fb', radius:12 }),
        mk({ type:'shape', x:66, y:27, w:27, h:49, fill:'#eef6fc', radius:12 }),
        mk({ type:'text', text:'Punto clave 1', x:9, y:33, w:23, h:10, fontSize:18, bold:true }),
        mk({ type:'text', text:'Punto clave 2', x:38.5, y:33, w:23, h:10, fontSize:18, bold:true }),
        mk({ type:'text', text:'Punto clave 3', x:68, y:33, w:23, h:10, fontSize:18, bold:true })
      ],
      timeline: [
        mk({ type:'text', text:titleText, x:7, y:8, w:86, h:12, fontSize:30, bold:true }),
        mk({ type:'shape', x:10, y:48, w:80, h:2, fill:accent, radius:2 }),
        mk({ type:'text', text:'01', x:10, y:39, w:12, h:9, fontSize:22, bold:true, color:accent }),
        mk({ type:'text', text:'02', x:40, y:39, w:12, h:9, fontSize:22, bold:true, color:accent }),
        mk({ type:'text', text:'03', x:70, y:39, w:12, h:9, fontSize:22, bold:true, color:accent }),
        mk({ type:'text', text:'Inicio', x:9, y:53, w:23, h:10, fontSize:17, bold:true }),
        mk({ type:'text', text:'Desarrollo', x:39, y:53, w:25, h:10, fontSize:17, bold:true }),
        mk({ type:'text', text:'Resultado', x:69, y:53, w:25, h:10, fontSize:17, bold:true })
      ],
      quote: [
        mk({ type:'text', text:'“', x:12, y:18, w:12, h:25, fontSize:52, bold:true, color:accent }),
        mk({ type:'text', text:text || titleText, x:20, y:28, w:64, h:25, fontSize:28, bold:true }),
        mk({ type:'text', text:'Autor / fuente', x:20, y:61, w:55, h:10, fontSize:16, color:accent })
      ],
      compare: [
        mk({ type:'text', text:titleText, x:7, y:8, w:86, h:12, fontSize:30, bold:true }),
        mk({ type:'shape', x:8, y:27, w:40, h:52, fill:'#e8f3fb', radius:12 }),
        mk({ type:'shape', x:52, y:27, w:40, h:52, fill:'#f4f7f9', radius:12 }),
        mk({ type:'text', text:'Opción A', x:12, y:33, w:32, h:10, fontSize:20, bold:true }),
        mk({ type:'text', text:'Opción B', x:56, y:33, w:32, h:10, fontSize:20, bold:true }),
        mk({ type:'text', text:'Ventaja principal', x:12, y:48, w:31, h:12, fontSize:16 }),
        mk({ type:'text', text:'Ventaja principal', x:56, y:48, w:31, h:12, fontSize:16 })
      ]
    };
    s.objects = layouts[layout] || layouts.title;
    s.title = titleText;
    this._selectedObjectId = null;
    this.commit();
    this.renderObjects();
    this.renderInspector();
    this.renderSlides(false);
  }

  updateSelectedFromInspector(e) {
    const o = this.selectedObject(); if (!o) return;
    const map = { slideObjX: 'x', slideObjY: 'y', slideObjW: 'w', slideObjH: 'h', slideObjFont: 'fontSize', slideObjColor: 'color', slideObjFill: 'fill' };
    const key = map[e.target.id]; if (!key) return;
    o[key] = ['color','fill'].includes(key) ? e.target.value : Number(e.target.value);
    this.commit(); this.renderObjects(); this.updateRibbonState();
  }

  async handleImage(event) {
    const file = event?.target?.files?.[0]; if (!file) return;
    const buffer = await file.arrayBuffer(); const bytes = new Uint8Array(buffer); let binary = ''; const chunk = 0x8000;
    for (let i = 0; i < bytes.length; i += chunk) binary += String.fromCharCode(...bytes.subarray(i, i + chunk));
    const base64 = btoa(binary); const path = `slides/assets/${crypto.randomUUID()}.${extFromMime(file.type)}`;
    const result = await this.storage.writeBytes(path, base64, file.type || 'application/octet-stream');
    if (!result?.ok) return this.announce('No se pudo guardar la imagen.');
    this.addObject({ type: 'image', src: path, x: 18, y: 52, w: 64, h: 34 });
    event.target.value = '';
  }

  makeThumb(s, i) {
    const b = this.button('', `slide-thumb${i === this.state.selected ? ' is-active' : ''}`, { 'aria-label': `Diapositiva ${i + 1}` });
    const mini = this.el('div', { className: 'slide-thumb-preview' }); mini.style.background = s.background;
    const title = s.objects?.find(o => o.type === 'text'); if (title) { const t = this.el('span', { text: title.text || 'Sin título' }); t.style.fontSize = '9px'; t.style.fontWeight = '700'; t.style.display = 'block'; t.style.padding = '8px'; mini.appendChild(t); }
    const n = this.el('span', { className: 'slide-thumb-number', text: String(i + 1) }); b.append(mini, n); b.addEventListener('click', async () => { this.state.selected = i; this._selectedObjectId = null; await this.renderSlides(); });
    return b;
  }

  async renderSlides(edit = true) {
    const box = this.byId('slideThumbs'); if (!box) return;
    this.clear(box); this.state.items.forEach((s, i) => box.appendChild(this.makeThumb(s, i)));
    const s = this.current();
    const theme = this.byId('slideThemeSelect'), bg = this.byId('slideBgColor');
    if (edit) { if (theme) theme.value = s.theme; if (bg) bg.value = s.background; const notes=this.byId('pptNotesInput'); if(notes) notes.value=s.notes||''; const label=this.byId('pptCurrentSlideLabel'); if(label) label.textContent=String(this.state.selected+1); }
    const canvas = this.byId('slideCanvas');
    if (canvas) { canvas.className = `slide-canvas theme-${s.theme}`; canvas.style.background = s.background; }
    const status = this.byId('slidesStatus'); if (status) status.textContent = `Presentación · ${this.state.items.length} diapositiva${this.state.items.length === 1 ? '' : 's'}`;
    this.renderObjects(); this.renderInspector();
  }

  async renderObjects() {
    const canvas = this.byId('slideCanvas'); if (!canvas) return;
    canvas.querySelectorAll('.slide-object').forEach(e => e.remove());
    const s = this.current();
    for (const o of s.objects || []) {
      const el = this.el('div', { className: `slide-object slide-object--${o.type}${o.id === this._selectedObjectId ? ' is-selected' : ''}` });
      el.dataset.objectId = o.id; el.style.left = `${o.x}%`; el.style.top = `${o.y}%`; el.style.width = `${o.w}%`; el.style.height = `${o.h}%`;
      el.style.color = o.color; el.style.fontSize = `${o.fontSize}px`; el.style.fontWeight = o.bold ? '700' : '400'; el.style.fontStyle = o.italic ? 'italic' : 'normal'; el.style.textDecoration = o.underline ? 'underline' : 'none'; el.style.fontFamily = o.fontFamily || 'Aptos'; el.style.textAlign = o.align || 'left'; el.style.zIndex = String(o.zIndex || 0); if(o.animation && o.animation!=='Sin animación') el.style.animation=`pptAnim${o.animation.replace(/[^a-z]/gi,'')} .45s ease`;
      if (o.type === 'shape') { el.style.background = o.fill; el.style.borderRadius = o.shapeKind === 'ellipse' ? '50%' : `${o.radius}px`; if(o.shapeKind==='line'){el.style.height='2px';} if(o.shapeKind==='arrow'){el.style.clipPath='polygon(0 30%,82% 30%,82% 0,100% 50%,82% 100%,82% 70%,0 70%)';} }
      if (o.type === 'text') { el.textContent = o.text; el.title = 'Doble clic para editar'; }
      if (o.type === 'image' && o.src) { const img = this.el('img', { attrs: { alt: '', decoding: 'async' } }); img.src = await this.storage.readDataURL(o.src); el.appendChild(img); }
      el.addEventListener('pointerdown', e => this.startDrag(e, o.id));
      el.addEventListener('click', e => { e.stopPropagation(); this.selectObject(o.id); });
      el.addEventListener('dblclick', e => { if (o.type === 'text') { const value = prompt('Texto del objeto', o.text); if (value !== null) { o.text = value; this.commit(); this.renderObjects(); this.renderInspector(); } } });
      canvas.appendChild(el);
    }
    canvas.onclick = () => { this._selectedObjectId = null; this.renderObjects(); this.renderInspector(); };
  }

  startDrag(event, id) {
    if (event.button !== 0) return;
    event.preventDefault(); event.stopPropagation(); this.selectObject(id);
    const canvas = this.byId('slideCanvas'); const o = this.selectedObject(); if (!canvas || !o) return;
    const rect = canvas.getBoundingClientRect();
    this.drag = { id, startX: event.clientX, startY: event.clientY, x: o.x, y: o.y, rect };
    const move = e => {
      if (!this.drag) return; const dx = ((e.clientX - this.drag.startX) / this.drag.rect.width) * 100; const dy = ((e.clientY - this.drag.startY) / this.drag.rect.height) * 100;
      o.x = Math.max(0, Math.min(100 - o.w, this.drag.x + dx)); o.y = Math.max(0, Math.min(100 - o.h, this.drag.y + dy)); this.renderObjects();
    };
    const up = () => { this.drag = null; window.removeEventListener('pointermove', move); window.removeEventListener('pointerup', up); this.commit(); this.renderInspector(); };
    window.addEventListener('pointermove', move); window.addEventListener('pointerup', up);
  }

  renderInspector() {
    const panel = this.byId('slideInspector'); if (!panel) return;
    const o = this.selectedObject(); panel.hidden = !o;
    if (!o) return;
    const set = (id, value) => { const e = this.byId(id); if (e) e.value = value; };
    set('slideObjX', Math.round(o.x)); set('slideObjY', Math.round(o.y)); set('slideObjW', Math.round(o.w)); set('slideObjH', Math.round(o.h)); set('slideObjFont', Math.round(o.fontSize)); set('slideObjColor', o.color); set('slideObjFill', o.fill); set('slideObjText', o.text || ''); set('slideObjFontFamily',o.fontFamily||'Aptos'); set('slideObjAlign',o.align||'left');
    const bold = this.byId('slideObjBold'); if (bold) bold.checked = Boolean(o.bold); const italic=this.byId('slideObjItalic'); if(italic) italic.checked=Boolean(o.italic); const underline=this.byId('slideObjUnderline'); if(underline) underline.checked=Boolean(o.underline);
    const kind = this.byId('slideInspectorType'); if (kind) kind.textContent = o.type === 'shape' ? 'Forma' : o.type === 'image' ? 'Imagen' : 'Texto';
    const textRow = this.byId('slideTextInspector'); if (textRow) textRow.hidden = o.type !== 'text';
    const fillRow = this.byId('slideFillInspector'); if (fillRow) fillRow.hidden = o.type !== 'shape';
  }

  async renderPresent() {
    const s = this.current(); const c = this.byId('slidePresentCard'); if (!c || !s) return;
    this.clear(c); c.style.background = s.background || '';
    for (const o of s.objects || []) {
      const el = this.el('div', { className: `present-object present-object--${o.type}` }); el.style.left = `${o.x}%`; el.style.top = `${o.y}%`; el.style.width = `${o.w}%`; el.style.height = `${o.h}%`; el.style.color = o.color; el.style.fontSize = `${Math.max(12, o.fontSize)}px`; el.style.fontWeight = o.bold ? '700' : '400'; el.style.fontStyle=o.italic?'italic':'normal'; el.style.textDecoration=o.underline?'underline':'none'; el.style.fontFamily=o.fontFamily||'Aptos'; el.style.textAlign=o.align||'left'; el.style.zIndex=String(o.zIndex||0);
      if (o.type === 'shape') { el.style.background = o.fill; el.style.borderRadius = o.shapeKind === 'ellipse' ? '50%' : `${o.radius}px`; if(o.shapeKind==='line'){el.style.height='2px';} if(o.shapeKind==='arrow'){el.style.clipPath='polygon(0 30%,82% 30%,82% 0,100% 50%,82% 100%,82% 70%,0 70%)';} }
      if (o.type === 'text') el.textContent = o.text;
      if (o.type === 'image' && o.src) { const img = this.el('img', { attrs: { alt: '', decoding: 'async' } }); img.src = await this.storage.readDataURL(o.src); el.appendChild(img); }
      c.appendChild(el);
    }
  }

  async present() { await this.renderPresent(); const o = this.byId('slidePresentOverlay'); if (o) o.hidden = false; }
  async slideNavigate(delta) { this.state.selected = Math.max(0, Math.min(this.state.items.length - 1, this.state.selected + delta)); this._selectedObjectId = null; await this.renderSlides(false); await this.renderPresent(); }
}

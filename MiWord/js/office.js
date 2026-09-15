import { SlidesModule } from '../src/renderer/modules/slides.js';
import { WriterModule } from '../src/renderer/modules/writer.js';
import { CalcModule } from '../src/renderer/modules/calc.js';
import { FilesModule } from '../src/renderer/modules/files.js';
import { MessengerModule } from '../src/renderer/modules/messenger.js';
import { CloudSyncModule } from '../src/renderer/modules/cloud-sync.js';
import { CollaborationModule } from '../src/renderer/modules/collaboration.js';
import { storage } from '../src/renderer/core/storage.js';

(function(){
  'use strict';
  const DEFAULT_CALC={rows:20,cols:12,data:{},styles:{},selected:'A1',sheetIndex:0,sheets:[{name:'Hoja 1',rows:20,cols:12,data:{},styles:{}}]};
  const O={
    app:'home', fileFilter:'all', fileSearch:'', fileSort:'updated', fileFolder:'root', fileSelected:null, calc:DEFAULT_CALC, slides:{items:[{title:'Título de la presentación',body:'Escribe el contenido de esta diapositiva.'}],selected:0}, chats:[{name:'Notas de TETORD',messages:[{from:'other',text:'Messenger local listo. Esta conversación permanece en este dispositivo.'}]}], chat:0,
    init(){
      this.modules={};
      this.shell=document.getElementById('officeShell'); if(!this.shell)return;
      this.calcPath = this.calcPath || 'calc/Libro de cálculo.tetord'; this.calc=this.normalizeCalc(this.load('calc',this.calc)); this.slides=this.load('slides',this.slides); this.chats=this.load('messenger',this.chats);
      document.getElementById('officeLauncherBtn')?.addEventListener('click',()=>this.open('home'));
      document.getElementById('officeBackToWriterBtn')?.addEventListener('click',()=>this.close());
      this.shell.querySelectorAll('[data-office-app]').forEach(b=>b.addEventListener('click',()=>this.open(b.dataset.officeApp)));
      this.initCalc(); this.initMessenger(); this.initFiles(); this.initHome(); this.open('home');
    },
    load(k,f){try{const v=localStorage.getItem('tetord.office.'+k);return v?JSON.parse(v):f}catch(_){return f}},
    save(k,v){
      try { localStorage.setItem('tetord.office.'+k,JSON.stringify(v)); } catch(_) {}
      if (storage.desktop) { const paths={calc:this.calcPath||'calc/Libro de cálculo.tetord',messenger:'messenger/chats.json',filesMeta:'files/index.json'}; if(paths[k]) storage.writeJSON(paths[k],v).catch(()=>{}); }
    },
    async hydrateDiskState(){
      if(!storage.desktop)return;
      const calc=await storage.readJSON(this.calcPath || 'calc/Libro de cálculo.tetord',null); if(calc){this.calc=this.normalizeCalc(calc); this.commitCalc();}
      const chats=await storage.readJSON('messenger/chats.json',null); if(Array.isArray(chats)){this.chats=chats; this.save('messenger',chats);}
      const files=await storage.readJSON('files/index.json',null); if(files&&typeof files==='object'){localStorage.setItem('tetord.office.filesMeta',JSON.stringify(files));}
    },
    open(app){if(app==='writer'){this.close();return} this.app=app;document.body.classList.add('office-mode');document.body.classList.remove('office-calc-mode','office-slides-mode');if(app==='calc')document.body.classList.add('office-calc-mode');if(app==='slides')document.body.classList.add('office-slides-mode');this.shell.hidden=false;const map={home:'officeHome',files:'officeFiles',calc:'officeCalc',slides:'officeSlides',messenger:'officeMessenger'};Object.entries(map).forEach(([k,id])=>{const e=document.getElementById(id);if(e)e.hidden=k!==app});this.shell.querySelectorAll('[data-office-app]').forEach(b=>b.classList.toggle('is-active',b.dataset.officeApp===app));if(app==='home')this.renderHome();if(app==='files')this.renderFiles();if(app==='calc')this.renderCalc();if(app==='slides'){ if(this.slidesModule?.homeVisible) this.slidesModule.showHome(); else this.renderSlides(); }if(app==='messenger')this.renderMessenger()},
    close(){document.body.classList.remove('office-mode','office-calc-mode','office-slides-mode');if(this.shell)this.shell.hidden=true;this.app='writer';this.shell?.querySelectorAll('[data-office-app]').forEach(b=>b.classList.toggle('is-active',b.dataset.officeApp==='writer'));window.MiWord?.Editor?.focus?.()},
    col(n){let s='';n++;while(n){let r=(n-1)%26;s=String.fromCharCode(65+r)+s;n=Math.floor((n-1)/26)}return s},ref(c,r){return this.col(c)+(r+1)},
    normalizeCalc(v){
      const c=v&&typeof v==='object'?v:{}; const sheets=Array.isArray(c.sheets)&&c.sheets.length?c.sheets:[{name:'Hoja 1',rows:c.rows||20,cols:c.cols||12,data:c.data||{},styles:c.styles||{}}];
      sheets.forEach((s,i)=>{s.name=String(s.name||('Hoja '+(i+1)));s.rows=Math.max(1,Number(s.rows)||20);s.cols=Math.max(1,Number(s.cols)||12);s.data=s.data&&typeof s.data==='object'?s.data:{};s.styles=s.styles&&typeof s.styles==='object'?s.styles:{}});
      c.sheets=sheets;c.sheetIndex=Math.min(Math.max(0,Number(c.sheetIndex)||0),sheets.length-1);this.syncCalcSheet(c);return c;
    },
    syncCalcSheet(c=this.calc){const s=c.sheets[c.sheetIndex];c.rows=s.rows;c.cols=s.cols;c.data=s.data;c.styles=s.styles;c.selected=c.selected||'A1'},
    commitCalc(){const s=this.calc.sheets[this.calc.sheetIndex];s.rows=this.calc.rows;s.cols=this.calc.cols;s.data=this.calc.data;s.styles=this.calc.styles;this.save('calc',this.calc)},

    initHome(){
      document.getElementById('officeQuickWriter')?.addEventListener('click',()=>{this.close();window.MiWord?.DocumentManager?.newDocument?.()});
      document.getElementById('officeQuickCalc')?.addEventListener('click',()=>this.open('calc'));
      document.getElementById('officeQuickSlides')?.addEventListener('click',()=>{this.slidesModule?.showHome?.();this.open('slides');});
      document.getElementById('officeQuickFiles')?.addEventListener('click',()=>this.open('files'));
      document.getElementById('officeHomeRefresh')?.addEventListener('click',()=>this.renderHome());
      document.getElementById('officeAccountLogin')?.addEventListener('click',async()=>{const u=prompt('Usuario TETORD');const pw=u?prompt('Contraseña (mínimo 8 caracteres)'):'';if(!u||!pw)return;try{const user=await this.modules.cloudSync.login(u,pw);this.modules.collaboration.setToken(this.modules.cloudSync.token);this.modules.messenger.onlineAdapter.token=this.modules.cloudSync.token;this.renderHome();this.showToast?.('Sesión iniciada.');}catch(e){alert(e.message||'No se pudo iniciar sesión.');}});
      document.getElementById('officeAccountRefresh')?.addEventListener('click',async()=>{try{const user=await this.modules.cloudSync.refreshToken();this.renderHome();this.showToast?.('Sesión renovada.');}catch(e){alert(e.message||'No se pudo renovar la sesión.');}});
      document.getElementById('officeAccountLogout')?.addEventListener('click',()=>{this.modules.cloudSync.logout();this.modules.collaboration.setToken('');this.modules.messenger.onlineAdapter.token='';this.renderHome();});
    },
    initFiles(){
      document.getElementById('officeFilesRefresh')?.addEventListener('click',()=>this.renderFiles());
      document.getElementById('officeFilesSearch')?.addEventListener('input',e=>{this.fileSearch=e.target.value;this.renderFilesList()});
      document.getElementById('officeFilesSort')?.addEventListener('change',e=>{this.fileSort=e.target.value;this.renderFilesList()});
      document.querySelectorAll('[data-file-filter]').forEach(b=>b.addEventListener('click',()=>{this.fileFilter=b.dataset.fileFilter;document.querySelectorAll('[data-file-filter]').forEach(x=>x.classList.toggle('is-active',x===b));this.renderFilesList()}));
      document.getElementById('officeFilesNewFolder')?.addEventListener('click',()=>this.createFolder());
      document.getElementById('officeFilesNewWriter')?.addEventListener('click',()=>this.newOfficeFile('writer'));
      document.getElementById('officeFilesNewCalc')?.addEventListener('click',()=>this.newOfficeFile('calc'));
      document.getElementById('officeFilesNewSlides')?.addEventListener('click',()=>this.newOfficeFile('slides'));
      document.getElementById('officeFilesList')?.addEventListener('contextmenu',e=>{const row=e.target.closest('[data-file-index]');if(row){e.preventDefault();this.selectFile(Number(row.dataset.fileIndex));}});
    },
    filesMeta(){const f=this.load('filesMeta',null);return f&&typeof f==='object'?{folders:Array.isArray(f.folders)?f.folders:[],items:Array.isArray(f.items)?f.items:[]}: {folders:[],items:[]}},
    saveFilesMeta(v){this.save('filesMeta',v)},
    ensureFilesMeta(){const m=this.filesMeta();if(!m.folders.some(x=>x.id==='root'))m.folders.unshift({id:'root',name:'Este dispositivo',parentId:null,createdAt:new Date().toISOString()});this.saveFilesMeta(m);return m},
    makeFileId(prefix){return prefix+'-'+Date.now().toString(36)+'-'+Math.random().toString(36).slice(2,8)},
    folderPath(id){const m=this.ensureFilesMeta(),out=[];let cur=m.folders.find(x=>x.id===id);while(cur){out.unshift(cur);cur=m.folders.find(x=>x.id===cur.parentId)}return out.length?out:[m.folders[0]]},
    createFolder(){const m=this.ensureFilesMeta(),name=prompt('Nombre de la carpeta','Nueva carpeta');if(!name||!name.trim())return;const clean=name.trim().slice(0,60);if(m.folders.some(x=>x.parentId===this.fileFolder&&x.name.toLowerCase()===clean.toLowerCase())){alert('Ya existe una carpeta con ese nombre.');return}m.folders.push({id:this.makeFileId('folder'),name:clean,parentId:this.fileFolder,createdAt:new Date().toISOString()});this.saveFilesMeta(m);this.renderFiles()},
    newOfficeFile(type){
      const m=this.ensureFilesMeta();
      if(type==='writer'){this.close();window.MiWord?.DocumentManager?.newDocument?.();return}
      if(type==='calc'){this.calcPath=`calc/${Date.now().toString(36)}-${crypto.randomUUID().slice(0,6)}.tetord`;this.calc=this.normalizeCalc({id:crypto.randomUUID(),name:'Libro de cálculo',rows:20,cols:12,data:{},styles:{},selected:'A1',sheetIndex:0,sheets:[{name:'Hoja 1',rows:20,cols:12,data:{},styles:{}}]});this.save('calc',this.calc);this.open('calc');return}
      if(this.slidesModule){ this.slidesModule.createFromTemplate('blank'); } else { this.save('slides',this.slides); this.open('slides'); }
    },
    fileRecords(){ return this.modules?.files?.records?.() || []; },
    openFile(item){ if (item?.isFolder) { this.modules.files.currentFolder=item.path; this.renderFiles(); return; } return this.modules.files.openRecord(item); },
    renameFile(item){ return this.modules.files.renameRecord(item); },
    toggleFavorite(item){ return this.modules.files.toggleFavorite(item); },
    deleteFile(item){ return this.modules.files.deleteRecord(item); },
    duplicateFile(item){ return this.modules.files.duplicateRecord(item); },
    moveFile(){ this.showToast?.('Mover archivos se gestiona mediante carpetas reales en Archivos.'); },
    renderHome(){
      const account=this.modules?.cloudSync; const status=document.getElementById('officeAccountStatus'), name=document.getElementById('officeAccountName'), login=document.getElementById('officeAccountLogin'), refresh=document.getElementById('officeAccountRefresh'), logout=document.getElementById('officeAccountLogout');
      if(status&&name&&login&&refresh&&logout){ const has=!!account?.token; status.textContent=has?'Sesión cloud disponible':'Modo local · sin sesión'; name.textContent=has?'Sesión autenticada':'Trabajando sin conexión'; login.hidden=has; refresh.hidden=!has; logout.hidden=!has; }
      const box=document.getElementById('officeHomeRecent'); if(!box)return;
      const rows=this.fileRecords().filter(f=>f.updatedAt).sort((a,b)=>String(b.updatedAt).localeCompare(String(a.updatedAt))).slice(0,6);
      while(box.firstChild) box.removeChild(box.firstChild);
      if(!rows.length){box.appendChild(this.modules.files.el('p',{className:'office-empty',text:'No hay archivos recientes todavía.'}));return;}
      rows.forEach(f=>{const b=this.modules.files.button('',`office-recent-item office-recent-item--${f.type}`);b.append(this.modules.files.el('span',{className:'office-recent-icon',text:({writer:'W',calc:'C',slides:'S'}[f.type]||'M')}),this.modules.files.el('span',{className:'office-recent-main',children:[this.modules.files.el('strong',{text:f.name}),this.modules.files.el('small',{text:`${f.label} · TETORD · ${f.updatedAt?new Date(f.updatedAt).toLocaleString('es-MX'):''}`})]}),this.modules.files.el('span',{className:'office-recent-arrow',text:'›'}));b.addEventListener('click',()=>this.openFile(f));box.appendChild(b);});
    },
    async renderFiles(){
      const input=document.getElementById('officeFilesSearch'); if(input){input.value=this.fileSearch; this.modules.files.search=this.fileSearch;}
      const sort=document.getElementById('officeFilesSort'); if(sort){sort.value=this.fileSort; this.modules.files.sort=this.fileSort;}
      this.modules.files.filter=this.fileFilter; this.modules.files.currentFolder=this.fileFolder||''; await this.modules.files.refresh(); this.modules.files.render();
      this.fileFolder=this.modules.files.currentFolder; const status=document.getElementById('officeFilesStatus'); if(status)status.textContent=`${this.modules.files.visibleRecords().length} elementos · ${this.modules.files.folderRecords().length} carpetas`;
    },
    renderFilesList(){ this.modules?.files?.render?.(); },
    fileMenu(item){ return this.modules.files.fileMenu(item); },
    renderFileBreadcrumbs(){ return this.modules.files.renderBreadcrumbs(); },
    touchOfficeMeta(kind){ const paths={calc:this.calcPath||'calc/Libro de cálculo.tetord',slides:this.slidesModule?.storagePath||'slides/Presentación.tetord'}; if(paths[kind]) this.modules?.files?.refresh?.(); },
    initCalc(){
      document.getElementById('calcNewBtn')?.addEventListener('click',()=>{this.calcPath=`calc/${Date.now().toString(36)}-${crypto.randomUUID().slice(0,6)}.tetord`;this.calc=this.normalizeCalc({rows:20,cols:12,data:{},styles:{},selected:'A1',sheetIndex:0,sheets:[{name:'Hoja 1',rows:20,cols:12,data:{},styles:{}}]});this.save('calc',this.calc);this.touchOfficeMeta('calc');this.renderCalc()});
      document.getElementById('calcAddRowBtn')?.addEventListener('click',()=>{this.calc.rows++;this.commitCalc();this.renderCalc()});
      document.getElementById('calcAddColBtn')?.addEventListener('click',()=>{this.calc.cols++;this.commitCalc();this.renderCalc()});
      document.getElementById('calcExportBtn')?.addEventListener('click',()=>this.exportCalc());
      document.getElementById('calcImportBtn')?.addEventListener('click',()=>document.getElementById('calcImportInput')?.click());
      document.getElementById('calcImportInput')?.addEventListener('change',e=>this.importCalc(e.target.files[0]));
      document.getElementById('calcFormulaInput')?.addEventListener('input',e=>{this.calc.data[this.calc.selected]=e.target.value;this.commitCalc();this.renderCalc(false)});
      document.getElementById('calcFormulaInput')?.addEventListener('keydown',e=>{if(e.key==='Enter'){e.preventDefault();this.renderCalc();const i=document.querySelector('#calcGrid input[data-ref="'+CSS.escape(this.calc.selected)+'"]');i?.focus()}});
      document.getElementById('calcSheets')?.addEventListener('click',e=>{const b=e.target.closest('[data-sheet-index]');if(!b)return;this.calc.sheetIndex=Number(b.dataset.sheetIndex);this.calc.selected='A1';this.syncCalcSheet();this.commitCalc();this.renderCalc()});
      document.getElementById('calcSheets')?.addEventListener('dblclick',e=>{const b=e.target.closest('[data-sheet-index]');if(!b)return;const idx=Number(b.dataset.sheetIndex),n=prompt('Nombre de la hoja',this.calc.sheets[idx].name);if(n&&n.trim()){this.calc.sheets[idx].name=n.trim().slice(0,40);this.commitCalc();this.renderCalc()}});
      document.getElementById('calcFillColor')?.addEventListener('input',e=>this.formatSelected({backgroundColor:e.target.value}));
      document.getElementById('calcNumberFormat')?.addEventListener('change',e=>this.formatSelected({numberFormat:e.target.value}));
      document.querySelectorAll('[data-calc-format]').forEach(b=>b.addEventListener('click',()=>this.formatSelected({[b.dataset.calcFormat]:true})));
      document.querySelectorAll('[data-calc-align]').forEach(b=>b.addEventListener('click',()=>this.formatSelected({align:b.dataset.calcAlign})));
      const add=document.getElementById('calcAddSheetBtn'); add?.addEventListener('click',()=>this.addSheet());
    },
    addSheet(){const n='Hoja '+(this.calc.sheets.length+1);this.calc.sheets.push({name:n,rows:20,cols:12,data:{},styles:{}});this.calc.sheetIndex=this.calc.sheets.length-1;this.syncCalcSheet();this.calc.selected='A1';this.commitCalc();this.renderCalc()},
    formatSelected(style){const ref=this.calc.selected;this.calc.styles[ref]={...(this.calc.styles[ref]||{}),...style};this.commitCalc();this.renderCalc(false)},
    formatValue(v,fmt){if(fmt==='currency'&&v!==''&&!isNaN(Number(v)))return new Intl.NumberFormat('es-MX',{style:'currency',currency:'MXN'}).format(Number(v));if(fmt==='percent'&&v!==''&&!isNaN(Number(v)))return (Number(v)*100).toFixed(2)+'%';if(fmt==='number'&&v!==''&&!isNaN(Number(v)))return new Intl.NumberFormat('es-MX',{maximumFractionDigits:6}).format(Number(v));return String(v)},
    value(ref,stack=[]){const raw=String(this.calc.data[ref]??'');if(raw.charAt(0)!=='=')return this.formatValue(raw,this.calc.styles[ref]?.numberFormat);if(stack.includes(ref))return '#CICLO!';let x=raw.slice(1).toUpperCase().trim();
      x=x.replace(/(SUM|AVERAGE|MIN|MAX|COUNT)\(([A-Z]+\d+):([A-Z]+\d+)\)/g,(_,fn,a,b)=>String(this.rangeFn(fn,a,b,stack.concat(ref))));
      x=x.replace(/IF\(([^,]+),([^,]+),([^\)]+)\)/g,(_,cond,a,b)=>{try{return Function('"use strict";return ('+this.replaceRefs(cond,stack.concat(ref))+')')()?a:b}catch(_){return b}});
      x=this.replaceRefs(x,stack.concat(ref));
      if(!/^[0-9+\-*/().%\s]+$/.test(x))return '#VALOR!';try{const v=Function('"use strict";return ('+x+')')();const out=Number.isFinite(v)?String(Math.round(v*1e10)/1e10):'#VALOR!';return this.formatValue(out,this.calc.styles[ref]?.numberFormat)}catch(_){return '#VALOR!'}},
    replaceRefs(x,stack){return x.replace(/([A-Z]+\d+)/g,(_,r)=>{const v=Number(this.value(r,stack));return Number.isFinite(v)?v:0})},
    rangeCells(a,b){const p=a.match(/^([A-Z]+)(\d+)$/),q=b.match(/^([A-Z]+)(\d+)$/);if(!p||!q)return[];const ci=s=>{let n=0;for(const c of s)n=n*26+c.charCodeAt(0)-64;return n-1};const out=[];for(let r=+p[2]-1;r<=+q[2]-1;r++)for(let c=ci(p[1]);c<=ci(q[1]);c++)out.push(this.ref(c,r));return out},
    rangeFn(fn,a,b,stack){const vals=this.rangeCells(a,b).map(r=>Number(this.value(r,stack))).filter(Number.isFinite);if(fn==='SUM')return vals.reduce((a,v)=>a+v,0);if(fn==='AVERAGE')return vals.length?vals.reduce((a,v)=>a+v,0)/vals.length:0;if(fn==='MIN')return vals.length?Math.min(...vals):0;if(fn==='MAX')return vals.length?Math.max(...vals):0;if(fn==='COUNT')return vals.length;return 0},
    renderSheets(){const box=document.getElementById('calcSheets');if(!box)return;box.innerHTML=this.calc.sheets.map((s,i)=>'<button type="button" role="tab" data-sheet-index="'+i+'" class="calc-sheet-tab '+(i===this.calc.sheetIndex?'is-active':'')+'">'+this.esc(s.name)+'</button>').join('')+'<button id="calcAddSheetBtn" class="calc-sheet-add" type="button" title="Agregar hoja">＋</button>';document.getElementById('calcAddSheetBtn').addEventListener('click',()=>this.addSheet())},
    renderCalc(focus=true){const t=document.getElementById('calcGrid');if(!t)return;this.syncCalcSheet();this.renderSheets();t.innerHTML='';let h='<thead><tr><th></th>';for(let c=0;c<this.calc.cols;c++)h+='<th>'+this.col(c)+'</th>';h+='</tr></thead><tbody>';for(let r=0;r<this.calc.rows;r++){h+='<tr><th>'+(r+1)+'</th>';for(let c=0;c<this.calc.cols;c++){const ref=this.ref(c,r),raw=this.calc.data[ref]??'',st=this.calc.styles[ref]||{},css='font-weight:'+(st.bold?'700':'400')+';font-style:'+(st.italic?'italic':'normal')+';text-align:'+(st.align||'left')+';background:'+(st.backgroundColor||'transparent')+';';h+='<td style="'+css+'"><input data-ref="'+ref+'" value="'+this.esc(raw)+'" title="'+ref+'"><span class="calc-value">'+this.esc(this.value(ref))+'</span></td>'}h+='</tr>'}t.innerHTML=h+'</tbody>';t.querySelectorAll('input').forEach(i=>{i.addEventListener('focus',()=>{this.calc.selected=i.dataset.ref;document.getElementById('calcNameBox').textContent=i.dataset.ref;document.getElementById('calcFormulaInput').value=this.calc.data[i.dataset.ref]??'';const st=this.calc.styles[i.dataset.ref]||{};document.getElementById('calcNumberFormat').value=st.numberFormat||'general';document.getElementById('calcFillColor').value=st.backgroundColor||'#fff3b8'});i.addEventListener('input',()=>{this.calc.data[i.dataset.ref]=i.value;this.commitCalc();document.getElementById('calcFormulaInput').value=i.value})});document.getElementById('calcStatus').textContent=this.calc.sheets[this.calc.sheetIndex].name+' · '+this.calc.cols+' × '+this.calc.rows;if(focus){const i=t.querySelector('input[data-ref="'+CSS.escape(this.calc.selected)+'"]');i?.focus()}}
    ,exportCalc(){const rows=[];for(let r=0;r<this.calc.rows;r++){const row=[];for(let c=0;c<this.calc.cols;c++){let v=String(this.calc.data[this.ref(c,r)]??'').replace(/"/g,'""');row.push(/[",\n]/.test(v)?'"'+v+'"':v)}rows.push(row.join(','))}this.download(this.calc.sheets[this.calc.sheetIndex].name+'.csv',rows.join('\n'),'text/csv')},
    async importCalc(file){if(!file)return;const lines=(await file.text()).split(/\r?\n/).filter(Boolean);const rows=Math.max(1,lines.length),cols=Math.max(1,...lines.map(x=>x.split(',').length));const sheet={name:file.name.replace(/\.[^.]+$/,'').slice(0,40)||'Importada',rows,cols,data:{},styles:{}};lines.forEach((line,r)=>line.split(',').forEach((v,c)=>sheet.data[this.ref(c,r)]=v.replace(/^"|"$/g,'').replace(/""/g,'"')));this.calc.sheets.push(sheet);this.calc.sheetIndex=this.calc.sheets.length-1;this.syncCalcSheet();this.save('calc',this.calc);this.touchOfficeMeta('calc');this.renderCalc()},
    initMessenger(){
      const module=this.modules?.messenger;
      document.getElementById('messengerComposer')?.addEventListener('submit',async e=>{e.preventDefault();const i=document.getElementById('messengerInput'),v=i.value.trim();if(!v)return;const ok=await module?.sendMessage?.(this.chat,v);if(ok!==false)i.value='';this.renderMessenger();});
      document.getElementById('messengerNewChatBtn')?.addEventListener('click',async()=>{const n=prompt('Nombre de la conversación','Nueva conversación');if(!n)return;const c=await module?.createConversation?.(n);if(c){this.chats=await module.listConversations();this.chat=Math.max(0,this.chats.length-1);this.renderMessenger();}});
    },
    updateMessengerStatus(detail={}){const head=document.getElementById('messengerConversationHead');if(!head)return;const status=head.querySelector('.messenger-status');if(status)status.textContent=detail.detail||detail.status||'';},
    renderMessenger(){const list=document.getElementById('messengerChats'),msgs=document.getElementById('messengerMessages'),head=document.getElementById('messengerConversationHead');if(!list||!msgs||!head)return;while(list.firstChild)list.removeChild(list.firstChild);const chats=this.modules?.messenger?.chats||this.chats||[];this.chats=chats;if(this.chat>=chats.length)this.chat=Math.max(0,chats.length-1);chats.forEach((c,i)=>{const b=document.createElement('button');b.className='messenger-chat'+(i===this.chat?' is-active':'');b.type='button';const strong=document.createElement('strong');strong.textContent=String(c.name||'Conversación');const small=document.createElement('small');const last=c.messages?.length?c.messages[c.messages.length-1].text:'Sin mensajes';small.textContent=String(last||'');b.append(strong,small);b.addEventListener('click',()=>{this.chat=i;this.renderMessenger()});list.appendChild(b)});while(head.firstChild)head.removeChild(head.firstChild);const c=chats[this.chat]||{name:'Messenger',messages:[]};const title=document.createElement('strong');title.textContent=String(c.name||'Conversación');const status=document.createElement('small');status.className='messenger-status';status.textContent=this.modules?.messenger?.onlineStatus==='online'?'En línea':this.modules?.messenger?.onlineStatus==='connecting'?'Conectando…':'Modo local';head.append(title,status);while(msgs.firstChild)msgs.removeChild(msgs.firstChild);(c.messages||[]).forEach(m=>{const d=document.createElement('div');d.className='message-bubble message-bubble--'+(m.from==='me'?'me':'other');d.textContent=String(m.text||'');msgs.appendChild(d)});msgs.scrollTop=msgs.scrollHeight},
    showToast(message,type='info',timeout=3200){const region=document.getElementById('tetordToastRegion');if(!region)return;const el=document.createElement('div');el.className='tetord-toast tetord-toast--'+type;el.textContent=String(message);region.appendChild(el);setTimeout(()=>el.remove(),timeout)},
    updateCollabUI(detail={}){const status=String(detail.status||'offline');const labels={online:'Sincronizado',reconnecting:'Reconectando…',offline:'Sin conexión',local:'Solo local'};document.querySelectorAll('.collab-statusbar').forEach(bar=>{bar.dataset.networkState=status;const label=bar.querySelector('[data-network-label]');if(label)label.textContent=labels[status]||status;});if(Array.isArray(detail.users))this.renderPresence(detail.users);},
    renderPresence(users=[]){document.querySelectorAll('[data-presence-list]').forEach(box=>{while(box.firstChild)box.removeChild(box.firstChild);users.slice(0,8).forEach(u=>{const chip=document.createElement('span');chip.className='presence-chip';chip.title=String(u.nick||u.username||'Usuario');const av=document.createElement('span');av.className='presence-avatar';av.textContent=String(u.nick||u.username||'?').trim().slice(0,1).toUpperCase();const name=document.createElement('span');name.className='presence-name';name.textContent=String(u.nick||u.username||'Usuario');chip.append(av,name);box.appendChild(chip)});});},
    scheduleWriterRemote(content){if(this._writerRemoteFrame)cancelAnimationFrame(this._writerRemoteFrame);this._writerRemoteFrame=requestAnimationFrame(()=>{this._writerRemoteFrame=null;if(window.MiWord?.Editor?.el&&window.MiWord.Editor.el.innerHTML!==content)window.MiWord.Editor.load(content);});},
    scheduleCalcRemote(value){if(this._calcRemoteFrame)cancelAnimationFrame(this._calcRemoteFrame);this._calcRemoteFrame=requestAnimationFrame(()=>{this._calcRemoteFrame=null;try{const next=this.normalizeCalc(JSON.parse(value));if(JSON.stringify(this.calc)!==JSON.stringify(next)){this.calc=next;this.commitCalc();this.renderCalc(false);}}catch(_){} });},
    esc(v){return String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]))},
    download(name,text,type){const a=document.createElement('a'),u=URL.createObjectURL(new Blob([text],{type:type+';charset=utf-8'}));a.href=u;a.download=name;document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(u),500)}
  };
  window.TETORDOffice=O;
  window.addEventListener('DOMContentLoaded', async ()=>{
    O.init();
    O.modules.writer = new WriterModule({ office: O });
    O.modules.calc = new CalcModule({ office: O });
    O.modules.files = new FilesModule({ office: O });
    O.modules.messenger = new MessengerModule({ office: O });
    O.modules.cloudSync = new CloudSyncModule({ office: O });
    O.modules.slides = new SlidesModule({ office: O });
    O.modules.collaboration = new CollaborationModule({ office: O });
    O.slidesModule = O.modules.slides;
    await Promise.all([O.hydrateDiskState(), O.slidesModule.init(), O.modules.files.init(), window.MiWord?.LocalDocuments?.hydrateNative?.()]);
    O.slides = O.slidesModule.state;
    await O.modules.messenger.load();
    O.modules.collaboration.setToken(localStorage.getItem('tetord.cloud.jwt') || '');
    O.modules.collaboration.setListeners({ onStatus: detail => O.updateCollabUI(detail), onCursors: users => O.renderPresence(Array.isArray(users)?users:[]) });
    document.getElementById('editor')?.addEventListener('input', () => { if (O.modules.collaboration.docId && O.modules.collaboration.module === 'writer') O.modules.collaboration.submitText(document.getElementById('editor').innerHTML); });
    document.getElementById('calcGrid')?.addEventListener('input', () => { if (O.modules.collaboration.docId && O.modules.collaboration.module === 'calc') O.modules.collaboration.submitCalcSnapshot(O.calc); });
    if(O.modules.messenger.onlineEnabled) O.modules.messenger.connectOnline().catch(()=>{});
    await O.renderFiles();
    O.modules.collaboration.listeners = O.modules.collaboration.listeners || {};
    O.modules.collaboration.setListeners({ onChange: change => { if (change?.users) O.renderPresence(change.users); if (change?.type === 'remote' && change.content != null) { if (O.modules.collaboration.module === 'writer') O.scheduleWriterRemote(change.content); if (O.modules.collaboration.module === 'calc') O.scheduleCalcRemote(change.content); O.showToast('Documento actualizado por otro usuario.','info',1800); } if(change?.type==='resync'&&change.content!=null){ if(O.modules.collaboration.module==='writer')O.scheduleWriterRemote(change.content); if(O.modules.collaboration.module==='calc')O.scheduleCalcRemote(change.content); O.showToast('Documento resincronizado.','success'); } }, onStatus: detail => O.updateCollabUI(detail), onCursors: users => O.renderPresence(Array.isArray(users)?users:[]) });
    O.renderSlides = (...args)=>O.slidesModule.renderSlides(...args);
    O.renderPresent = (...args)=>O.slidesModule.renderPresent(...args);
    O.present = (...args)=>O.slidesModule.present(...args);
    O.slideNavigate = (...args)=>O.slidesModule.slideNavigate(...args);
    if(O.app==='home') O.renderHome();
  });
})();

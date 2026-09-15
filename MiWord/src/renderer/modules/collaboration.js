import { AppModule } from './app-module.js';

/**
 * Lightweight revisioned collaboration client.
 * Writer uses the document HTML string as the collaborative text stream;
 * Calc uses a JSON snapshot. The server serializes operations and broadcasts
 * remote cursor/selection metadata. Local editing never depends on this module.
 */
export class CollaborationModule extends AppModule {
  constructor(ctx = {}) {
    super(ctx);
    const config = window.TETORD_COLLABORATION || {};
    this.enabled = config.enabled !== false;
    this.url = String(config.url || 'ws://127.0.0.1:8787');
    this.token = String(config.token || localStorage.getItem('tetord.cloud.jwt') || '');
    this.socket = null;
    this.connected = false;
    this.docId = null;
    this.module = null;
    this.rev = 0;
    this.content = '';
    this.pending = [];
    this.outbox = [];
    this.inFlight = new Map();
    this.reconnect = true;
    this.retryCount = 0;
    this.reconnectTimer = null;
    this.maxRetries = Infinity;
    this.cursors = new Map();
    this.clientId = globalThis.crypto?.randomUUID?.() || `${Date.now()}-${Math.random()}`;
    this.listeners = { onChange: null, onStatus: null, onCursors: null };
    this.lastRemoteSignature = '';
    this.lastCursorSent = 0;
    this.cursorThrottleMs = 50;
  }

  setListeners(listeners = {}) { Object.assign(this.listeners, listeners); }
  emitStatus(status, detail = '') { this.listeners.onStatus?.({ status, detail }); }
  emitChange(change) { this.listeners.onChange?.(change); }
  setToken(token) { this.token = String(token || ''); }

  async connect() {
    if (!this.enabled || !this.token || typeof WebSocket === 'undefined') throw new Error('Colaboración no disponible.');
    if (this.socket?.readyState === WebSocket.OPEN) return true;
    return new Promise((resolve, reject) => {
      const ws = new WebSocket(this.url);
      this.socket = ws;
      const timer = setTimeout(() => { try { ws.close(); } catch (_) {} reject(new Error('Tiempo de espera agotado.')); }, 6000);
      ws.addEventListener('open', () => {
        clearTimeout(timer); this.connected = true;
        this.retryCount = 0;
        ws.send(JSON.stringify({ type: 'auth', token: this.token, room: 'general' }));
      }, { once: true });
      ws.addEventListener('message', event => {
        let packet; try { packet = JSON.parse(event.data); } catch (_) { return; }
        if (packet.type === 'auth_ok') { this.emitStatus('online', 'Colaboración en línea'); resolve(true); }
        this.handlePacket(packet);
      });
      ws.addEventListener('error', () => { clearTimeout(timer); this.emitStatus('offline', 'Servidor no disponible'); reject(new Error('No se pudo conectar.')); }, { once: true });
      ws.addEventListener('close', () => { clearTimeout(timer); this.connected = false; this.socket = null; this.emitStatus('offline', 'Colaboración desconectada'); });
    });
  }

  handlePacket(p) {
    if (p.type === 'collab_state' && p.docId === this.docId) {
      this.rev = Number(p.rev) || 0; this.content = String(p.content ?? '');
      this.lastRemoteSignature = `${this.rev}:${this.content.length}`;
      this.emitChange({ type: 'state', content: this.content, rev: this.rev, users: p.users || [] });
      return;
    }
    if (p.type === 'collab_op' && p.docId === this.docId) {
      const opId = String(p.op?.opId || '');
      if (opId && this.inFlight.has(opId)) { this.inFlight.delete(opId); }
      this.rev = Number(p.rev) || this.rev;
      if (String(p.userId || '') === String(this.userId || this.clientId)) return;
      this.applyRemoteOp(p.op);
      return;
    }
    if (p.type === 'collab_ack') { const id=String(p.opId||''); if(id) this.inFlight.delete(id); this.rev = Math.max(this.rev, Number(p.rev) || 0); this.flush(); return; }
    if (p.type === 'collab_resync' && p.docId === this.docId) { this.rev=Number(p.rev)||0; this.content=String(p.content??''); this.inFlight.clear(); this.outbox.length=0; this.emitChange({type:'resync',content:this.content,rev:this.rev}); return; }
    if (p.type === 'cursor' && p.docId === this.docId && p.userId) {
      this.cursors.set(p.userId, p.cursor || {}); this.listeners.onCursors?.([...this.cursors.entries()]);
      return;
    }
    if (p.type === 'collab_presence' && p.docId === this.docId) this.listeners.onCursors?.(p.users || []);
  }

  async joinDocument(docId, module = 'writer', initialContent = '') {
    if (!this.enabled) return false;
    if (!this.connected) await this.connect();
    this.docId = String(docId || 'untitled').slice(0, 160);
    this.userId = String(this.userId || this.clientId);
    this.module = module === 'calc' ? 'calc' : 'writer';
    this.content = String(initialContent ?? ''); this.rev = 0; this.pending.length = 0; this.cursors.clear();
    this.send({ type: 'collab_join', docId: this.docId, module: this.module, content: this.content });
    return true;
  }

  send(packet, { queue = true } = {}) {
    if (this.socket?.readyState === WebSocket.OPEN) {
      try { this.socket.send(JSON.stringify(packet)); return true; } catch (_) {}
    }
    if (queue) this.pending.push(packet);
    return false;
  }

  flush() {
    if (this.socket?.readyState !== WebSocket.OPEN) return;
    while (this.pending.length) {
      const packet = this.pending.shift();
      if (!this.send(packet, { queue: false })) { this.pending.unshift(packet); break; }
    }
    while (this.outbox.length) {
      const packet = this.outbox.shift();
      if (!this.send(packet, { queue: false })) { this.outbox.unshift(packet); break; }
    }
  }

  queueReconnect() {
    if (!this.reconnect || this.reconnectTimer || this.retryCount >= this.maxRetries || !this.docId) return;
    const delay = Math.min(30000, 500 * (2 ** this.retryCount));
    this.retryCount += 1;
    this.reconnectTimer = setTimeout(() => {
      this.reconnectTimer = null;
      this.connect().then(() => this.rejoinAfterReconnect()).catch(() => this.queueReconnect());
    }, delay);
    this.emitStatus('reconnecting', `Reintentando colaboración en ${Math.ceil(delay / 1000)} s`);
  }

  async rejoinAfterReconnect() {
    if (!this.docId) return;
    const docId = this.docId, module = this.module, content = this.content;
    this.send({ type: 'collab_join', docId, module, content }, { queue: false });
    this.flush();
  }

  submitText(nextContent) {
    if (!this.docId || this.module !== 'writer') return false;
    const next = String(nextContent ?? ''), prev = this.content;
    if (next === prev) return false;
    const prefix = commonPrefix(prev, next), suffix = commonSuffix(prev, next, prefix);
    const deleted = prev.slice(prefix, prev.length - suffix);
    const inserted = next.slice(prefix, next.length - suffix);
    const op = { kind: 'text', index: prefix, deleteCount: deleted.length, insertText: inserted, clientId: this.clientId, opId: crypto.randomUUID() };
    this.content = next;
    const packet = { type: 'collab_op', docId: this.docId, baseRev: this.rev, op };
    this.outbox.push(packet); this.inFlight.set(op.opId, packet); this.flush();
    return true;
  }

  submitCalcSnapshot(nextValue) {
    if (!this.docId || this.module !== 'calc') return false;
    const next = typeof nextValue === 'string' ? nextValue : JSON.stringify(nextValue);
    if (next === this.content) return false;
    this.content = next;
    const op = { kind: 'replace', value: next, clientId: this.clientId, opId: crypto.randomUUID() };
    const packet = { type: 'collab_op', docId: this.docId, baseRev: this.rev, op };
    this.outbox.push(packet); this.inFlight.set(op.opId, packet); this.flush();
    return true;
  }

  applyRemoteOp(op = {}) {
    if (op.kind === 'replace') this.content = String(op.value ?? '');
    if (op.kind === 'text') this.content = this.content.slice(0, op.index) + String(op.insertText || '') + this.content.slice(op.index + Number(op.deleteCount || 0));
    this.emitChange({ type: 'remote', content: this.content, rev: this.rev, op });
  }

  sendCursor(cursor = {}) {
    if (!this.docId) return;
    const now=Date.now(); if(now-this.lastCursorSent<this.cursorThrottleMs)return; this.lastCursorSent=now; this.send({ type: 'cursor', docId: this.docId, cursor: { start: Number(cursor.start) || 0, end: Number(cursor.end) || 0 } });
  }

  leaveDocument() { if (this.docId) this.send({ type: 'collab_leave', docId: this.docId }, { queue: false }); this.docId = null; this.module = null; this.rev = 0; this.pending.length = 0; this.outbox.length = 0; this.inFlight.clear(); this.cursors.clear(); }
  async disconnect() { this.reconnect = false; if (this.reconnectTimer) { clearTimeout(this.reconnectTimer); this.reconnectTimer = null; } this.leaveDocument(); try { this.socket?.close(1000, 'Client disconnect'); } catch (_) {} this.connected = false; this.socket = null; }
}

function commonPrefix(a, b) { const n = Math.min(a.length, b.length); let i = 0; while (i < n && a[i] === b[i]) i++; return i; }
function commonSuffix(a, b, prefix) { let i = 0; while (i < a.length - prefix && i < b.length - prefix && a[a.length - 1 - i] === b[b.length - 1 - i]) i++; return i; }

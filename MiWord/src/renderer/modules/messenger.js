import { AppModule } from './app-module.js';

const DEFAULT_URL = 'ws://127.0.0.1:8787';
const DEFAULT_NICK = 'TETORD User';
const STORAGE_PATH = 'messenger/chats.json';

/**
 * Transport boundary used by the Messenger UI.
 * Implementations expose the same async API so the UI does not depend on the network transport.
 */
export class MessengerAdapter {
  constructor() { this.onChange = null; this.onStatus = null; }
  setListeners({ onChange, onStatus } = {}) { this.onChange = onChange || null; this.onStatus = onStatus || null; return this; }
  emitChange() { this.onChange?.(); }
  emitStatus(status, detail = '') { this.onStatus?.({ status, detail }); }
  async connect() { throw new Error('MessengerAdapter.connect() no implementado'); }
  async disconnect() {}
  async listConversations() { throw new Error('MessengerAdapter.listConversations() no implementado'); }
  async sendMessage() { throw new Error('MessengerAdapter.sendMessage() no implementado'); }
  async createConversation() { throw new Error('MessengerAdapter.createConversation() no implementado'); }
}

export class LocalMessengerAdapter extends MessengerAdapter {
  constructor(storage) { super(); this.storage = storage; this.path = STORAGE_PATH; this.chats = []; }

  async connect() {
    this.chats = await this.storage.readJSON(this.path, null);
    if (!Array.isArray(this.chats) || !this.chats.length) {
      this.chats = [{ id: crypto.randomUUID(), name: 'Notas de TETORD', messages: [{ id: crypto.randomUUID(), from: 'other', text: 'Messenger local listo. Esta conversación permanece en este dispositivo.', at: Date.now() }] }];
      await this.persist();
    }
    this.emitStatus('local', 'Modo local');
    this.emitChange();
    return this.chats;
  }

  async listConversations() { return this.chats; }

  async persist() { return this.storage.writeJSON(this.path, this.chats); }

  async sendMessage(index, text) {
    const chat = this.chats[index];
    if (!chat || !String(text || '').trim()) return false;
    chat.messages.push({ id: crypto.randomUUID(), from: 'me', text: String(text).trim().slice(0, 4000), at: Date.now() });
    const result = await this.persist();
    this.emitChange();
    return !!result?.ok;
  }

  async createConversation(name) {
    const chat = { id: crypto.randomUUID(), name: String(name || 'Nueva conversación').trim().slice(0, 80) || 'Nueva conversación', messages: [] };
    this.chats.push(chat);
    await this.persist();
    this.emitChange();
    return chat;
  }
}

/**
 * WebSocket transport for the future/current Messenger backend.
 * Protocol:
 * client -> auth {type:'auth', nick, token, room}
 * client -> send {type:'message', room, clientId, text}
 * server -> message {type:'message', room, message}
 * server -> history {type:'history', room, messages}
 */
export class WebSocketMessengerAdapter extends MessengerAdapter {
  constructor({ url = DEFAULT_URL, nick = DEFAULT_NICK, token = '', room = 'general', reconnect = true, maxRetries = Infinity } = {}) {
    super();
    this.url = url;
    this.httpUrl = this.url.replace(/^ws/, 'http').replace(/\/$/, '');
    this.nick = nick || DEFAULT_NICK;
    this.token = token || localStorage.getItem('tetord.cloud.jwt') || '';
    this.room = room || 'general';
    this.reconnect = reconnect;
    this.maxRetries = maxRetries;
    this.socket = null;
    this.connected = false;
    this.authenticated = false;
    this.retryCount = 0;
    this.reconnectTimer = null;
    this.pending = [];
    this.presence = [];
    this.directMessages = new Map();
    this.seenMessageIds = new Set();
    this.seenDmIds = new Set();
    this.chats = [{ id: this.room, name: this.room, messages: [] }];
    this.connectPromise = null;
  }

  get chat() { return this.chats[0]; }

  async connect() {
    if (this.connected && this.authenticated) return this.chats;
    if (this.connectPromise) return this.connectPromise;
    this.connectPromise = new Promise((resolve, reject) => {
      if (typeof WebSocket === 'undefined') { reject(new Error('WebSocket no disponible.')); return; }
      this.emitStatus('connecting', `Conectando a ${this.url}`);
      let settled = false;
      const socket = new WebSocket(this.url);
      this.socket = socket;
      const fail = (error) => { if (!settled) { settled = true; reject(error); } };
      socket.addEventListener('open', () => {
        this.connected = true;
        this.retryCount = 0;
        this.emitStatus('online', 'Conectado');
        socket.send(JSON.stringify({ type: 'auth', nick: this.nick, token: this.token, room: this.room }));
      });
      socket.addEventListener('message', event => {
        let packet;
        try { packet = JSON.parse(event.data); } catch (_) { return; }
        this.handlePacket(packet, resolve, fail);
      });
      socket.addEventListener('error', () => {
        this.emitStatus('offline', 'No se pudo conectar al servidor');
        fail(new Error('No se pudo conectar al servidor WebSocket.'));
      });
      socket.addEventListener('close', () => {
        this.connected = false;
        this.authenticated = false;
        this.socket = null;
        this.connectPromise = null;
        this.emitStatus('offline', 'Servidor desconectado');
        if (this.reconnect) this.scheduleReconnect();
      });
      setTimeout(() => fail(new Error('Tiempo de espera agotado al conectar.')), 6000);
    }).finally(() => { this.connectPromise = null; });
    return this.connectPromise;
  }

  handlePacket(packet, resolve, reject) {
    if (packet.type === 'auth_ok') {
      this.authenticated = true;
      try { localStorage.setItem('tetord.cloud.jwt', this.token); } catch (_) {}
      const history = Array.isArray(packet.messages) ? packet.messages : [];
      this.chat.messages = history;
      this.flushPending();
      this.emitChange();
      if (resolve) resolve(this.chats);
      return;
    }
    if (packet.type === 'presence') { this.presence = Array.isArray(packet.users) ? packet.users : []; this.emitChange(); return; }
    if (packet.type === 'dm') { const m = packet.message; if (m && m.id && !this.seenDmIds.has(String(m.id))) { this.seenDmIds.add(String(m.id)); const key = String(m.fromUserId === this.userId ? m.toUserId : m.fromUserId); const list = this.directMessages.get(key) || []; if (!list.some(x => x.id === m.id)) list.push(m); this.directMessages.set(key, list); this.emitChange(); } return; }
    if (packet.type === 'dm_history') { const incoming = Array.isArray(packet.messages) ? packet.messages : []; const key = String(packet.toUserId); const current = this.directMessages.get(key) || []; const merged = new Map([...current, ...incoming].filter(m => m?.id).map(m => [String(m.id), m])); for (const m of incoming) if (m?.id) this.seenDmIds.add(String(m.id)); this.directMessages.set(key, [...merged.values()]); this.emitChange(); return; }
    if (packet.type === 'history') {
      if (Array.isArray(packet.messages)) { const merged = new Map(this.chat.messages.filter(m => m?.id).map(m => [String(m.id), m])); for (const m of packet.messages) { if (m?.id) this.seenMessageIds.add(String(m.id)); if (m?.id) merged.set(String(m.id), m); } this.chat.messages = [...merged.values()]; }
      this.emitChange();
      return;
    }
    if (packet.type === 'message' && packet.message) {
      if (packet.message.id && this.seenMessageIds.has(String(packet.message.id))) return;
      if (packet.message.id) this.seenMessageIds.add(String(packet.message.id));
      if (!this.chat.messages.some(m => m.id && packet.message.id && m.id === packet.message.id)) this.chat.messages.push(packet.message);
      this.emitChange();
      return;
    }
    if (packet.type === 'error') {
      this.emitStatus('error', packet.message || 'Error del servidor');
      if (packet.fatal && reject) reject(new Error(packet.message || 'Autenticación rechazada.'));
    }
  }

  scheduleReconnect() {
    if (this.reconnectTimer || this.retryCount >= this.maxRetries) return;
    const delay = Math.min(30000, 1000 * (2 ** this.retryCount));
    this.retryCount += 1;
    this.reconnectTimer = setTimeout(() => {
      this.reconnectTimer = null;
      this.connect().catch(() => {});
    }, delay);
    this.emitStatus('reconnecting', `Reintentando en ${Math.ceil(delay / 1000)} s`);
  }

  async flushPending() {
    if (!this.authenticated || !this.socket || this.socket.readyState !== WebSocket.OPEN) return;
    const queue = this.pending.splice(0);
    for (const item of queue) {
      try { this.socket.send(JSON.stringify(item)); } catch (_) { this.pending.unshift(item); break; }
    }
  }

  async sendMessage(_index, text) {
    const clean = String(text || '').trim().slice(0, 4000);
    if (!clean) return false;
    const packet = { type: 'message', room: this.room, clientId: crypto.randomUUID(), text: clean };
    if (!this.authenticated || !this.socket || this.socket.readyState !== WebSocket.OPEN) {
      this.pending.push(packet);
      this.emitStatus('offline', 'Mensaje en cola; se enviará al reconectar');
      return false;
    }
    try { this.socket.send(JSON.stringify(packet)); return true; } catch (_) { this.pending.push(packet); return false; }
  }

  async listConversations() { return this.chats; }

  async listUsers() {
    if (!this.token) return [];
    const response = await fetch(`${this.httpUrl}/users`, { headers: { authorization: `Bearer ${this.token}` } });
    if (!response.ok) throw new Error('No se pudieron cargar los usuarios.');
    const body = await response.json(); return body.users || [];
  }

  async loadDirectMessages(userId) {
    if (!this.authenticated) return [];
    this.sendPacket?.({ type: 'dm_history', toUserId: String(userId) });
    if (this.socket?.readyState === WebSocket.OPEN) this.socket.send(JSON.stringify({ type: 'dm_history', toUserId: String(userId) }));
    return this.directMessages.get(String(userId)) || [];
  }

  async sendDirectMessage(userId, text) {
    const clean = String(text || '').trim().slice(0, 4000); if (!clean) return false;
    const packet = { type: 'dm', toUserId: String(userId), clientId: crypto.randomUUID(), text: clean };
    if (!this.authenticated || !this.socket || this.socket.readyState !== WebSocket.OPEN) { this.pending.push(packet); return false; }
    this.socket.send(JSON.stringify(packet)); return true;
  }

  async createConversation(name) {
    const room = String(name || 'general').trim().toLowerCase().replace(/[^a-z0-9_-]+/g, '-').slice(0, 40) || 'general';
    this.room = room;
    this.chats = [{ id: room, name: room, messages: [] }];
    if (this.authenticated && this.socket?.readyState === WebSocket.OPEN) this.socket.send(JSON.stringify({ type: 'join', room }));
    this.emitChange();
    return this.chat;
  }

  async disconnect() {
    this.reconnect = false;
    if (this.reconnectTimer) { clearTimeout(this.reconnectTimer); this.reconnectTimer = null; }
    if (this.socket) this.socket.close(1000, 'Client disconnect');
    this.socket = null; this.connected = false; this.authenticated = false;
    this.emitStatus('offline', 'Desconectado');
  }
}

export class MessengerModule extends AppModule {
  constructor(ctx) {
    super(ctx);
    this.localAdapter = new LocalMessengerAdapter(this.storage);
    const config = window.TETORD_MESSENGER || {};
    this.onlineAdapter = new WebSocketMessengerAdapter(config);
    this.adapter = this.localAdapter;
    this.onlineEnabled = config.enabled !== false;
    this.onlineStatus = 'local';
    this.chats = [];
    this.chat = 0;
    this.bindAdapter(this.localAdapter);
  }

  bindAdapter(adapter) {
    adapter.setListeners({
      onChange: async () => { this.chats = await adapter.listConversations(); if (adapter === this.onlineAdapter && this.onlineAdapter.authenticated) this.adapter = this.onlineAdapter; this.office && (this.office.chats = this.chats); this.office?.renderMessenger?.(); },
      onStatus: detail => { this.onlineStatus = detail.status; if (adapter === this.onlineAdapter && detail.status === 'online') this.adapter = this.onlineAdapter; this.office?.updateMessengerStatus?.(detail); }
    });
  }

  async load() {
    this.bindAdapter(this.adapter);
    this.chats = await this.adapter.connect();
    this.office && (this.office.chats = this.chats);
    return this.chats;
  }

  async connectOnline() {
    this.bindAdapter(this.onlineAdapter);
    try {
      this.chats = await this.onlineAdapter.connect();
      this.adapter = this.onlineAdapter;
      this.office && (this.office.chats = this.chats);
      return true;
    } catch (error) {
      this.bindAdapter(this.localAdapter);
      this.adapter = this.localAdapter;
      this.chats = await this.localAdapter.connect();
      this.office && (this.office.chats = this.chats);
      this.announce(`Servidor no disponible. Messenger continúa en modo local.`);
      return false;
    }
  }

  async sendMessage(index, text) { return this.adapter.sendMessage(index, text); }
  async createConversation(name) { return this.adapter.createConversation(name); }
  async save() { return this.localAdapter.persist(); }
}

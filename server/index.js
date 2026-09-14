'use strict';

const fs = require('node:fs/promises');
const path = require('node:path');
const crypto = require('node:crypto');
const http = require('node:http');
const { WebSocketServer } = require('ws');

const PORT = Number(process.env.PORT || 8787);
const HOST = process.env.HOST || '127.0.0.1';
const JWT_SECRET = process.env.TETORD_JWT_SECRET || (process.env.NODE_ENV === 'production' ? '' : 'tetord-dev-change-this-secret');
const JWT_TTL_SEC = Number(process.env.TETORD_JWT_TTL || 60 * 60 * 24 * 7);
const DATA_DIR = path.resolve(process.env.TETORD_SERVER_DATA || path.join(process.cwd(), 'data'));
const MESSENGER_FILE = path.join(DATA_DIR, 'messenger.json');
const DM_FILE = path.join(DATA_DIR, 'direct-messages.json');
const USERS_FILE = path.join(DATA_DIR, 'users.json');
const DOCS_DIR = path.join(DATA_DIR, 'documents');
const MAX_MESSAGES = 500;
const MAX_TEXT = 4000;
const MAX_DOC_BYTES = 2 * 1024 * 1024;
const RATE_WINDOW_MS = 10_000;
const RATE_LIMIT = 30;
const WS_RATE_LIMIT = 20;

if (!JWT_SECRET) throw new Error('TETORD_JWT_SECRET es obligatorio en producción.');

const state = { rooms: new Map(), users: new Map(), dms: new Map(), collab: new Map() };
const presence = new Map(); // room -> Map(userId, {userId,nick,status,lastSeen})
const rateBuckets = new Map();
const seenMessageIds = new Map();
const seenDmIds = new Map();

const cleanNick = v => String(v || '').trim().replace(/[^\p{L}\p{N}_ .-]/gu, '').slice(0, 40) || 'TETORD User';
const cleanRoom = v => String(v || 'general').trim().toLowerCase().replace(/[^a-z0-9_-]/g, '-').slice(0, 40) || 'general';
const cleanText = v => String(v || '').trim().slice(0, MAX_TEXT);
const cleanUsername = v => String(v || '').trim().toLowerCase().replace(/[^a-z0-9_.-]/g, '').slice(0, 32);
const cleanDocName = v => String(v || 'document').replace(/[^a-zA-Z0-9._ -]/g, '-').slice(0, 120);
const now = () => Date.now();

function hashPassword(password, salt = crypto.randomBytes(16).toString('hex')) {
  const hash = crypto.scryptSync(String(password), salt, 32).toString('hex');
  return { salt, hash };
}
function verifyPassword(password, record) {
  const actual = crypto.scryptSync(String(password), record.salt, 32);
  const expected = Buffer.from(record.hash, 'hex');
  return actual.length === expected.length && crypto.timingSafeEqual(actual, expected);
}
function b64url(input) { return Buffer.from(input).toString('base64url'); }
function signJwt(payload) {
  const header = b64url(JSON.stringify({ alg: 'HS256', typ: 'JWT' }));
  const body = b64url(JSON.stringify(payload));
  const signature = crypto.createHmac('sha256', JWT_SECRET).update(`${header}.${body}`).digest('base64url');
  return `${header}.${body}.${signature}`;
}
function verifyJwt(token) {
  const parts = String(token || '').split('.');
  if (parts.length !== 3) throw new Error('JWT inválido.');
  const [header, body, signature] = parts;
  const expected = crypto.createHmac('sha256', JWT_SECRET).update(`${header}.${body}`).digest('base64url');
  if (signature.length !== expected.length || !crypto.timingSafeEqual(Buffer.from(signature), Buffer.from(expected))) throw new Error('JWT inválido.');
  const parsedHeader = JSON.parse(Buffer.from(header, 'base64url').toString('utf8'));
  if (parsedHeader.alg !== 'HS256') throw new Error('Algoritmo JWT no permitido.');
  const payload = JSON.parse(Buffer.from(body, 'base64url').toString('utf8'));
  if (!payload.sub || !payload.exp || payload.exp < Math.floor(now() / 1000)) throw new Error('JWT expirado.');
  return payload;
}
function issueToken(user) {
  const iat = Math.floor(now() / 1000);
  return signJwt({ sub: user.id, username: user.username, nick: user.nick, iat, exp: iat + JWT_TTL_SEC });
}
function bearer(req) {
  const value = String(req.headers.authorization || '');
  return value.startsWith('Bearer ') ? value.slice(7) : '';
}
function authUser(req) {
  const payload = verifyJwt(bearer(req));
  const user = state.users.get(payload.sub);
  if (!user) throw new Error('Usuario no encontrado.');
  return user;
}
function clientAddress(reqOrSocket) {
  return String(reqOrSocket.headers?.['x-forwarded-for'] || reqOrSocket.socket?.remoteAddress || reqOrSocket._socket?.remoteAddress || 'unknown').split(',')[0].trim();
}
function allowRate(key, limit = RATE_LIMIT) {
  const t = now();
  const bucket = rateBuckets.get(key) || [];
  const fresh = bucket.filter(x => t - x < RATE_WINDOW_MS);
  if (fresh.length >= limit) { rateBuckets.set(key, fresh); return false; }
  fresh.push(t); rateBuckets.set(key, fresh); return true;
}
function sendJson(res, status, value) {
  const body = JSON.stringify(value);
  res.writeHead(status, { 'content-type': 'application/json; charset=utf-8', 'cache-control': 'no-store', 'x-content-type-options': 'nosniff' });
  res.end(body);
}
function safeDocPath(userId, name) {
  const safeUser = userId.replace(/[^a-zA-Z0-9_-]/g, '_');
  const safeName = cleanDocName(name).toLowerCase().endsWith('.tetord') ? cleanDocName(name) : `${cleanDocName(name)}.tetord`;
  const dir = path.join(DOCS_DIR, safeUser);
  const file = path.resolve(dir, safeName);
  if (!file.startsWith(path.resolve(dir) + path.sep)) throw new Error('Ruta inválida.');
  return { dir, file, name: safeName };
}
function validTetord(value) {
  return value && typeof value === 'object' && typeof value.schema === 'string' && /^tetord\.(writer|calc|slides)$/.test(value.schema);
}

function dmKey(a, b) { return [String(a), String(b)].sort().join(':'); }
function rememberId(map, scope, id, max = 1000) {
  const key = `${scope}:${String(id || '')}`;
  if (!id) return false;
  if (map.has(key)) return true;
  map.set(key, now());
  if (map.size > max) {
    const oldest = map.keys().next().value;
    if (oldest) map.delete(oldest);
  }
  return false;
}
function publicDm(message) { return { ...message }; }
function collabRoom(id) { if (!state.collab.has(id)) state.collab.set(id, { module: 'writer', content: '', rev: 0, ops: [], clients: new Set(), cursors: new Map() }); return state.collab.get(id); }
function collabUsers(room) { return [...room.clients].map(ws => ({ userId: ws.tetordUser.id, nick: ws.tetordUser.nick, cursor: room.cursors.get(ws.tetordUser.id) || null })); }
function broadcastCollab(docId, packet, except = null) { const room = state.collab.get(docId); if (!room) return; const data = JSON.stringify(packet); for (const ws of room.clients) if (ws !== except && ws.readyState === 1) ws.send(data); }
function transformOp(op, against) {
  if (!op || !against || op.kind !== 'text' || against.kind !== 'text') return op;
  const out = { ...op }; const a = Math.max(0, Number(against.index) || 0), ac = Math.max(0, Number(against.deleteCount) || 0), ai = String(against.insertText || '').length;
  if (against.clientId && op.clientId && against.clientId === op.clientId) return out;
  if (ai) {
    if (a < out.index || (a === out.index && String(against.clientId || '') < String(op.clientId || ''))) out.index += ai;
    else if (a < out.index + out.deleteCount) out.deleteCount += ai;
  }
  if (ac) {
    const start = out.index, end = out.index + out.deleteCount, delEnd = a + ac;
    if (delEnd <= start) out.index -= ac;
    else if (a >= end) {}
    else { const overlap = Math.max(0, Math.min(end, delEnd) - Math.max(start, a)); out.deleteCount = Math.max(0, out.deleteCount - overlap); if (a < start) out.index = a; }
  }
  return out;
}

async function loadState() {
  await fs.mkdir(DATA_DIR, { recursive: true });
  await fs.mkdir(DOCS_DIR, { recursive: true });
  try {
    const x = JSON.parse(await fs.readFile(MESSENGER_FILE, 'utf8'));
    for (const [room, msgs] of Object.entries(x.rooms || {})) if (Array.isArray(msgs)) state.rooms.set(room, msgs.slice(-MAX_MESSAGES));
  } catch (e) { if (e.code !== 'ENOENT') throw e; }
  try {
    const x = JSON.parse(await fs.readFile(DM_FILE, 'utf8'));
    for (const [key, msgs] of Object.entries(x.dms || {})) if (Array.isArray(msgs)) state.dms.set(key, msgs.slice(-500));
  } catch (e) { if (e.code !== 'ENOENT') throw e; }
  try {
    const x = JSON.parse(await fs.readFile(USERS_FILE, 'utf8'));
    for (const user of Array.isArray(x.users) ? x.users : []) state.users.set(user.id, user);
  } catch (e) { if (e.code !== 'ENOENT') throw e; }
}
let saveTimer = null;
function persistMessenger() {
  if (saveTimer) clearTimeout(saveTimer);
  saveTimer = setTimeout(async () => {
    saveTimer = null;
    const tmp = `${MESSENGER_FILE}.tmp-${process.pid}-${Date.now()}`;
    try { await fs.writeFile(tmp, JSON.stringify({ version: 2, rooms: Object.fromEntries(state.rooms) }, null, 2)); await fs.rename(tmp, MESSENGER_FILE); }
    finally { await fs.rm(tmp, { force: true }).catch(() => {}); }
  }, 50);
}
let dmSaveTimer = null;
function persistDms() { if (dmSaveTimer) clearTimeout(dmSaveTimer); dmSaveTimer = setTimeout(async () => { dmSaveTimer = null; const tmp = `${DM_FILE}.tmp-${process.pid}-${Date.now()}`; try { await fs.writeFile(tmp, JSON.stringify({ version: 1, dms: Object.fromEntries(state.dms) }, null, 2)); await fs.rename(tmp, DM_FILE); } finally { await fs.rm(tmp, { force: true }).catch(() => {}); } }, 50); }
let userSaveTimer = null;
function persistUsers() {
  if (userSaveTimer) clearTimeout(userSaveTimer);
  userSaveTimer = setTimeout(async () => {
    userSaveTimer = null;
    const tmp = `${USERS_FILE}.tmp-${process.pid}-${Date.now()}`;
    try { await fs.writeFile(tmp, JSON.stringify({ version: 1, users: [...state.users.values()] }, null, 2)); await fs.rename(tmp, USERS_FILE); }
    finally { await fs.rm(tmp, { force: true }).catch(() => {}); }
  }, 50);
}
function messages(room) { if (!state.rooms.has(room)) state.rooms.set(room, []); return state.rooms.get(room); }
function roomPresence(room) { if (!presence.has(room)) presence.set(room, new Map()); return presence.get(room); }
function presencePacket(room) { return { type: 'presence', room, users: [...roomPresence(room).values()] }; }
function broadcast(room, packet) {
  const serialized = JSON.stringify(packet);
  for (const client of wss.clients) if (client.readyState === 1 && client.tetordAuthed && client.tetordRoom === room) client.send(serialized);
}
function setPresence(room, user, status) {
  const map = roomPresence(room);
  const existing = map.get(user.id) || { userId: user.id, nick: user.nick };
  existing.status = status; existing.lastSeen = now(); map.set(user.id, existing);
  broadcast(room, presencePacket(room));
}
function removePresence(room, userId) {
  const map = roomPresence(room);
  const existing = map.get(userId);
  if (existing) { existing.status = 'offline'; existing.lastSeen = now(); map.set(userId, existing); broadcast(room, presencePacket(room)); }
}

const server = http.createServer(async (req, res) => {
  res.setHeader('access-control-allow-origin', process.env.TETORD_CORS_ORIGIN || '*');
  res.setHeader('access-control-allow-headers', 'Authorization, Content-Type');
  res.setHeader('access-control-allow-methods', 'GET, POST, PUT, OPTIONS');
  if (req.method === 'OPTIONS') return res.writeHead(204).end();
  if (!allowRate(`ip:${clientAddress(req)}`)) return sendJson(res, 429, { error: 'Demasiadas solicitudes. Intenta de nuevo más tarde.' });

  const url = new URL(req.url, `http://${HOST}:${PORT}`);
  try {
    if (url.pathname === '/health' && req.method === 'GET') return sendJson(res, 200, { ok: true, version: '4.0.0' });
    if (url.pathname === '/auth/register' && req.method === 'POST') {
      const body = await readBody(req, 16 * 1024);
      const username = cleanUsername(body.username), nick = cleanNick(body.nick || body.username), password = String(body.password || '');
      if (!/^[a-z0-9_.-]{3,32}$/.test(username) || password.length < 8 || password.length > 128) return sendJson(res, 400, { error: 'Usuario inválido o contraseña de 8–128 caracteres requerida.' });
      if ([...state.users.values()].some(u => u.username === username)) return sendJson(res, 409, { error: 'Ese usuario ya existe.' });
      const id = crypto.randomUUID(), credentials = hashPassword(password);
      const user = { id, username, nick, salt: credentials.salt, hash: credentials.hash, createdAt: now() };
      state.users.set(id, user); persistUsers();
      return sendJson(res, 201, { user: publicUser(user), token: issueToken(user) });
    }
    if (url.pathname === '/auth/login' && req.method === 'POST') {
      const body = await readBody(req, 16 * 1024);
      const username = cleanUsername(body.username), user = [...state.users.values()].find(u => u.username === username);
      if (!user || !verifyPassword(String(body.password || ''), user)) return sendJson(res, 401, { error: 'Credenciales inválidas.' });
      user.lastLoginAt = now(); persistUsers();
      return sendJson(res, 200, { user: publicUser(user), token: issueToken(user) });
    }
    if (url.pathname === '/auth/refresh' && req.method === 'POST') {
      const user = authUser(req); return sendJson(res, 200, { user: publicUser(user), token: issueToken(user) });
    }

    const user = authUser(req);
    const authKey = `user:${user.id}`;
    if (!allowRate(authKey)) return sendJson(res, 429, { error: 'Límite de solicitudes alcanzado.' });

    if (url.pathname === '/auth/me' && req.method === 'GET') return sendJson(res, 200, { user: publicUser(user) });
    if (url.pathname === '/presence' && req.method === 'GET') return sendJson(res, 200, { rooms: Object.fromEntries([...presence].map(([r, map]) => [r, [...map.values()]])) });
    if (url.pathname === '/users' && req.method === 'GET') return sendJson(res, 200, { users: [...state.users.values()].map(publicUser) });
    const dmMatch = url.pathname.match(/^\/dm\/([^/]+)$/);
    if (dmMatch && (req.method === 'GET' || req.method === 'POST')) {
      const otherId = decodeURIComponent(dmMatch[1]); if (!state.users.has(otherId) || otherId === user.id) return sendJson(res, 400, { error: 'Usuario de DM inválido.' });
      const key = dmKey(user.id, otherId); if (!state.dms.has(key)) state.dms.set(key, []);
      if (req.method === 'GET') return sendJson(res, 200, { messages: state.dms.get(key).slice(-100).map(publicDm) });
      const body = await readBody(req, 16 * 1024); const text = cleanText(body.text); if (!text) return sendJson(res, 400, { error: 'Mensaje vacío.' });
      const message = { id: crypto.randomUUID(), fromUserId: user.id, toUserId: otherId, from: user.nick, text, at: now() }; state.dms.get(key).push(message); if (state.dms.get(key).length > 500) state.dms.get(key).splice(0, state.dms.get(key).length - 500); persistDms();
      for (const ws of wss.clients) if (ws.readyState === 1 && ws.tetordAuthed && (ws.tetordUser.id === user.id || ws.tetordUser.id === otherId)) send(ws, { type: 'dm', message });
      return sendJson(res, 201, { message });
    }

    if (url.pathname === '/sync/documents' && req.method === 'GET') {
      const dir = path.join(DOCS_DIR, user.id); await fs.mkdir(dir, { recursive: true });
      const names = (await fs.readdir(dir, { withFileTypes: true })).filter(x => x.isFile() && x.name.toLowerCase().endsWith('.tetord'));
      const docs = [];
      for (const entry of names) { const stat = await fs.stat(path.join(dir, entry.name)); docs.push({ name: entry.name, updatedAt: stat.mtimeMs, size: stat.size }); }
      return sendJson(res, 200, { documents: docs });
    }
    const docMatch = url.pathname.match(/^\/sync\/documents\/(.+)$/);
    if (docMatch && (req.method === 'GET' || req.method === 'PUT')) {
      const info = safeDocPath(user.id, decodeURIComponent(docMatch[1]));
      if (req.method === 'GET') {
        const stat = await fs.stat(info.file).catch(() => null); if (!stat) return sendJson(res, 404, { error: 'Documento no encontrado.' });
        const raw = await fs.readFile(info.file, 'utf8'); return sendJson(res, 200, { name: info.name, updatedAt: stat.mtimeMs, document: JSON.parse(raw) });
      }
      const body = await readBody(req, MAX_DOC_BYTES + 1024);
      if (!validTetord(body.document)) return sendJson(res, 400, { error: 'Documento .tetord inválido.' });
      const requestedAt = Number(body.updatedAt) || now();
      const existingStat = await fs.stat(info.file).catch(() => null);
      if (existingStat && existingStat.mtimeMs > requestedAt + 1000) return sendJson(res, 409, { error: 'Conflicto: el documento remoto es más reciente.', updatedAt: existingStat.mtimeMs });
      await fs.mkdir(info.dir, { recursive: true });
      const tmp = `${info.file}.tmp-${process.pid}-${Date.now()}`;
      await fs.writeFile(tmp, JSON.stringify(body.document, null, 2), 'utf8');
      await fs.rename(tmp, info.file);
      return sendJson(res, 200, { ok: true, name: info.name, updatedAt: (await fs.stat(info.file)).mtimeMs });
    }
    return sendJson(res, 404, { error: 'Ruta no encontrada.' });
  } catch (error) {
    const status = /JWT|Usuario no encontrado|Credenciales/.test(error.message) ? 401 : error.code === 'ENOENT' ? 404 : 400;
    return sendJson(res, status, { error: error.message || 'Error interno.' });
  }
});

async function readBody(req, maxBytes) {
  let size = 0; const chunks = [];
  for await (const chunk of req) { size += chunk.length; if (size > maxBytes) throw new Error('Payload demasiado grande.'); chunks.push(chunk); }
  if (!chunks.length) return {};
  const raw = Buffer.concat(chunks).toString('utf8');
  const body = JSON.parse(raw); if (!body || typeof body !== 'object' || Array.isArray(body)) throw new Error('JSON inválido.'); return body;
}
function publicUser(user) { return { id: user.id, username: user.username, nick: user.nick, createdAt: user.createdAt, lastLoginAt: user.lastLoginAt || null }; }

const wss = new WebSocketServer({ server, maxPayload: 16 * 1024 });
wss.on('connection', ws => {
  ws.tetordAuthed = false; ws.tetordUser = null; ws.tetordRoom = 'general'; ws.tetordIp = clientAddress(ws);
  ws.on('message', async raw => {
    if (!allowRate(`ws:${ws.tetordIp}`, WS_RATE_LIMIT)) return send(ws, { type: 'error', fatal: false, message: 'Demasiados mensajes. Reduce la frecuencia.' });
    if (raw.length > 16 * 1024) return send(ws, { type: 'error', fatal: false, message: 'Payload demasiado grande.' });
    let p; try { p = JSON.parse(raw.toString()); } catch { return send(ws, { type: 'error', fatal: false, message: 'JSON inválido.' }); }
    try {
      if (p.type === 'auth') {
        const payload = verifyJwt(p.token);
        const user = state.users.get(payload.sub); if (!user) throw new Error('Usuario no encontrado.');
        ws.tetordAuthed = true; ws.tetordUser = user; ws.tetordRoom = cleanRoom(p.room);
        setPresence(ws.tetordRoom, user, 'online');
        send(ws, { type: 'auth_ok', user: publicUser(user), nick: user.nick, room: ws.tetordRoom, messages: messages(ws.tetordRoom).slice(-100), presence: presencePacket(ws.tetordRoom).users });
        return;
      }
      if (!ws.tetordAuthed) throw new Error('Autenticación requerida.');
      if (p.type === 'join') {
        removePresence(ws.tetordRoom, ws.tetordUser.id); ws.tetordRoom = cleanRoom(p.room); setPresence(ws.tetordRoom, ws.tetordUser, 'online');
        return send(ws, { type: 'history', room: ws.tetordRoom, messages: messages(ws.tetordRoom).slice(-100), presence: presencePacket(ws.tetordRoom).users });
      }
      if (p.type === 'dm') {
        const otherId = String(p.toUserId || ''); if (!state.users.has(otherId) || otherId === ws.tetordUser.id) throw new Error('Destinatario inválido.');
        const text = cleanText(p.text); if (!text) return; if (rememberId(seenDmIds, dmKey(ws.tetordUser.id, otherId), p.clientId)) return send(ws, { type: 'dm_ack', clientId: String(p.clientId), duplicate: true }); const key = dmKey(ws.tetordUser.id, otherId); if (!state.dms.has(key)) state.dms.set(key, []);
        const message = { id: String(p.clientId || crypto.randomUUID()).slice(0, 100), fromUserId: ws.tetordUser.id, toUserId: otherId, from: ws.tetordUser.nick, text, at: now() }; state.dms.get(key).push(message); if (state.dms.get(key).length > 500) state.dms.get(key).splice(0, state.dms.get(key).length - 500); persistDms();
        for (const client of wss.clients) if (client.readyState === 1 && client.tetordAuthed && (client.tetordUser.id === ws.tetordUser.id || client.tetordUser.id === otherId)) send(client, { type: 'dm', message });
        return;
      }
      if (p.type === 'dm_history') { const otherId = String(p.toUserId || ''); const key = dmKey(ws.tetordUser.id, otherId); return send(ws, { type: 'dm_history', toUserId: otherId, messages: (state.dms.get(key) || []).slice(-100) }); }
      if (p.type === 'collab_join') {
        const docId = String(p.docId || '').slice(0, 160); if (!docId) throw new Error('Documento colaborativo inválido.');
        const room = collabRoom(docId); room.module = p.module === 'calc' ? 'calc' : 'writer'; if (!room.rev && typeof p.content === 'string') room.content = p.content.slice(0, MAX_DOC_BYTES);
        room.clients.add(ws); ws.tetordCollabDocs = ws.tetordCollabDocs || new Set(); ws.tetordCollabDocs.add(docId);
        send(ws, { type: 'collab_state', docId, module: room.module, content: room.content, rev: room.rev, users: collabUsers(room) }); broadcastCollab(docId, { type: 'collab_presence', docId, users: collabUsers(room) }, ws); return;
      }
      if (p.type === 'collab_op') {
        const docId = String(p.docId || ''); const room = state.collab.get(docId); if (!room || !room.clients.has(ws)) throw new Error('No estás en esta sala de edición.');
        let op = p.op && typeof p.op === 'object' ? { ...p.op } : null; if (!op || !['text','replace'].includes(op.kind)) throw new Error('Operación colaborativa inválida.');
        if (op.kind === 'replace') op.value = String(op.value || '').slice(0, MAX_DOC_BYTES);
        if (op.kind === 'text') { op.index = Math.max(0, Math.min(room.content.length, Number(op.index) || 0)); op.deleteCount = Math.max(0, Math.min(room.content.length - op.index, Number(op.deleteCount) || 0)); op.insertText = String(op.insertText || '').slice(0, MAX_TEXT); const baseRev = Math.max(0, Number(p.baseRev) || 0); if (baseRev < room.rev - room.ops.length) return send(ws, { type: 'collab_resync', docId, content: room.content, rev: room.rev, module: room.module }); for (const historic of room.ops.slice(Math.max(0, baseRev - (room.rev - room.ops.length)))) op = transformOp(op, historic); room.content = room.content.slice(0, op.index) + op.insertText + room.content.slice(op.index + op.deleteCount); }
        else room.content = op.value; room.rev += 1; op.clientId = String(op.clientId || '').slice(0, 100); room.ops.push(op); if (room.ops.length > 1000) room.ops.splice(0, room.ops.length - 1000);
        send(ws, { type: 'collab_ack', docId, rev: room.rev, opId: op.opId || null }); broadcastCollab(docId, { type: 'collab_op', docId, rev: room.rev, userId: ws.tetordUser.id, senderId: ws.tetordUser.id, op }, ws); return;
      }
      if (p.type === 'cursor') { const docId = String(p.docId || ''); const room = state.collab.get(docId); if (!room || !room.clients.has(ws)) return; const cursor = { start: Math.max(0, Number(p.cursor?.start) || 0), end: Math.max(0, Number(p.cursor?.end) || 0) }; room.cursors.set(ws.tetordUser.id, cursor); broadcastCollab(docId, { type: 'cursor', docId, userId: ws.tetordUser.id, cursor }, ws); return; }
      if (p.type === 'collab_leave') { const docId = String(p.docId || ''); const room = state.collab.get(docId); if (room) { room.clients.delete(ws); room.cursors.delete(ws.tetordUser.id); broadcastCollab(docId, { type: 'collab_presence', docId, users: collabUsers(room) }, ws); if (!room.clients.size) state.collab.delete(docId); } return; }
      if (p.type === 'presence') {
        const status = ['online', 'away', 'offline'].includes(p.status) ? p.status : 'online'; setPresence(ws.tetordRoom, ws.tetordUser, status); return;
      }
      if (p.type === 'message') {
        const text = cleanText(p.text); if (!text) return;
        const room = cleanRoom(p.room || ws.tetordRoom); ws.tetordRoom = room; if (rememberId(seenMessageIds, room, p.clientId)) return send(ws, { type: 'message_ack', clientId: String(p.clientId), duplicate: true }); setPresence(room, ws.tetordUser, 'online');
        const msg = { id: String(p.clientId || `${now()}-${crypto.randomBytes(4).toString('hex')}`).slice(0, 100), userId: ws.tetordUser.id, from: ws.tetordUser.nick, text, at: now() };
        const list = messages(room); list.push(msg); if (list.length > MAX_MESSAGES) list.splice(0, list.length - MAX_MESSAGES); persistMessenger(); broadcast(room, { type: 'message', room, message: msg });
      }
    } catch (error) { send(ws, { type: 'error', fatal: true, message: error.message || 'Error.' }); }
  });
  ws.on('close', () => { if (ws.tetordUser) removePresence(ws.tetordRoom, ws.tetordUser.id); for (const docId of ws.tetordCollabDocs || []) { const room = state.collab.get(docId); if (room) { room.clients.delete(ws); room.cursors.delete(ws.tetordUser?.id); broadcastCollab(docId, { type: 'collab_presence', docId, users: collabUsers(room) }, ws); if (!room.clients.size) state.collab.delete(docId); } } });
});
function send(ws, packet) { if (ws.readyState === 1) ws.send(JSON.stringify(packet)); }

server.listen(PORT, HOST, () => console.log(`TETORD 4.0.1 server listening on http://${HOST}:${PORT} (WebSocket + REST)`));
server.on('error', error => console.error(`TETORD server error: ${error.message}`));
async function shutdown() { server.close(); if (saveTimer) { clearTimeout(saveTimer); saveTimer = null; persistMessenger(); } if (userSaveTimer) { clearTimeout(userSaveTimer); userSaveTimer = null; persistUsers(); } if (dmSaveTimer) { clearTimeout(dmSaveTimer); dmSaveTimer = null; persistDms(); } }
process.on('SIGINT', shutdown); process.on('SIGTERM', shutdown);
loadState().catch(error => { console.error(error); process.exitCode = 1; });

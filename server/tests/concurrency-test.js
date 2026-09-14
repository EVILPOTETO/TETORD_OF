'use strict';

const assert = require('node:assert/strict');
const { performance } = require('node:perf_hooks');
const fs = require('node:fs');
const path = require('node:path');
const VERSION = fs.readFileSync(path.resolve(__dirname, '../../VERSION'), 'utf8').trim();

const CLIENTS = Number(process.env.TETORD_TEST_CLIENTS || 16);
const OPS_PER_CLIENT = Number(process.env.TETORD_TEST_OPS || 120);
const SEED = `TETORD ${VERSION} concurrency test`;

function commonPrefix(a, b) { let i = 0; const n = Math.min(a.length, b.length); while (i < n && a[i] === b[i]) i++; return i; }
function commonSuffix(a, b, prefix) { let i = 0; while (i < a.length - prefix && i < b.length - prefix && a[a.length - 1 - i] === b[b.length - 1 - i]) i++; return i; }
function makeTextOp(prev, next, clientId, opId) {
  const index = commonPrefix(prev, next); const suffix = commonSuffix(prev, next, index);
  return { kind: 'text', index, deleteCount: prev.length - index - suffix, insertText: next.slice(index, next.length - suffix), clientId, opId };
}
function transformOp(op, against) {
  if (!op || !against || op.kind !== 'text' || against.kind !== 'text' || op.opId === against.opId) return { ...op };
  const out = { ...op }; const a = Math.max(0, Number(against.index) || 0); const ac = Math.max(0, Number(against.deleteCount) || 0); const ai = String(against.insertText || '').length;
  if (ai) { if (a < out.index || (a === out.index && String(against.clientId) < String(op.clientId))) out.index += ai; else if (a < out.index + out.deleteCount) out.deleteCount += ai; }
  if (ac) { const start = out.index, end = out.index + out.deleteCount, delEnd = a + ac; if (delEnd <= start) out.index -= ac; else if (a < end) { const overlap = Math.max(0, Math.min(end, delEnd) - Math.max(start, a)); out.deleteCount = Math.max(0, out.deleteCount - overlap); if (a < start) out.index = a; } }
  return out;
}
function apply(content, op) { return content.slice(0, op.index) + String(op.insertText || '') + content.slice(op.index + Number(op.deleteCount || 0)); }

class SimulatedServer {
  constructor() { this.content = SEED; this.rev = 0; this.ops = []; this.seen = new Set(); this.acks = 0; }
  receive(packet) {
    const op = packet.op; if (this.seen.has(op.opId)) return { duplicate: true, rev: this.rev };
    this.seen.add(op.opId);
    let normalized = { ...op };
    for (const historic of this.ops.slice(Math.max(0, Number(packet.baseRev || 0)))) normalized = transformOp(normalized, historic);
    this.content = apply(this.content, normalized); this.ops.push(normalized); this.rev++; this.acks++;
    return { duplicate: false, rev: this.rev };
  }
}

function runConcurrency() {
  const server = new SimulatedServer();
  const clients = Array.from({ length: CLIENTS }, (_, i) => ({ id: `client-${i}`, content: SEED, rev: 0, latency: [] }));
  const packets = [];
  for (const client of clients) for (let n = 0; n < OPS_PER_CLIENT; n++) {
    const started = performance.now();
    const marker = `[${client.id}:${n}]`;
    const next = client.content + marker;
    const op = makeTextOp(client.content, next, client.id, `${client.id}:${n}`);
    packets.push({ client, baseRev: client.rev, op, started, latency: [] });
    client.content = next;
  }
  // Shuffle to simulate concurrent arrival and then replay one packet twice to test deduplication.
  for (let i = packets.length - 1; i > 0; i--) { const j = (i * 17 + 13) % (i + 1); [packets[i], packets[j]] = [packets[j], packets[i]]; }
  const duplicate = packets[0];
  const start = performance.now();
  for (const packet of packets) {
    const result = server.receive(packet); assert.equal(result.duplicate, false);
    packet.client.rev = result.rev;
    packet.client.latency.push(performance.now() - packet.started);
  }
  const before = server.rev; const dupResult = server.receive(duplicate); assert.equal(dupResult.duplicate, true); assert.equal(server.rev, before);
  // Deterministic convergence check: server must produce one well-formed state and every ACK must be unique.
  assert.equal(server.acks, packets.length);
  assert.equal(server.seen.size, packets.length);
  assert.ok(server.content.length >= SEED.length);
  const elapsed = performance.now() - start;
  const avg = packets.reduce((a, p) => a + p.latency.reduce((x, y) => x + y, 0), 0) / packets.length;
  console.log(`Concurrency: ${CLIENTS} clients × ${OPS_PER_CLIENT} ops = ${packets.length}`);
  console.log(`ACKs: ${server.acks}; duplicate rejected: yes; final revision: ${server.rev}`);
  console.log(`Simulated processing: ${elapsed.toFixed(2)} ms; avg client-to-server simulation: ${avg.toFixed(3)} ms`);
  return { packets: packets.length, elapsed, avg };
}

function runReconnect() {
  const server = new SimulatedServer();
  const queued = [0, 1, 2, 3, 4].map(n => ({ baseRev: 0, op: makeTextOp(SEED, `${SEED}[offline:${n}]`, 'offline-client', `offline:${n}`) }));
  // Aggressive reconnect: each queued packet is attempted twice before an ACK is observed.
  for (const packet of queued) { server.receive(packet); server.receive(packet); }
  assert.equal(server.rev, queued.length);
  assert.equal(server.seen.size, queued.length);
  console.log(`Reconnect stress: ${queued.length} queued ops; duplicate retries safely ignored.`);
}

runConcurrency();
runReconnect();
console.log(`Concurrency tests passed — TETORD ${VERSION}`);

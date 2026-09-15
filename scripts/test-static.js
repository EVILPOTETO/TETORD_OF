#!/usr/bin/env node
const fs = require('node:fs');
const path = require('node:path');
const { execFileSync } = require('node:child_process');

const root = path.resolve(__dirname, '..');
const jsFiles = [
  'electron/main.js', 'electron/preload.js', 'scripts/sync-version.js',
  'scripts/test-static.js', 'scripts/build-release.js', 'server/index.js', 'MiWord/js/app.js', 'MiWord/js/office.js',
  ...fs.readdirSync(path.join(root, 'MiWord/src/renderer/modules')).map(f => `MiWord/src/renderer/modules/${f}`),
  'MiWord/src/renderer/core/storage.js', 'MiWord/src/renderer/modules/cloud-sync.js', 'MiWord/src/renderer/modules/collaboration.js', 'server/tests/concurrency-test.js'
];
for (const rel of jsFiles) execFileSync(process.execPath, ['--check', path.join(root, rel)], { stdio: 'inherit' });

const html = fs.readFileSync(path.join(root, 'MiWord/index.html'), 'utf8');
for (const tag of ['div', 'section']) {
  const open = (html.match(new RegExp(`<${tag}\\b`, 'g')) || []).length;
  const close = (html.match(new RegExp(`</${tag}>`, 'g')) || []).length;
  if (open !== close) throw new Error(`${tag} imbalance: ${open}/${close}`);
}
const slides = fs.readFileSync(path.join(root, 'MiWord/src/renderer/modules/slides.js'), 'utf8');
if (/renderSlides[\s\S]*?innerHTML/.test(slides) || /renderPresent[\s\S]*?innerHTML/.test(slides)) throw new Error('Unsafe innerHTML remains in Slides render methods.');
if (fs.existsSync(path.join(root, 'patch.py')) || fs.existsSync(path.join(root, 'upgrade30.py'))) throw new Error('Legacy Python patch script remains.');
const version = fs.readFileSync(path.join(root, 'VERSION'), 'utf8').trim();
const pkg = JSON.parse(fs.readFileSync(path.join(root, 'package.json'), 'utf8'));
if (pkg.version !== version) throw new Error(`Version mismatch: VERSION=${version}, package=${pkg.version}`);
const serverPkg = JSON.parse(fs.readFileSync(path.join(root, 'server/package.json'), 'utf8'));
if (serverPkg.version !== version) throw new Error(`Version mismatch: VERSION=${version}, server/package=${serverPkg.version}`);
const cloud = fs.readFileSync(path.join(root, 'MiWord/src/renderer/modules/cloud-sync.js'), 'utf8');
if (!cloud.includes('/auth/login') || !cloud.includes('/sync/documents')) throw new Error('Cloud Sync endpoints missing.');
const server = fs.readFileSync(path.join(root, 'server/index.js'), 'utf8');
for (const marker of ['verifyJwt', 'scryptSync', 'allowRate', "type: 'presence'", '/sync/documents', 'collab_join', 'collab_op', "type: 'dm'", "'/auth/refresh'"]) if (!server.includes(marker)) throw new Error(`3.9 security/sync marker missing: ${marker}`);
if (!server.includes('seenDmIds') || !server.includes('seenMessageIds')) throw new Error('Server deduplication markers missing.');
const collab = fs.readFileSync(path.join(root, 'MiWord/src/renderer/modules/collaboration.js'), 'utf8');
for (const marker of ['outbox', 'inFlight', 'queueReconnect', 'rejoinAfterReconnect']) if (!collab.includes(marker)) throw new Error(`Collaboration stability marker missing: ${marker}`);
const buildRelease = fs.readFileSync(path.join(root, 'scripts/build-release.js'), 'utf8');
for (const marker of ['preflight', 'electron-builder', 'nsis', 'portable', 'RELEASE_MANIFEST.json']) if (!buildRelease.includes(marker)) throw new Error(`Release build marker missing: ${marker}`);
console.log(`Static tests passed — TETORD ${version}`);

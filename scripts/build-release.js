#!/usr/bin/env node
/**
 * TETORD release gate.
 * Runs preflight, removes stale build output, validates the electron-builder
 * allowlist, builds NSIS + portable targets, and verifies the resulting names.
 */
const fs = require('node:fs/promises');
const path = require('node:path');
const { spawnSync } = require('node:child_process');

const ROOT = path.resolve(__dirname, '..');
const DIST = path.join(ROOT, 'dist');
const VERSION_FILE = path.join(ROOT, 'VERSION');
const PACKAGE_FILE = path.join(ROOT, 'package.json');

const REQUIRED_DEV_EXCLUDES = [
  'node_modules', '.git', '.github', 'server', 'scripts', 'dist',
  'server/tests', 'test', 'tests', '*.log'
];

function npmCommand() {
  return process.platform === 'win32' ? 'npm.cmd' : 'npm';
}

function run(command, args) {
  console.log(`\\n> ${command} ${args.join(' ')}`);
  const result = spawnSync(command, args, { cwd: ROOT, stdio: 'inherit', shell: process.platform === 'win32' });
  if (result.status !== 0) throw new Error(`Comando falló (${result.status ?? 'signal'}): ${command}`);
}

async function exists(file) {
  try { await fs.access(file); return true; } catch { return false; }
}

async function main() {
  const version = (await fs.readFile(VERSION_FILE, 'utf8')).trim();
  const pkg = JSON.parse(await fs.readFile(PACKAGE_FILE, 'utf8'));
  if (pkg.version !== version) throw new Error(`Version mismatch: VERSION=${version}, package.json=${pkg.version}`);

  console.log(`TETORD ${version} — release preflight`);
  run(npmCommand(), ['run', 'preflight']);

  const build = pkg.build || {};
  const files = Array.isArray(build.files) ? build.files : [];
  if (!files.includes('MiWord/**/*') || !files.includes('electron/**/*')) {
    throw new Error('electron-builder allowlist incompleta: faltan MiWord/**/* o electron/**/*');
  }
  for (const forbidden of REQUIRED_DEV_EXCLUDES) {
    if (files.some(entry => entry === forbidden || entry === `${forbidden}/**/*`)) {
      throw new Error(`La configuración de producción incluye contenido de desarrollo prohibido: ${forbidden}`);
    }
  }
  if (!Array.isArray(build.win?.target) || !build.win.target.includes('nsis') || !build.win.target.includes('portable')) {
    throw new Error('La build de Windows debe producir NSIS y portable.');
  }

  await fs.rm(DIST, { recursive: true, force: true });
  await fs.mkdir(DIST, { recursive: true });

  run(npmCommand(), ['exec', '--', 'electron-builder', '--win', 'nsis', 'portable', '--x64', '--publish', 'never']);

  const entries = await fs.readdir(DIST);
  const setup = `TETORD-Setup-${version}-x64.exe`;
  const portable = `TETORD-${version}-x64.exe`;
  if (!entries.includes(setup)) throw new Error(`No se encontró el instalador NSIS esperado: ${setup}`);
  if (!entries.includes(portable)) throw new Error(`No se encontró el portable esperado: ${portable}`);

  const allowedBuilderOutputs = new Set([
    setup,
    portable,
    'latest.yml',
    'RELEASE_MANIFEST.json'
  ]);
  const suspicious = entries.filter(name => {
    if (allowedBuilderOutputs.has(name)) return false;
    return ['test', 'tests', 'server', 'scripts'].some(part =>
      name.toLowerCase().includes(part)
    );
  });
  if (suspicious.length) throw new Error(`Artefactos sospechosos en dist/: ${suspicious.join(', ')}`);

  const manifest = {
    product: 'TETORD',
    version,
    generatedAt: new Date().toISOString(),
    artifacts: [setup, portable],
    targets: ['nsis', 'portable'],
    preflight: 'passed'
  };
  await fs.writeFile(path.join(DIST, 'RELEASE_MANIFEST.json'), JSON.stringify(manifest, null, 2) + '\n');

  console.log(`\\nRelease build verified: TETORD ${version}`);
  console.log(`NSIS:     ${setup}`);
  console.log(`Portable: ${portable}`);
}

main().catch(error => {
  console.error(`\\nRELEASE BUILD FAILED: ${error.stack || error.message}`);
  process.exitCode = 1;
});

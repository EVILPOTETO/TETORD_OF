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
  'node_modules',
  '.git',
  '.github',
  'server',
  'scripts',
  'dist',
  'server/tests',
  'test',
  'tests',
  '*.log'
];

function npmCommand() {
  return process.platform === 'win32' ? 'npm.cmd' : 'npm';
}

function run(command, args) {
  console.log(`\n> ${command} ${args.join(' ')}`);

  const result = spawnSync(command, args, {
    cwd: ROOT,
    stdio: 'inherit',
    shell: process.platform === 'win32',
    windowsHide: false
  });

  if (result.error) {
    throw new Error(
      `No se pudo ejecutar el comando.\n` +
      `Comando: ${command} ${args.join(' ')}\n` +
      `Error: ${result.error.message}\n` +
      `Código: ${result.status ?? 'N/A'}\n` +
      `Señal: ${result.signal ?? 'N/A'}`
    );
  }

  if (result.status !== 0) {
    throw new Error(
      `El comando terminó con error.\n` +
      `Comando: ${command} ${args.join(' ')}\n` +
      `Código de salida: ${result.status ?? 'N/A'}\n` +
      `Señal: ${result.signal ?? 'N/A'}`
    );
  }

  console.log(`✓ Comando completado correctamente: ${command} ${args.join(' ')}`);
}

async function exists(file) {
  try {
    await fs.access(file);
    return true;
  } catch {
    return false;
  }
}

async function main() {
  const version = (await fs.readFile(VERSION_FILE, 'utf8')).trim();
  const pkg = JSON.parse(await fs.readFile(PACKAGE_FILE, 'utf8'));

  if (pkg.version !== version) {
    throw new Error(
      `Version mismatch: VERSION=${version}, package.json=${pkg.version}`
    );
  }

  console.log(`TETORD ${version} — release preflight`);

  // ------------------------------------------------------------
  // 1. PRE-FLIGHT
  // ------------------------------------------------------------
  console.log('\n[1/5] Ejecutando preflight...');

  try {
    run(npmCommand(), ['run', 'preflight']);
  } catch (error) {
    console.error('\n❌ PRE-FLIGHT FALLÓ');
    console.error(error.message);
    throw error;
  }

  console.log('✓ Preflight completado.');

  // ------------------------------------------------------------
  // 2. VALIDATE ELECTRON-BUILDER CONFIG
  // ------------------------------------------------------------
  console.log('\n[2/5] Validando configuración de electron-builder...');

  const build = pkg.build || {};
  const files = Array.isArray(build.files) ? build.files : [];

  if (!files.includes('MiWord/**/*') || !files.includes('electron/**/*')) {
    throw new Error(
      'electron-builder allowlist incompleta: faltan MiWord/**/* o electron/**/*'
    );
  }

  for (const forbidden of REQUIRED_DEV_EXCLUDES) {
    if (
      files.some(
        entry => entry === forbidden || entry === `${forbidden}/**/*`
      )
    ) {
      throw new Error(
        `La configuración de producción incluye contenido de desarrollo prohibido: ${forbidden}`
      );
    }
  }

  if (
    !Array.isArray(build.win?.target) ||
    !build.win.target.includes('nsis') ||
    !build.win.target.includes('portable')
  ) {
    throw new Error(
      'La build de Windows debe producir NSIS y portable.'
    );
  }

  console.log('✓ Allowlist de producción correcta.');
  console.log('✓ Targets NSIS + Portable encontrados.');

  // ------------------------------------------------------------
  // 3. CLEAN DIST
  // ------------------------------------------------------------
  console.log('\n[3/5] Limpiando dist/...');

  await fs.rm(DIST, {
    recursive: true,
    force: true
  });

  await fs.mkdir(DIST, {
    recursive: true
  });

  console.log('✓ dist/ preparada.');

  // ------------------------------------------------------------
  // 4. BUILD
  // ------------------------------------------------------------
  console.log('\n[4/5] Generando instalador Windows...');

  run(
    npmCommand(),
    [
      'exec',
      '--',
      'electron-builder',
      '--win',
      'nsis',
      'portable',
      '--x64'
    ]
  );

  console.log('✓ electron-builder terminó correctamente.');

  // ------------------------------------------------------------
  // 5. VERIFY OUTPUT
  // ------------------------------------------------------------
  console.log('\n[5/5] Verificando artefactos...');

  const entries = await fs.readdir(DIST);

  const setup = `TETORD-Setup-${version}-x64.exe`;
  const portable = `TETORD-${version}-x64.exe`;

  if (!entries.includes(setup)) {
    throw new Error(
      `No se encontró el instalador NSIS esperado: ${setup}\n` +
      `Contenido actual de dist/: ${entries.join(', ')}`
    );
  }

  if (!entries.includes(portable)) {
    throw new Error(
      `No se encontró el portable esperado: ${portable}\n` +
      `Contenido actual de dist/: ${entries.join(', ')}`
    );
  }

  const forbiddenOutputNames = [
    'test',
    'tests',
    'server',
    'scripts'
  ];

  const suspicious = entries.filter(name =>
    forbiddenOutputNames.some(part =>
      name.toLowerCase().includes(part)
    )
  );

  if (suspicious.length) {
    throw new Error(
      `Artefactos sospechosos en dist/: ${suspicious.join(', ')}`
    );
  }

  const manifest = {
    product: 'TETORD',
    version,
    generatedAt: new Date().toISOString(),
    artifacts: [
      setup,
      portable
    ],
    targets: [
      'nsis',
      'portable'
    ],
    preflight: 'passed'
  };

  await fs.writeFile(
    path.join(DIST, 'RELEASE_MANIFEST.json'),
    JSON.stringify(manifest, null, 2) + '\n'
  );

  console.log('\n========================================');
  console.log(`Release build verified: TETORD ${version}`);
  console.log(`NSIS:     ${setup}`);
  console.log(`Portable: ${portable}`);
  console.log('========================================');
}

main().catch(error => {
  console.error('\n========================================');
  console.error('RELEASE BUILD FAILED');
  console.error('========================================');
  console.error(error.stack || error.message);
  process.exitCode = 1;
});

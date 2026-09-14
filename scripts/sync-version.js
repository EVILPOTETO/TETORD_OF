#!/usr/bin/env node
const fs = require('node:fs/promises');
const path = require('node:path');

const ROOT = path.resolve(__dirname, '..');
const VERSION_FILE = path.join(ROOT, 'VERSION');
const DIST_TEMPLATE = path.join(ROOT, 'dist-template');

async function atomicWrite(file, content) {
  const tmp = `${file}.tmp-${process.pid}-${Date.now()}`;
  await fs.writeFile(tmp, content, 'utf8');
  await fs.rename(tmp, file);
}

async function replaceIfExists(file, replacer) {
  try {
    const text = await fs.readFile(file, 'utf8');
    await atomicWrite(file, replacer(text));
  } catch (error) {
    if (error.code !== 'ENOENT') throw error;
  }
}

async function main() {
  const version = (await fs.readFile(VERSION_FILE, 'utf8')).trim();
  if (!/^\d+\.\d+\.\d+$/.test(version)) throw new Error(`VERSION inválida: ${version}`);

  const packagePath = path.join(ROOT, 'package.json');
  const pkg = JSON.parse(await fs.readFile(packagePath, 'utf8'));
  pkg.version = version;
  await atomicWrite(packagePath, JSON.stringify(pkg, null, 2) + '\n');

  const serverPackagePath = path.join(ROOT, 'server', 'package.json');
  try {
    const serverPkg = JSON.parse(await fs.readFile(serverPackagePath, 'utf8'));
    serverPkg.version = version;
    await atomicWrite(serverPackagePath, JSON.stringify(serverPkg, null, 2) + '\n');
  } catch (error) {
    if (error.code !== 'ENOENT') throw error;
  }

  await replaceIfExists(path.join(ROOT, 'README.md'), text =>
    text.replace(/(## Versión actual\s*\n\s*)[^\n]+/i, `$1TETORD ${version} — Final Build & Release Edition.`));

  await replaceIfExists(path.join(ROOT, 'GIT_UPLOAD.md'), text =>
    text.replace(/TETORD \d+\.\d+\.\d+/g, `TETORD ${version}`));

  await fs.mkdir(DIST_TEMPLATE, { recursive: true });
  await atomicWrite(path.join(DIST_TEMPLATE, 'VERSION'), version + '\n');
  await atomicWrite(path.join(DIST_TEMPLATE, 'BUILD_INFO.txt'), `TETORD ${version}\nGenerated: ${new Date().toISOString()}\n`);

  console.log(`TETORD version synchronized: ${version}`);
}

main().catch(error => {
  console.error(error.stack || error.message);
  process.exitCode = 1;
});

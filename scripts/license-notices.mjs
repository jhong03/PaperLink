import fs from 'node:fs/promises';
import path from 'node:path';
const root = process.cwd(), packageJson = JSON.parse(await fs.readFile(path.join(root, 'package.json'), 'utf8'));
const queue = Object.keys(packageJson.dependencies).map((name) => ({ name, from: root })), seen = new Set(), notices = [];
async function packageDirectory(name, from) {
  for (let current = from; ; current = path.dirname(current)) {
    const candidate = path.join(current, 'node_modules', name);
    try { await fs.access(path.join(candidate, 'package.json')); return candidate; } catch {}
    if (current === root || current === path.dirname(current)) return null;
  }
}
while (queue.length) {
  const { name, from } = queue.shift(), directory = await packageDirectory(name, from);
  if (!directory || seen.has(directory)) continue; seen.add(directory);
  let meta; try { meta = JSON.parse(await fs.readFile(path.join(directory, 'package.json'), 'utf8')); } catch { continue; }
  queue.push(...Object.keys(meta.dependencies || {}).map((dependency) => ({ name: dependency, from: directory })));
  let license = '';
  for (const file of ['LICENSE', 'LICENSE.md', 'LICENSE.txt', 'license', 'license.md', 'LICENSE-MIT']) {
    try { license = await fs.readFile(path.join(directory, file), 'utf8'); break; } catch {}
  }
  notices.push(name + ' ' + meta.version + '\nLicense: ' + (typeof meta.license === 'string' ? meta.license : JSON.stringify(meta.license || 'See package source')) + '\nSource: ' + (meta.homepage || meta.repository?.url || 'https://www.npmjs.com/package/' + name) + '\n\n' + (license || 'Refer to the package source for its license text.') + '\n');
}
const fontboxLicense = (await fs.readFile(path.join(root, 'node_modules', 'adler-32', 'LICENSE'), 'utf8')).replace('Copyright (C) 2014-present   SheetJS LLC', 'Copyright [yyyy] [name of copyright owner]');
notices.push('@libpdf/core fontbox components\nLicense: Apache-2.0\nDerived from Apache PDFBox / FontBox, Apache Software Foundation.\nSource: https://github.com/LibPDF-js/core/tree/main/src/fontbox\nSource: https://pdfbox.apache.org/\nThese components are distributed under Apache-2.0 rather than the library\'s MIT license.\n\n' + fontboxLicense + '\n');
await fs.mkdir(path.join(root, 'docs'), { recursive: true });
await fs.writeFile(path.join(root, 'docs', 'THIRD-PARTY-NOTICES.txt'), 'PaperLink third-party notices\n\nRuntime library licenses are reproduced below. Electron and Chromium notices are also shipped with the Electron runtime.\n\n' + notices.sort().join('\n' + '='.repeat(72) + '\n\n'));
console.log('License notices written for ' + notices.length + ' runtime packages.');

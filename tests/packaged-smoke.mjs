import { chromium } from 'playwright';
import { spawn } from 'node:child_process';
import { createServer } from 'node:net';
import fs from 'node:fs/promises';
import path from 'node:path';
import assert from 'node:assert/strict';
const root = process.cwd(), binary = path.join(root, 'release', 'win-unpacked', 'PaperLink.exe');
const env = { ...process.env }; delete env.ELECTRON_RUN_AS_NODE; delete env.NODE_OPTIONS;
const server = createServer(); await new Promise((resolve) => server.listen(0, '127.0.0.1', resolve));
const port = server.address().port; await new Promise((resolve) => server.close(resolve));
const child = spawn(binary, ['--smoke', '--remote-debugging-address=127.0.0.1', '--remote-debugging-port=' + port], { env, cwd: root, windowsHide: true, stdio: ['ignore', 'pipe', 'pipe'] });
let logs = '', browser;
child.stdout.on('data', (chunk) => { logs += chunk; }); child.stderr.on('data', (chunk) => { logs += chunk; });
try {
  let ready = false;
  for (let i = 0; i < 100; i++) {
    if (child.exitCode != null) throw new Error('Packaged app exited: ' + child.exitCode + '\n' + logs);
    try { const response = await fetch('http://127.0.0.1:' + port + '/json/version'); if (response.ok) { ready = true; break; } } catch {}
    await new Promise((resolve) => setTimeout(resolve, 200));
  }
  if (!ready) throw new Error('Packaged app did not expose its test browser.\n' + logs);
  browser = await chromium.connectOverCDP('http://127.0.0.1:' + port);
  const context = browser.contexts()[0], page = context.pages()[0] || await context.waitForEvent('page'), errors = [];
  page.on('pageerror', (error) => errors.push(error.message));
  await page.waitForFunction(() => Boolean(window.__paperlinkTest));
  await page.getByRole('button', { name: 'Try sample project' }).click();
  await page.getByLabel('Increment linked value', { exact: true }).fill('150');
  const result = await page.evaluate(async () => {
    const t = window.__paperlinkTest, p = t.project(), layout = t.preview();
    const pdf = await window.paperlink.renderPDF(t.printHtml(layout.pages, p.settings));
    return { salary: p.cell('Employees', 'E2').v, pages: layout.pages.length, magic: new TextDecoder().decode(new Uint8Array(pdf).slice(0, 4)) };
  });
  assert.deepEqual(result, { salary: 4350, pages: 1, magic: '%PDF' });
  await fs.access(path.join(root, 'release', 'win-unpacked', 'resources', 'samples', 'payroll-fictional.xlsx'));
  await fs.access(path.join(root, 'release', 'win-unpacked', 'resources', 'docs', 'MANUAL-TESTING.md'));
  assert.deepEqual(errors, []);
  console.log('Packaged app passed: offline startup, linked editing/formula recalculation, native PDF output, and bundled test resources.');
} catch (error) { console.error(logs.slice(-6000)); throw error; }
finally { await browser?.close().catch(() => {}); child.kill(); }

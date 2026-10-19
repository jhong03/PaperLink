import { build } from 'esbuild';
await build({ entryPoints: ['desktop/protection-worker.mjs'], outfile: '.cache/protection-worker.cjs', bundle: true, platform: 'node', format: 'cjs', target: 'node22', legalComments: 'inline' });
console.log('Document protection worker bundled.');

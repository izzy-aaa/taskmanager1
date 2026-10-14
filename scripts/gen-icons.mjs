// Bundles the Lucide icons the app uses into src/icons.json, so icons don't
// depend on Iconify's API (which embedded previews may block).
// Run after adding an icon:  npm run icons
import fs from 'node:fs';
import { createRequire } from 'node:module';

const require = createRequire(import.meta.url);
const lucide = JSON.parse(fs.readFileSync(require.resolve('@iconify-json/lucide/icons.json'), 'utf8'));
const src = fs.readdirSync('src').filter((f) => f.endsWith('.ts')).map((f) => fs.readFileSync(`src/${f}`, 'utf8')).join('\n');
const names = new Set();
for (const m of src.matchAll(/['"`](?:lucide:)?([a-z0-9]+(?:-[a-z0-9]+)*)['"`]/g)) if (lucide.icons[m[1]]) names.add(m[1]);
for (const m of src.matchAll(/lucide:([a-z0-9-]+)/g)) if (lucide.icons[m[1]]) names.add(m[1]);
const icons = {};
for (const n of [...names].sort()) icons[n] = lucide.icons[n];
const out = { prefix: 'lucide', width: lucide.width ?? 24, height: lucide.height ?? 24, icons };
fs.writeFileSync('src/icons.json', JSON.stringify(out));
console.log(`${names.size} icons -> src/icons.json`);

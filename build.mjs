// Stitches src/page.js into src/mod.cpp as a wide raw string literal and writes the
// single-file mod that Windhawk installs: mods/outlook-refined.wh.cpp.
// Usage: node build.mjs
import { readFileSync, writeFileSync, mkdirSync } from 'node:fs';

const root = new URL('./', import.meta.url);
const PLACEHOLDER = '/*@@PAGE@@*/';
const DELIMITER = ')olkt"';

const template = readFileSync(new URL('src/mod.cpp', root), 'utf8');
const page = readFileSync(new URL('src/page.js', root), 'utf8').replace(/\r\n/g, '\n').trimEnd();

const fail = msg => { console.error(`build failed: ${msg}`); process.exit(1); };
if (template.split(PLACEHOLDER).length !== 2) fail(`src/mod.cpp must hold ${PLACEHOLDER} exactly once`);
if (page.includes(DELIMITER)) fail(`src/page.js contains the raw string delimiter ${DELIMITER}`);

const out = template.replace(PLACEHOLDER, () => page);
mkdirSync(new URL('mods/', root), { recursive: true });
writeFileSync(new URL('mods/outlook-refined.wh.cpp', root), out);
console.log(`mods/outlook-refined.wh.cpp written (${out.length} chars, page ${page.length})`);

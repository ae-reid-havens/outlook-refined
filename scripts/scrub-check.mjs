// Fails when a file that git tracks or would track holds personal data: email
// addresses, GUIDs in text form, Graph item ids, local paths, plus every term in
// scrub-terms.local.txt (one per line, case-insensitive; that file is git-ignored,
// so the terms themselves never reach the repo).
// Usage: node scripts/scrub-check.mjs   (exit 1 on any hit)
import { readFileSync, existsSync } from 'node:fs';
import { execFileSync } from 'node:child_process';

const root = new URL('../', import.meta.url);
const SELF = 'scripts/scrub-check.mjs';
const LOCAL = 'scrub-terms.local.txt';

const checks = [
  ['email address', /[A-Za-z0-9._%+-]+@[A-Za-z0-9-]+(\.[A-Za-z0-9-]+)*\.[A-Za-z]{2,}/g],
  ['GUID', /\b[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\b/gi],
  ['Graph-style id', /\bAA[MQk][A-Za-z0-9+/_-]{10,}/g],
  ['local path', /[A-Za-z]:[\\/](Users|GitHub)\b/gi],
];

const localPath = new URL(LOCAL, root);
if (existsSync(localPath)) {
  for (const term of readFileSync(localPath, 'utf8').split(/\r?\n/).map(s => s.trim()).filter(Boolean)) {
    checks.push([`local term "${term}"`, new RegExp(term.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'gi')]);
  }
} else {
  console.warn(`warning: ${LOCAL} not found; checking built-in patterns only`);
}

const files = execFileSync('git', ['ls-files', '--cached', '--others', '--exclude-standard'], { cwd: root, encoding: 'utf8' })
  .split('\n').filter(f => f && f !== SELF);

let hits = 0;
for (const file of files) {
  const text = readFileSync(new URL(file, root), 'utf8');
  text.split('\n').forEach((line, i) => {
    for (const [name, re] of checks) {
      for (const m of line.matchAll(re)) {
        hits++;
        console.error(`${file}:${i + 1}: ${name}: ${m[0]}`);
      }
    }
  });
}
if (hits) { console.error(`scrub check failed: ${hits} hit(s)`); process.exit(1); }
console.log(`scrub check passed: ${files.length} files, ${checks.length} patterns`);

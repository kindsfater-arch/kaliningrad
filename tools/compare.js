#!/usr/bin/env node
/**
 * Сверяет данные в собранном дашборде с эталонным (обычно — с текущим index.html).
 * Запуск: node tools/compare.js dist/index.html index.html
 */

'use strict';
const fs = require('fs');

function extract(file) {
  const line = fs.readFileSync(file, 'utf8').split('\n').find(l => l.startsWith('const DATA = '));
  if (!line) throw new Error(`в ${file} нет строки с данными`);
  return JSON.parse(line.replace(/^const DATA = /, '').replace(/;\s*$/, '').replace(/<\\\//g, '</'));
}

const diffs = [];
function walk(a, b, p) {
  if (a === b) return;
  const ta = Array.isArray(a) ? 'array' : a === null ? 'null' : typeof a;
  const tb = Array.isArray(b) ? 'array' : b === null ? 'null' : typeof b;
  if (ta !== tb) return diffs.push(`${p}: тип ${tb} → ${ta}`);
  if (ta === 'array') {
    if (a.length !== b.length) diffs.push(`${p}: длина ${b.length} → ${a.length}`);
    for (let i = 0; i < Math.max(a.length, b.length); i++) walk(a[i], b[i], `${p}[${i}]`);
    return;
  }
  if (ta === 'object') {
    for (const k of new Set([...Object.keys(a), ...Object.keys(b)])) walk(a[k], b[k], p ? `${p}.${k}` : k);
    return;
  }
  diffs.push(`${p}: ${JSON.stringify(b)} → ${JSON.stringify(a)}`);
}

const [mine, ref] = [extract(process.argv[2]), extract(process.argv[3])];
walk(mine.city, ref.city, 'city');
walk(mine.oblast, ref.oblast, 'oblast');

if (!diffs.length) {
  console.log('Различий нет — сборщик воспроизводит эталон точь-в-точь.');
} else {
  console.log(`Различий: ${diffs.length}\n`);
  diffs.slice(0, 60).forEach(d => console.log('  ' + d));
  if (diffs.length > 60) console.log(`  … и ещё ${diffs.length - 60}`);
}

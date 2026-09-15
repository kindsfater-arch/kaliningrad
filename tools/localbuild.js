#!/usr/bin/env node
/**
 * Локальный прогон appsscript/Code.gs без Google Apps Script.
 *
 * Подменяет DriveApp, Utilities, UrlFetchApp и PropertiesService заглушками:
 *   — «файлы на Диске» берутся из аргументов командной строки;
 *   — «репозиторий на GitHub» читается из рабочей копии, а запись идёт в отдельную папку.
 *
 * Запуск:
 *   node tools/localbuild.js <операторы.xlsx> <отчёт-области.xlsx> [--out dist] [--updated 15.09.2026]
 *
 * Прогон дважды подряд должен во второй раз сообщить, что публиковать нечего.
 */

'use strict';

const fs = require('fs');
const path = require('path');
const zlib = require('zlib');
const vm = require('vm');
const crypto = require('crypto');

// ───────────────────────── аргументы ─────────────────────────

const argv = process.argv.slice(2);
const opt = { out: 'dist', updated: null, files: [] };
for (let i = 0; i < argv.length; i++) {
  if (argv[i] === '--out') opt.out = argv[++i];
  else if (argv[i] === '--updated') opt.updated = argv[++i];
  else opt.files.push(argv[i]);
}
if (opt.files.length === 0) {
  console.error('Укажите хотя бы один xlsx. См. комментарий в начале файла.');
  process.exit(2);
}

const REPO_ROOT = path.resolve(__dirname, '..');
const OUT_DIR = path.resolve(REPO_ROOT, opt.out);

// ───────────────────────── чтение zip ─────────────────────────
// Свой разбор, чтобы не тянуть зависимости: центральный каталог → локальные заголовки.

function unzip(buf) {
  let eocd = -1;
  for (let i = buf.length - 22; i >= 0 && i > buf.length - 66000; i--) {
    if (buf.readUInt32LE(i) === 0x06054b50) { eocd = i; break; }
  }
  if (eocd < 0) throw new Error('не найден конец архива zip');

  const count = buf.readUInt16LE(eocd + 10);
  let p = buf.readUInt32LE(eocd + 16);
  const out = [];

  for (let n = 0; n < count; n++) {
    if (buf.readUInt32LE(p) !== 0x02014b50) throw new Error('повреждён центральный каталог zip');
    const method = buf.readUInt16LE(p + 10);
    const compSize = buf.readUInt32LE(p + 20);
    const nameLen = buf.readUInt16LE(p + 28);
    const extraLen = buf.readUInt16LE(p + 30);
    const commentLen = buf.readUInt16LE(p + 32);
    const localOff = buf.readUInt32LE(p + 42);
    const name = buf.toString('utf8', p + 46, p + 46 + nameLen);

    const lNameLen = buf.readUInt16LE(localOff + 26);
    const lExtraLen = buf.readUInt16LE(localOff + 28);
    const start = localOff + 30 + lNameLen + lExtraLen;
    const raw = buf.slice(start, start + compSize);
    out.push(makeBlob(method === 8 ? zlib.inflateRawSync(raw) : raw, name));

    p += 46 + nameLen + extraLen + commentLen;
  }
  return out;
}

// ───────────────────────── заглушки Apps Script ─────────────────────────

function makeBlob(buffer, name) {
  return {
    _buf: Buffer.isBuffer(buffer) ? buffer : Buffer.from(buffer),
    getName() { return name || ''; },
    setContentType() { return this; },
    copyBlob() { return makeBlob(this._buf, name); },
    getDataAsString(enc) { return this._buf.toString(enc && /utf-?8/i.test(enc) ? 'utf8' : 'utf8'); },
    // Apps Script отдаёт байты знаковыми (-128..127) — воспроизводим, чтобы хеши совпадали
    getBytes() { return Array.from(this._buf).map(b => (b > 127 ? b - 256 : b)); }
  };
}

const Utilities = {
  DigestAlgorithm: { SHA_1: 'sha1' },
  unzip(blob) { return unzip(blob._buf); },
  // как в Apps Script: принимает и строку, и массив байтов
  newBlob(data) {
    return makeBlob(Array.isArray(data) ? Buffer.from(data.map(b => b & 0xff)) : Buffer.from(String(data), 'utf8'));
  },
  base64Encode(bytes) {
    const buf = Array.isArray(bytes) ? Buffer.from(bytes.map(b => b & 0xff)) : Buffer.from(String(bytes), 'utf8');
    return buf.toString('base64');
  },
  base64Decode(s) {
    return Array.from(Buffer.from(s, 'base64')).map(b => (b > 127 ? b - 256 : b));
  },
  computeDigest(alg, bytes) {
    const buf = Buffer.from(bytes.map(b => b & 0xff));
    return Array.from(crypto.createHash(alg).update(buf).digest()).map(b => (b > 127 ? b - 256 : b));
  }
};

function parseDdMmYyyy(s) {
  const m = /^(\d{2})\.(\d{2})\.(\d{4})$/.exec(s || '');
  return m ? new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]), 12, 0, 0) : null;
}

const driveFiles = opt.files.map(f => {
  const abs = path.resolve(f);
  const buf = fs.readFileSync(abs);
  const when = parseDdMmYyyy(opt.updated) || fs.statSync(abs).mtime;
  return {
    getName() { return path.basename(abs); },
    getBlob() { return makeBlob(buf, path.basename(abs)); },
    getLastUpdated() { return when; }
  };
});

const DriveApp = {
  getFolderById() {
    let i = 0;
    return { getFiles: () => ({ hasNext: () => i < driveFiles.length, next: () => driveFiles[i++] }) };
  }
};

// ───────────── эмуляция GitHub Contents API поверх файловой системы ─────────────
// Чтение: сначала папка вывода, затем рабочая копия. Запись: только папка вывода.

function gitBlobSha(buf) {
  return crypto.createHash('sha1').update(Buffer.concat([Buffer.from(`blob ${buf.length}\0`), buf])).digest('hex');
}

function resolveRead(rel) {
  const inOut = path.join(OUT_DIR, rel);
  if (fs.existsSync(inOut)) return inOut;
  const inRepo = path.join(REPO_ROOT, rel);
  return fs.existsSync(inRepo) ? inRepo : null;
}

const writes = [];

const UrlFetchApp = {
  fetch(url, options) {
    const m = /api\.github\.com\/repos\/[^/]+\/[^/]+\/contents\/(.+)$/.exec(url);
    if (!m) return { getResponseCode: () => 404, getContentText: () => '' };
    const rel = decodeURI(m[1].split('?')[0]);

    if ((options.method || 'GET').toUpperCase() === 'GET') {
      const file = resolveRead(rel);
      if (!file) return { getResponseCode: () => 404, getContentText: () => '' };
      const buf = fs.readFileSync(file);
      const body = JSON.stringify({ sha: gitBlobSha(buf), content: buf.toString('base64') });
      return { getResponseCode: () => 200, getContentText: () => body };
    }

    const payload = JSON.parse(options.payload);
    const dest = path.join(OUT_DIR, rel);
    fs.mkdirSync(path.dirname(dest), { recursive: true });
    fs.writeFileSync(dest, Buffer.from(payload.content, 'base64'));
    writes.push(rel);
    return { getResponseCode: () => 200, getContentText: () => '{}' };
  }
};

const PropertiesService = {
  getScriptProperties: () => ({ getProperty: key => (key === 'GITHUB_TOKEN' ? 'локальный-прогон' : null) })
};

const Logger = { log: msg => console.log('  ' + msg) };
const MailApp = { sendEmail() {} };
const ScriptApp = {
  getProjectTriggers: () => [],
  newTrigger() { throw new Error('setup() локально не запускается'); }
};

// ───────────────────────── запуск ─────────────────────────

const sandbox = { Utilities, DriveApp, UrlFetchApp, PropertiesService, Logger, MailApp, ScriptApp, console, Date, Math, JSON, String, Number, Array, Object, RegExp, isNaN, parseInt, parseFloat, Error };
vm.createContext(sandbox);
vm.runInContext(fs.readFileSync(path.join(REPO_ROOT, 'appsscript', 'Code.gs'), 'utf8'), sandbox, { filename: 'Code.gs' });

fs.mkdirSync(OUT_DIR, { recursive: true });
console.log('Сборка…');
const report = vm.runInContext('build()', sandbox);
console.log(report);
console.log(writes.length ? `Записано в ${path.relative(process.cwd(), OUT_DIR)}: ${writes.join(', ')}` : 'Файлы не записывались.');

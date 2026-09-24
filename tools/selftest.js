#!/usr/bin/env node
/**
 * Проверки разбора, не требующие файлов: поиск колонок, разбор ячеек готовности,
 * заслон от порчи данных. Запуск: node tools/selftest.js
 */

'use strict';
const vm = require('vm');
const fs = require('fs');
const path = require('path');

const logged = [];
const Logger = { log: msg => logged.push(String(msg)) };
const sandbox = { console, Logger, Math, JSON, String, Number, Array, Object, RegExp, Date, isNaN, parseInt, parseFloat, Error };
vm.createContext(sandbox);
vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'appsscript', 'Code.gs'), 'utf8'), sandbox, { filename: 'Code.gs' });
const g = name => vm.runInContext(name, sandbox);

let failed = 0;
function check(name, got, want) {
  const ok = JSON.stringify(got) === JSON.stringify(want);
  if (!ok) failed++;
  console.log(`  ${ok ? 'OK  ' : 'СБОЙ'}  ${name}${ok ? '' : `\n         получено ${JSON.stringify(got)}, ожидалось ${JSON.stringify(want)}`}`);
}
function throws(name, fn, fragment) {
  let msg = null;
  try { fn(); } catch (e) { msg = e.message; }
  const ok = msg !== null && msg.indexOf(fragment) >= 0;
  if (!ok) failed++;
  console.log(`  ${ok ? 'OK  ' : 'СБОЙ'}  ${name}${ok ? '' : `\n         получено: ${msg === null ? 'ошибки не было' : msg}`}`);
}

// ── поиск колонок: точное совпадение должно побеждать подстроку ──
console.log('\nПоиск колонок по заголовку');
const findCol = g('findCol_');
check('«Городской округ» не перехватывает «Город»', findCol(['Дата', 'Городской округ', 'Город'], ['город']), 2);
check('«Город» без соседей', findCol(['Дата', 'Город', 'ИНН'], ['город']), 1);
check('обрезанный заголовок ловится подстрокой', findCol(['Признак наличия активного меню на'], ['признак наличия активного меню']), 0);
check('«Подключено» не путается с «Признак подключения…»', findCol(['Признак подключения интеграции с', 'Установлено', 'Подключено'], ['подключено']), 2);
check('«Установлено» не путается с «Обновлено оборудование СБС»', findCol(['Обновлено оборудование СБС', 'Установлено'], ['установлено']), 1);
check('колонки нет', findCol(['Дата', 'ИНН'], ['город']), -1);

// ── разбор ячейки «название - да/нет» ──
console.log('\nРазбор ячеек готовности');
const split = g('splitFlag_');
check('обычная запись', split('МАОУ СОШ № 24 - Да'), { name: 'МАОУ СОШ № 24', flag: true });
check('без пробела перед флагом', split('МАОУ СОШ № 57 -да'), { name: 'МАОУ СОШ № 57', flag: true });
check('без пробела перед дефисом', split('МАОУ СОШ № 2- да'), { name: 'МАОУ СОШ № 2', flag: true });
check('отрицание', split('МАОУ гимназия № 32 - нет'), { name: 'МАОУ гимназия № 32', flag: false });
check('перенос строки в конце', split('МАОУ СОШ № 19 - нет\n'), { name: 'МАОУ СОШ № 19', flag: false });
check('точка в названии', split('МАОУ СОШ №9 им. Дьякова П.М. - да'), { name: 'МАОУ СОШ №9 им. Дьякова П.М.', flag: true });
check('дефис в названии, флаг берётся последний', split('МБОУ Школа - Интернат - да'), { name: 'МБОУ Школа - Интернат', flag: true });
check('название без флага', split('МАОУ лицей 35'), { name: 'МАОУ лицей 35', flag: null });
check('пустая ячейка', split(''), { name: '', flag: null });

// ── короткое имя оператора ──
console.log('\nКороткое имя оператора');
const short = g('kshpShort_');
check('ёлочки', short('Общество с ограниченной ответственностью «Новый уровень»'), 'Новый уровень');
check('прямые кавычки', short('ОБЩЕСТВО С ОГРАНИЧЕННОЙ ОТВЕТСТВЕННОСТЬЮ "МИТПРОДТОРГ"'), 'МИТПРОДТОРГ');
check('без кавычек — обрезка', short('Индивидуальный предприниматель Дараган Людмила Юрьевна'), 'Индивидуальный предприниматель Дараган Л');

// ── нормализация муниципалитета ──
console.log('\nНазвание муниципалитета');
const city = g('cityShort_');
check('город', city('Багратионовск город'), 'Багратионовск');
check('пгт', city('Янтарный поселок городского типа'), 'Янтарный (пгт)');

// ── заслон от порчи данных ──
console.log('\nЗаслон от порчи данных');
const sanityOblast = g('sanityOblast_');
const sanityCity = g('sanityCity_');
const school = extra => Object.assign({ pupils: 100, ls: 90, cards: 10, parents: 5, terminals: 1, menu: true }, extra);
const prevObl = { schools: [school(), school(), school(), school(), school()] };

check('нормальный срез проходит', sanityOblast({ schools: prevObl.schools.slice() }, prevObl), undefined);
check('первый запуск — сравнивать не с чем', sanityOblast({ schools: [] }, null), undefined);
throws('колонка учащихся перестала читаться',
  () => sanityOblast({ schools: prevObl.schools.map(s => Object.assign({}, s, { pupils: 0 })) }, prevObl),
  'Кол-во учащихся');
throws('школы массово пропали',
  () => sanityOblast({ schools: [school()] }, prevObl),
  'школ было 5, стало 1');

// структурный заслон: колонка читалась вчера и не нашлась сегодня.
// Публикацию это не отменяет — страница сама скажет «нет в выгрузке».
const withFound = (found, extra) => ({ schools: prevObl.schools.map(s => Object.assign({}, s, extra)), found: found });
const foundAll = { pupils: true, skudConn: true, skudSoft: true };
logged.length = 0;
check('колонку СКУД переименовали — публикация продолжается',
  sanityOblast(withFound({ pupils: true, skudConn: false, skudSoft: true }, { skudConn: false }), withFound(foundAll)),
  undefined);
check('о пропаже сообщается в лог',
  logged.length === 1 && /Контроллеры подключены/.test(logged[0]), true);
check('о пропаже сообщается один раз, а не дважды',
  (logged[0].match(/Контроллеры подключены/g) || []).length, 1);
check('колонка на месте — претензий нет',
  sanityOblast(withFound(foundAll), withFound(foundAll)), undefined);
check('новая колонка, которой раньше не было, не считается пропажей',
  sanityOblast(withFound(foundAll), withFound({ pupils: true, skudConn: true, skudSoft: false })), undefined);
check('спад в пределах нормы не мешает',
  sanityOblast({ schools: prevObl.schools.slice(0, 4) }, prevObl), undefined);

// а вот живая колонка, где всё обнулилось, сборку останавливает
throws('колонка на месте, но обнулилась — сборка встаёт',
  () => sanityOblast(withFound(foundAll, { cards: 0 }), withFound(foundAll)),
  'колонка «Общее кол-во привязанных карт» больше не читается');
throws('транзакции обнулились — сборка встаёт',
  () => sanityOblast(withFound(foundAll, { spend: 0 }), withFound(foundAll, { spend: 500 })),
  'колонка «Сумма списаний за питание» больше не читается');
check('устаревший Code.gs в облаке опознаётся',
  (() => { try { sanityOblast({ schools: prevObl.schools.slice() }, withFound(foundAll)); } catch (e) { return /устаревшая версия Code.gs/.test(e.message); } })(), true);

const op = n => ({ schools: Array.from({ length: n }, () => ({ sales: [10, 20] })) });
const zeroed = n => ({ schools: Array.from({ length: n }, () => ({ sales: [0, 0] })) });
const prevCity = { operators: [op(5), op(5), op(5)] };
check('нормальный город проходит', sanityCity({ operators: [op(5), op(5), op(5)] }, prevCity), undefined);
throws('операторы пропали', () => sanityCity({ operators: [op(5)] }, prevCity), 'операторов было 3, стало 1');
throws('продажи обнулились',
  () => sanityCity({ operators: [zeroed(5), zeroed(5), zeroed(5)] }, prevCity),
  'все продажи обнулились');

console.log(failed ? `\nПровалено проверок: ${failed}` : '\nВсе проверки пройдены.');
process.exit(failed ? 1 : 0);

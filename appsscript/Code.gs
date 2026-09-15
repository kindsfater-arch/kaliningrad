/**
 * Дашборд «Карта школьника» — автосборка из двух Excel-файлов на Google Диске.
 *
 * Что делает:
 *   1. Находит в папке на Диске два свежих xlsx (тип определяется по содержимому);
 *   2. читает их без сторонних библиотек (Utilities.unzip + разбор XML);
 *   3. собирает модель данных для двух контуров — город и область;
 *   4. берёт шаблон build/template.html из репозитория на GitHub;
 *   5. подставляет данные и публикует index.html обратно в репозиторий.
 *
 * Настройка — см. README.md в репозитории.
 * Точки входа: setup() — один раз при установке, build() — ежедневный прогон.
 */

// ══════════════════════════ Конфигурация ══════════════════════════

var DEFAULTS = {
  FOLDER_ID:     '1q8qx9jrjup346OllQqBrOyKU0ozlM1_w',
  REPO:          'kindsfater-arch/kaliningrad',
  BRANCH:        'main',
  TEMPLATE_PATH: 'build/template.html',
  AVG_CHECK:     '150',   // рублей на одну транзакцию
  BUILD_HOUR:    '6',     // час ежедневного запуска
  NOTIFY_EMAIL:  ''       // пусто — письма об ошибках не отправляются
};

var PLACEHOLDER = '<!--DATA-->';

function cfg_(key) {
  var v = PropertiesService.getScriptProperties().getProperty(key);
  return (v === null || v === undefined || v === '') ? DEFAULTS[key] : v;
}

// ══════════════════════════ Точки входа ══════════════════════════

/** Выполняется один раз при установке: создаёт ежедневный триггер и делает первую сборку. */
function setup() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'build') ScriptApp.deleteTrigger(triggers[i]);
  }
  ScriptApp.newTrigger('build').timeBased().everyDays(1).atHour(Number(cfg_('BUILD_HOUR'))).create();
  Logger.log('Ежедневный триггер создан на ' + cfg_('BUILD_HOUR') + ':00. Запускаю первую сборку…');
  Logger.log(build());
}

/** Основной прогон: собрать дашборд и, если он изменился, опубликовать. */
function build() {
  try {
    var data = collect_();
    var html = render_(fetchTemplate_(), data);
    var report = publish_(html, data);
    Logger.log(report);
    return report;
  } catch (err) {
    notify_(err);
    throw err;
  }
}

// ══════════════════════════ Сбор данных ══════════════════════════

function collect_() {
  var src = findSources_(cfg_('FOLDER_ID'));

  var city = src.city
    ? parseCity_(src.city.book, src.city.updated)
    : loadPublished_('data/city.json');
  var oblast = src.oblast
    ? parseOblast_(src.oblast.book, src.oblast.updated)
    : loadPublished_('data/oblast.json');

  if (!city) {
    throw new Error('В папке нет файла операторов питания (ищу шапку «Получены доступы»), ' +
                    'и в репозитории нет ранее опубликованного data/city.json. Публикация отменена.');
  }
  if (!oblast) {
    throw new Error('В папке нет отчёта по области (ищу шапку «ID школы» и «Наименование КШП»), ' +
                    'и в репозитории нет ранее опубликованного data/oblast.json. Публикация отменена.');
  }

  return {
    city: city,
    oblast: oblast,
    updated: laterDate_(city.updated, oblast.updated)
  };
}

/** Перебирает файлы папки и выбирает самый свежий файл каждого типа. */
function findSources_(folderId) {
  var folder = DriveApp.getFolderById(folderId);
  var it = folder.getFiles();
  var best = { city: null, oblast: null };
  var seen = 0;

  while (it.hasNext()) {
    var f = it.next();
    var name = f.getName();
    if (!/\.xlsx$/i.test(name)) continue;
    seen++;
    var book;
    try {
      book = readWorkbook_(f.getBlob());
    } catch (e) {
      Logger.log('Пропущен «' + name + '»: не читается как xlsx (' + e.message + ')');
      continue;
    }
    var kind = classify_(book);
    if (!kind) {
      Logger.log('Пропущен «' + name + '»: не похож ни на один из двух ожидаемых отчётов');
      continue;
    }
    var when = f.getLastUpdated();
    if (!best[kind] || when.getTime() > best[kind].updated.getTime()) {
      best[kind] = { book: book, updated: when, name: name };
    }
  }

  Logger.log('Просмотрено xlsx: ' + seen +
             '; город: ' + (best.city ? best.city.name : '—') +
             '; область: ' + (best.oblast ? best.oblast.name : '—'));
  return best;
}

/** Определяет тип отчёта по тексту шапки, а не по имени файла. */
function classify_(book) {
  for (var s = 0; s < book.sheets.length; s++) {
    var rows = book.sheets[s].rows;
    var head = '';
    for (var i = 0; i < Math.min(rows.length, 8); i++) head += ' ' + rows[i].join(' ');
    head = head.toLowerCase();
    if (head.indexOf('получены доступы') >= 0) return 'city';
    if (head.indexOf('id школы') >= 0 && head.indexOf('наименование кшп') >= 0) return 'oblast';
  }
  return null;
}

// ══════════════════════════ Чтение XLSX ══════════════════════════
//
// Свой разбор вместо готовой библиотеки — по четырём причинам, каждая проверена
// на реальных файлах:
//   1. Два файла сделаны разными генераторами: в одном теги с префиксом
//      (<s:row>, <s:c>), в другом без (<row>, <c>). Все выражения ниже
//      безразличны к префиксу пространства имён.
//   2. В словаре строк одного из файлов 569 записей, но 941 тег <t> — часть
//      строк разбита на фрагменты. Склеиваем все <t> внутри одного <si>.
//   3. Номер строки берём из атрибута r: пустые строки в XML отсутствуют,
//      и счёт по порядку сдвигает данные.
//   4. Стили не читаем вовсе — в одном из файлов xl/styles.xml битый.

/** Читает xlsx из Blob. Возвращает { sheets: [ { name, rows } ] }, rows[i] — строка Excel i+1. */
function readWorkbook_(blob) {
  // Декодируем в текст только нужные записи: в архиве бывает двоичное
  // содержимое (например, xl/printerSettings/*.bin), гнать его через UTF-8 незачем.
  var parts = Utilities.unzip(blob.setContentType('application/zip'));
  var raw = {};
  for (var i = 0; i < parts.length; i++) raw[parts[i].getName()] = parts[i];
  var cache = {};
  var files = function (name) {
    if (!raw[name]) return null;
    if (!(name in cache)) cache[name] = raw[name].getDataAsString('UTF-8');
    return cache[name];
  };
  if (!files('xl/workbook.xml')) throw new Error('нет xl/workbook.xml');

  var shared = parseSharedStrings_(files('xl/sharedStrings.xml') || '');

  var rels = {};
  var relRe = /<Relationship\b([^>]*?)(?:\/>|>)/g, m;
  var relXml = files('xl/_rels/workbook.xml.rels') || '';
  while ((m = relRe.exec(relXml))) {
    var ra = attrs_(m[1]);
    if (ra.Id) rels[ra.Id] = ra.Target;
  }

  var sheets = [];
  var shRe = /<(?:\w+:)?sheet\b([^>]*?)(?:\/>|>)/g;
  while ((m = shRe.exec(files('xl/workbook.xml')))) {
    var sa = attrs_(m[1]);
    var target = sa['r:id'] ? rels[sa['r:id']] : null;
    if (!target) continue;
    var path = target.charAt(0) === '/' ? target.substring(1)
             : (target.indexOf('xl/') === 0 ? target : 'xl/' + target);
    var xml = files(path);
    if (!xml) continue;
    sheets.push({ name: sa.name || '', rows: parseSheet_(xml, shared) });
  }
  if (!sheets.length) throw new Error('в книге нет читаемых листов');
  return { sheets: sheets };
}

function parseSharedStrings_(xml) {
  var out = [];
  if (!xml) return out;
  var re = /<(?:\w+:)?si\b[^>]*?(?:\/>|>([\s\S]*?)<\/(?:\w+:)?si>)/g, m;
  while ((m = re.exec(xml))) out.push(m[1] === undefined ? '' : collectText_(m[1]));
  return out;
}

/** Склеивает все <t> внутри фрагмента — ловушка №2. */
function collectText_(frag) {
  var out = '', re = /<(?:\w+:)?t\b[^>]*?(?:\/>|>([\s\S]*?)<\/(?:\w+:)?t>)/g, m;
  while ((m = re.exec(frag))) out += (m[1] === undefined ? '' : unesc_(m[1]));
  return out;
}

function parseSheet_(xml, shared) {
  var rows = [];
  var rowRe = /<(?:\w+:)?row\b([^>]*?)(?:\/>|>([\s\S]*?)<\/(?:\w+:)?row>)/g;
  var cellRe = /<(?:\w+:)?c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/(?:\w+:)?c>)/g;
  var valRe = /<(?:\w+:)?v\b[^>]*?(?:\/>|>([\s\S]*?)<\/(?:\w+:)?v>)/;
  var rm, maxRow = 0;

  while ((rm = rowRe.exec(xml))) {
    var rn = parseInt(attrs_(rm[1]).r, 10);      // номер строки — ловушка №3
    if (!rn) continue;
    var cells = [];
    var body = rm[2] || '';
    var cm;
    cellRe.lastIndex = 0;
    while ((cm = cellRe.exec(body))) {
      var ca = attrs_(cm[1]);
      var ci = colNum_(ca.r);
      if (ci < 0) continue;
      var inner = cm[2];
      if (inner === undefined || inner === '') continue;   // пустая ячейка <c .../>
      var val;
      if (ca.t === 'inlineStr') {
        val = collectText_(inner);
      } else {
        var vm = valRe.exec(inner);
        if (!vm) continue;
        var raw = vm[1] === undefined ? '' : unesc_(vm[1]);
        if (ca.t === 's') {
          val = shared[parseInt(raw, 10)];
          if (val === undefined) val = '';
        } else if (ca.t === 'b') {
          val = (raw === '1');
        } else if (ca.t === 'str' || ca.t === 'e') {
          val = raw;
        } else {
          var n = Number(raw);
          val = (raw !== '' && !isNaN(n)) ? n : raw;
        }
      }
      if (val === '' || val === undefined || val === null) continue;
      cells[ci] = val;
    }
    rows[rn - 1] = cells;
    if (rn > maxRow) maxRow = rn;
  }
  for (var i = 0; i < maxRow; i++) if (!rows[i]) rows[i] = [];
  return rows;
}

function attrs_(s) {
  var out = {}, re = /([\w:.\-]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g, m;
  while ((m = re.exec(s || ''))) out[m[1]] = unesc_(m[2] !== undefined ? m[2] : m[3]);
  return out;
}

/** "AB12" → 27 (нумерация колонок с нуля). */
function colNum_(ref) {
  if (!ref) return -1;
  var n = 0, any = false;
  for (var i = 0; i < ref.length; i++) {
    var ch = ref.charCodeAt(i);
    if (ch >= 65 && ch <= 90) { n = n * 26 + (ch - 64); any = true; }
    else if (ch >= 97 && ch <= 122) { n = n * 26 + (ch - 96); any = true; }
    else break;
  }
  return any ? n - 1 : -1;
}

function unesc_(s) {
  if (s === undefined || s === null) return '';
  s = String(s);
  if (s.indexOf('&') < 0) return s;
  return s
    .replace(/&#x([0-9a-fA-F]+);/g, function (_, h) { return String.fromCharCode(parseInt(h, 16)); })
    .replace(/&#(\d+);/g, function (_, d) { return String.fromCharCode(parseInt(d, 10)); })
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"').replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&');
}

// ══════════════════════════ Общие помощники ══════════════════════════

function txt_(v) { return (v === undefined || v === null) ? '' : String(v).trim(); }
function norm_(v) { return txt_(v).toLowerCase().replace(/\s+/g, ' '); }

function num_(v) {
  if (typeof v === 'number') return v;
  var s = txt_(v).replace(/ |\s/g, '').replace(',', '.');
  if (!s) return 0;
  var n = Number(s);
  return isNaN(n) ? 0 : n;
}

function yes_(v) { return /^да$/i.test(txt_(v)); }

/** Ищет колонку по подстроке заголовка. Заголовки в источнике бывают обрезаны — сверяем по началу. */
function findCol_(row, needles) {
  for (var i = 0; i < row.length; i++) {
    var h = norm_(row[i]);
    if (!h) continue;
    for (var j = 0; j < needles.length; j++) if (h.indexOf(needles[j]) >= 0) return i;
  }
  return -1;
}

function need_(idx, what) {
  if (idx < 0) throw new Error('В файле не найдена колонка «' + what + '» — структура отчёта изменилась');
  return idx;
}

function pad2_(n) { return (n < 10 ? '0' : '') + n; }

/** Серийный номер Excel → дата (база 1899-12-30, считаем в UTC, чтобы не зависеть от пояса). */
function serialDate_(n) {
  return new Date(Date.UTC(1899, 11, 30) + Math.round(n) * 86400000);
}

function ddmm_(d) { return pad2_(d.getUTCDate()) + '.' + pad2_(d.getUTCMonth() + 1); }

function ddmmyyyy_(d) {
  return pad2_(d.getDate()) + '.' + pad2_(d.getMonth() + 1) + '.' + d.getFullYear();
}

function toIso_(ddmmyyyy) {
  var m = /^(\d{2})\.(\d{2})\.(\d{4})$/.exec(txt_(ddmmyyyy));
  return m ? m[3] + '-' + m[2] + '-' + m[1] : null;
}

function laterDate_(a, b) {
  var ia = toIso_(a), ib = toIso_(b);
  if (!ia) return b || '';
  if (!ib) return a || '';
  return ia >= ib ? a : b;
}

// ══════════════════════════ Разбор: город ══════════════════════════

/**
 * «Операторы питания»: 12 операторов, 43 школы Калининграда, продажи по дням.
 * Особенности листа: блок оператора начинается строкой с номером, между блоками
 * пустая строка; готовность записана текстом прямо в ячейке («МАОУ СОШ № 24 - Да»);
 * колонки продаж дописываются вправо по мере появления новых дней.
 */
function parseCity_(book, updated) {
  var rows = null;
  for (var s = 0; s < book.sheets.length && !rows; s++) {
    if (findCol_(book.sheets[s].rows[1] || [], ['получены доступы']) >= 0) rows = book.sheets[s].rows;
  }
  if (!rows) rows = book.sheets[0].rows;
  var H = rows[1] || [];   // строка 2 — заголовки
  var S = rows[2] || [];   // строка 3 — подзаголовки: корпуса и даты продаж
  var avg = Number(cfg_('AVG_CHECK')) || 150;

  var cHot   = need_(findCol_(H, ['меню горячего']), 'Подготовлено меню горячего питания');
  var cBuf   = need_(findCol_(H, ['меню буфета']), 'Подготовлено меню буфета');
  var cSets  = need_(findCol_(H, ['составлены комлекс', 'составлены комплекс']), 'Составлены комплексы');
  var cStart = need_(findCol_(H, ['начаты продажи']), 'Начаты продажи');
  var cCash  = findCol_(H, ['зарегистрированы кассиры']);
  var cAssrt = findCol_(H, ['заведен ассортимент']);
  var cNote  = findCol_(H, ['примечания']);
  var cIssue = findCol_(H, ['замечания']);
  var cAdded = -1;
  for (var i = 0; i < H.length; i++) if (/^на .*школ/i.test(txt_(H[i]))) { cAdded = i; break; }

  // Второй корпус существует, только если подзаголовок это подтверждает
  var cHot2 = /корпус\s*2/i.test(txt_(S[cHot + 1])) ? cHot + 1 : -1;
  var cBuf2 = /корпус\s*2/i.test(txt_(S[cBuf + 1])) ? cBuf + 1 : -1;

  // Колонки дат: серийники Excel в подзаголовке, но только те, где есть хоть одно значение
  // (в файле заранее заготовлены пустые колонки под будущие дни).
  var dateCols = [];
  for (var c = 0; c < S.length; c++) {
    var sv = S[c];
    if (typeof sv !== 'number' || sv < 40000 || sv > 80000) continue;
    var filled = false;
    for (var r = 3; r < rows.length && !filled; r++) {
      if (txt_(rows[r][c]) !== '') filled = true;
    }
    if (filled) dateCols.push({ col: c, date: serialDate_(sv) });
  }
  if (!dateCols.length) throw new Error('В файле операторов не найдено ни одной заполненной колонки продаж');

  var ops = [];
  var cur = null;
  var held = { note: '', added: '', issues: '' };   // примечания первого оператора лежат выше его блока

  for (var rr = 2; rr < rows.length; rr++) {
    var row = rows[rr] || [];

    // Начало блока оператора
    var no = row[0];
    if (txt_(no) !== '' && !isNaN(Number(no))) {
      cur = {
        num: Number(no),
        name: txt_(row[1]),
        cashiers: txt_(row[cCash]),
        assortment: txt_(row[cAssrt]),
        cashier_flag: null,
        note: '', added: '', issues: '',
        schools: []
      };
      ops.push(cur);
      if (ops.length === 1) {
        cur.note = held.note; cur.added = held.added; cur.issues = held.issues;
      }
    }

    // Примечания оператора (объединённая ячейка отдаёт значение в первой строке блока,
    // а у первого оператора она начинается ещё на строке подзаголовков)
    var note = txt_(row[cNote]), added = txt_(row[cAdded]), issues = txt_(row[cIssue]);
    if (note || added || issues) {
      var sink = cur || held;
      if (note) sink.note = note;
      if (added) sink.added = added;
      if (issues) sink.issues = issues;
    }

    if (rr < 3 || !cur) continue;   // строки 1–3 — шапка, школ там нет

    var cells = [row[cHot], cHot2 >= 0 ? row[cHot2] : undefined,
                 row[cBuf], cBuf2 >= 0 ? row[cBuf2] : undefined,
                 row[cSets], row[cStart]];
    var name = '';
    for (var k = 0; k < cells.length && !name; k++) name = splitFlag_(cells[k]).name;
    if (!name) continue;           // строка-разделитель между операторами

    var sales = [];
    for (var d = 0; d < dateCols.length; d++) {
      sales.push(Math.round(num_(row[dateCols[d].col]) / avg));
    }

    cur.schools.push({
      name: name,
      hot1: splitFlag_(row[cHot]).flag,
      hot2: cHot2 >= 0 ? splitFlag_(row[cHot2]).flag : null,
      buf1: splitFlag_(row[cBuf]).flag,
      buf2: cBuf2 >= 0 ? splitFlag_(row[cBuf2]).flag : null,
      sets: splitFlag_(row[cSets]).flag,
      started: splitFlag_(row[cStart]).flag,
      sales: sales
    });
  }

  // Признак «кассиров меньше, чем терминалов» в выгрузке отдельным полем не приходит —
  // он выводится из примечания оператора.
  for (var o = 0; o < ops.length; o++) {
    if (/терминал/i.test(ops[o].note) && /кассир/i.test(ops[o].note)) {
      ops[o].cashier_flag = 'кассиров меньше, чем терминалов';
    }
  }

  var dates = [];
  for (var dd = 0; dd < dateCols.length; dd++) dates.push(ddmm_(dateCols[dd].date));

  return { updated: ddmmyyyy_(updated), dates: dates, operators: ops };
}

/**
 * «МАОУ СОШ № 24 - Да» → { name: 'МАОУ СОШ № 24', flag: true }.
 * В источнике встречаются варианты «- Да», «- да», «-да», «№ 2- да» и перенос строки в конце.
 * Дефис берём последний: в названиях школ этого файла дефисов нет (проверено на всех 43).
 */
function splitFlag_(v) {
  var s = txt_(v).replace(/\s+/g, ' ');
  if (!s) return { name: '', flag: null };
  var m = /^(.*)-\s*(да|нет)$/i.exec(s);
  if (!m) return { name: s, flag: null };
  return { name: txt_(m[1]), flag: /^да$/i.test(m[2]) };
}

// ══════════════════════════ Разбор: область ══════════════════════════

/**
 * «Отчет Калининград …_Dashboard»: срез на один день по школам области.
 * Лист назван датой среза. Колонки ищем по тексту заголовка, а не по номеру:
 * в выгрузке они уже приходят обрезанными и могут переставляться.
 */
function parseOblast_(book, updated) {
  var sheet = null, bestKey = -2;
  for (var s = 0; s < book.sheets.length; s++) {
    var iso = toIso_(book.sheets[s].name);
    var key = iso ? Number(iso.replace(/-/g, '')) : -1;
    if (key > bestKey) { bestKey = key; sheet = book.sheets[s]; }
  }
  var rows = sheet.rows;

  var hr = -1;
  for (var i = 0; i < Math.min(rows.length, 12); i++) {
    if (findCol_(rows[i], ['id школы']) >= 0) { hr = i; break; }
  }
  if (hr < 0) throw new Error('В отчёте по области не найдена строка заголовков (нет колонки «ID школы»)');
  var H = rows[hr];

  var C = {
    city:          need_(findCol_(H, ['город']), 'Город'),
    school:        need_(findCol_(H, ['краткое наименование школы']), 'Краткое наименование школы'),
    kshp:          need_(findCol_(H, ['наименование кшп']), 'Наименование КШП'),
    pupils:        findCol_(H, ['кол-во учащихся']),
    ls:            findCol_(H, ['кол-во привязанных лс']),
    cards:         findCol_(H, ['общее кол-во привязанных карт']),
    parents:       findCol_(H, ['кол-во созданных лк родителя']),
    tariffs:       findCol_(H, ['кол-во тарифов в школе']),
    tariffsActive: findCol_(H, ['количество активных тарифов']),
    lsTariff:      findCol_(H, ['кол-во лс, привязанных к тарифам']),
    menu:          findCol_(H, ['признак наличия активного меню']),
    complexes:     findCol_(H, ['кол-во активных комплексов в меню']),
    dishes:        findCol_(H, ['кол-во активных блюд']),
    orders:        findCol_(H, ['кол-во заявок на питание']),
    sbs:           findCol_(H, ['обновлено оборудование сбс']),
    terminals:     findCol_(H, ['кол-во терминалов']),
    skudInst:      findCol_(H, ['установлено']),
    skudConn:      findCol_(H, ['подключено'])
  };

  var schools = [];
  for (var r = hr + 1; r < rows.length; r++) {
    var row = rows[r] || [];
    var school = txt_(row[C.school]);
    if (!school) continue;
    schools.push({
      city:          cityShort_(row[C.city]),
      school:        school,
      pupils:        num_(row[C.pupils]),
      ls:            num_(row[C.ls]),
      cards:         num_(row[C.cards]),
      parents:       num_(row[C.parents]),
      tariffs:       num_(row[C.tariffs]),
      tariffsActive: num_(row[C.tariffsActive]),
      lsTariff:      num_(row[C.lsTariff]),
      menu:          yes_(row[C.menu]),
      complexes:     num_(row[C.complexes]),
      dishes:        num_(row[C.dishes]),
      orders:        num_(row[C.orders]),
      sbs:           yes_(row[C.sbs]),
      terminals:     num_(row[C.terminals]),
      skudInst:      yes_(row[C.skudInst]),
      skudConn:      yes_(row[C.skudConn]),
      kshpShort:     kshpShort_(row[C.kshp])
    });
  }
  if (!schools.length) throw new Error('В отчёте по области не найдено ни одной строки со школой');

  var date = toIso_(sheet.name) ? txt_(sheet.name) : ddmmyyyy_(updated);
  return { updated: ddmmyyyy_(updated), date: date, schools: schools };
}

/** «Багратионовск город» → «Багратионовск»; «Янтарный поселок городского типа» → «Янтарный (пгт)». */
function cityShort_(v) {
  var t = txt_(v);
  if (/поселок городского типа/i.test(t)) {
    return txt_(t.replace(/\s*поселок городского типа\s*/i, '')) + ' (пгт)';
  }
  return txt_(t.replace(/\s+город\s*$/i, ''));
}

/** Короткое имя оператора: то, что в кавычках, иначе первые 40 символов. */
function kshpShort_(v) {
  var t = txt_(v);
  var m = /[«"]([^»"]+)[»"]/.exec(t);
  if (m) return txt_(m[1]);
  return t.length > 40 ? t.substring(0, 40) : t;
}

// ══════════════════════════ GitHub ══════════════════════════

function gh_(method, path, payload) {
  var token = cfg_('GITHUB_TOKEN');
  if (!token) throw new Error('Не задано свойство скрипта GITHUB_TOKEN');
  var options = {
    method: method,
    headers: {
      Authorization: 'Bearer ' + token,
      Accept: 'application/vnd.github+json',
      'X-GitHub-Api-Version': '2022-11-28'
    },
    muteHttpExceptions: true
  };
  if (payload) {
    options.contentType = 'application/json';
    options.payload = JSON.stringify(payload);
  }
  var res = UrlFetchApp.fetch('https://api.github.com/repos/' + cfg_('REPO') + '/' + path, options);
  var code = res.getResponseCode();
  var body = res.getContentText();
  if (code === 404) return null;
  if (code >= 300) {
    throw new Error('GitHub ' + method + ' ' + path + ' → ' + code + ': ' + body.substring(0, 400));
  }
  return body ? JSON.parse(body) : {};
}

function ghRead_(path) {
  var j = gh_('GET', 'contents/' + encodeURI(path) + '?ref=' + cfg_('BRANCH'));
  if (!j || !j.content) return null;
  var bytes = Utilities.base64Decode(j.content.replace(/\s/g, ''));
  return { sha: j.sha, text: Utilities.newBlob(bytes).getDataAsString('UTF-8') };
}

/** Записывает файл, только если содержимое действительно изменилось. Возвращает true, если был коммит. */
function ghWrite_(path, text, message) {
  var bytes = Utilities.newBlob(text).getBytes();
  var mine = gitBlobSha_(bytes);
  var cur = gh_('GET', 'contents/' + encodeURI(path) + '?ref=' + cfg_('BRANCH'));
  if (cur && cur.sha === mine) return false;
  var payload = {
    message: message,
    content: Utilities.base64Encode(bytes),
    branch: cfg_('BRANCH')
  };
  if (cur && cur.sha) payload.sha = cur.sha;
  gh_('PUT', 'contents/' + encodeURI(path), payload);
  return true;
}

/** SHA-1 в том виде, в каком его считает git: sha1 от "blob <длина>" + нулевой байт + содержимое. */
function gitBlobSha_(bytes) {
  var head = 'blob ' + bytes.length;
  var all = [];
  for (var i = 0; i < head.length; i++) all.push(head.charCodeAt(i));
  all.push(0);
  for (var j = 0; j < bytes.length; j++) all.push(bytes[j]);
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_1, all);
  var hex = '';
  for (var k = 0; k < digest.length; k++) {
    hex += ('0' + (digest[k] & 0xFF).toString(16)).slice(-2);
  }
  return hex;
}

function fetchTemplate_() {
  var f = ghRead_(cfg_('TEMPLATE_PATH'));
  if (!f) throw new Error('В репозитории нет файла ' + cfg_('TEMPLATE_PATH'));
  return f.text;
}

function loadPublished_(path) {
  var f = ghRead_(path);
  if (!f) return null;
  try { return JSON.parse(f.text); } catch (e) { return null; }
}

// ══════════════════════════ Сборка страницы ══════════════════════════

function render_(tpl, data) {
  if (tpl.indexOf(PLACEHOLDER) < 0) {
    throw new Error('В шаблоне нет строки ' + PLACEHOLDER + ' — некуда подставлять данные');
  }
  var json = JSON.stringify(data)
    .replace(/<\//g, '<\\/')          // чтобы «</script>» внутри данных не закрыл тег
    .replace(new RegExp('\\u2028', 'g'), '\\u2028')   // разделители строк ломают JS-литерал
    .replace(new RegExp('\\u2029', 'g'), '\\u2029');
  return tpl.replace(PLACEHOLDER, function () { return 'const DATA = ' + json + ';'; });
}

/** Агрегаты среза по области — то, из чего строятся графики динамики. */
function snapshot_(o) {
  var t = { schools: 0, pupils: 0, ls: 0, cards: 0, parents: 0, terminals: 0,
            menu: 0, sbs: 0, skudInst: 0, skudConn: 0, tariffsActive: 0, orders: 0 };
  var byCity = {}, byKshp = {};

  for (var i = 0; i < o.schools.length; i++) {
    var s = o.schools[i];
    t.schools++; t.pupils += s.pupils; t.ls += s.ls; t.cards += s.cards;
    t.parents += s.parents; t.terminals += s.terminals; t.orders += s.orders;
    if (s.menu) t.menu++;
    if (s.sbs) t.sbs++;
    if (s.skudInst) t.skudInst++;
    if (s.skudConn) t.skudConn++;
    if (s.tariffsActive > 0) t.tariffsActive++;

    var g = byCity[s.city] || (byCity[s.city] = { name: s.city, schools: 0, pupils: 0, ls: 0, cards: 0, menu: 0 });
    g.schools++; g.pupils += s.pupils; g.ls += s.ls; g.cards += s.cards; if (s.menu) g.menu++;

    var k = byKshp[s.kshpShort] || (byKshp[s.kshpShort] = { name: s.kshpShort, schools: 0, pupils: 0, ls: 0, cards: 0, menu: 0 });
    k.schools++; k.pupils += s.pupils; k.ls += s.ls; k.cards += s.cards; if (s.menu) k.menu++;
  }

  return { date: o.date, totals: t, cities: values_(byCity), kshp: values_(byKshp) };
}

function values_(obj) {
  var out = [];
  for (var k in obj) if (obj.hasOwnProperty(k)) out.push(obj[k]);
  out.sort(function (a, b) { return b.schools - a.schools || (a.name < b.name ? -1 : 1); });
  return out;
}

// ══════════════════════════ Публикация ══════════════════════════

function publish_(html, data) {
  var stamp = data.updated;
  var written = [];

  if (ghWrite_('index.html', html, 'Дашборд: данные на ' + stamp)) written.push('index.html');
  if (ghWrite_('data/city.json', json_(data.city), 'Данные города на ' + stamp)) written.push('data/city.json');
  if (ghWrite_('data/oblast.json', json_(data.oblast), 'Данные области на ' + stamp)) written.push('data/oblast.json');

  var iso = toIso_(data.oblast.date);
  if (iso) {
    var path = 'data/history/' + iso + '.json';
    if (ghWrite_(path, json_(snapshot_(data.oblast)), 'Срез по области за ' + data.oblast.date)) written.push(path);
  }

  return written.length
    ? 'Опубликовано (' + stamp + '): ' + written.join(', ')
    : 'Данные не изменились — публиковать нечего (' + stamp + ')';
}

function json_(obj) { return JSON.stringify(obj, null, 1) + '\n'; }

function notify_(err) {
  var to = cfg_('NOTIFY_EMAIL');
  if (!to) return;
  try {
    MailApp.sendEmail(to, 'Дашборд «Карта школьника»: сборка не удалась',
      'Ошибка: ' + (err && err.message ? err.message : String(err)) +
      '\n\nДашборд не обновлялся, на сайте осталась предыдущая версия.' +
      '\nЖурнал выполнения: https://script.google.com/home/executions');
  } catch (e) {
    Logger.log('Не удалось отправить письмо: ' + e.message);
  }
}

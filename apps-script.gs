// Бэкенд опроса «Дата корпоратива».
// Вставьте этот код в Google Таблице: Расширения → Apps Script, затем Развернуть → Новое развертывание → Веб-приложение.

const DAYS = ['2026-12-08','2026-12-09','2026-12-10','2026-12-11','2026-12-12',
              '2026-12-15','2026-12-16','2026-12-17','2026-12-18','2026-12-19',
              '2026-12-22','2026-12-23','2026-12-24','2026-12-25','2026-12-26'];
const HEAD = ['id', 'Имя', 'Отдел', 'Комментарий', 'Не может', 'Обновлено', 'Не придёт совсем'];

function sheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function answers_() {
  const sh = sheet_('Ответы');
  if (sh.getLastRow() === 0) sh.appendRow(HEAD);
  else if (sh.getRange(1, HEAD.length).getValue() !== HEAD[HEAD.length - 1]) sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]);
  return sh;
}

function readAll_() {
  const sh = answers_();
  const rows = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, HEAD.length).getValues() : [];
  const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  return rows.filter(r => r[0]).map(r => ({
    id: String(r[0]), name: String(r[1]), dept: String(r[2]), comment: String(r[3]),
    none: String(r[6]).trim().toLowerCase() === 'да',
    busy: String(r[6]).trim().toLowerCase() === 'да' ? [] : busyList_(r[4], tz)
  }));
}

// Таблица может сама превратить одиночную дату «2026-12-09» в дату-ячейку: приводим обратно к тексту
function busyList_(v, tz) {
  if (v instanceof Date) return [Utilities.formatDate(v, tz, 'yyyy-MM-dd')];
  return String(v).split(',').map(s => s.trim()).filter(s => DAYS.indexOf(s) >= 0);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ---- Голосование за формат (лист «Формат») ----
const FORMAT_FIXED = [
  'Общий корпоратив с другими коллективами (шоу-программа, социализация)',
  'Свой корпоратив в VIP-зале заведения с шоу-программой от Show Today и т.п.'
];
const FHEAD = ['id', 'Имя', 'Выбор', 'Свой вариант', 'Обновлено'];

function formatSheet_() {
  const sh = sheet_('Формат');
  if (sh.getLastRow() === 0) sh.appendRow(FHEAD);
  return sh;
}

const nameKey_ = s => String(s || '').toLowerCase().replace(/ё/g, 'е').replace(/\s+/g, ' ').trim();
const clean_ = s => String(s || '').replace(/[\r\n]+/g, ' ').trim();
// Человек = первое слово имени: «Айнура К.», «айнура» и «Айнура Кенжебаева» — один и тот же ответ.
// (одинаковых имён в коллективе нет)
const personKey_ = s => nameKey_(s).replace(/[^\p{L}\p{N} ]/gu, ' ').trim().split(/\s+/)[0] || '';

// Номер строки для ответа: сначала ищем этого человека по имени, потом это устройство по id
function findRow_(sh, id, name) {
  const n = sh.getLastRow() - 1;
  if (n < 1) return sh.getLastRow() + 1;
  const vals = sh.getRange(2, 1, n, 2).getValues();
  const pk = personKey_(name);
  for (let i = n - 1; i >= 0; i--) if (vals[i][0] && personKey_(vals[i][1]) === pk) return i + 2;
  for (let i = n - 1; i >= 0; i--) if (String(vals[i][0]) === id) return i + 2;
  return sh.getLastRow() + 1;
}

// Убирает дубли: у каждого человека остаётся самый свежий ответ
function dedupe_(sh, timeCol) {
  const n = sh.getLastRow() - 1;
  if (n < 2) return 0;
  const vals = sh.getRange(2, 1, n, timeCol).getValues();
  const best = {}, drop = [];
  vals.forEach((r, i) => {
    if (!r[0]) return;
    const k = personKey_(r[1]);
    if (!k) return;
    const t = r[timeCol - 1] instanceof Date ? r[timeCol - 1].getTime() : 0;
    const prev = best[k];
    if (!prev) { best[k] = { i, t }; return; }
    if (t >= prev.t) { drop.push(prev.i); best[k] = { i, t }; } else drop.push(i);
  });
  drop.sort((a, b) => b - a).forEach(i => sh.deleteRow(i + 2));
  return drop.length;
}

function cleanDuplicates_() {
  return dedupe_(answers_(), 6) + dedupe_(formatSheet_(), 5);
}

function readFormat_() {
  const sh = formatSheet_();
  const rows = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, FHEAD.length).getValues() : [];
  return rows.filter(r => r[0]).map(r => ({
    id: String(r[0]), name: String(r[1]),
    choices: String(r[2]).split('\n').map(clean_).filter(Boolean),
    own: clean_(r[3])
  }));
}

// Все варианты: два фиксированных + «свои варианты» людей (без дублей по регистру)
function formatOptions_(votes) {
  const out = [], seen = {};
  const add = s => { const k = nameKey_(s); if (s && !seen[k]) { seen[k] = true; out.push(s); } };
  FORMAT_FIXED.forEach(add);
  votes.forEach(v => add(v.own));
  votes.forEach(v => v.choices.forEach(add));
  return out;
}

function saveFormat_(b) {
  const id = String(b.id || '').slice(0, 64);
  const name = clean_(b.name).slice(0, 60);
  if (!id || !name) return json_({ ok: false, error: 'Нужны имя и id' });
  const own = clean_(b.own).slice(0, 120);
  const votes = readFormat_();
  const allowed = formatOptions_(votes);
  if (own) allowed.push(own);
  const keys = allowed.map(nameKey_);
  const choices = [];
  (Array.isArray(b.choices) ? b.choices : []).forEach(c => {
    const i = keys.indexOf(nameKey_(c));
    if (i >= 0 && choices.indexOf(allowed[i]) < 0) choices.push(allowed[i]);
  });
  if (own && choices.map(nameKey_).indexOf(nameKey_(own)) < 0) choices.push(own);
  if (!choices.length) return json_({ ok: false, error: 'Выберите хотя бы один вариант' });

  // один голос на человека: ищем по имени, потом по устройству
  const sh = formatSheet_();
  const r = findRow_(sh, id, name);
  sh.getRange(r, 1, 1, 4).setNumberFormat('@');
  sh.getRange(r, 1, 1, FHEAD.length).setValues([[id, name, choices.join('\n'), own, new Date()]]);
  return null;
}

// GET: все ответы (id не отдаём наружу целиком — только его хэш, чтобы нельзя было переписать чужой ответ)
function doGet() {
  const list = readAll_().map(r => ({ ...r, id: hash_(r.id) }));
  const votes = readFormat_();
  const format = {
    options: formatOptions_(votes),
    votes: votes.map(v => ({ id: hash_(v.id), choices: v.choices, own: v.own }))  // без имён
  };
  return json_({ ok: true, responses: list, format });
}

// POST: сохранить или обновить свой ответ
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const b = JSON.parse(e.postData.contents);
    if (b.kind === 'format') {
      const err = saveFormat_(b);
      if (err) return err;
      updateSummary_();
      return json_({ ok: true });
    }
    const id = String(b.id || '').slice(0, 64);
    const name = String(b.name || '').trim().slice(0, 60);
    if (!id || !name) return json_({ ok: false, error: 'Нужны имя и id' });
    const none = b.none === true;
    const busy = none ? [] : (Array.isArray(b.busy) ? b.busy : []).filter(d => DAYS.indexOf(d) >= 0).sort();
    const row = [id, name, String(b.dept || '').trim().slice(0, 40), String(b.comment || '').trim().slice(0, 300),
                 busy.join(', '), new Date(), none ? 'да' : ''];
    const sh = answers_();
    const r = findRow_(sh, id, name);   // тот же человек с другого устройства — заменяем его ответ
    sh.getRange(r, 1, 1, 5).setNumberFormat('@'); // текст, чтобы таблица не превращала даты в даты-ячейки
    sh.getRange(r, 1, 1, HEAD.length).setValues([row]);
    updateSummary_();
    return json_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

// Лист «Итоги»: сколько могут в каждый день и кто не может
function updateSummary_() {
  cleanDuplicates_();
  const everyone = readAll_();
  const out = everyone.filter(r => r.none);      // не придут ни в какой день — в статистику дней не входят
  const all = everyone.filter(r => !r.none);
  const n = all.length;
  const wd = ['вс','пн','вт','ср','чт','пт','сб'];
  const rows = DAYS.map(k => {
    const no = all.filter(r => r.busy.indexOf(k) >= 0);
    const d = new Date(k + 'T12:00:00');
    return [Number(k.slice(-2)) + ' дек, ' + wd[d.getDay()], n - no.length, n, n ? (n - no.length) / n : 0,
            no.map(r => r.name).join(', ')];
  }).sort((a, b) => b[1] - a[1]);
  const sh = sheet_('Итоги');
  sh.clear();
  sh.getRange(1, 1, 1, 5).setValues([['День', 'Могут', 'Ответили', 'Доля', 'Не могут']]).setFontWeight('bold');
  if (rows.length) {
    sh.getRange(2, 1, rows.length, 5).setValues(rows);
    sh.getRange(2, 4, rows.length, 1).setNumberFormat('0%');
  }
  const last = rows.length + 3;
  sh.getRange(last, 1, 1, 2).setValues([['Не придут совсем', out.length]]).setFontWeight('bold');
  if (out.length) sh.getRange(last, 5).setValue(out.map(r => r.name).join(', '));

  // Формат
  const votes = readFormat_();
  const opts = formatOptions_(votes);
  const frows = opts.map(o => {
    const who = votes.filter(v => v.choices.map(nameKey_).indexOf(nameKey_(o)) >= 0);
    return [o, who.length, votes.length, votes.length ? who.length / votes.length : 0, who.map(v => v.name).join(', ')];
  }).sort((a, b) => b[1] - a[1]);
  const f0 = last + 2;
  sh.getRange(f0, 1, 1, 5).setValues([['Формат', 'Голосов', 'Проголосовали', 'Доля', 'Кто']]).setFontWeight('bold');
  if (frows.length) {
    sh.getRange(f0 + 1, 1, frows.length, 5).setValues(frows);
    sh.getRange(f0 + 1, 4, frows.length, 1).setNumberFormat('0%');
  }
  sh.autoResizeColumns(1, 5);
}

// ---- Автопересчёт «Итогов» при ручной правке таблицы ----

// Запустите ОДИН раз вручную (выберите setupTrigger вверху и нажмите «Выполнить»).
// После этого «Итоги» пересчитываются при любом изменении таблицы, включая удаление строк.
function setupTrigger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'onSheetChange')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('onSheetChange').forSpreadsheet(ss).onChange().create();
  updateSummary_();
}

// Изменения, сделанные самим скриптом, триггер не вызывают, так что зацикливания нет.
function onSheetChange(e) {
  updateSummary_();
}

// Меню «Опрос → Пересчитать итоги» на всякий случай
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Опрос').addItem('Пересчитать итоги', 'recalc').addToUi();
}

function recalc() {
  updateSummary_();
}

function hash_(s) {
  const raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s);
  return raw.slice(0, 8).map(b => ((b + 256) % 256).toString(16).padStart(2, '0')).join('');
}

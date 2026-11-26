// Бэкенд опроса «Дата корпоратива».
// Вставьте этот код в Google Таблице: Расширения → Apps Script, затем Развернуть → Новое развертывание → Веб-приложение.

const DAYS = ['2026-12-08','2026-12-09','2026-12-10','2026-12-11','2026-12-12',
              '2026-12-15','2026-12-16','2026-12-17','2026-12-18','2026-12-19',
              '2026-12-22','2026-12-23','2026-12-24','2026-12-25','2026-12-26'];
const HEAD = ['id', 'Имя', 'Отдел', 'Комментарий', 'Не может', 'Обновлено'];

function sheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function answers_() {
  const sh = sheet_('Ответы');
  if (sh.getLastRow() === 0) sh.appendRow(HEAD);
  return sh;
}

function readAll_() {
  const sh = answers_();
  const rows = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, HEAD.length).getValues() : [];
  return rows.filter(r => r[0]).map(r => ({
    id: String(r[0]), name: String(r[1]), dept: String(r[2]), comment: String(r[3]),
    busy: String(r[4]).split(',').map(s => s.trim()).filter(Boolean)
  }));
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// GET: все ответы (id не отдаём наружу целиком — только его хэш, чтобы нельзя было переписать чужой ответ)
function doGet() {
  const list = readAll_().map(r => ({ ...r, id: hash_(r.id) }));
  return json_({ ok: true, responses: list });
}

// POST: сохранить или обновить свой ответ
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const b = JSON.parse(e.postData.contents);
    const id = String(b.id || '').slice(0, 64);
    const name = String(b.name || '').trim().slice(0, 60);
    if (!id || !name) return json_({ ok: false, error: 'Нужны имя и id' });
    const busy = (Array.isArray(b.busy) ? b.busy : []).filter(d => DAYS.indexOf(d) >= 0).sort();
    const row = [id, name, String(b.dept || '').trim().slice(0, 40), String(b.comment || '').trim().slice(0, 300),
                 busy.join(', '), new Date()];
    const sh = answers_();
    const ids = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(r => String(r[0])) : [];
    const i = ids.indexOf(id);
    if (i >= 0) sh.getRange(i + 2, 1, 1, HEAD.length).setValues([row]);
    else sh.appendRow(row);
    updateSummary_();
    return json_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

// Лист «Итоги»: сколько могут в каждый день и кто не может
function updateSummary_() {
  const all = readAll_();
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
  sh.autoResizeColumns(1, 5);
}

function hash_(s) {
  const raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s);
  return raw.slice(0, 8).map(b => ((b + 256) % 256).toString(16).padStart(2, '0')).join('');
}

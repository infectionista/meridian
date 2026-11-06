/**
 * Meridian Science Map — progress backend (Google Apps Script).
 *
 * Привязан к Google-таблице (создаётся через Extensions → Apps Script).
 * Листы создаёт сам при первом обращении:
 *   Students — Name | Code          (заполняет учитель)
 *   Events   — лог всех событий     (append-only, аудит)
 *   State    — матрица достижений:  строка = ученик, столбец = ключ ачивки,
 *              в ячейке дата первого выполнения. Можно править вручную:
 *              поставить дату/любой текст = засчитать, очистить = снять.
 *
 * API (Web app, Execute as: Me, Access: Anyone):
 *   GET  ?mode=ping                          → {ok:true, v:1}
 *   GET  ?name=Ira&code=7342                 → {ok, name, keys:[...]}
 *   GET  ?name=..&code=..&action=mark&type=trainer&id=eatsun  → как POST
 *   POST text/plain JSON {name, code, type, id}
 *        → {ok, key, new, keys:[...]}
 *   Всё поддерживает &callback=fn (JSONP) на случай CORS-капризов.
 *
 * Типы событий: trainer | trail | step | open
 *   trainer:eatsun        — пройден квиз тренажёра
 *   trail:eatsun          — пройден маршрут целиком
 *   step:eatsun:U5.K1     — шаг маршрута
 *   open:U2.K1            — открыт материал узла
 */

const TYPES = ['trainer', 'trail', 'step', 'open', 'quiz'];

/* ---------- entry points ---------- */

function doGet(e) {
  const p = (e && e.parameter) || {};
  const cb = sanitizeCb_(p.callback);
  try {
    if (p.mode === 'ping') return out_({ ok: true, v: 1 }, cb);
    const auth = auth_(p.name, p.code);
    if (!auth.ok) return out_(auth, cb);
    if (p.action === 'mark') return out_(mark_(auth.name, p.type, p.id), cb);
    return out_({ ok: true, name: auth.name, keys: stateKeys_(auth.name) }, cb);
  } catch (err) {
    return out_({ ok: false, error: String(err) }, cb);
  }
}

function doPost(e) {
  let b = {};
  try { b = JSON.parse(e.postData.contents); } catch (err) {
    return out_({ ok: false, error: 'bad json' }, null);
  }
  try {
    const auth = auth_(b.name, b.code);
    if (!auth.ok) return out_(auth, null);
    return out_(mark_(auth.name, b.type, b.id), null);
  } catch (err) {
    return out_({ ok: false, error: String(err) }, null);
  }
}

/* ---------- core ---------- */

function mark_(name, type, id) {
  type = String(type || '').trim().toLowerCase();
  id = String(id || '').trim().replace(/[^\w:.,\/ \-]/g, '').slice(0, 64);
  if (TYPES.indexOf(type) < 0 || !id) return { ok: false, error: 'bad request' };
  const key = type + ':' + id;

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (type === 'quiz') {           // every attempt is logged, no achievement
      logQuiz_(name, id);
      return { ok: true, quiz: true };
    }
    logEvent_(name, type, id);
    const isNew = setState_(name, key);
    return { ok: true, key: key, new: isNew, keys: stateKeys_(name) };
  } finally {
    lock.releaseLock();
  }
}

/* Quiz attempts: raw rows in 'Quiz log', aggregates in 'Quiz stats'
   (per trainer: Attempts | Perfect | Q<n> errors columns). */
function logQuiz_(name, id) {
  const sh = sheet_('Quiz log', ['When', 'Name', 'Trainer', 'Score', 'Total', 'Wrong questions']);
  const m = String(id).match(/^([\w-]+):(\d+)\/(\d+):(.*)$/);
  if (!m) { sh.appendRow([new Date(), name, id, '', '', '']); return; }
  const slug = m[1], score = Number(m[2]), total = Number(m[3]);
  const wrong = m[4] === 'ok' ? [] :
    m[4].split(',').map(function (x) { return Number(x); })
        .filter(function (x) { return x >= 1 && x <= 40; });
  sh.appendRow([new Date(), name, slug, score, total, wrong.join(', ')]);

  const st = sheet_('Quiz stats', ['Trainer', 'Attempts', 'Perfect']);
  let row = findRow_(st, slug);
  if (row === 0) { row = st.getLastRow() + 1; st.getRange(row, 1).setValue(slug); }
  bump_(st, row, 2);
  if (score === total) bump_(st, row, 3);
  const header = st.getRange(1, 1, 1, Math.max(st.getLastColumn(), 3)).getValues()[0];
  wrong.forEach(function (q) {
    const key = 'Q' + q + ' errors';
    let col = header.indexOf(key) + 1;
    if (col === 0) { col = header.length + 1; header.push(key); st.getRange(1, col).setValue(key).setFontWeight('bold'); }
    bump_(st, row, col);
  });
}
function bump_(sh, row, col) {
  const c = sh.getRange(row, col);
  c.setValue((Number(c.getValue()) || 0) + 1);
}

function auth_(name, code) {
  name = String(name || '').trim();
  code = String(code || '').trim();
  if (!name || !code) return { ok: false, error: 'auth' };
  const reg = students_();
  const rec = reg[name.toLowerCase()];
  if (!rec || rec.code !== code) return { ok: false, error: 'auth' };
  return { ok: true, name: rec.name };
}

/* ---------- sheets ---------- */

function ss_() { return SpreadsheetApp.getActive(); }

function sheet_(title, headers) {
  const ss = ss_();
  let sh = ss.getSheetByName(title);
  if (!sh) {
    sh = ss.insertSheet(title);
    sh.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function students_() {
  const sh = sheet_('Students', ['Name', 'Code']);
  const vals = sh.getDataRange().getValues();
  const reg = {};
  for (let i = 1; i < vals.length; i++) {
    const name = String(vals[i][0] || '').trim();
    const code = String(vals[i][1] || '').trim();
    if (name && code) reg[name.toLowerCase()] = { name: name, code: code };
  }
  return reg;
}

function logEvent_(name, type, id) {
  sheet_('Events', ['When', 'Name', 'Type', 'Id'])
    .appendRow([new Date(), name, type, id]);
}

/* State: row per student (col A), column per achievement key (row 1). */
function setState_(name, key) {
  const sh = sheet_('State', ['Name']);
  const header = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0];
  let col = header.indexOf(key) + 1;
  if (col === 0) {
    col = header.length + 1;
    sh.getRange(1, col).setValue(key).setFontWeight('bold');
  }
  let row = findRow_(sh, name);
  if (row === 0) {
    row = sh.getLastRow() + 1;
    sh.getRange(row, 1).setValue(name);
  }
  const cell = sh.getRange(row, col);
  if (String(cell.getValue()) !== '') return false; // already earned
  cell.setValue(new Date());
  return true;
}

function stateKeys_(name) {
  const sh = sheet_('State', ['Name']);
  const lastCol = sh.getLastColumn();
  if (lastCol < 2) return [];
  const header = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  const row = findRow_(sh, name);
  if (row === 0) return [];
  const vals = sh.getRange(row, 1, 1, lastCol).getValues()[0];
  const keys = [];
  for (let c = 1; c < lastCol; c++) {
    if (String(vals[c]) !== '') keys.push(String(header[c]));
  }
  return keys;
}

function findRow_(sh, name) {
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const names = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < names.length; i++) {
    if (String(names[i][0]).trim().toLowerCase() === name.toLowerCase()) return i + 2;
  }
  return 0;
}

/* ---------- output ---------- */

function sanitizeCb_(cb) {
  cb = String(cb || '').replace(/[^\w.$]/g, '');
  return cb || null;
}

function out_(obj, cb) {
  const s = JSON.stringify(obj);
  if (cb) {
    return ContentService.createTextOutput(cb + '(' + s + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(s)
    .setMimeType(ContentService.MimeType.JSON);
}

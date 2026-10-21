/**
 * 未来運動会 準備スケジュール — Googleスプレッドシート連携（Google Apps Script）
 *
 * スケジュールのタスクをスプレッドシートに保存し、安部さん・鈴木さんなど複数人で同じデータを使えるようにします。
 *   「タスク」シート … 1行=1タスク（スプレッドシートを直接見たり印刷したりもできます）
 *   「設定」シート   … 開催日・表示開始日・表示する週数
 *
 * 【設置手順（最初の1回だけ）】
 *  1. Googleドライブで新しいスプレッドシートを作成（名前は「未来運動会_スケジュール」など）
 *  2. メニュー「拡張機能」→「Apps Script」を開き、このファイルの内容を全部貼り付けて保存
 *  3. 下の SHARED_KEY を、関係者だけが知る合言葉に変更して保存
 *  4. 「デプロイ」→「新しいデプロイ」→ 種類「ウェブアプリ」
 *       次のユーザーとして実行：自分
 *       アクセスできるユーザー：全員
 *     → 「デプロイ」（スプレッドシートへのアクセス許可を求められたら許可）
 *  5. 表示された「ウェブアプリのURL」と SHARED_KEY を、アプリの「設定・連携」に入力
 *     （安部さん・鈴木さんそれぞれの端末で、同じURLと合言葉を入力します）
 *
 * コードを変更したときは「デプロイを管理」→ 編集（鉛筆）→ バージョン「新バージョン」→ デプロイ。
 * （URLは変わりません）
 */

const SHARED_KEY = 'change-me'; // ← 必ず変更してください

const TASK_SHEET = 'タスク';
const SETTING_SHEET = '設定';
const HEADERS = ['ID', 'タスク', '担当', '開始日', '終了日', '時期', '状況', 'メモ', '更新日時', '更新者'];
const FIELDS = ['id', 'name', 'owner', 'start', 'end', 'certainty', 'status', 'memo', 'updatedAt', 'updatedBy'];
const SETTING_KEYS = { eventDate: '開催日', viewStart: '表示開始日', weeks: '表示する週数' };

function doGet() {
  return json_({ ok: true, message: 'undoukai schedule api' });
}

function doPost(e) {
  let b;
  try { b = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: '不正なリクエストです' }); }
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    auth_(b.key);
    switch (b.action) {
      case 'ping': return json_({ ok: true, sheetUrl: SpreadsheetApp.getActive().getUrl() });
      case 'list': return json_({ ok: true, tasks: readTasks_(), settings: readSettings_() });
      case 'upsert': upsert_(b.task, b.by); return json_({ ok: true, tasks: readTasks_() });
      case 'delete': remove_(b.id); return json_({ ok: true, tasks: readTasks_() });
      case 'replaceAll': replaceAll_(b.tasks || [], b.by); if (b.settings) writeSettings_(b.settings); return json_({ ok: true, tasks: readTasks_(), settings: readSettings_() });
      case 'saveSettings': writeSettings_(b.settings || {}); return json_({ ok: true, settings: readSettings_() });
      default: return json_({ ok: false, error: '不明な操作です' });
    }
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

function auth_(key) {
  if (SHARED_KEY === 'change-me') throw new Error('Code.gs の SHARED_KEY を変更してから、もう一度デプロイしてください');
  if (key !== SHARED_KEY) throw new Error('合言葉が違います');
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function taskSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(TASK_SHEET);
  if (!sh) {
    sh = ss.insertSheet(TASK_SHEET, 0);
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold').setBackground('#D9E2F3');
    sh.setFrozenRows(1);
    sh.getRange('A:J').setNumberFormat('@'); // 日付は「2027-04-24」の文字列のまま保存
    [80, 260, 100, 100, 100, 60, 70, 360, 160, 80].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  }
  return sh;
}

function readTasks_() {
  const sh = taskSheet_();
  const n = sh.getLastRow() - 1;
  if (n <= 0) return [];
  return sh.getRange(2, 1, n, FIELDS.length).getDisplayValues()
    .filter(function (r) { return r[0]; })
    .map(function (r) { const o = {}; FIELDS.forEach(function (f, i) { o[f] = r[i]; }); return o; });
}

function row_(t, by) {
  const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm');
  return FIELDS.map(function (f) {
    if (f === 'updatedAt') return now;
    if (f === 'updatedBy') return String(by || t.updatedBy || '');
    return String(t[f] == null ? '' : t[f]);
  });
}

function upsert_(t, by) {
  if (!t || !t.id) throw new Error('タスクのIDがありません');
  const sh = taskSheet_();
  const n = sh.getLastRow() - 1;
  const ids = n > 0 ? sh.getRange(2, 1, n, 1).getDisplayValues().map(function (r) { return r[0]; }) : [];
  const i = ids.indexOf(String(t.id));
  if (i >= 0) sh.getRange(i + 2, 1, 1, FIELDS.length).setValues([row_(t, by)]);
  else sh.appendRow(row_(t, by));
  sort_();
}

function remove_(id) {
  const sh = taskSheet_();
  const n = sh.getLastRow() - 1;
  if (n <= 0) return;
  const ids = sh.getRange(2, 1, n, 1).getDisplayValues().map(function (r) { return r[0]; });
  const i = ids.indexOf(String(id));
  if (i >= 0) sh.deleteRow(i + 2);
}

function replaceAll_(tasks, by) {
  const sh = taskSheet_();
  const n = sh.getLastRow() - 1;
  if (n > 0) sh.getRange(2, 1, n, FIELDS.length).clearContent();
  if (tasks.length) sh.getRange(2, 1, tasks.length, FIELDS.length).setValues(tasks.map(function (t) { return row_(t, by); }));
  sort_();
}

function sort_() {
  const sh = taskSheet_();
  const n = sh.getLastRow() - 1;
  if (n > 1) sh.getRange(2, 1, n, FIELDS.length).sort([{ column: 4, ascending: true }, { column: 5, ascending: true }]);
}

function settingSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SETTING_SHEET);
  if (!sh) {
    sh = ss.insertSheet(SETTING_SHEET);
    sh.getRange('A:C').setNumberFormat('@');
    sh.getRange(1, 1, 1, 3).setValues([['項目', '値', 'キー']]).setFontWeight('bold').setBackground('#D9E2F3');
  }
  return sh;
}

function readSettings_() {
  const sh = settingSheet_();
  const n = sh.getLastRow() - 1;
  const o = {};
  if (n > 0) sh.getRange(2, 1, n, 3).getDisplayValues().forEach(function (r) { if (r[2]) o[r[2]] = r[1]; });
  if (o.weeks) o.weeks = Number(o.weeks);
  return o;
}

function writeSettings_(s) {
  const sh = settingSheet_();
  const cur = readSettings_();
  Object.keys(SETTING_KEYS).forEach(function (k) { if (s[k] != null && s[k] !== '') cur[k] = s[k]; });
  const rows = Object.keys(SETTING_KEYS).filter(function (k) { return cur[k] != null; })
    .map(function (k) { return [SETTING_KEYS[k], String(cur[k]), k]; });
  const n = sh.getLastRow() - 1;
  if (n > 0) sh.getRange(2, 1, n, 3).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, 3).setValues(rows);
}

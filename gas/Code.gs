/**
 * 調性判断テスト：データ受信用 Google Apps Script
 *
 * 受け取ったデータを，指定の Google ドライブフォルダ（FOLDER_ID）に保存します
 *   https://drive.google.com/drive/folders/1JfMDViUAi2KsJQW34p40-pra87X6ns5U
 *   1. スプレッドシート「調性判断テスト_集計」… trials シート（1試行1行）と summary シート（1セッション1行）
 *   2. セッションごとのCSVファイル（test_参加者ID_日時.csv）
 * 同じセッションを2回受け取った場合は，2回目を無視します（再送しても重複しません）。
 *
 * 【設定のしかた】
 *   1. https://script.google.com で「新しいプロジェクト」を作り，このコードを全部貼り付けて保存
 *   2. 右上「デプロイ」→「新しいデプロイ」→ 種類の選択（歯車）で「ウェブアプリ」
 *        次のユーザーとして実行：自分
 *        アクセスできるユーザー：全員   ← ここが「全員」でないと届きません
 *   3. エディタ上部の関数の選択で「setup」を選んで「実行」→ アクセスを承認
 *      （「詳細」→「（プロジェクト名）に移動（安全ではないページ）」→ すべて「許可」）
 *      ※ コードを貼り替えたあとは必ずもう一度 setup を実行してください（新しく必要になった権限を承認するため）
 *   4. 表示された「ウェブアプリのURL」（末尾が /exec）をテストの設定画面に貼り付け，「接続テスト」
 *   ※ コードを書き換えたら「デプロイ」→「デプロイを管理」→ 鉛筆 → バージョン「新しいバージョン」→「デプロイ」
 *   ※ 大学など組織のアカウントで「全員」が選べない場合は，個人の Gmail アカウントで作ってください
 *   ※ スクリプトを作ったアカウントが，保存先フォルダの「編集者」になっている必要があります
 *      （フォルダの持ち主と別のアカウントで作る場合は，フォルダの「共有」でそのアカウントを編集者に追加）
 */

// 保存先フォルダのID（フォルダURLの /folders/ の後ろの部分）
const FOLDER_ID = '1JfMDViUAi2KsJQW34p40-pra87X6ns5U';
const SHEET_NAME = '調性判断テスト_集計';
// 文字列のまま残す列（日付や数値に自動変換されないように）
const TEXT_COLS = ['session_id', 'participant', 'group', 'start_time', 'end_time', 'answered_at',
  'base_key', 'comp_key', 'original_chord', 'changed_chord', 'base_seq', 'comp_seq',
  'base_voicing', 'comp_voicing', 'stratify', 'difficulty', 'condition', 'response', 'sound'];

/** 権限の承認と動作確認用：エディタで一度「実行」してください */
function setup() {
  const folder = folder_();
  const ss = spreadsheet_(folder);
  Logger.log('準備OK\n保存先フォルダ：' + folder.getName() + '\n' + folder.getUrl() + '\n集計シート：' + ss.getUrl());
}

function doGet() {
  return json_({ ok: true, message: '調性判断テストの受信スクリプトは動作中です' });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const p = JSON.parse(e.postData.contents);
    const folder = folder_();
    if (p.action === 'ping') {
      let sheetOk = '';
      try { spreadsheet_(folder); } catch (err) { sheetOk = String(err && err.message || err); }
      return json_({ ok: true, ping: true, folder: folder.getName() + '（' + folder.getUrl() + '）',
        warning: sheetOk ? '集計シートを作れません（CSVは保存されます）：' + sheetOk + ' → エディタで setup を実行して権限を承認し，新しいバージョンでデプロイしてください' : '' });
    }
    if (!p.session_id || !p.trial_columns || !p.summary_columns) throw new Error('データの形式が違います');

    // 1. CSVファイル（ドライブの権限だけで保存できるので先に保存。同じ名前があれば作らない）
    const name = p.csv_name || ('test_' + p.session_id + '.csv');
    let csvExists = false;
    const it = folder.getFilesByName(name);
    while (it.hasNext()) if (!it.next().isTrashed()) csvExists = true;
    if (!csvExists) folder.createFile(name, '\ufeff' + csv_(p), MimeType.CSV);

    // 2. 集計シート（失敗してもCSVは残る）
    let warning = '';
    try {
      const ss = spreadsheet_(folder);
      const sum = sheet_(ss, 'summary', p.summary_columns);
      const tr = sheet_(ss, 'trials', p.trial_columns);
      let dup = false;
      if (sum.getLastRow() > 1) dup = sum.getRange(2, 1, sum.getLastRow() - 1, 1).getDisplayValues().some(r => r[0] === String(p.session_id));
      if (!dup) { write_(tr, p.trial_columns, p.trials || []); write_(sum, p.summary_columns, [p.summary]); }
    } catch (err) {
      warning = 'CSVは保存しましたが，集計シートに書き込めませんでした：' + String(err && err.message || err) + ' → エディタで setup を実行して権限を承認し，新しいバージョンでデプロイしてください';
    }
    return json_({ ok: true, saved: (p.trials || []).length, warning: warning });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

function folder_() {
  let f;
  try { f = DriveApp.getFolderById(FOLDER_ID); f.getName(); }
  catch (e) { throw new Error('保存先フォルダを開けません。スクリプトを作ったアカウント（' + Session.getEffectiveUser().getEmail() + '）にフォルダの編集権限があるか確認してください'); }
  if (f.isTrashed()) throw new Error('保存先フォルダがゴミ箱に入っています');
  let canEdit = true;
  try {
    const a = f.getAccess(Session.getEffectiveUser());
    canEdit = a === DriveApp.Permission.OWNER || a === DriveApp.Permission.EDIT || a === DriveApp.Permission.ORGANIZER || a === DriveApp.Permission.FILE_ORGANIZER;
  } catch (e) {}   // 共有ドライブなどで判定できない場合は書き込みを試す
  if (!canEdit) throw new Error('保存先フォルダへの編集権限がありません。フォルダの「共有」で ' + Session.getEffectiveUser().getEmail() + ' を「編集者」に追加してください');
  return f;
}

function spreadsheet_(folder) {
  const props = PropertiesService.getScriptProperties(), key = 'SHEET_ID_' + FOLDER_ID;
  const id = props.getProperty(key);
  if (id) { try { const file = DriveApp.getFileById(id); if (!file.isTrashed()) return SpreadsheetApp.openById(id); } catch (e) {} }
  // フォルダ内に同じ名前の集計シートがあればそれを使う
  const it = folder.getFilesByName(SHEET_NAME);
  while (it.hasNext()) {
    const file = it.next();
    if (!file.isTrashed() && file.getMimeType() === MimeType.GOOGLE_SHEETS) { props.setProperty(key, file.getId()); return SpreadsheetApp.openById(file.getId()); }
  }
  const ss = SpreadsheetApp.create(SHEET_NAME);
  DriveApp.getFileById(ss.getId()).moveTo(folder);
  props.setProperty(key, ss.getId());
  return ss;
}

function sheet_(ss, name, cols) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    const first = ss.getSheets()[0];
    if (first.getName() !== name && first.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(first);
  }
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function write_(sh, cols, rows) {
  if (!rows.length) return;
  const start = sh.getLastRow() + 1;
  cols.forEach((c, i) => { if (TEXT_COLS.indexOf(c) >= 0) sh.getRange(start, i + 1, rows.length, 1).setNumberFormat('@'); });
  sh.getRange(start, 1, rows.length, cols.length).setValues(rows.map(r => cols.map((_, i) => r[i] == null ? '' : r[i])));
}

function csv_(p) {
  const esc = v => { const s = String(v == null ? '' : v); return /[",\n\r]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s; };
  const lines = [p.trial_columns.join(',')];
  (p.trials || []).forEach(r => lines.push(r.map(esc).join(',')));
  lines.push('', p.summary_columns.join(','), p.summary.map(esc).join(','));
  return lines.join('\r\n');
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/*************************************************
 * Mitsume Coffee シフト希望提出システム（受け側）
 * Google Apps Script
 *
 * 仕組み:
 *  - LIFF画面(doGet)がこのスクリプトから配信される
 *  - 画面は LINE の IDトークンを送ってくるので、LINEの検証APIで
 *    本人確認してから、その人の分だけ読み書きする
 *  - 提出は「対象月」単位。締切までは何度でも上書き更新できる
 *  - データはこのスプレッドシートに貯まる
 *************************************************/

/***** 設定（ここだけ編集すればOK） *****/
const CONFIG = {
  // ↓ LINE Developersで取得して記入する（手順書参照）
  LIFF_ID: '2012734423-wNqztLhN',
  LINE_CHANNEL_ID: '2012734423', // LINEログインチャネルの「チャネルID」

  // 顧客ID・店舗ID（サービス提供に備えて全テーブルに持たせる。2026-11-24 ユーザー指示）
  CUSTOMER_ID: 'M10001', // 合同会社Giorno
  STORE_ID: 'S001',      // 三ツ目珈琲

  // 締切：対象月の「前月」の DEADLINE_DAY 日 23:59 まで（例: 12月分は11月25日締切）
  DEADLINE_DAY: 25,

  // 個人別締切例外： { 'LINEのuserId': 'yyyy-MM-dd HH:mm' } その人だけ、その日時まで提出できる
  // userId は「アクセス記録」シートで確認する（フォームを開いた人が自動で記録される）
  DEADLINE_EXCEPTIONS: {},

  // 定休日の曜日（0=日〜6=土）。水曜=3
  CLOSED_WEEKDAYS: [3],
  // 年末休み：12月は29日以降を休業にする
  YEAR_END_CLOSED: { month: 12, fromDay: 29 },

  // プリセットボタン（画面と集計の表示に使う）
  PRESETS: {
    full:  { label: '全部OK', start: '08:30', end: '17:00' },
    early: { label: '早番OK', start: '08:30', end: '14:30' },
    late:  { label: '遅番OK', start: '11:00', end: '17:00' }
  }
};

const TZ = 'Asia/Tokyo';
const SHEET_DATA = '提出データ';   // 1日 = 1行
const SHEET_STATUS = '提出状況';   // 1人×1月 = 1行（誰が出したか・メモ・更新日時）
const SHEET_ACCESS = 'アクセス記録'; // フォームを開いた人の userId を自動記録（許可IDリスト・例外設定用）
const SHEET_ROSTER = '確定シフト';   // 確定した割当表（1行目:対象月/ラベル、2行目:見出し、3行目〜:日別）
const SHEET_STAFF = 'スタッフ';      // Line_ID → 苗字（確定シフトの列名）の対応表

/***** 画面の配信 *****/
function doGet(e) {
  // カレンダー追加用ページ（外部ブラウザで開いて .ics をダウンロードさせる）
  if (e && e.parameter && e.parameter.ics === '1') return serveIcs_(e.parameter);
  // GitHub Pages 上の画面からは JSONP API として呼ばれる
  if (e && e.parameter && e.parameter.api === '1') return handleApi_(e.parameter);
  const t = HtmlService.createTemplateFromFile('index');
  t.liffId = CONFIG.LIFF_ID;
  t.presetsJson = JSON.stringify(CONFIG.PRESETS);
  return t.evaluate()
    .setTitle('シフト希望提出 | Mitsume Coffee')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function handleApi_(p) {
  let out;
  try {
    if (p.action === 'getInitialData') {
      out = { ok: true, data: getInitialData(p.idToken || '', p.name || '') };
    } else if (p.action === 'saveSubmission') {
      out = { ok: true, data: saveSubmission(p.idToken || '', p.name || '', JSON.parse(p.payload || '{}')) };
    } else if (p.action === 'getRoster') {
      out = { ok: true, data: getRoster2(p.idToken || '', p.name || '', p.period || '') };
    } else if (p.action === 'getIcs') {
      out = { ok: true, data: getIcs2(p.idToken || '', p.period || '') };
    } else {
      out = { ok: false, error: 'unknown action' };
    }
  } catch (err) {
    out = { ok: false, error: String(err && err.message ? err.message : err) };
  }
  const body = JSON.stringify(out);
  if (p.callback) {
    return ContentService.createTextOutput(p.callback + '(' + body + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(body).setMimeType(ContentService.MimeType.JSON);
}

/***** 対象月と締切 *****/
// 対象月 = 翌月。例: 11月に開いたら「12月分」、締切は11月25日23:59
function targetMonthInfo() {
  const now = new Date();
  const y = now.getFullYear();
  const m = now.getMonth() + 1; // 今月(1-12)
  // 翌月
  const tDate = new Date(y, m, 1); // m は0始まり換算で翌月1日
  const ty = tDate.getFullYear();
  const tm = tDate.getMonth() + 1;
  const month = Utilities.formatDate(tDate, TZ, 'yyyy-MM');
  // 締切 = 対象月の前月の DEADLINE_DAY 日 23:59:59
  const prevLast = new Date(ty, tm - 1, 0).getDate(); // 前月の末日
  const dDay = Math.min(CONFIG.DEADLINE_DAY, prevLast);
  const deadline = new Date(ty, tm - 2, dDay, 23, 59, 59);
  return {
    month: month,
    monthLabel: tm + '月分',
    deadline: deadline,
    deadlineText: Utilities.formatDate(deadline, TZ, 'M月d日(E) HH:mm'),
    locked: now.getTime() > deadline.getTime()
  };
}

/***** 本人確認（LINE IDトークン検証） *****/
function verifyIdToken(idToken) {
  if (!idToken) throw new Error('ログイン情報が取得できませんでした。LINEアプリから開き直してください。');
  const res = UrlFetchApp.fetch('https://api.line.me/oauth2/v2.1/verify', {
    method: 'post',
    payload: { id_token: idToken, client_id: CONFIG.LINE_CHANNEL_ID },
    muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) {
    throw new Error('本人確認に失敗しました。もう一度開き直してください。');
  }
  const j = JSON.parse(res.getContentText());
  return { userId: j.sub, name: j.name || '' };
}

/***** 対象月セルの正規化（シートが日付に自動変換した場合も 'yyyy-MM' に戻す） *****/
function monthKey_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM');
  return String(v || '');
}

/***** シートの準備 *****/
function ensureSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let data = ss.getSheetByName(SHEET_DATA);
  if (!data) {
    data = ss.insertSheet(SHEET_DATA);
    data.appendRow(['ユーザーID', '表示名', '対象月', '日付', '曜日', '区分', '開始', '終了', '更新日時', '顧客ID', '店舗ID']);
    data.setFrozenRows(1);
  } else if (!data.getRange(1, 10).getValue()) {
    data.getRange(1, 10).setValue('顧客ID');
    data.getRange(1, 11).setValue('店舗ID');
  }
  let status = ss.getSheetByName(SHEET_STATUS);
  if (!status) {
    status = ss.insertSheet(SHEET_STATUS);
    status.appendRow(['ユーザーID', '表示名', '対象月', 'メモ', '更新日時', '顧客ID', '店舗ID']);
    status.setFrozenRows(1);
  } else if (!status.getRange(1, 6).getValue()) {
    status.getRange(1, 6).setValue('顧客ID');
    status.getRange(1, 7).setValue('店舗ID');
  }
  let access = ss.getSheetByName(SHEET_ACCESS);
  if (!access) {
    access = ss.insertSheet(SHEET_ACCESS);
    access.appendRow(['ユーザーID', '表示名', '初回アクセス', '最終アクセス', '最終アクション', '顧客ID', '店舗ID']);
    access.setFrozenRows(1);
  } else if (!access.getRange(1, 6).getValue()) {
    access.getRange(1, 6).setValue('顧客ID');
    access.getRange(1, 7).setValue('店舗ID');
  }
  return { data: data, status: status, access: access };
}

/***** 個人別の実効締切（例外リスト優先） *****/
function effectiveDeadline_(info, userId) {
  const ex = (CONFIG.DEADLINE_EXCEPTIONS || {})[userId];
  if (ex) {
    const d = new Date(String(ex).replace(' ', 'T') + '+09:00');
    if (!isNaN(d.getTime())) return d;
  }
  return info.deadline;
}

/***** アクセス記録（開いた人の userId を残す：許可IDリスト・例外設定用） *****/
function logAccess_(user, displayName, action) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_ACCESS);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_ACCESS);
      sheet.appendRow(['ユーザーID', '表示名', '初回アクセス', '最終アクセス', '最終アクション', '顧客ID', '店舗ID']);
      sheet.setFrozenRows(1);
    }
    const nowText = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === user.userId) {
        sheet.getRange(i + 1, 2).setValue(displayName || user.name || rows[i][1] || '');
        sheet.getRange(i + 1, 4).setValue(nowText);
        sheet.getRange(i + 1, 5).setValue(action);
        sheet.getRange(i + 1, 6).setValue(CONFIG.CUSTOMER_ID);
        sheet.getRange(i + 1, 7).setValue(CONFIG.STORE_ID);
        return;
      }
    }
    sheet.appendRow([user.userId, displayName || user.name || '', nowText, nowText, action, CONFIG.CUSTOMER_ID, CONFIG.STORE_ID]);
  } catch (e) { /* 記録の失敗で本体を止めない */ }
}

/***** 区分ラベル *****/
function typeLabel(type, start, end) {
  if (type === 'off') return '休み';
  const p = CONFIG.PRESETS;
  for (const k in p) {
    if (start === p[k].start && end === p[k].end) return p[k].label;
  }
  return 'カスタム';
}

/***** 画面の初期データ（本人の分だけ返す） *****/
function getInitialData(idToken, displayName) {
  const user = verifyIdToken(idToken);
  const info = targetMonthInfo();
  logAccess_(user, displayName, '表示');
  requireAccess_(user);
  const myDeadline = effectiveDeadline_(info, user.userId);
  const myLocked = new Date().getTime() > myDeadline.getTime();
  ensureSheets();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 本人の提出済みデータ
  const saved = {};
  const rows = ss.getSheetByName(SHEET_DATA).getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === user.userId && monthKey_(rows[i][2]) === info.month) {
      saved[rows[i][3]] = { type: rows[i][5], start: rows[i][6], end: rows[i][7] };
    }
  }
  // メモと更新日時
  let memo = '', updatedAt = '';
  const srows = ss.getSheetByName(SHEET_STATUS).getDataRange().getValues();
  for (let i = 1; i < srows.length; i++) {
    if (srows[i][0] === user.userId && monthKey_(srows[i][2]) === info.month) {
      memo = srows[i][3] || '';
      updatedAt = srows[i][4] ? Utilities.formatDate(new Date(srows[i][4]), TZ, 'M/d HH:mm') : '';
    }
  }

  // 対象月の日リスト
  const parts = info.month.split('-');
  const y = parseInt(parts[0], 10), m = parseInt(parts[1], 10);
  const lastDay = new Date(y, m, 0).getDate();
  const dowNames = ['日', '月', '火', '水', '木', '金', '土'];
  const days = [];
  for (let d = 1; d <= lastDay; d++) {
    const date = new Date(y, m - 1, d);
    const key = Utilities.formatDate(date, TZ, 'yyyy-MM-dd');
    const dow = date.getDay();
    let savedDay = null;
    if (saved[key]) {
      const s = saved[key];
      // 保存済みの区分ラベルから画面用タイプへ戻す
      let t = 'custom';
      if (s.type === '休み') t = 'off';
      else {
        for (const k in CONFIG.PRESETS) {
          if (CONFIG.PRESETS[k].label === s.type) t = k;
        }
      }
      savedDay = { type: t, start: s.start, end: s.end };
    }
    // 休業日（定休日・年末休み）は入力できないよう closed を付ける
    let closed = null;
    const ye = CONFIG.YEAR_END_CLOSED;
    if (ye && m === ye.month && d >= ye.fromDay) closed = '年末休み';
    else if (CONFIG.CLOSED_WEEKDAYS.indexOf(dow) >= 0) closed = '定休日';
    days.push({
      date: key, day: d, dow: dowNames[dow],
      isSun: dow === 0, isSat: dow === 6,
      saved: savedDay, closed: closed
    });
  }

  return {
    ok: true,
    name: displayName || user.name || '',
    month: info.month,
    monthLabel: info.monthLabel,
    deadlineText: Utilities.formatDate(myDeadline, TZ, 'M月d日(E) HH:mm'),
    locked: myLocked,
    days: days,
    memo: memo,
    updatedAt: updatedAt
  };
}

/***** 提出の保存（締切までは上書き更新） *****/
function saveSubmission(idToken, displayName, payload) {
  const user = verifyIdToken(idToken);
  requireAccess_(user);
  const info = targetMonthInfo();
  if (payload.month !== info.month) throw new Error('対象月が変わっています。開き直してください。');
  const myDeadline = effectiveDeadline_(info, user.userId);
  if (new Date().getTime() > myDeadline.getTime()) {
    logAccess_(user, displayName, '締切エラー');
    throw new Error('締切を過ぎているため、提出・変更はできません。');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sheets = ensureSheets();
    const dataSheet = sheets.data;
    const nowText = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');

    // 既存の「本人×対象月」の行を削除（上書き更新のため）
    const rows = dataSheet.getDataRange().getValues();
    const keep = [rows[0]];
    for (let i = 1; i < rows.length; i++) {
      if (!(rows[i][0] === user.userId && monthKey_(rows[i][2]) === info.month)) keep.push(rows[i]);
    }
    dataSheet.clearContents();
    dataSheet.getRange(1, 1, keep.length, keep[0].length).setValues(keep);

    // 新しい行を追加
    const dowNames = ['日', '月', '火', '水', '木', '金', '土'];
    const newRows = [];
    const dates = Object.keys(payload.days).sort();
    dates.forEach(function (key) {
      const d = payload.days[key];
      if (!d || !d.type) return;
      const dt = new Date(key + 'T00:00:00+09:00');
      const label = d.type === 'off' ? '休み' : typeLabel(d.type, d.start, d.end);
      newRows.push([
        user.userId, displayName || user.name || '', info.month, key,
        dowNames[dt.getDay()], label,
        d.type === 'off' ? '' : d.start, d.type === 'off' ? '' : d.end,
        nowText, CONFIG.CUSTOMER_ID, CONFIG.STORE_ID
      ]);
    });
    if (newRows.length > 0) {
      const startRow = dataSheet.getLastRow() + 1;
      const range = dataSheet.getRange(startRow, 1, newRows.length, newRows[0].length);
      range.setNumberFormat('@'); // 時刻を文字列のまま保持
      range.setValues(newRows);
    }

    // 提出状況（メモ・更新日時）を upsert
    const statusSheet = sheets.status;
    const srows = statusSheet.getDataRange().getValues();
    let found = -1;
    for (let i = 1; i < srows.length; i++) {
      if (srows[i][0] === user.userId && monthKey_(srows[i][2]) === info.month) { found = i + 1; break; }
    }
    const sRow = [user.userId, displayName || user.name || '', info.month, payload.memo || '', nowText, CONFIG.CUSTOMER_ID, CONFIG.STORE_ID];
    const targetRow = found > 0 ? found : statusSheet.getLastRow() + 1;
    const sRange = statusSheet.getRange(targetRow, 1, 1, sRow.length);
    sRange.setNumberFormat('@'); // 「2026-12」が日付に自動変換されるのを防ぐ
    sRange.setValues([sRow]);

    logAccess_(user, displayName, '保存');
    return { ok: true, updatedAt: Utilities.formatDate(new Date(), TZ, 'M/d HH:mm') };
  } finally {
    lock.releaseLock();
  }
}

/***** 集計（日付×スタッフのマトリクスを作る） *****/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('シフト集計')
    .addItem('翌月分を集計', 'buildSummaryForTarget')
    .addItem('月を指定して集計', 'buildSummaryPrompt')
    .addToUi();
}

function buildSummaryForTarget() {
  buildSummary(targetMonthInfo().month);
}

function buildSummaryPrompt() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('集計する月', '例: 2026-12 のように入力してください', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const month = res.getResponseText().trim();
  if (!/^\d{4}-\d{2}$/.test(month)) { ui.alert('形式が違います。例: 2026-12'); return; }
  buildSummary(month);
}

function buildSummary(month) {
  ensureSheets();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataRows = ss.getSheetByName(SHEET_DATA).getDataRange().getValues();
  const statusRows = ss.getSheetByName(SHEET_STATUS).getDataRange().getValues();

  // 対象月に提出したスタッフ一覧
  const staff = [];
  const staffName = {};
  for (let i = 1; i < statusRows.length; i++) {
    if (monthKey_(statusRows[i][2]) === month) {
      staff.push(statusRows[i][0]);
      staffName[statusRows[i][0]] = statusRows[i][1] || '名前なし';
    }
  }
  // セル値: userId|date -> 表示文字列
  const cell = {};
  for (let i = 1; i < dataRows.length; i++) {
    if (monthKey_(dataRows[i][2]) === month) {
      const label = dataRows[i][5];
      cell[dataRows[i][0] + '|' + dataRows[i][3]] =
        label === '休み' ? '休' : (dataRows[i][6] + '-' + dataRows[i][7]);
    }
  }

  const parts = month.split('-');
  const y = parseInt(parts[0], 10), m = parseInt(parts[1], 10);
  const lastDay = new Date(y, m, 0).getDate();
  const dowNames = ['日', '月', '火', '水', '木', '金', '土'];
  const out = [['日付', '曜日'].concat(staff.map(function (u) { return staffName[u]; }))];
  for (let d = 1; d <= lastDay; d++) {
    const date = new Date(y, m - 1, d);
    const key = Utilities.formatDate(date, TZ, 'yyyy-MM-dd');
    const row = [key, dowNames[date.getDay()]];
    staff.forEach(function (u) { row.push(cell[u + '|' + key] || ''); });
    out.push(row);
  }

  const name = '集計_' + month;
  let sheet = ss.getSheetByName(name);
  if (sheet) sheet.clearContents(); else sheet = ss.insertSheet(name);
  sheet.getRange(1, 1, out.length, out[0].length).setValues(out);
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2);
  ss.setActiveSheet(sheet);
}

/***** 確定シフト照会（本人の列を特定して全員分＋本人分を返す） *****/
function getRoster(idToken, displayName) {
  const user = verifyIdToken(idToken);
  logAccess_(user, displayName, '照会');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // Line_ID → 苗字
  let me = '';
  const staffSheet = ss.getSheetByName(SHEET_STAFF);
  if (staffSheet) {
    const srows = staffSheet.getDataRange().getValues();
    for (let i = 1; i < srows.length; i++) {
      if (srows[i][0] === user.userId) { me = String(srows[i][1] || ''); break; }
    }
  }
  const sheet = ss.getSheetByName(SHEET_ROSTER);
  if (!sheet) throw new Error('確定シフトがまだ登録されていません。');
  const values = sheet.getDataRange().getValues();
  const month = monthKey_(values[0][0]);
  const monthLabel = String(values[0][1] || '');
  const headers = values[1].map(function (v) { return String(v || ''); });
  const columns = headers.slice(2).filter(function (h) { return h !== ''; });
  const rows = [];
  for (let i = 2; i < values.length; i++) {
    const r = values[i];
    if (r[0] === '' || r[0] === null) continue;
    const cells = {};
    for (let c = 0; c < columns.length; c++) {
      const v = r[2 + c];
      cells[columns[c]] = (v === '' || v === null) ? '' : String(v);
    }
    rows.push({ day: Number(r[0]), wd: String(r[1] || ''), cells: cells });
  }
  return { month: month, monthLabel: monthLabel, columns: columns, rows: rows, me: me };
}

/***** カレンダー追加用ページ（外部ブラウザで開かれる前提。LINE内ブラウザではダウンロードが動かないため） *****/
function serveIcs_(p) {
  try {
    const user = verifyIdToken(p.idToken || '');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const me = mySurname_(ss, user.userId);
    if (!me) throw new Error('スタッフ登録が見つかりませんでした。');
    const sheet = ss.getSheetByName(SHEET_ROSTER);
    if (!sheet) throw new Error('確定シフトがまだ登録されていません。');
    const values = sheet.getDataRange().getValues();
    const month = monthKey_(values[0][0]);
    const label = String(values[0][1] || '');
    const headers = values[1].map(function (v) { return String(v || ''); });
    const col = headers.indexOf(me);
    if (col < 2) throw new Error('シフト表にあなたの列がありません。');
    const ym = month.split('-');
    const ymd = ym[0] + ('0' + ym[1]).slice(-2);
    const lines = ['BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//Mitsume Coffee//Shift//JA'];
    for (let i = 2; i < values.length; i++) {
      const r = values[i];
      const v = r[col];
      if (!v) continue;
      const parts = String(v).split('-');
      if (parts.length !== 2) continue;
      const dstr = ymd + ('0' + Number(r[0])).slice(-2);
      lines.push('BEGIN:VEVENT');
      lines.push('UID:shift-' + month + '-' + r[0] + '-' + me + '@mitsume-coffee');
      lines.push('DTSTART;TZID=Asia/Tokyo:' + dstr + 'T' + parts[0].replace(':', '') + '00');
      lines.push('DTEND;TZID=Asia/Tokyo:' + dstr + 'T' + parts[1].replace(':', '') + '00');
      lines.push('SUMMARY:三ツ目珈琲 シフト');
      lines.push('END:VEVENT');
    }
    lines.push('END:VCALENDAR');
    const ics = lines.join('\r\n');
    const html = '<!DOCTYPE html><html lang="ja"><head><meta charset="utf-8">' +
      '<meta name="viewport" content="width=device-width, initial-scale=1"><title>カレンダーに追加</title></head>' +
      '<body style="font-family:-apple-system,sans-serif;padding:24px;text-align:center">' +
      '<h1 style="font-size:18px;color:#0A5A0B">カレンダーに追加</h1>' +
      '<p>' + me + ' さんのシフト（' + label + '）</p>' +
      '<button onclick="dl()" style="font-size:16px;padding:14px 28px;border:none;border-radius:999px;background:#0A5A0B;color:#fff;font-weight:700">ダウンロード</button>' +
      '<p style="font-size:12px;color:#888;line-height:1.7">ダウンロードが始まらない時は上のボタンを押してください。<br>ダウンロードしたファイル（shift.ics）を開くと、カレンダーにまとめて追加できます。</p>' +
      '<script>var ICS=' + JSON.stringify(ics) + ';' +
      'function dl(){var b=new Blob([ICS],{type:"text/calendar"});var a=document.createElement("a");a.href=URL.createObjectURL(b);a.download="shift.ics";document.body.appendChild(a);a.click();a.remove();}' +
      'window.onload=function(){setTimeout(dl,500);};</script>' +
      '</body></html>';
    return HtmlService.createHtmlOutput(html).setTitle('カレンダーに追加');
  } catch (err) {
    return HtmlService.createHtmlOutput('<p style="font-family:sans-serif;padding:24px">エラー：' +
      String(err && err.message ? err.message : err) + '</p>');
  }
}

/***** カレンダー追加用ICSデータ（API版。ページ側でダウンロードさせる） *****/
function getIcs(idToken) {
  const user = verifyIdToken(idToken);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let me = '';
  const staffSheet = ss.getSheetByName(SHEET_STAFF);
  if (staffSheet) {
    const srows = staffSheet.getDataRange().getValues();
    for (let i = 1; i < srows.length; i++) {
      if (srows[i][0] === user.userId) { me = String(srows[i][1] || ''); break; }
    }
  }
  if (!me) throw new Error('スタッフ登録が見つかりませんでした。');
  const sheet = ss.getSheetByName(SHEET_ROSTER);
  if (!sheet) throw new Error('確定シフトがまだ登録されていません。');
  const values = sheet.getDataRange().getValues();
  const month = monthKey_(values[0][0]);
  const label = String(values[0][1] || '');
  const headers = values[1].map(function (v) { return String(v || ''); });
  const col = headers.indexOf(me);
  if (col < 2) throw new Error('シフト表にあなたの列がありません。');
  const ym = month.split('-');
  const ymd = ym[0] + ('0' + ym[1]).slice(-2);
  const lines = ['BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//Mitsume Coffee//Shift//JA'];
  for (let i = 2; i < values.length; i++) {
    const r = values[i];
    const v = r[col];
    if (!v) continue;
    const parts = String(v).split('-');
    if (parts.length !== 2) continue;
    const dstr = ymd + ('0' + Number(r[0])).slice(-2);
    lines.push('BEGIN:VEVENT');
    lines.push('UID:shift-' + month + '-' + r[0] + '-' + me + '@mitsume-coffee');
    lines.push('DTSTART;TZID=Asia/Tokyo:' + dstr + 'T' + parts[0].replace(':', '') + '00');
    lines.push('DTEND;TZID=Asia/Tokyo:' + dstr + 'T' + parts[1].replace(':', '') + '00');
    lines.push('SUMMARY:三ツ目珈琲 シフト');
    lines.push('END:VEVENT');
  }
  lines.push('END:VCALENDAR');
  return { me: me, monthLabel: label, ics: lines.join('\r\n') };
}

/***** 期間対応（v11）: 「確定シフト_YYYYMMDD」シートを1期間として扱う *****
   シート形式: 1行目 A=開始日(yyyy-MM-dd) B=終了日 C=ラベル / 2行目 見出し（日・曜・苗字…） / 3行目〜 A=日付(yyyy-MM-dd) B=曜日 */
const SHEET_ROSTER_PREFIX = '確定シフト_';
// 確定シフトのシートは「割当・確定」側のブックにある（提出ブックとは分ける運用。2026-11-26）
const ROSTER_SPREADSHEET_ID = '1gvwZbHGkYHpuJ8czj7NR2mW7kTySd7M3M3xaB04X-bQ';
function rosterBook_() { return SpreadsheetApp.openById(ROSTER_SPREADSHEET_ID); }
// 共通マスター（従業員・期間の正。2026-11-26 v14で参照開始）
const MASTER_SPREADSHEET_ID = '1UowDnTtWjvcWEy0VBDpEQUMeP0SyZ5yDekpztmyWjZQ';
function masterBook_() { return SpreadsheetApp.openById(MASTER_SPREADSHEET_ID); }

/***** アクセス権（共通マスターの従業員マスターに Line_ID 登録がある人だけ使える） *****/
function isRegisteredStaff_(userId) {
  const msh = masterBook_().getSheetByName('従業員マスター');
  if (!msh) return false;
  const rows = msh.getDataRange().getValues();
  const today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][10]) === String(userId)) {
      const t = rows[i][8]; // Termination_Date
      const term = t instanceof Date ? Utilities.formatDate(t, TZ, 'yyyy-MM-dd') : String(t || '').slice(0, 10);
      if (term && term < today) return false; // 退職日を過ぎたらアクセス不可
      return true;
    }
  }
  return false;
}
function requireAccess_(user) {
  if (!isRegisteredStaff_(user.userId)) throw new Error('NO_ACCESS');
}
// 動作確認用（エディタから実行）: owner=true / 未登録=false のはず
function checkAccessGate() {
  Logger.log('owner: ' + isRegisteredStaff_('Ude0725279e144285a91121e3c24653ab'));
  Logger.log('fake: ' + isRegisteredStaff_('Uffffffffffffffffffffffffffffffff'));
}
// エディタから実行して共通マスター接続を確認するための公開関数
function checkMaster() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log('periods: ' + JSON.stringify(listPeriods_()));
  Logger.log('owner surname: ' + mySurname_(ss, 'Ude0725279e144285a91121e3c24653ab'));
  Logger.log('staff surname: ' + mySurname_(ss, 'U85c75169ce5f54a4d7eaa24b03f495d6'));
}
// エディタから実行して接続確認するための公開関数（期間一覧をログに出す）
function checkRosterBook() {
  const ps = listPeriods_();
  Logger.log('periods: ' + JSON.stringify(ps));
}

function dateStr_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy-MM-dd');
  return String(v || '').slice(0, 10);
}
function listPeriods_() {
  // v14: 期間の正は共通マスターの「期間マスター」。確定シフトの実データは割当案ブック側のシートを読む
  const sh = masterBook_().getSheetByName('期間マスター');
  const out = [];
  if (sh) {
    const rows = sh.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      if (!r[7]) continue;
      const sheetName = String(r[7]);
      out.push({ key: sheetName.slice(SHEET_ROSTER_PREFIX.length), sheetName: sheetName,
        start: dateStr_(r[1]), end: dateStr_(r[2]), label: String(r[3] || '') });
    }
  }
  out.sort(function (a, b) { return a.start < b.start ? -1 : (a.start > b.start ? 1 : 0); });
  return out;
}
function pickPeriod_(periods, key) {
  if (!periods.length) return null;
  if (key) { for (let i = 0; i < periods.length; i++) if (periods[i].key === key) return periods[i]; }
  const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  for (let i = 0; i < periods.length; i++) if (periods[i].start <= today && today <= periods[i].end) return periods[i];
  for (let i = 0; i < periods.length; i++) if (periods[i].start > today) return periods[i];
  return periods[periods.length - 1];
}
function mySurname_(ss, userId) {
  // v14: 人の正は共通マスターの「従業員マスター」（Line_ID列→Family_J列）
  const msh = masterBook_().getSheetByName('従業員マスター');
  if (msh) {
    const mrows = msh.getDataRange().getValues();
    for (let i = 1; i < mrows.length; i++) {
      if (mrows[i][10] === userId) return String(mrows[i][1] || '');
    }
  }
  // 移行中のフォールバック: 提出ブックの「スタッフ」シート（オーナーのテスト紐付け用に残す）
  const staffSheet = ss.getSheetByName(SHEET_STAFF);
  if (!staffSheet) return '';
  const srows = staffSheet.getDataRange().getValues();
  for (let i = 1; i < srows.length; i++) {
    if (srows[i][0] === userId) return String(srows[i][1] || '');
  }
  return '';
}
function readPeriodRows_(sheet) {
  const values = sheet.getDataRange().getValues();
  const headers = values[2].map(function (v) { return String(v || ''); });
  const columns = headers.slice(2).filter(function (h) { return h !== ''; });
  const rows = [];
  for (let i = 3; i < values.length; i++) {
    const r = values[i];
    if (r[0] === '' || r[0] === null) continue;
    const ds = dateStr_(r[0]);
    const parts = ds.split('-');
    const cells = {};
    for (let c = 0; c < columns.length; c++) {
      const v = r[2 + c];
      cells[columns[c]] = (v === '' || v === null) ? '' : String(v);
    }
    rows.push({ date: ds, day: Number(parts[2]), dayLabel: Number(parts[1]) + '/' + Number(parts[2]), wd: String(r[1] || ''), cells: cells });
  }
  return { columns: columns, rows: rows };
}
function getRoster2(idToken, displayName, periodKey) {
  const user = verifyIdToken(idToken);
  logAccess_(user, displayName, '照会');
  requireAccess_(user);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const periods = listPeriods_();
  const period = pickPeriod_(periods, periodKey || '');
  if (!period) throw new Error('確定シフトがまだ登録されていません。');
  const data = readPeriodRows_(rosterBook_().getSheetByName(period.sheetName));
  return {
    month: period.start.slice(0, 7), monthLabel: period.label, label: period.label,
    periodKey: period.key,
    periods: periods.map(function (p) { return { key: p.key, label: p.label }; }),
    columns: data.columns, rows: data.rows, me: mySurname_(ss, user.userId)
  };
}
function getIcs2(idToken, periodKey) {
  const user = verifyIdToken(idToken);
  requireAccess_(user);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const me = mySurname_(ss, user.userId);
  if (!me) throw new Error('スタッフ登録が見つかりませんでした。');
  const periods = listPeriods_();
  const period = pickPeriod_(periods, periodKey || '');
  if (!period) throw new Error('確定シフトがまだ登録されていません。');
  const data = readPeriodRows_(rosterBook_().getSheetByName(period.sheetName));
  if (data.columns.indexOf(me) < 0) throw new Error('シフト表にあなたの列がありません。');
  const lines = ['BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//Mitsume Coffee//Shift//JA'];
  data.rows.forEach(function (r) {
    const v = r.cells[me];
    if (!v) return;
    const parts = String(v).split('-');
    if (parts.length !== 2) return;
    const dstr = r.date.replace(/-/g, '');
    lines.push('BEGIN:VEVENT');
    lines.push('UID:shift-' + r.date + '-' + me + '@mitsume-coffee');
    lines.push('DTSTART;TZID=Asia/Tokyo:' + dstr + 'T' + parts[0].replace(':', '') + '00');
    lines.push('DTEND;TZID=Asia/Tokyo:' + dstr + 'T' + parts[1].replace(':', '') + '00');
    lines.push('SUMMARY:三ツ目珈琲 シフト');
    lines.push('END:VEVENT');
  });
  lines.push('END:VCALENDAR');
  return { me: me, monthLabel: period.label, ics: lines.join('\r\n') };
}

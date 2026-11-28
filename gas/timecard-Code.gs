/*************************************************
 * Mitsume Coffee タイムカード（出退勤）システム（受け側）
 * Google Apps Script
 *
 * 仕組み:
 *  - LIFF画面(GitHub Pages)から JSONP API として呼ばれる
 *  - 画面は LINE の IDトークンを送ってくるので、LINEの検証APIで
 *    本人確認してから、その人の分だけ読み書きする（名前選択は不要）
 *  - 出社・退社の打刻は GPS で店から RADIUS_M 以内かを確認してから記録
 *  - 時刻はサーバー（このスクリプト）の時刻を使う
 *  - 1人×1日 = 1行（出社・退社・休憩・休暇）
 *************************************************/

/***** 設定（ここだけ編集すればOK） *****/
const CONFIG = {
  LINE_CHANNEL_ID: '2012734423', // LINEログインチャネル（シフト提出と同じ）

  // 店の位置と打刻できる半径（2026-11-23 ユーザー決定：150m）
  SHOP_LAT: 35.60142,
  SHOP_LNG: 139.35262,
  RADIUS_M: 150,

  // 顧客ID・店舗ID（サービス提供に備えて全テーブルに持たせる。2026-11-24 ユーザー指示）
  CUSTOMER_ID: 'M10001', // 合同会社Giorno
  STORE_ID: 'S001',      // 三ツ目珈琲

  // 共通マスター（従業員マスターでアクセス権を判定する）
  MASTER_SPREADSHEET_ID: '1UowDnTtWjvcWEy0VBDpEQUMeP0SyZ5yDekpztmyWjZQ'
};

const TZ = 'Asia/Tokyo';
const SHEET_LOG = '打刻記録'; // 1人×1日 = 1行

/***** アクセス権（共通マスターの従業員マスターに Line_ID 登録がある人だけ使える） *****/
function masterBook_() { return SpreadsheetApp.openById(CONFIG.MASTER_SPREADSHEET_ID); }
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

/***** アクセス記録（開いた人を残す：管理画面の「候補」に出すため） *****/
function logAccess_(user, displayName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('アクセス記録');
    if (!sheet) {
      sheet = ss.insertSheet('アクセス記録');
      sheet.appendRow(['ユーザーID', '表示名', '初回アクセス', '最終アクセス', '顧客ID', '店舗ID']);
      sheet.setFrozenRows(1);
    }
    const now = new Date();
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(user.userId)) { sheet.getRange(i + 1, 4).setValue(now); return; }
    }
    sheet.appendRow([user.userId, displayName || user.name || '', now, now, CONFIG.CUSTOMER_ID, CONFIG.STORE_ID]);
  } catch (e) { /* 記録に失敗しても本処理は止めない */ }
}

// 休暇の種類（2026-11-23 ユーザー指定）。半休だけは半日勤務なので打刻できる
const LEAVE_TYPES = ['年休', '半休', '忌引', '欠勤', '他種'];
function isFullLeave_(leave) { return !!leave && leave !== '半休'; }

/***** API 入口 *****/
function doGet(e) {
  if (e && e.parameter && e.parameter.api === '1') return handleApi_(e.parameter);
  return ContentService.createTextOutput('Mitsume Coffee タイムカード API');
}

function handleApi_(p) {
  let out;
  try {
    if (p.action === 'getStatus') {
      out = { ok: true, data: getStatus(p.idToken || '', p.name || '') };
    } else if (p.action === 'punch') {
      out = { ok: true, data: punch(p.idToken || '', p.name || '', JSON.parse(p.payload || '{}')) };
    } else if (p.action === 'getMonth') {
      out = { ok: true, data: getMonth(p.idToken || '', p.name || '', p.month || '') };
    } else if (p.action === 'saveEdit') {
      out = { ok: true, data: saveEdit(p.idToken || '', p.name || '', JSON.parse(p.payload || '{}')) };
    } else if (p.action === 'deleteRecord') {
      out = { ok: true, data: deleteRecord(p.idToken || '', p.name || '', p.date || '') };
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

/***** シートの準備 *****/
function ensureSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let log = ss.getSheetByName(SHEET_LOG);
  if (!log) {
    log = ss.insertSheet(SHEET_LOG);
    log.appendRow(['ユーザーID', '表示名', '日付', '出社', '退社', '休憩（分）', '休暇', '出社距離m', '退社距離m', '更新日時', '打刻出社', '打刻退社', '顧客ID', '店舗ID']);
    log.setFrozenRows(1);
  } else {
    if (!log.getRange(1, 11).getValue()) {
      // v3移行（2026-11-24）: 元の打刻時刻を保持する列を追加する
      log.getRange(1, 11).setValue('打刻出社');
      log.getRange(1, 12).setValue('打刻退社');
    }
    if (!log.getRange(1, 13).getValue()) {
      // 顧客ID・店舗ID列を追加する（2026-11-24）
      log.getRange(1, 13).setValue('顧客ID');
      log.getRange(1, 14).setValue('店舗ID');
    }
  }
  return { log: log };
}

/***** 日付ユーティリティ *****/
function todayKey_() {
  return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
}
function nowTime_() {
  return Utilities.formatDate(new Date(), TZ, 'HH:mm');
}
function nowText_() {
  return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');
}
function dateLabel_(key) {
  const d = new Date(key + 'T00:00:00+09:00');
  return Utilities.formatDate(d, TZ, 'yyyy年M月d日(E)');
}

/***** 距離計算（ハバサイン、メートル） *****/
function distanceM_(lat1, lng1, lat2, lng2) {
  const R = 6371000;
  const rad = function (x) { return x * Math.PI / 180; };
  const dLat = rad(lat2 - lat1);
  const dLng = rad(lng2 - lng1);
  const a = Math.sin(dLat / 2) * Math.sin(dLat / 2) +
    Math.cos(rad(lat1)) * Math.cos(rad(lat2)) * Math.sin(dLng / 2) * Math.sin(dLng / 2);
  return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
}

/***** 本人のその日の行を探す（無ければ null、行番号は1始まり） *****/
function findRow_(sheet, userId, dateKey) {
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === userId && String(rows[i][2]) === dateKey) {
      return { row: i + 1, values: rows[i] };
    }
  }
  return null;
}

function recordOf_(values) {
  if (!values) return { inTime: '', outTime: '', breakMin: 0, leave: '', rawIn: '', rawOut: '' };
  return {
    inTime: values[3] ? String(values[3]) : '',
    outTime: values[4] ? String(values[4]) : '',
    breakMin: Number(values[5] || 0),
    leave: values[6] ? String(values[6]) : '',
    // 打刻出社・打刻退社 = ボタンを押した時の元の時刻（編集しても変わらない）
    rawIn: values[10] ? String(values[10]) : '',
    rawOut: values[11] ? String(values[11]) : ''
  };
}

function workMin_(rec) {
  if (!rec.inTime || !rec.outTime) return null;
  const p = function (s) { const a = s.split(':'); return Number(a[0]) * 60 + Number(a[1]); };
  const w = p(rec.outTime) - p(rec.inTime) - (rec.breakMin || 0);
  return w < 0 ? 0 : w;
}

/***** 状態取得（画面の初期表示） *****/
function getStatus(idToken, displayName) {
  const user = verifyIdToken(idToken);
  logAccess_(user, displayName);
  requireAccess_(user);
  ensureSheets();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_LOG);
  const today = todayKey_();
  const found = findRow_(sheet, user.userId, today);
  const rec = recordOf_(found ? found.values : null);
  return {
    name: displayName || user.name || '',
    date: today,
    dateLabel: dateLabel_(today),
    record: rec,
    canIn: !rec.inTime && !isFullLeave_(rec.leave),
    canOut: !!rec.inTime && !rec.outTime,
    openDays: openDays_(sheet, user.userId, today)
  };
}

/***** 未退社のまま残っている過去日（出社だけあって退社が無い日） *****/
function openDays_(sheet, userId, today) {
  const rows = sheet.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] !== userId) continue;
    const key = String(rows[i][2]);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(key) || key >= today) continue;
    const rec = recordOf_(rows[i]);
    if (rec.inTime && !rec.outTime && !isFullLeave_(rec.leave)) {
      out.push({
        date: key, label: dateLabel_(key),
        inTime: rec.inTime, outTime: '', breakMin: rec.breakMin, leave: rec.leave,
        rawIn: rec.rawIn, rawOut: rec.rawOut
      });
    }
  }
  out.sort(function (a, b) { return a.date < b.date ? -1 : 1; });
  return out;
}

/***** 打刻（出社・退社） *****/
function punch(idToken, displayName, payload) {
  const user = verifyIdToken(idToken);
  requireAccess_(user);
  const kind = payload.kind; // 'in' | 'out'
  if (kind !== 'in' && kind !== 'out') throw new Error('打刻の種別が不正です。');

  // GPS チェック（店から RADIUS_M 以内だけ打刻できる）
  const lat = Number(payload.lat), lng = Number(payload.lng);
  if (!isFinite(lat) || !isFinite(lng)) {
    throw new Error('位置情報が取れませんでした。位置情報を許可して、もう一度お試しください。');
  }
  const dist = Math.round(distanceM_(CONFIG.SHOP_LAT, CONFIG.SHOP_LNG, lat, lng));
  if (dist > CONFIG.RADIUS_M) {
    throw new Error('店から ' + dist + 'm 離れています。打刻は店の近く（' + CONFIG.RADIUS_M + 'm以内）でのみできます。');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    ensureSheets();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_LOG);
    const today = todayKey_();
    const found = findRow_(sheet, user.userId, today);
    const rec = recordOf_(found ? found.values : null);
    const time = nowTime_();
    const name = displayName || user.name || '';

    if (kind === 'in') {
      if (isFullLeave_(rec.leave)) throw new Error('今日は' + rec.leave + 'になっています。変更は鉛筆ボタンからどうぞ。');
      if (rec.inTime) throw new Error('すでに出社済みです。');
      if (found) {
        sheet.getRange(found.row, 4).setValue(time);
        sheet.getRange(found.row, 11).setValue(time); // 打刻出社（元の時刻）
        sheet.getRange(found.row, 8).setValue(dist);
        sheet.getRange(found.row, 2).setValue(name);
        sheet.getRange(found.row, 10).setValue(nowText_());
        sheet.getRange(found.row, 13).setValue(CONFIG.CUSTOMER_ID);
        sheet.getRange(found.row, 14).setValue(CONFIG.STORE_ID);
      } else {
        const range = sheet.getRange(sheet.getLastRow() + 1, 1, 1, 14);
        range.setNumberFormat('@');
        range.setValues([[user.userId, name, today, time, '', 0, '', dist, '', nowText_(), time, '', CONFIG.CUSTOMER_ID, CONFIG.STORE_ID]]);
      }
    } else { // out
      if (!rec.inTime) throw new Error('出社が記録されていません。');
      if (rec.outTime) throw new Error('すでに退社済みです。');
      const breakMin = Math.max(0, Math.min(720, Number(payload.breakMin || 0)));
      sheet.getRange(found.row, 5).setValue(time);
      sheet.getRange(found.row, 12).setValue(time); // 打刻退社（元の時刻）
      sheet.getRange(found.row, 13).setValue(CONFIG.CUSTOMER_ID);
      sheet.getRange(found.row, 14).setValue(CONFIG.STORE_ID);
      sheet.getRange(found.row, 6).setValue(breakMin);
      sheet.getRange(found.row, 9).setValue(dist);
      sheet.getRange(found.row, 2).setValue(name);
      sheet.getRange(found.row, 10).setValue(nowText_());
    }
    const after = recordOf_(sheet.getRange(found ? found.row : sheet.getLastRow(), 1, 1, 12).getValues()[0]);
    return { record: after, canIn: !after.inTime && !isFullLeave_(after.leave), canOut: !!after.inTime && !after.outTime, time: time };
  } finally {
    lock.releaseLock();
  }
}

/***** 月の一覧（本人の分だけ） *****/
function getMonth(idToken, displayName, month) {
  const user = verifyIdToken(idToken);
  requireAccess_(user);
  ensureSheets();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_LOG);
  const ym = /^\d{4}-\d{2}$/.test(month || '') ? month : todayKey_().slice(0, 7);
  const rows = sheet.getDataRange().getValues();
  const byDate = {};
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === user.userId && String(rows[i][2]).slice(0, 7) === ym) {
      byDate[String(rows[i][2])] = recordOf_(rows[i]);
    }
  }
  const parts = ym.split('-');
  const y = Number(parts[0]), m = Number(parts[1]);
  const lastDay = new Date(y, m, 0).getDate();
  const dowNames = ['日', '月', '火', '水', '木', '金', '土'];
  const days = [];
  let totalWorkMin = 0;
  for (let d = 1; d <= lastDay; d++) {
    const key = ym + '-' + ('0' + d).slice(-2);
    const rec = byDate[key] || { inTime: '', outTime: '', breakMin: 0, leave: '', rawIn: '', rawOut: '' };
    const wm = workMin_(rec);
    if (wm !== null) totalWorkMin += wm;
    days.push({
      date: key, day: d, dow: dowNames[new Date(key + 'T00:00:00+09:00').getDay()],
      inTime: rec.inTime, outTime: rec.outTime, breakMin: rec.breakMin, leave: rec.leave, workMin: wm,
      rawIn: rec.rawIn, rawOut: rec.rawOut
    });
  }
  return { month: ym, days: days, totalWorkMin: totalWorkMin };
}

/***** 編集（鉛筆ボタン：自分の記録を直す） *****/
function saveEdit(idToken, displayName, payload) {
  const user = verifyIdToken(idToken);
  requireAccess_(user);
  const date = String(payload.date || '');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('日付が不正です。');
  const inTime = String(payload.inTime || '');
  const outTime = String(payload.outTime || '');
  const timeOk = function (s) { return s === '' || /^([01]\d|2[0-3]):[0-5]\d$/.test(s); };
  if (!timeOk(inTime) || !timeOk(outTime)) throw new Error('時刻の形式が違います。');
  if (inTime && outTime && outTime <= inTime) throw new Error('退社は出社より後の時刻にしてください。');
  const breakMin = Math.max(0, Math.min(720, Number(payload.breakMin || 0)));
  const leave = String(payload.leave || '');
  if (leave && LEAVE_TYPES.indexOf(leave) < 0) throw new Error('休暇の種類が正しくありません。');
  const fullLeave = isFullLeave_(leave);

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    ensureSheets();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_LOG);
    const found = findRow_(sheet, user.userId, date);
    const name = displayName || user.name || '';
    const values = [user.userId, name, date, fullLeave ? '' : inTime, fullLeave ? '' : outTime, fullLeave ? 0 : breakMin, leave, '', '', nowText_(), '', '', CONFIG.CUSTOMER_ID, CONFIG.STORE_ID];
    if (found) {
      // 距離と元の打刻時刻は、元の打刻のものを残す（編集で消さない）
      const old = found.values;
      values[7] = old[7]; values[8] = old[8];
      values[10] = old[10] || ''; values[11] = old[11] || '';
      const range = sheet.getRange(found.row, 1, 1, 14);
      range.setNumberFormat('@');
      range.setValues([values]);
    } else {
      const range = sheet.getRange(sheet.getLastRow() + 1, 1, 1, 14);
      range.setNumberFormat('@');
      range.setValues([values]);
    }
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

/***** 記録の削除（編集画面から） *****/
function deleteRecord(idToken, displayName, date) {
  const user = verifyIdToken(idToken);
  requireAccess_(user);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(date || ''))) throw new Error('日付が不正です。');
  ensureSheets();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_LOG);
  const found = findRow_(sheet, user.userId, date);
  if (found) sheet.deleteRow(found.row);
  return { ok: true };
}

/***** 集計（月を指定して、日付×スタッフの勤務表を作る） *****/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('タイムカード集計')
    .addItem('今月分を集計', 'buildSummaryThisMonth')
    .addItem('月を指定して集計', 'buildSummaryPrompt')
    .addToUi();
}

function buildSummaryThisMonth() {
  buildSummary(todayKey_().slice(0, 7));
}

function buildSummaryPrompt() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('集計する月', '例: 2026-12 のように入力してください', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const month = res.getResponseText().trim();
  if (!/^\d{4}-\d{2}$/.test(month)) { ui.alert('形式が違います。例: 2026-12'); return; }
  buildSummary(month);
}

function fmtHours_(min) {
  if (min === null || min === undefined) return '';
  return (Math.round(min / 6) / 10).toFixed(1); // 6分=0.1時間
}

function buildSummary(month) {
  ensureSheets();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_LOG);
  const rows = sheet.getDataRange().getValues();

  const staff = [];
  const staffName = {};
  const cell = {}; // userId|date -> {in,out,brk,leave}
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][2]).slice(0, 7) === month) {
      const uid = rows[i][0];
      if (!staffName[uid]) { staff.push(uid); staffName[uid] = rows[i][1] || '名前なし'; }
      cell[uid + '|' + String(rows[i][2])] = recordOf_(rows[i]);
    }
  }
  const parts = month.split('-');
  const y = Number(parts[0]), m = Number(parts[1]);
  const lastDay = new Date(y, m, 0).getDate();
  const dowNames = ['日', '月', '火', '水', '木', '金', '土'];
  const out = [['日付', '曜日'].concat(staff.map(function (u) { return staffName[u]; }))];
  const totals = staff.map(function () { return 0; });
  for (let d = 1; d <= lastDay; d++) {
    const key = month + '-' + ('0' + d).slice(-2);
    const row = [key, dowNames[new Date(key + 'T00:00:00+09:00').getDay()]];
    staff.forEach(function (u, idx) {
      const rec = cell[u + '|' + key];
      if (!rec) { row.push(''); return; }
      if (rec.leave && !rec.inTime) { row.push(rec.leave); return; }
      const wm = workMin_(rec);
      if (wm !== null) totals[idx] += wm;
      row.push(rec.inTime ? (rec.inTime + '-' + (rec.outTime || '') + (wm !== null ? ' (' + fmtHours_(wm) + 'h)' : '')) : '');
    });
    out.push(row);
  }
  out.push(['合計（時間）', ''].concat(totals.map(function (t) { return fmtHours_(t); })));

  const name = '集計_' + month;
  let s = ss.getSheetByName(name);
  if (s) s.clearContents(); else s = ss.insertSheet(name);
  s.getRange(1, 1, out.length, out[0].length).setValues(out);
  s.setFrozenRows(1);
  s.setFrozenColumns(2);
  ss.setActiveSheet(s);
}

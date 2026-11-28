/*************************************************
 * 店長代理 管理API（共通マスターの読み書き）
 * Google Apps Script（スタンドアロン）
 *
 * 仕組み:
 *  - 管理LIFF画面とポータル画面から JSONP API として呼ばれる
 *  - LINE の IDトークンで本人確認し、共通マスターの「従業員マスター」の
 *    ロール（admin）で管理者かどうかを毎回チェックする（当たらないリクエストは拒否）
 *  - 変更はすべて共通マスターの「変更履歴」に残す
 *  - ポータル用に checkAdmin（管理者かどうかだけを返す）も用意する
 *************************************************/

const CONFIG = {
  LINE_CHANNEL_ID: '2012734423', // LINEログインチャネル（他のアプリと同じ）
  MASTER_SPREADSHEET_ID: '1UowDnTtWjvcWEy0VBDpEQUMeP0SyZ5yDekpztmyWjZQ', // 共通マスター（ホーム店舗）
  HOME_STORE_ID: 'S001', // この管理APIがホームとする店舗
  PARENT_FOLDER_ID: '16K4FHfKihhnUr-43SwK9-eKwFv8IQqk1', // 店長代理M（店舗フォルダの親）

  // 候補（まだ従業員マスターに無い人）の収集元
  SHIFT_SPREADSHEET_ID: '1kxtknUONwVKgvLeoqcW_OqD1cX2vN43yHtE3u24JI68',    // シフト提出（アクセス記録）
  TIMECARD_SPREADSHEET_ID: '1N_8wL2WFEMFfrDuDzVDwRPlUq-e-RkicPe_hrPyk7Rk', // タイムカード（打刻記録）
  REPORT_SPREADSHEET_ID: '1rAB6sxFmrm3UXvRS1BCjADuYKG7NDVoshkipR6QETtQ'     // 壊れ物報告（報告一覧）
};

const TZ = 'Asia/Tokyo';
const SHEET_EMP = '従業員マスター';
const SHEET_PERIOD = '期間マスター';
const SHEET_SETTING = '設定';
const SHEET_HISTORY = '変更履歴';
const SHEET_STORE = '店舗マスター';

/***** API 入口 *****/
function doGet(e) {
  if (e && e.parameter && e.parameter.api === '1') return handleApi_(e.parameter);
  return ContentService.createTextOutput('店長代理 管理 API');
}

function handleApi_(p) {
  let out;
  try {
    if (p.action === 'checkAdmin') {
      out = { ok: true, data: checkAdmin(p.idToken || '') };
    } else if (p.action === 'getAdmin') {
      out = { ok: true, data: getAdmin(p.idToken || '', p.name || '', JSON.parse(p.payload || '{}')) };
    } else if (p.action === 'saveEmployee') {
      out = { ok: true, data: saveEmployee(p.idToken || '', p.name || '', JSON.parse(p.payload || '{}')) };
    } else if (p.action === 'savePeriod') {
      out = { ok: true, data: savePeriod(p.idToken || '', p.name || '', JSON.parse(p.payload || '{}')) };
    } else if (p.action === 'deletePeriod') {
      out = { ok: true, data: deletePeriod(p.idToken || '', p.name || '', JSON.parse(p.payload || '{}')) };
    } else if (p.action === 'saveSettings') {
      out = { ok: true, data: saveSettings(p.idToken || '', p.name || '', JSON.parse(p.payload || '{}')) };
    } else if (p.action === 'getStore') {
      out = { ok: true, data: getStore(p.idToken || '', p.name || '', JSON.parse(p.payload || '{}')) };
    } else if (p.action === 'addStore') {
      out = { ok: true, data: addStore(p.idToken || '', p.name || '', JSON.parse(p.payload || '{}')) };
    } else if (p.action === 'getTimecard') {
      out = { ok: true, data: getTimecard(p.idToken || '', p.name || '', JSON.parse(p.payload || '{}')) };
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

/***** 動作確認用（エディタから実行すると権限承認も済む） *****/
function setupCheck() {
  const ss = masterBook_();
  const emp = ss.getSheetByName(SHEET_EMP).getDataRange().getValues();
  let admins = 0;
  for (let i = 1; i < emp.length; i++) { if (String(emp[i][13] || '') === 'admin') admins++; }
  Logger.log('従業員マスター: ' + (emp.length - 1) + '名 / 管理者: ' + admins + '名 / 店舗: ' + readStores_(ss).length + '店');
}

// 店舗追加（Drive操作）の権限承認用：親フォルダが見えるかだけ確認する
function checkDrive() {
  Logger.log('親フォルダ: ' + DriveApp.getFolderById(CONFIG.PARENT_FOLDER_ID).getName());
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

/***** 共通マスター *****/
function masterBook_() {
  return SpreadsheetApp.openById(CONFIG.MASTER_SPREADSHEET_ID);
}

function nowText_() {
  return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm');
}

/***** 権限チェック *****/
// 従業員マスターで Line_ID が一致し、ロールが admin の人が管理者。
// 所属店舗（O列）がブランクなら「全社管理」（全店舗を管理できる）。
// 所属店舗に店舗IDが入っていれば、その店舗だけの「店舗管理者」。
// 列: 1 Family_J / 2 First_J / 10 Line_ID / 13 ロール / 14 所属店舗
function findAdmin_(userId) {
  const rows = masterBook_().getSheetByName(SHEET_EMP).getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][10]) === userId && String(rows[i][13] || '') === 'admin') {
      const name = (String(rows[i][1] || '') + ' ' + String(rows[i][2] || '')).trim();
      const belong = String(rows[i][14] || '');
      if (!belong) return { name: name, role: '全社管理', scope: 'all', storeId: '' };
      if (belong === CONFIG.HOME_STORE_ID) return { name: name, role: '店舗管理者', scope: 'store', storeId: belong };
      return null; // 他店の店舗管理者は、この管理画面の対象外
    }
  }
  return null;
}

function requireAdmin_(userId) {
  const admin = findAdmin_(userId);
  if (!admin) throw new Error('管理者として登録されていません');
  return admin;
}

// ポータル用：管理者かどうかだけ返す（違ってもエラーにしない）
function checkAdmin(idToken) {
  const user = verifyIdToken(idToken);
  const admin = findAdmin_(user.userId);
  if (!admin) return { admin: false };
  return { admin: true, role: admin.role, scope: admin.scope, name: admin.name };
}

/***** 変更履歴 *****/
function logChange_(adminName, target, detail, ss) {
  const sheet = (ss || masterBook_()).getSheetByName(SHEET_HISTORY);
  sheet.appendRow([nowText_(), adminName, target, detail]);
}

/***** 店舗マスター *****/
// 列: 0 店舗ID / 1 店名 / 2 共通マスターID / 3 フォルダID
function readStores_(ss) {
  const sheet = ss.getSheetByName(SHEET_STORE);
  if (!sheet) return [];
  const rows = sheet.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    out.push({ storeId: String(rows[i][0]), name: String(rows[i][1] || ''), masterId: String(rows[i][2] || ''), folderId: String(rows[i][3] || '') });
  }
  return out;
}

function storeBook_(storeId) {
  const stores = readStores_(masterBook_());
  for (let i = 0; i < stores.length; i++) {
    if (stores[i].storeId === storeId && stores[i].masterId) return SpreadsheetApp.openById(stores[i].masterId);
  }
  throw new Error('店舗が見つかりません: ' + storeId);
}

// その店舗を触ってよい管理者か（全社管理は全店可、店舗管理者は自店のみ）
function requireStoreAccess_(admin, storeId) {
  if (admin.scope === 'all') return;
  if (admin.scope === 'store' && admin.storeId === storeId) return;
  throw new Error('その店舗の管理権限がありません');
}

/***** 設定 *****/
function readSettings_(ss) {
  const rows = ss.getSheetByName(SHEET_SETTING).getDataRange().getValues();
  const map = {};
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0]) map[String(rows[i][0])] = String(rows[i][1] || '');
  }
  return map;
}

/***** 従業員マスター *****/
// 列: 0 Employee_No / 1 Family_J / 2 First_J / 3 Family_E / 4 First_E / 5 DispOrder /
//     6 Birthday / 7 Employment_Date / 8 Termination_Date / 9 Type / 10 Line_ID / 11 顧客ID / 12 店舗ID / 13 ロール / 14 所属店舗
function employeeOf_(values, row) {
  return {
    row: row,
    employeeNo: String(values[0] || ''),
    familyJ: String(values[1] || ''), firstJ: String(values[2] || ''),
    familyE: String(values[3] || ''), firstE: String(values[4] || ''),
    dispOrder: values[5] === '' || values[5] === null ? '' : Number(values[5]),
    birthday: String(values[6] || ''), employmentDate: String(values[7] || ''),
    terminationDate: String(values[8] || ''), type: String(values[9] || ''),
    lineId: String(values[10] || ''), roll: String(values[13] || ''),
    belongStore: String(values[14] || '')
  };
}

function readEmployees_(ss) {
  const rows = ss.getSheetByName(SHEET_EMP).getDataRange().getValues();
  const out = [];
  for (let i = 1; i < rows.length; i++) out.push(employeeOf_(rows[i], i + 1));
  out.sort(function (a, b) {
    const ao = a.dispOrder === '' ? 9999 : a.dispOrder;
    const bo = b.dispOrder === '' ? 9999 : b.dispOrder;
    return ao - bo;
  });
  return out;
}

/***** 期間マスター *****/
// 列: 0 期間ID / 1 対象期間開始日 / 2 対象期間終了日 / 3 ラベル / 4 受付開始日時 / 5 受付終了日時 /
//     6 個別延長 / 7 確定シフトシート名 / 8 顧客ID / 9 店舗ID / 10 期間No
function periodOf_(values, row) {
  return {
    row: row,
    periodId: String(values[0] || ''),
    start: String(values[1] || ''), end: String(values[2] || ''),
    label: String(values[3] || ''),
    acceptStart: String(values[4] || ''), acceptEnd: String(values[5] || ''),
    sheetName: String(values[7] || ''),
    periodNo: values[10] === '' || values[10] === null || values[10] === undefined ? '' : Number(values[10])
  };
}

function readPeriods_(ss) {
  const rows = ss.getSheetByName(SHEET_PERIOD).getDataRange().getValues();
  const out = [];
  for (let i = 1; i < rows.length; i++) out.push(periodOf_(rows[i], i + 1));
  out.sort(function (a, b) { return a.start < b.start ? -1 : 1; });
  return out;
}

/***** 候補（各アプリの記録に現れるのに、従業員マスターに無い人） *****/
function readCandidates_(ss) {
  const known = {};
  readEmployees_(ss).forEach(function (e) { if (e.lineId) known[e.lineId] = true; });
  const found = {}; // userId -> {userId, name, sources:[], lastSeen}
  function add(userId, name, source, lastSeen) {
    if (!userId || known[userId]) return;
    const cur = found[userId] || { userId: userId, name: '', sources: [], lastSeen: '' };
    if (name) cur.name = name;
    if (cur.sources.indexOf(source) < 0) cur.sources.push(source);
    if (String(lastSeen || '') > cur.lastSeen) cur.lastSeen = String(lastSeen || '');
    found[userId] = cur;
  }
  try { // シフト提出「アクセス記録」: 0 ユーザーID / 1 表示名 / 3 最終アクセス
    const rows = SpreadsheetApp.openById(CONFIG.SHIFT_SPREADSHEET_ID).getSheetByName('アクセス記録').getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) add(String(rows[i][0]), String(rows[i][1] || ''), 'シフト', rows[i][3]);
  } catch (e) { /* 読めない収集元は飛ばす */ }
  try { // タイムカード「打刻記録」: 0 ユーザーID / 1 表示名 / 9 更新日時
    const rows = SpreadsheetApp.openById(CONFIG.TIMECARD_SPREADSHEET_ID).getSheetByName('打刻記録').getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) add(String(rows[i][0]), String(rows[i][1] || ''), 'タイムカード', rows[i][9]);
  } catch (e) { /* 読めない収集元は飛ばす */ }
  try { // 壊れ物報告「報告一覧」: 1 報告日時 / 2 報告者 / 3 ユーザーID
    const rows = SpreadsheetApp.openById(CONFIG.REPORT_SPREADSHEET_ID).getSheetByName('報告一覧').getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) add(String(rows[i][3]), String(rows[i][2] || ''), '壊れ物報告', rows[i][1]);
  } catch (e) { /* 読めない収集元は飛ばす */ }
  try { // タイムカード「アクセス記録」: 0 ユーザーID / 1 表示名 / 3 最終アクセス
    const rows = SpreadsheetApp.openById(CONFIG.TIMECARD_SPREADSHEET_ID).getSheetByName('アクセス記録').getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) add(String(rows[i][0]), String(rows[i][1] || ''), 'タイムカード', rows[i][3]);
  } catch (e) { /* 読めない収集元は飛ばす */ }
  try { // 壊れ物報告「アクセス記録」: 0 ユーザーID / 1 表示名 / 3 最終アクセス
    const rows = SpreadsheetApp.openById(CONFIG.REPORT_SPREADSHEET_ID).getSheetByName('アクセス記録').getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) add(String(rows[i][0]), String(rows[i][1] || ''), '壊れ物報告', rows[i][3]);
  } catch (e) { /* 読めない収集元は飛ばす */ }
  const out = Object.keys(found).map(function (k) { return found[k]; });
  out.sort(function (a, b) { return a.lastSeen < b.lastSeen ? 1 : -1; });
  return out;
}

/***** 初期表示（管理者のみ） *****/
function getAdmin(idToken, displayName, payload) {
  const user = verifyIdToken(idToken);
  const admin = requireAdmin_(user.userId);
  const ss = masterBook_();
  const settings = readSettings_(ss);
  let stores = readStores_(ss);
  if (admin.scope === 'store') stores = stores.filter(function (s) { return s.storeId === admin.storeId; });
  // 店舗フィルタ: 指定があればその店のマスター、なければ全体管理者は全店分を集約、店舗管理者は自店分
  const wantStore = String((payload && payload.storeId) || '');
  let employees = [], periods = [];
  if (wantStore) {
    requireStoreAccess_(admin, wantStore);
    const b = storeBook_(wantStore);
    employees = readEmployees_(b).map(function (e) { e.storeId = wantStore; return e; });
    periods = readPeriods_(b).map(function (p) { p.storeId = wantStore; return p; });
  } else if (admin.scope === 'store') {
    const b = storeBook_(admin.storeId);
    employees = readEmployees_(b).map(function (e) { e.storeId = admin.storeId; return e; });
    periods = readPeriods_(b).map(function (p) { p.storeId = admin.storeId; return p; });
  } else {
    stores.forEach(function (s) {
      const b = storeBook_(s.storeId);
      readEmployees_(b).forEach(function (e) { e.storeId = s.storeId; e.storeName = s.name; employees.push(e); });
      readPeriods_(b).forEach(function (p) { p.storeId = s.storeId; p.storeName = s.name; periods.push(p); });
    });
    periods.sort(function (a, b) { return a.start < b.start ? -1 : 1; });
  }
  return {
    admin: { name: admin.name || displayName || user.name || '', role: admin.role, scope: admin.scope, storeId: admin.storeId },
    customerName: settings['顧客名'] || '',
    customerId: settings['顧客ID'] || '',
    homeStoreId: CONFIG.HOME_STORE_ID,
    stores: stores,
    employees: employees,
    periods: periods,
    settings: settings,
    candidates: readCandidates_(ss),
    viewStoreId: wantStore
  };
}

/***** バリデーション *****/
function dateOk_(s) { return s === '' || /^\d{4}-\d{2}-\d{2}$/.test(s); }
function dateTimeOk_(s) { return s === '' || /^\d{4}-\d{2}-\d{2} \d{2}:\d{2}$/.test(s); }

/***** 従業員の保存（追加・編集） *****/
function saveEmployee(idToken, displayName, payload) {
  const user = verifyIdToken(idToken);
  const admin = requireAdmin_(user.userId);
  const familyJ = String(payload.familyJ || '').trim();
  if (!familyJ) throw new Error('姓（漢字）は必須です。');
  const type = String(payload.type || '');
  if (type && type !== 'FullTime' && type !== 'Parttime') throw new Error('種別が正しくありません。');
  const roll = String(payload.roll || '');
  if (roll && roll !== 'admin') throw new Error('ロールが正しくありません。');
  const belongStore = String(payload.belongStore || '');
  if (belongStore && !readStores_(masterBook_()).some(function (s) { return s.storeId === belongStore; })) {
    throw new Error('所属店舗が正しくありません。');
  }
  const birthday = String(payload.birthday || ''), employmentDate = String(payload.employmentDate || ''), terminationDate = String(payload.terminationDate || '');
  if (!dateOk_(birthday) || !dateOk_(employmentDate) || !dateOk_(terminationDate)) throw new Error('日付の形式が違います。');
  const dispOrder = payload.dispOrder === '' || payload.dispOrder === null || payload.dispOrder === undefined ? '' : Number(payload.dispOrder);
  if (dispOrder !== '' && (!isFinite(dispOrder) || dispOrder < 0)) throw new Error('並び順が正しくありません。');

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const targetStoreE = String(payload.storeId || '');
    if (targetStoreE) requireStoreAccess_(admin, targetStoreE);
    const ss = targetStoreE ? storeBook_(targetStoreE) : masterBook_();
    const sheet = ss.getSheetByName(SHEET_EMP);
    const row = Number(payload.row || 0);
    let old = null, employeeNo = '', customerId = '', storeId = '';
    if (row >= 2) {
      const values = sheet.getRange(row, 1, 1, 15).getValues()[0];
      old = employeeOf_(values, row);
      employeeNo = String(values[0] || ''); customerId = String(values[11] || ''); storeId = String(values[12] || '');
    } else {
      const settings = readSettings_(ss);
      customerId = settings['顧客ID'] || ''; storeId = settings['店舗ID'] || '';
    }
    const values = [employeeNo, familyJ, String(payload.firstJ || ''), String(payload.familyE || ''), String(payload.firstE || ''),
      dispOrder, birthday, employmentDate, terminationDate, type, String(payload.lineId || ''), customerId, storeId, roll, belongStore];
    let targetRow = row;
    if (row >= 2) {
      const range = sheet.getRange(row, 1, 1, 15);
      range.setNumberFormat('@');
      range.setValues([values]);
    } else {
      targetRow = sheet.getLastRow() + 1;
      const range = sheet.getRange(targetRow, 1, 1, 15);
      range.setNumberFormat('@');
      range.setValues([values]);
    }
    const name = familyJ + ' ' + String(payload.firstJ || '');
    let detail;
    if (!old) {
      detail = '従業員を追加: ' + name + (values[10] ? '（LINE紐付け済み）' : '（LINE未登録）');
    } else {
      const diffs = [];
      if (old.familyJ + ' ' + old.firstJ !== name) diffs.push('氏名');
      if (old.type !== type) diffs.push('種別');
      if (old.lineId !== values[10]) diffs.push(values[10] ? 'LINE紐付け' : 'LINE解除');
      if (old.terminationDate !== terminationDate) diffs.push(terminationDate ? '退職日を設定' : '退職日を解除');
      if (old.roll !== roll) diffs.push(roll === 'admin' ? '管理者に設定' : '管理者を解除');
      if (old.belongStore !== belongStore) diffs.push('所属店舗');
      if (String(old.dispOrder) !== String(dispOrder)) diffs.push('並び順');
      detail = '従業員を更新: ' + name + (diffs.length ? '（' + diffs.join('・') + '）' : '');
    }
    logChange_(admin.name || displayName || '', SHEET_EMP, detail);
    return { employees: readEmployees_(ss), candidates: readCandidates_(ss) };
  } finally {
    lock.releaseLock();
  }
}

/***** 期間の保存（追加・編集） *****/
function savePeriod(idToken, displayName, payload) {
  const user = verifyIdToken(idToken);
  const admin = requireAdmin_(user.userId);
  const start = String(payload.start || ''), end = String(payload.end || '');
  if (!dateOk_(start) || !start || !dateOk_(end) || !end) throw new Error('対象期間の開始日・終了日を入れてください。');
  const targetStoreP = String(payload.storeId || '');
  if (targetStoreP) requireStoreAccess_(admin, targetStoreP);
  const periodBook = targetStoreP ? storeBook_(targetStoreP) : masterBook_();
  let periodId = String(payload.periodId || '').trim();
  if (Number(payload.row || 0) < 2 && !periodId) {
    // 新規作成で期間IDが空なら、開始日の年月（YYYYMM）から自動で付ける。重なったら -2, -3… を足す
    const base = start.slice(0, 4) + start.slice(5, 7);
    const existing = {};
    periodBook.getSheetByName(SHEET_PERIOD).getDataRange().getValues().forEach(function (r, i) {
      if (i > 0 && r[0]) existing[String(r[0])] = true;
    });
    periodId = base;
    let n = 2;
    while (existing[periodId]) { periodId = base + '-' + n; n++; }
  }
  if (!periodId) throw new Error('期間IDは必須です。');
  const acceptStart = String(payload.acceptStart || '').replace('T', ' ');
  const acceptEnd = String(payload.acceptEnd || '').replace('T', ' ');
  if (!dateTimeOk_(acceptStart) || !dateTimeOk_(acceptEnd)) throw new Error('受付期間の形式が違います。');

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = periodBook;
    const sheet = ss.getSheetByName(SHEET_PERIOD);
    const row = Number(payload.row || 0);
    let extensions = '', customerId = '', storeId = '', periodNo = '';
    if (row >= 2) {
      const values = sheet.getRange(row, 1, 1, 11).getValues()[0];
      extensions = String(values[6] || ''); customerId = String(values[8] || ''); storeId = String(values[9] || '');
      periodNo = values[10] === '' || values[10] === null ? '' : Number(values[10]);
    } else {
      const settings = readSettings_(ss);
      customerId = settings['顧客ID'] || ''; storeId = settings['店舗ID'] || '';
      const all = sheet.getDataRange().getValues();
      let maxNo = 0;
      for (let i = 1; i < all.length; i++) { const n = Number(all[i][10] || 0); if (n > maxNo) maxNo = n; }
      periodNo = maxNo + 1;
    }
    const values = [periodId, start, end, String(payload.label || ''), acceptStart, acceptEnd, extensions,
      String(payload.sheetName || ''), customerId, storeId, periodNo];
    if (row >= 2) {
      const range = sheet.getRange(row, 1, 1, 11);
      range.setNumberFormat('@');
      range.setValues([values]);
    } else {
      const range = sheet.getRange(sheet.getLastRow() + 1, 1, 1, 11);
      range.setNumberFormat('@');
      range.setValues([values]);
    }
    logChange_(admin.name || displayName || '', SHEET_PERIOD,
      (row >= 2 ? '期間を更新: ' : '期間を追加: ') + periodId + '（' + start + '〜' + end + '）');
    return { periods: readPeriods_(ss), savedId: periodId };
  } finally {
    lock.releaseLock();
  }
}

/***** 期間の削除 *****/
function deletePeriod(idToken, displayName, payload) {
  const user = verifyIdToken(idToken);
  const admin = requireAdmin_(user.userId);
  const row = Number(payload.row || 0);
  if (row < 2) throw new Error('削除する期間が指定されていません。');
  const targetStoreD = String(payload.storeId || '');
  if (targetStoreD) requireStoreAccess_(admin, targetStoreD);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = targetStoreD ? storeBook_(targetStoreD) : masterBook_();
    const sheet = ss.getSheetByName(SHEET_PERIOD);
    const values = sheet.getRange(row, 1, 1, 11).getValues()[0];
    const p = periodOf_(values, row);
    sheet.deleteRow(row);
    logChange_(admin.name || displayName || '', SHEET_PERIOD,
      '期間を削除: ' + p.periodId + '（' + p.start + '〜' + p.end + '）');
    return { periods: readPeriods_(ss) };
  } finally {
    lock.releaseLock();
  }
}

/***** 店舗設定の保存（店舗ごとの共通マスターに書く） *****/
const EDITABLE_SETTINGS = ['店名', '定休日（曜日）', '店舗緯度', '店舗経度', 'GPS判定半径m', '特別休日'];
function saveSettings(idToken, displayName, payload) {
  const user = verifyIdToken(idToken);
  const admin = requireAdmin_(user.userId);
  const storeId = String(payload.storeId || CONFIG.HOME_STORE_ID);
  requireStoreAccess_(admin, storeId);
  const incoming = payload.settings || {};
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = storeBook_(storeId);
    const sheet = ss.getSheetByName(SHEET_SETTING);
    const rows = sheet.getDataRange().getValues();
    const changed = [];
    const seen = {};
    for (let i = 1; i < rows.length; i++) {
      const key = String(rows[i][0] || '');
      if (!key) continue;
      seen[key] = true;
      if (EDITABLE_SETTINGS.indexOf(key) >= 0 && incoming[key] !== undefined) {
        const v = String(incoming[key]);
        if (v !== String(rows[i][1] || '')) {
          sheet.getRange(i + 1, 2).setValue(v);
          changed.push(key);
        }
      }
    }
    // 設定シートにまだ無い項目（特別休日など）は行を足す
    EDITABLE_SETTINGS.forEach(function (key) {
      if (!seen[key] && incoming[key] !== undefined && String(incoming[key]) !== '') {
        sheet.appendRow([key, String(incoming[key]), '']);
        changed.push(key);
      }
    });
    if (changed.length) {
      logChange_(admin.name || displayName || '', SHEET_SETTING, '設定を更新（' + storeId + '）: ' + changed.join('・'), ss);
      if (storeId !== CONFIG.HOME_STORE_ID) {
        logChange_(admin.name || displayName || '', SHEET_SETTING, '設定を更新（' + storeId + '）: ' + changed.join('・'));
      }
    }
    return { settings: readSettings_(ss), storeId: storeId };
  } finally {
    lock.releaseLock();
  }
}

/***** 店舗の設定を読む（店舗管理の画面用） *****/
function getStore(idToken, displayName, payload) {
  const user = verifyIdToken(idToken);
  const admin = requireAdmin_(user.userId);
  const storeId = String(payload.storeId || '');
  requireStoreAccess_(admin, storeId);
  const ss = storeBook_(storeId);
  return { storeId: storeId, settings: readSettings_(ss) };
}

/***** 店舗の追加（保存時にフォルダと共通マスターまで作る。全社管理者のみ） *****/
function addStore(idToken, displayName, payload) {
  const user = verifyIdToken(idToken);
  const admin = requireAdmin_(user.userId);
  if (admin.scope !== 'all') throw new Error('店舗の追加は全社管理者だけができます。');
  const name = String(payload.name || '').trim();
  if (!name) throw new Error('店名を入れてください。');
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const home = masterBook_();
    const homeSettings = readSettings_(home);
    const stores = readStores_(home);
    let maxN = 0;
    stores.forEach(function (s) {
      const m = /^S(\d+)$/.exec(s.storeId);
      if (m) maxN = Math.max(maxN, Number(m[1]));
    });
    const newId = 'S' + ('00' + (maxN + 1)).slice(-3);

    // 店舗フォルダ（店長代理S00X）を親フォルダの下に作る
    const folder = DriveApp.getFolderById(CONFIG.PARENT_FOLDER_ID).createFolder('店長代理' + newId);
    // 共通マスターを作って店舗フォルダへ移す
    const book = SpreadsheetApp.create('共通マスター');
    const bookFile = DriveApp.getFileById(book.getId());
    folder.addFile(bookFile);
    DriveApp.getRootFolder().removeFile(bookFile);

    // シート構成はホーム店舗の共通マスターに合わせる（見出し行をコピー）
    const sheetNames = [SHEET_EMP, SHEET_PERIOD, SHEET_SETTING, '権限マスター', SHEET_HISTORY, SHEET_STORE];
    const first = book.getSheets()[0];
    first.setName(sheetNames[0]);
    for (let i = 1; i < sheetNames.length; i++) book.insertSheet(sheetNames[i]);
    sheetNames.forEach(function (sn) {
      const srcSheet = home.getSheetByName(sn);
      const dstSheet = book.getSheetByName(sn);
      if (srcSheet && dstSheet) {
        const lastCol = srcSheet.getLastColumn();
        if (lastCol > 0) dstSheet.getRange(1, 1, 1, lastCol).setValues([srcSheet.getRange(1, 1, 1, lastCol).getValues()[0]]);
      }
    });
    // 設定の初期値
    const setSheet = book.getSheetByName(SHEET_SETTING);
    [['店名', name], ['顧客ID', homeSettings['顧客ID'] || ''], ['店舗ID', newId],
     ['定休日（曜日）', ''], ['店舗緯度', ''], ['店舗経度', ''], ['GPS判定半径m', '150'],
     ['スキーマ版数', '1'], ['特別休日', ''], ['顧客名', homeSettings['顧客名'] || '']
    ].forEach(function (kv) { setSheet.appendRow([kv[0], kv[1], '']); });
    // 自店の店舗マスター行と変更履歴の初期記録
    book.getSheetByName(SHEET_STORE).appendRow([newId, name, book.getId(), folder.getId(), '']);
    book.getSheetByName(SHEET_HISTORY).appendRow([nowText_(), admin.name || displayName || '', '共通マスター', '店舗追加時に新規作成（' + newId + ' ' + name + '）']);
    // ホームの店舗マスターに登録
    home.getSheetByName(SHEET_STORE).appendRow([newId, name, book.getId(), folder.getId(), '']);
    logChange_(admin.name || displayName || '', SHEET_STORE, '店舗を追加: ' + newId + ' ' + name + '（共通マスターを作成）');
    return {
      stores: readStores_(home),
      created: { storeId: newId, name: name, masterId: book.getId(), folderId: folder.getId() }
    };
  } finally {
    lock.releaseLock();
  }
}

/***** タイムカード全員分（管理画面の閲覧用） *****/
// 打刻記録の列: 0 ユーザーID / 1 表示名 / 2 日付 / 3 出社 / 4 退社 / 5 休憩（分） / 6 休暇 /
//               7 出社距離m / 8 退社距離m / 9 更新日時 / 10 打刻出社 / 11 打刻退社
function timeToMin_(t) {
  const p = String(t).split(':');
  return Number(p[0]) * 60 + Number(p[1] || 0);
}
function dateKey_(v) {
  if (Object.prototype.toString.call(v) === '[object Date]') return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  return String(v || '');
}
function getTimecard(idToken, displayName, payload) {
  const user = verifyIdToken(idToken);
  const admin = requireAdmin_(user.userId);
  const month = /^\d{4}-\d{2}$/.test(String(payload.month || '')) ? String(payload.month)
    : Utilities.formatDate(new Date(), TZ, 'yyyy-MM');
  const wantStore = String(payload.storeId || '');
  if (wantStore) requireStoreAccess_(admin, wantStore);
  const storeNameOf = {};
  const empByLine = {};
  readStores_(masterBook_()).forEach(function (s) {
    storeNameOf[s.storeId] = s.name;
    if (admin.scope === 'store' && s.storeId !== admin.storeId) return;
    readEmployees_(storeBook_(s.storeId)).forEach(function (e) { if (e.lineId && !empByLine[e.lineId]) empByLine[e.lineId] = e; });
  });
  const rows = SpreadsheetApp.openById(CONFIG.TIMECARD_SPREADSHEET_ID).getSheetByName('打刻記録').getDataRange().getValues();
  const byPerson = {};
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    const date = dateKey_(r[2]);
    if (date.slice(0, 7) !== month) continue;
    if (wantStore && String(r[13] || '') !== wantStore) continue;
    const lineId = String(r[0] || '');
    const emp = empByLine[lineId];
    const name = emp ? (emp.familyJ + ' ' + emp.firstJ).trim() : String(r[1] || '');
    const key = lineId || name;
    if (!byPerson[key]) {
      byPerson[key] = { name: name, order: emp && emp.dispOrder !== '' ? emp.dispOrder : 9999, records: [], totalMin: 0 };
    }
    const inTime = String(r[3] || ''), outTime = String(r[4] || '');
    const breakMin = Number(r[5] || 0);
    const leave = String(r[6] || '');
    let workMin = null;
    if (inTime && outTime) {
      const mins = timeToMin_(outTime) - timeToMin_(inTime) - breakMin;
      workMin = mins >= 0 ? mins : null;
    }
    if (workMin !== null) byPerson[key].totalMin += workMin;
    byPerson[key].records.push({
      date: date, inTime: inTime, outTime: outTime, breakMin: breakMin, leave: leave,
      rawIn: String(r[10] || ''), rawOut: String(r[11] || ''), workMin: workMin,
      storeId: String(r[13] || ''), storeName: storeNameOf[String(r[13] || '')] || '',
      open: !!inTime && !outTime && !leave
    });
  }
  const people = Object.keys(byPerson).map(function (k) { return byPerson[k]; });
  people.forEach(function (p) { p.records.sort(function (a, b) { return a.date < b.date ? -1 : 1; }); });
  people.sort(function (a, b) { return (a.order - b.order) || (a.name < b.name ? -1 : 1); });
  return { month: month, people: people };
}

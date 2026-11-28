/*************************************************
* Mitsume Coffee 壊れ物報告システム（受け側）
* Google Apps Script
*
* 仕組み:
* - 画面は GitHub Pages から配信し、このスクリプトが受け口になる
* - 報告の送信は POST（doPost）で受ける。写真は Base64 で一緒に届く
* - LINE の IDトークンを検証して本人確認（シフト提出と同じチャネル）
* - 写真は Drive の専用フォルダに保存し、シートにはリンクを書く
* - 報告が入ったらオーナーにメールで通知する
* - 送信確認用に、クライアントトークンを CacheService に記録し、
* 画面側が JSONP（doGet?api=1&action=check）で到達確認できる
*************************************************/

const CONFIG = {
LINE_CHANNEL_ID: '2012734423',
NOTIFY_EMAIL: 'mitsume.coffee@gmail.com',
PHOTO_FOLDER_NAME: '壊れ物報告の写真',
SHEET_NAME: '報告一覧',
CUSTOMER_ID: 'M10001',
STORE_ID: 'S001',
MASTER_SPREADSHEET_ID: '1UowDnTtWjvcWEy0VBDpEQUMeP0SyZ5yDekpztmyWjZQ',
CATEGORIES: ['設備・機器', '水まわり', '電気', '食器・備品', '店内・建物', 'その他']
};

const TZ = 'Asia/Tokyo';

/* ---------- エントリポイント ---------- */

function doPost(e) {
let out;
try {
const payload = JSON.parse(e.postData.contents);
out = { ok: true, data: saveReport(payload)};
} catch (err) {
out = { ok: false, error: String(err && err.message? err.message: err)};
}
return ContentService.createTextOutput(JSON.stringify(out))
.setMimeType(ContentService.MimeType.JSON);
}

// 送信の到達確認（JSONP）とヘルスチェック
function doGet(e) {
const p = (e && e.parameter) || {};
if (p.api === '1' && p.action === 'access') {
// 開いた人のアクセス権チェック（画面の初期表示用）。開いたことは記録に残す
let body;
try {
const user = verifyIdToken(p.idToken || '');
logAccess_(user, p.name || '');
body = JSON.stringify({ ok: true, data: { allowed: isRegisteredStaff_(user.userId) } });
} catch (err) {
body = JSON.stringify({ ok: false, error: String(err && err.message ? err.message : err) });
}
if (p.callback) {
return ContentService.createTextOutput(p.callback + '(' + body + ')')
.setMimeType(ContentService.MimeType.JAVASCRIPT);
}
return ContentService.createTextOutput(body).setMimeType(ContentService.MimeType.JSON);
}
if (p.api === '1' && p.action === 'check' && p.token) {
const num = CacheService.getScriptCache().get('done_' + p.token);
const body = JSON.stringify({ ok: true, data: { saved:!!num, number: num || ''}});
if (p.callback) {
return ContentService.createTextOutput(p.callback + '(' + body + ')')
.setMimeType(ContentService.MimeType.JAVASCRIPT);
}
return ContentService.createTextOutput(body).setMimeType(ContentService.MimeType.JSON);
}
return ContentService.createTextOutput('report api ok');
}

/* ---------- 本人確認（シフト提出と共通） ---------- */

function verifyIdToken(idToken) {
if (!idToken) throw new Error('ログイン情報が取得できませんでした。LINEアプリから開き直してください。');
const res = UrlFetchApp.fetch('https://api.line.me/oauth2/v2.1/verify', {
method: 'post',
payload: { id_token: idToken, client_id: CONFIG.LINE_CHANNEL_ID},
muteHttpExceptions: true
});
if (res.getResponseCode()!== 200) {
throw new Error('本人確認に失敗しました。もう一度開き直してください。');
}
const j = JSON.parse(res.getContentText());
return { userId: j.sub, name: j.name || ''};
}

/* ---------- アクセス権（共通マスターの従業員マスターに Line_ID 登録がある人だけ） ---------- */

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

/* ---------- アクセス記録（開いた人を残す：管理画面の「候補」に出すため） ---------- */

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

/* ---------- 報告の保存 ---------- */

function ensureSheet() {
const ss = SpreadsheetApp.getActiveSpreadsheet();
let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
if (!sheet) {
sheet = ss.insertSheet(CONFIG.SHEET_NAME);
sheet.appendRow(['受付番号', '報告日時', '報告者', 'ユーザーID', 'カテゴリ', '何が壊れたか', '症状・状況', '緊急度', '写真', 'ステータス', '顧客ID', '店舗ID']);
sheet.setFrozenRows(1);
} else if (!sheet.getRange(1, 11).getValue()) {
sheet.getRange(1, 11).setValue('顧客ID');
sheet.getRange(1, 12).setValue('店舗ID');
}
return sheet;
}

function photoFolder() {
const it = DriveApp.getFoldersByName(CONFIG.PHOTO_FOLDER_NAME);
if (it.hasNext()) return it.next();
return DriveApp.createFolder(CONFIG.PHOTO_FOLDER_NAME);
}

function savePhoto(dataUrl, number) {
const m = /^data:image\/(png|jpe?g);base64,(.+)$/.exec(dataUrl || '');
if (!m) return '';
const ext = m[1].replace('jpeg', 'jpg');
const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), 'image/' + (m[1] === 'jpg'? 'jpeg': m[1]), number + '.' + ext);
const file = photoFolder().createFile(blob);
file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
return file.getUrl();
}

function saveReport(p) {
const user = verifyIdToken(p.idToken || '');
requireAccess_(user);
const category = CONFIG.CATEGORIES.indexOf(p.category) >= 0? p.category: 'その他';
const item = String(p.item || '').trim();
const detail = String(p.detail || '').trim();
const urgency = p.urgency === 'high'? '営業に支障あり': '急がない';
if (!item) throw new Error('「何が壊れたか」が未入力です。');

const lock = LockService.getScriptLock();
lock.waitLock(20000);
try {
const sheet = ensureSheet();
const now = new Date();
const dayKey = Utilities.formatDate(now, TZ, 'yyyyMMdd');
// その日の連番
const rows = sheet.getDataRange().getValues();
let seq = 1;
for (let i = 1; i < rows.length; i++) {
if (String(rows[i][0]).indexOf('R-' + dayKey + '-') === 0) seq++;
}
const number = 'R-' + dayKey + '-' + ('0' + seq).slice(-2);
const photoUrl = p.photo? savePhoto(p.photo, number): '';
const name = p.name || user.name || '';
sheet.appendRow([
number,
Utilities.formatDate(now, TZ, 'yyyy-MM-dd HH:mm:ss'),
name, user.userId, category, item, detail, urgency,
photoUrl, '未対応', CONFIG.CUSTOMER_ID, CONFIG.STORE_ID
]);

// オーナーへ通知メール
const ssUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
const subject = '' + item + '（' + urgency + '）';
const body =
'壊れ物の報告が入りました。\n\n' +
'受付番号: ' + number + '\n' +
'報告者: ' + name + '\n' +
'日時: ' + Utilities.formatDate(now, TZ, 'yyyy-MM-dd HH:mm') + '\n' +
'カテゴリ: ' + category + '\n' +
'何が壊れたか: ' + item + '\n' +
'症状・状況: ' + detail + '\n' +
'緊急度: ' + urgency + '\n' +
(photoUrl? '写真: ' + photoUrl + '\n': '写真: なし\n') +
'\n一覧（ステータス管理）: ' + ssUrl;
MailApp.sendEmail(CONFIG.NOTIFY_EMAIL, subject, body);

if (p.clientToken) {
CacheService.getScriptCache().put('done_' + p.clientToken, number, 600);
}
return { number: number};
} finally {
lock.releaseLock();
}
}

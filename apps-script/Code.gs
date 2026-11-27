/**
 * 식군스인테리어 홈페이지 서버 (Google Apps Script) — sikgoons.kr
 *
 * 하는 일
 *  1) 관리자 페이지에서 고친 내용(문구·현장·사진 등)을 이 시트의 '내용' 탭에 저장하고, 홈페이지가 읽어 갑니다.
 *  2) 저장할 때마다 '기록' 탭에 바로 전 내용을 남겨 언제든 되돌릴 수 있습니다.
 *  3) 관리자 페이지에서 올린 사진은 드라이브 '식군스 홈페이지 사진' 폴더에 저장합니다.
 *  4) 홈페이지 상담 신청을 '상담접수' 탭에 저장합니다. (카톡·문자 알림은 DNA LABS 상담알림 허브가 보냅니다)
 *
 * 처음 한 번: 위쪽 함수 목록에서 setup 을 골라 ▶실행 → 권한 허용
 *   - 시트 탭이 만들어지고, 첫 관리자 비밀번호가 '관리자 안내' 탭 B3 칸에 적힙니다.
 *   - 고객이 관리자 페이지 [비밀번호 설정]에서 비밀번호를 바꾸면 그 칸은 저절로 지워집니다.
 * 그다음: 배포 > 새 배포 > 웹 앱 / 실행: 나 / 액세스: 모든 사용자 → 나온 주소(…/exec)를 홈페이지에 넣습니다.
 *
 * 카톡·문자 알림은 이 스크립트가 아니라 DNA LABS 상담알림 허브가 보냅니다(받을 번호는 관리자 페이지 [상담 알림]).
 */

var DATA_SHEET = '내용', LOG_SHEET = '기록', CONTACT_SHEET = '상담접수', INFO_SHEET = '관리자 안내';
var CHUNK = 40000, MAX_CELLS = 20;        // 내용이 길면 여러 칸(A2~A21)에 나눠 저장
var MAX_FAILS = 8;                       // 10분 동안 이만큼 틀리면 잠시 막음
var MAX_UPLOAD_MB = 8;
var HEAD = ['접수 시각', '성함', '연락처', '통화 가능 시간', '단지명', '평형', '희망 시기', '예산', '메시지', '마케팅 동의', '접수 경로', '알림 결과', '처리', '번호'];

/* ---------------- 처음 설정 ---------------- */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('이 스크립트는 구글 시트의 [확장 프로그램 > Apps Script]에서 만들어야 합니다.');
  var data = ss.getSheetByName(DATA_SHEET) || ss.insertSheet(DATA_SHEET);
  var log = ss.getSheetByName(LOG_SHEET) || ss.insertSheet(LOG_SHEET);
  var info = ss.getSheetByName(INFO_SHEET) || ss.insertSheet(INFO_SHEET, 0);
  contactSheet_();
  data.getRange('A1').setValue('홈페이지 내용 (직접 고치지 마세요. 관리자 페이지에서 고칩니다)');
  if (log.getLastRow() === 0) log.appendRow(['저장 시각', '메모', '저장 전 내용(되돌리기용)']);
  log.getRange('A:A').setNumberFormat('@');
  var props = PropertiesService.getScriptProperties();
  info.getRange('A1').setValue('식군스인테리어 홈페이지 관리자 안내').setFontWeight('bold');
  info.getRange('A2').setValue('관리자 페이지');
  info.getRange('B2').setValue('https://sikgoons.kr/admin');
  info.getRange('A3').setValue('첫 비밀번호');
  if (!props.getProperty('ADMIN_HASH')) {
    var pw = randomPw_();
    setPw_(pw);
    info.getRange('B3').setValue(pw);
    info.getRange('C3').setValue('← 고객에게 전달한 뒤, 고객이 [비밀번호 설정]에서 바꾸면 이 칸은 저절로 지워집니다.');
  }
  info.setColumnWidth(1, 120); info.setColumnWidth(2, 260);
  photoFolder_();
  return '준비 완료';
}

/* ---------------- 시트 메뉴: [관리자] → 비밀번호 새로 정하기 ----------------
 * 고객이 관리자 비밀번호를 잊었을 때 제작사가 이 시트에서 새 비밀번호를 정해 줍니다.
 * 새 비밀번호는 시트에 적지 않습니다. 정한 뒤 고객에게 직접 알려 주세요. */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('관리자').addItem('비밀번호 새로 정하기', 'resetPasswordFromSheet').addToUi();
}
function resetPasswordFromSheet() {
  var ui = SpreadsheetApp.getUi();
  var r1 = ui.prompt('관리자 비밀번호 새로 정하기', '새 비밀번호를 입력하세요 (8자 이상, 영어·숫자 권장).', ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  var p1 = String(r1.getResponseText() || '').trim();
  if (p1.length < 8) { ui.alert('8자 이상으로 정해 주세요. 비밀번호는 바뀌지 않았습니다.'); return; }
  if (/[ㄱ-ㅎㅏ-ㅣ가-힣]/.test(p1)) { ui.alert('한글이 섞여 있어요. 한/영 키를 눌러 영어로 바꾼 뒤 다시 해 주세요. 비밀번호는 바뀌지 않았습니다.'); return; }
  var r2 = ui.prompt('한 번 더 확인', '같은 비밀번호를 한 번 더 입력하세요.', ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK) return;
  if (String(r2.getResponseText() || '').trim() !== p1) { ui.alert('두 번 입력한 비밀번호가 서로 달라요. 비밀번호는 바뀌지 않았습니다.'); return; }
  setPw_(p1);
  CacheService.getScriptCache().remove('fails');          // 여러 번 틀려서 잠긴 상태도 풀기
  var info = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(INFO_SHEET);
  if (info) {
    info.getRange('A3').setValue('비밀번호');
    info.getRange('B3').setValue('(시트 메뉴로 새로 정함 · 여기에 적지 않음)');
    info.getRange('C3').setValue(Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm') + ' 변경');
  }
  ui.alert('관리자 비밀번호를 바꿨습니다. 새 비밀번호를 고객에게 직접 알려 주세요. (이전 비밀번호는 이제 쓸 수 없습니다)');
}

/* ---------------- 받는 곳 ---------------- */
function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || 'content';
  if (action === 'content') return out_(readContent_());
  return out_({ ok: true, service: 'sikgoons' });
}

function doPost(e) {
  var req = {};
  try { req = JSON.parse((e && e.postData && e.postData.contents) || '{}'); }
  catch (err) { return out_({ ok: false, err: 'bad_request', error: '요청 형식이 올바르지 않습니다.' }); }

  // 홈페이지 상담 신청은 비밀번호 없이 받습니다 (action 이 없으면 상담 신청으로 봄: 예전 형식 호환)
  if (!req.action || req.action === 'contact') return out_(contact_(req));

  var gate = checkPw_(req.password);
  if (gate) return out_({ ok: false, error: gate });
  try {
    switch (req.action) {
      case 'login': return out_({ ok: true });
      case 'save': return out_(save_(req.content, req.note));
      case 'history': return out_({ ok: true, items: history_() });
      case 'restore': return out_(restore_(req.time));
      case 'upload': return out_(upload_(req));
      case 'setpw': return out_(changePw_(req.newPassword));
      case 'inbox': return out_({ ok: true, items: inbox_() });
      case 'inboxDone': return out_(inboxSet_(req.id, req.done));
      case 'inboxDel': return out_(inboxDel_(req.id));
      // 관리자 페이지 [수정 요청] → 접수 화면에 비밀번호 없이 바로 들어가기 (값은 스크립트 속성 REQ_PW 에만 둔다)
      case 'reqpw': return out_({ ok: true, pw: PropertiesService.getScriptProperties().getProperty('REQ_PW') || '' });
      default: return out_({ ok: false, error: '알 수 없는 요청입니다.' });
    }
  } catch (err) {
    return out_({ ok: false, error: String(err.message || err) });
  }
}

/* ---------------- 비밀번호 ---------------- */
function randomPw_() {
  var c = 'abcdefghjkmnpqrstuvwxyz23456789', s = '';
  for (var i = 0; i < 10; i++) s += c.charAt(Math.floor(Math.random() * c.length));
  return s;
}
function hashPw_(salt, pw) { return hex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + ':' + pw, Utilities.Charset.UTF_8)); }
function setPw_(pw) {
  var salt = Utilities.getUuid();
  PropertiesService.getScriptProperties().setProperties({ ADMIN_SALT: salt, ADMIN_HASH: hashPw_(salt, pw) });
}
function checkPw_(pw) {
  var cache = CacheService.getScriptCache();
  var fails = Number(cache.get('fails') || 0);
  if (fails >= MAX_FAILS) return '비밀번호를 여러 번 틀려서 10분 동안 잠겼습니다. 잠시 후 다시 시도해 주세요.';
  var p = PropertiesService.getScriptProperties();
  var salt = p.getProperty('ADMIN_SALT'), h = p.getProperty('ADMIN_HASH');
  if (!h) return '관리자 비밀번호가 아직 설정되지 않았습니다. 제작사에 문의해 주세요.';
  if (hashPw_(salt, String(pw || '')) !== h) { cache.put('fails', String(fails + 1), 600); return '비밀번호가 맞지 않습니다.'; }
  return '';
}
function changePw_(np) {
  np = String(np || '');
  if (np.length < 8) return { ok: false, error: '새 비밀번호는 8자 이상으로 정해 주세요.' };
  setPw_(np);
  var info = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(INFO_SHEET);
  if (info) { info.getRange('B3').setValue('(고객이 바꿨습니다)'); info.getRange('C3').setValue(Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm') + ' 변경'); }
  return { ok: true };
}

/* ---------------- 내용 저장·읽기 ---------------- */
function readRaw_() {
  var d = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DATA_SHEET);
  var v = d.getRange(2, 1, MAX_CELLS, 1).getValues();
  return v.map(function (r) { return String(r[0] || ''); }).join('');
}
function writeRaw_(text) {
  var d = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DATA_SHEET);
  var cells = [];
  for (var i = 0; i < MAX_CELLS; i++) cells.push([text.slice(i * CHUNK, (i + 1) * CHUNK)]);
  d.getRange(2, 1, MAX_CELLS, 1).setNumberFormat('@').setValues(cells);
}
function readContent_() {
  var cache = CacheService.getScriptCache();
  var hit = cache.get('content');
  if (hit) return JSON.parse(hit);
  var raw = readRaw_();
  var c = raw ? JSON.parse(raw) : {};
  var t = JSON.stringify(c);
  if (t.length < 90000) cache.put('content', t, 300);
  return c;
}
function save_(content, note) {
  if (!content || typeof content !== 'object' || Array.isArray(content)) return { ok: false, error: '저장할 내용이 비어 있습니다.' };
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return { ok: false, error: '다른 저장이 진행 중입니다. 잠시 후 다시 눌러 주세요.' };
  try {
    content.updated = new Date().toISOString();
    var text = JSON.stringify(content);
    if (text.length > CHUNK * MAX_CELLS) return { ok: false, error: '내용이 너무 깁니다. 제작사에 문의해 주세요.' };
    var before = readRaw_();
    writeRaw_(text);
    var log = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET);
    log.appendRow([content.updated, String(note || '저장').slice(0, 200), before.length < 49000 ? before : '(너무 길어 생략)']);
    var rows = log.getLastRow() - 1;
    if (rows > 100) log.deleteRows(2, rows - 100);
    CacheService.getScriptCache().remove('content');
    return { ok: true, updated: content.updated };
  } finally { lock.releaseLock(); }
}
function history_() {
  var log = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET);
  var n = log.getLastRow() - 1;
  if (n <= 0) return [];
  var take = Math.min(n, 20);
  var vals = log.getRange(log.getLastRow() - take + 1, 1, take, 2).getValues();
  return vals.reverse().map(function (r) { return { time: String(r[0]), note: String(r[1]) }; });
}
/** time 에 저장하기 '바로 전' 내용으로 되돌립니다 */
function restore_(time) {
  var log = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET);
  var n = log.getLastRow() - 1;
  if (n <= 0) return { ok: false, error: '되돌릴 기록이 없습니다.' };
  var vals = log.getRange(2, 1, n, 3).getValues();
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][0]) === String(time)) {
      var snap = String(vals[i][2] || '');
      if (snap.charAt(0) !== '{' && snap !== '') return { ok: false, error: '이 기록은 너무 길어서 되돌릴 수 없습니다.' };
      return save_(snap ? JSON.parse(snap) : {}, '되돌리기: ' + time);
    }
  }
  return { ok: false, error: '그 기록을 찾지 못했습니다.' };
}

/* ---------------- 사진 ---------------- */
function photoFolder_() {
  var p = PropertiesService.getScriptProperties(), id = p.getProperty('PHOTO_FOLDER');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  var f = DriveApp.createFolder('식군스 홈페이지 사진');
  p.setProperty('PHOTO_FOLDER', f.getId());
  return f;
}
function upload_(req) {
  var mime = String(req.mime || '');
  if (['image/jpeg', 'image/png', 'image/webp'].indexOf(mime) < 0) return { ok: false, error: '사진(JPG·PNG) 파일만 올릴 수 있습니다.' };
  var bytes = Utilities.base64Decode(String(req.data || ''));
  if (bytes.length > MAX_UPLOAD_MB * 1024 * 1024) return { ok: false, error: '사진이 너무 큽니다.' };
  var name = String(req.name || '사진').replace(/[\\/:*?"<>|]/g, '_').slice(0, 120);
  var file = photoFolder_().createFile(Utilities.newBlob(bytes, mime, name));
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { ok: true, id: file.getId() };
}

/* ---------------- 상담 신청 ---------------- */
/* 상담 신청서의 도면·참고 파일: 드라이브 '식군스 상담 첨부' 폴더 아래 접수마다 폴더를 만들어 저장 (링크 공개 안 함) */
var ATTACH_MAX_FILES = 10, ATTACH_MAX_MB = 25;
function attachFolder_() {
  var p = PropertiesService.getScriptProperties(), id = p.getProperty('ATTACH_FOLDER');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  var f = DriveApp.createFolder('식군스 상담 첨부');
  p.setProperty('ATTACH_FOLDER', f.getId());
  return f;
}
function saveAttach_(files, d, at) {
  if (!files || !files.length) return null;
  files = files.slice(0, ATTACH_MAX_FILES);
  var ok = ['image/jpeg', 'image/png', 'image/webp', 'image/heic', 'image/heif', 'application/pdf'], total = 0, blobs = [];
  files.forEach(function (f) {
    var mime = String(f && f.mime || '').toLowerCase();
    if (ok.indexOf(mime) < 0) return;
    var bytes = Utilities.base64Decode(String(f.data || ''));
    total += bytes.length;
    if (total > ATTACH_MAX_MB * 1024 * 1024) return;
    blobs.push(Utilities.newBlob(bytes, mime, String(f.name || '첨부').replace(/[\\/:*?"<>|]/g, '_').slice(0, 120)));
  });
  if (!blobs.length) return null;
  var sub = attachFolder_().createFolder(at.replace(':', '') + ' ' + clean_(d.nm).slice(0, 20) + ' ' + String(d.ph || '').replace(/[^0-9]/g, '').slice(-4));
  blobs.forEach(function (b) { sub.createFile(b); });
  var share = String(PropertiesService.getScriptProperties().getProperty('ATTACH_SHARE') || '').split(',').map(function (x) { return x.trim(); }).filter(Boolean);
  share.forEach(function (m) { try { sub.addViewer(m); } catch (e) {} });
  return { n: blobs.length, url: sub.getUrl() };
}
function contactSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(CONTACT_SHEET) || ss.insertSheet(CONTACT_SHEET);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEAD); sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEAD.length).setFontWeight('bold').setBackground('#F3F1EC');
  }
  return sh;
}
function contact_(d) {
  if (d.hp) return { ok: true };                                     // 로봇 거르기
  if (!clean_(d.nm) || !clean_(d.ph)) return { ok: false, err: 'missing' };
  if (d.t && Date.now() - Number(d.t) < 3000) return { ok: false, err: 'too_fast' };
  var cache = CacheService.getScriptCache();
  var dupKey = 'dup_' + hex_(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, clean_(d.ph) + '|' + clean_(d.ms), Utilities.Charset.UTF_8));
  if (cache.get(dupKey)) return { ok: true, dup: true };
  var hourKey = 'h_' + Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyyMMddHH');
  var n = Number(cache.get(hourKey) || 0);
  if (n >= 30) return { ok: false, err: 'rate' };
  cache.put(hourKey, String(n + 1), 3600); cache.put(dupKey, '1', 120);

  var at = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm');
  var att = null, files = d.files; delete d.files;
  try { att = saveAttach_(files, d, at); } catch (err) { att = { n: 0, url: '저장 실패' }; }
  if (att) d.ms = (d.ms ? d.ms + '\n' : '') + '· 도면·참고 파일 ' + att.n + '개: ' + att.url;

  var lock = LockService.getScriptLock(); lock.waitLock(10000);
  var sh = contactSheet_();
  var id = Utilities.getUuid().slice(0, 8);
  sh.appendRow([at, d.nm, d.ph, d.ct, d.apt, d.py, d.wh, d.bd, d.ms, d.agM ? '동의' : '', d.src, '보내는 중', '', id].map(safe_));
  var row = sh.getLastRow();
  lock.releaseLock();
  sh.getRange(row, 12).setValue(notify_(d, at));
  return { ok: true, att: att };
}
function inbox_() {
  var sh = contactSheet_(), n = sh.getLastRow() - 1;
  if (n <= 0) return [];
  var v = sh.getRange(2, 1, n, HEAD.length).getDisplayValues();
  return v.reverse().slice(0, 300).map(function (r) {
    return { id: r[13], at: r[0], done: !!r[12], d: { nm: r[1], ph: r[2], ct: r[3], apt: r[4], py: r[5], wh: r[6], bd: r[7], ms: r[8], agM: r[9] === '동의', src: r[10] }, note: r[11] };
  });
}
function findRow_(id) {
  var sh = contactSheet_(), n = sh.getLastRow() - 1;
  if (n <= 0 || !id) return 0;
  var ids = sh.getRange(2, 14, n, 1).getValues();
  for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return 0;
}
function inboxSet_(id, done) {
  var r = findRow_(id); if (!r) return { ok: false, error: '그 신청을 찾지 못했습니다.' };
  contactSheet_().getRange(r, 13).setValue(done ? '완료' : '');
  return { ok: true };
}
function inboxDel_(id) {
  var r = findRow_(id); if (!r) return { ok: false, error: '그 신청을 찾지 못했습니다.' };
  contactSheet_().deleteRow(r);
  return { ok: true };
}

/* ---------------- 알림톡 (솔라피) ---------------- */
function notify_(d, at) {
  var p = PropertiesService.getScriptProperties().getProperties();
  if (!p.SOLAPI_KEY || !p.SOLAPI_SECRET || !p.ALERT_TO) return '알림 설정 전';
  var v = {
    '#{고객명}': or_(d.nm), '#{연락처}': or_(d.ph), '#{통화가능시간}': or_(d.ct),
    '#{단지명}': or_(d.apt), '#{평형}': or_(d.py), '#{희망시기}': or_(d.wh),
    '#{예산}': or_(d.bd), '#{메시지}': or_(d.ms), '#{접수시각}': at
  };
  var msg = { to: digits_(p.ALERT_TO) };
  if (p.PFID && p.TEMPLATE_ID) {
    msg.kakaoOptions = { pfId: p.PFID, templateId: p.TEMPLATE_ID, variables: v, disableSms: !p.SMS_FROM };
    if (p.SMS_FROM) msg.from = digits_(p.SMS_FROM);
  } else if (p.SMS_FROM) {
    msg.from = digits_(p.SMS_FROM);
    msg.subject = '[식군스] 새 상담 신청';
    msg.text = '[식군스] 홈페이지 상담 신청\n성함: ' + v['#{고객명}'] + '\n연락처: ' + v['#{연락처}'] + '\n통화 가능: ' + v['#{통화가능시간}'] +
      '\n단지: ' + v['#{단지명}'] + ' / ' + v['#{평형}'] + '\n시기: ' + v['#{희망시기}'] + '\n예산: ' + v['#{예산}'] + '\n메시지: ' + v['#{메시지}'] + '\n접수: ' + at;
  } else return '알림 설정 전';
  var date = new Date().toISOString(), salt = Utilities.getUuid().replace(/-/g, '');
  var sig = hex_(Utilities.computeHmacSha256Signature(date + salt, p.SOLAPI_SECRET));
  var res = UrlFetchApp.fetch('https://api.solapi.com/messages/v4/send-many/detail', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { Authorization: 'HMAC-SHA256 apiKey=' + p.SOLAPI_KEY + ', date=' + date + ', salt=' + salt + ', signature=' + sig },
    payload: JSON.stringify({ messages: [msg] })
  });
  var code = res.getResponseCode();
  if (code >= 200 && code < 300) {
    var j = {}; try { j = JSON.parse(res.getContentText()); } catch (e) {}
    var fail = j.failedMessageList && j.failedMessageList.length;
    return fail ? '실패: ' + JSON.stringify(j.failedMessageList[0]).slice(0, 200) : (msg.kakaoOptions ? '알림톡 보냄' : '문자 보냄');
  }
  return '실패 ' + code + ': ' + res.getContentText().slice(0, 200);
}

/* ---------------- 도우미 ---------------- */
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function clean_(s) { return String(s == null ? '' : s).replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, '').trim().slice(0, 1000); }
function or_(s) { s = clean_(s); return s || '-'; }
function digits_(s) { return String(s).replace(/[^0-9]/g, ''); }
function safe_(s) { s = clean_(s); return /^[=+\-@]/.test(s) ? "'" + s : s; }
function hex_(b) { return b.map(function (x) { return ('0' + (x & 255).toString(16)).slice(-2); }).join(''); }


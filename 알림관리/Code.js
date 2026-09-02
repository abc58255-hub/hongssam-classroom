// 알림관리 앱 — 교사가 학급알림을 작성·관리하고 푸시 알림을 전송하는 도구
// 데이터·전송 백엔드는 모두 ClassCore에 위임 (학급알림 시트 = ALARM_SHEET_ID, 푸시 = FCM)
// 하드코딩 없음

// ── 라우팅 ──
function doGet() {
  try { ClassCore.registerAppUrl('알림관리', _serviceUrl_()); } catch(_) {}
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('알림 관리')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
function _serviceUrl_() {
  try { return ScriptApp.getService().getUrl() || ''; } catch(_) { return ''; }
}

// ── 교사 인증 (ClassCore 위임) ──
function teacherLogin(pw)            { return ClassCore.teacherLogin(pw); }
function teacherLogout(token)        { return ClassCore.teacherLogout(token); }
function validateTeacherSession(token) { return ClassCore.verifyTeacher(token); }

// ── 학급알림 (ClassCore 위임) ──
function getNotices()         { return ClassCore.getNotices(); }
function saveNotice(data)     { return ClassCore.saveNotice(data); }
function deleteNotice(rowIdx) { return ClassCore.deleteNotice(rowIdx); }

// ── 푸시 알림 전송 (ClassCore 위임) ──
// 대상: filterClass 지정 시 해당 반만, 없으면 전체. studentIds=null → 전체 대상.
function sendPush(title, body, filterClass, clickUrl) {
  try {
    return ClassCore.sendPush(null, title, body, 'notice', filterClass || '', clickUrl || '');
  } catch(e) { return { success: false, message: e.toString() }; }
}

// ── 반 목록 / 담임반 (대상 선택용) ──
function getClassInfo() {
  try {
    var data = ClassCore.getRosterValues();
    var set = {};
    for (var i = 1; i < data.length; i++) {
      var sid = String(data[i][1] || '').trim();
      if (sid.length >= 2) set[sid.substring(0, 1) + '학년 ' + sid.substring(1, 2) + '반'] = true;
    }
    var classes = Object.keys(set).sort();
    return { success: true, classes: classes, homeroom: ClassCore.getConfig('담임반', '') };
  } catch(e) { return { success: false, message: e.toString() }; }
}

// ══════════════════════════════════════════════════════════════
//  교실TV알림판(학급칠판) 공지 — 교실-알림판 바운드 시트의 '공지사항' 시트에 기록
//  시트 구조: A=날짜 B=조회공지 C=조회전달 D=종례공지 E=종례전달 F=슬라이드URL G=유튜브URL H=영상모드
//  알림판 앱은 13시 기준으로 조회/종례 자동 전환. 보드 시트 ID = ClassCore '알림판시트ID'
// ══════════════════════════════════════════════════════════════
function _boardSheetId() {
  var id = '';
  try { id = ClassCore.getConfig('알림판시트ID', ''); } catch(_) {}
  if (id) return id;
  // 마스터 시스템설정에서 1회 자동 이전 (마커 파일로 원본 탐색)
  try {
    var it = DriveApp.getFilesByName('홍쌤교실시스템_SHEET_ID');
    if (it.hasNext()) {
      var masterId = String(it.next().getBlob().getDataAsString() || '').trim();
      var sh = SpreadsheetApp.openById(masterId).getSheetByName('시스템설정');
      if (sh && sh.getLastRow() >= 2) {
        var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
        for (var i = 0; i < rows.length; i++) {
          if (String(rows[i][0]).trim() === '알림판시트ID') { id = String(rows[i][1] || '').trim(); break; }
        }
      }
      if (id) { try { ClassCore.setConfig('알림판시트ID', id); } catch(_) {} }
    }
  } catch(_) {}
  return id;
}

// 보드 ID 수동 등록 (자동 이전이 안 될 때 UI에서 직접 입력)
function setBoardSheetId(id) {
  try { ClassCore.setConfig('알림판시트ID', String(id || '').trim()); return { success: true }; }
  catch(e) { return { success: false, message: e.toString() }; }
}

function getBoardNotices() {
  try {
    var boardId = _boardSheetId();
    if (!boardId) return { success: false, needId: true, message: '교실 알림판 스프레드시트 ID가 설정되지 않았습니다.' };
    var sheet = SpreadsheetApp.openById(boardId).getSheetByName('공지사항');
    if (!sheet) return { success: false, message: '교실 알림판 스프레드시트에 "공지사항" 시트가 없습니다.' };
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: true, boards: [] };
    var data = sheet.getRange(2, 1, lastRow - 1, 8).getDisplayValues();
    var boards = [];
    for (var i = 0; i < data.length; i++) {
      var r = data[i];
      if (!r[0]) continue;
      boards.push({
        rowIdx:         i + 2,
        date:           String(r[0] || '').trim(),
        announcement:   String(r[1] || '').trim(),
        notice:         String(r[2] || '').trim(),
        announcementPm: String(r[3] || '').trim(),
        noticePm:       String(r[4] || '').trim(),
        slideUrl:       String(r[5] || '').trim(),
        videoUrl:       String(r[6] || '').trim(),
        videoMode:      String(r[7] || '소리만').trim()
      });
    }
    boards.sort(function(a, b) { return b.date.localeCompare(a.date); });
    // 기본 슬라이드 URL: 설정값 우선, 없으면 가장 최근 공지의 슬라이드를 기본으로
    var defaultSlide = '';
    try { defaultSlide = ClassCore.getConfig('기본슬라이드URL', ''); } catch(_) {}
    if (!defaultSlide) {
      for (var j = 0; j < boards.length; j++) { if (boards[j].slideUrl) { defaultSlide = boards[j].slideUrl; break; } }
    }
    return { success: true, boards: boards, defaultSlide: defaultSlide };
  } catch(e) { return { success: false, message: e.toString() }; }
}

// 기본 슬라이드 URL 저장 (이후 새 칠판공지에 자동 입력)
function setDefaultSlide(url) {
  try { ClassCore.setConfig('기본슬라이드URL', String(url || '').trim()); return { success: true }; }
  catch(e) { return { success: false, message: e.toString() }; }
}

function saveBoardData(data) {
  try {
    var boardId = _boardSheetId();
    if (!boardId) return { success: false, needId: true, message: '교실 알림판 스프레드시트 ID가 없습니다.' };
    var sheet = SpreadsheetApp.openById(boardId).getSheetByName('공지사항');
    if (!sheet) return { success: false, message: '"공지사항" 시트를 찾을 수 없습니다.' };
    // 슬라이드 URL 비어있으면 기본 슬라이드로 채움 (TV에 항상 슬라이드 연결)
    var slideUrl = data.slideUrl || '';
    if (!slideUrl) { try { slideUrl = ClassCore.getConfig('기본슬라이드URL', ''); } catch(_) {} }
    var row = [
      data.date            || '',
      data.announcement    || '',
      data.notice          || '',
      data.announcementPm  || '',
      data.noticePm        || '',
      slideUrl,
      data.videoUrl        || '',
      data.videoMode       || '소리만'
    ];
    if (data.rowIdx && data.rowIdx > 1) {
      sheet.getRange(data.rowIdx, 1, 1, 8).setValues([row]);
    } else {
      sheet.getRange(Math.max(sheet.getLastRow(), 1) + 1, 1, 1, 8).setValues([row]);
    }
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}

function deleteBoardData(rowIdx) {
  try {
    var boardId = _boardSheetId();
    if (!boardId) return { success: false, message: 'ID 없음' };
    SpreadsheetApp.openById(boardId).getSheetByName('공지사항').deleteRow(rowIdx);
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}


// ── PWA(Pages) 프론트 → GAS 백엔드 (google.script.run 어댑터) + 토큰 게이트 ──
var RPC_WHITELIST = ["deleteBoardData", "deleteNotice", "getBoardNotices", "getClassInfo", "getNotices", "saveBoardData", "saveNotice", "sendPush", "setBoardSheetId", "setDefaultSlide", "sendTvFlash", "clearTvFlash", "getTvFlash", "teacherLogin", "teacherLogout", "validateTeacherSession"];
var RPC_NOAUTH = ['teacherLogin','teacherLogout','getTvFlash']; // getTvFlash=TV 화면 폴링(공개)
function _atOk_(token) { try { var v = ClassCore.verifyTeacher(token); return (v === true) || !!(v && (v.success || v.valid || v.ok)); } catch(_) { return false; } }
function doPost(e) {
  var out;
  try {
    var req = JSON.parse(e.postData.contents);
    if (RPC_WHITELIST.indexOf(req.fn) < 0) throw new Error('허용되지 않은 함수: ' + req.fn);
    if (RPC_NOAUTH.indexOf(req.fn) < 0 && !_atOk_(req.token)) {
      return ContentService.createTextOutput(JSON.stringify({ ok: false, error: '로그인이 필요해요', needAuth: true })).setMimeType(ContentService.MimeType.JSON);
    }
    var fn = globalThis[req.fn];
    if (typeof fn !== 'function') throw new Error('함수를 찾을 수 없음: ' + req.fn);
    out = { ok: true, result: fn.apply(null, req.args || []) };
  } catch (err) { out = { ok: false, error: String(err && err.message || err) }; }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

// ── 📺 TV 실시간 알림(팝업) ── 교실TV알림판 스프레드시트의 '_TV_FLASH' 탭에 저장(교실-알림판 GAS가 same-origin으로 읽음)
// 인증: doPost 게이트(_atOk_)가 send/clear 보호. getTvFlash는 RPC_NOAUTH(공개).
var TV_FLASH_TAB = '_TV_FLASH';
function _tvFlashSheet_() {
  var id = _boardSheetId();
  if (!id) return null;
  var ss = SpreadsheetApp.openById(id);
  var sh = ss.getSheetByName(TV_FLASH_TAB);
  if (!sh) { sh = ss.insertSheet(TV_FLASH_TAB); sh.getRange('A1:B1').setValues([['id', 'msg']]); }
  return sh;
}
function sendTvFlash(msg) {
  var text = String(msg == null ? '' : msg).trim();
  if (!text) return { success: false, message: '보낼 내용을 입력하세요.' };
  var sh = _tvFlashSheet_();
  if (!sh) return { success: false, message: '알림판 시트가 연결되지 않았어요. (칠판공지 설정을 먼저 해주세요)' };
  sh.getRange('A2:B2').setValues([[Date.now(), text]]);
  return { success: true };
}
function clearTvFlash() {
  var sh = _tvFlashSheet_();
  if (!sh) return { success: false, message: '알림판 시트가 연결되지 않았어요.' };
  sh.getRange('A2:B2').setValues([[Date.now(), '']]); // 빈 메시지 + 새 id → TV가 팝업 닫음
  return { success: true };
}
function getTvFlash() {
  try {
    var sh = _tvFlashSheet_();
    if (!sh) return { id: 0, msg: '' };
    var v = sh.getRange('A2:B2').getValues()[0];
    return { id: Number(v[0]) || 0, msg: String(v[1] || '') };
  } catch (_) { return { id: 0, msg: '' }; }
}

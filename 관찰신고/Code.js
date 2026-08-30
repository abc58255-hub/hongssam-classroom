// 관찰·신고 관리 앱 — 교사가 학생 관찰을 기록하고 학생 신고를 검토하는 도구
// 데이터: 전용 "관찰·신고 시트"(학생관찰기록·신고접수)
// 공통(학생 인증·설정·교사 인증)은 ClassCore 라이브러리에 위임 — 하드코딩 없음
//
// 컬럼 구조(홍쌤 교실 시스템 호환, 데이터 복붙 가능):
//   학생관찰기록 : 날짜·학번·이름·카테고리·내용·구분(담임학급/수학수업)
//   신고접수     : 접수일시·학번·이름·유형·내용·처리여부

// ── 관찰·신고 시트 (전용) ───────────────────────────────────
var _cachedSs = null;

function _obsSs() {
  if (_cachedSs) return _cachedSs;
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('OBS_SHEET_ID');
  var ss;
  if (id) { try { ss = SpreadsheetApp.openById(id); } catch(_) {} }
  if (!ss) {
    ss = SpreadsheetApp.create('관찰·신고 데이터');
    props.setProperty('OBS_SHEET_ID', ss.getId());
  }
  _ensureSheets(ss);
  // 관찰·신고 시트 ID를 ClassCore에 공유 → 학생-담임 등 다른 앱이 같은 시트를 참조
  try { if (ClassCore.getConfig('관찰신고시트ID', '') !== ss.getId()) ClassCore.setConfig('관찰신고시트ID', ss.getId()); } catch(_) {}
  _cachedSs = ss;
  return ss;
}

function _ensureSheets(ss) {
  function mk(name, headers, color) {
    if (ss.getSheetByName(name)) return;
    var sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold').setBackground(color).setFontColor('white');
    sh.setFrozenRows(1);
  }
  mk('학생관찰기록', ['날짜','학번','이름','카테고리','내용','구분'], '#10b981');
  mk('신고접수', ['접수일시','학번','이름','유형','내용','처리여부'], '#ef4444');
  var def = ss.getSheetByName('시트1') || ss.getSheetByName('Sheet1');
  if (def && ss.getSheets().length > 1) { try { ss.deleteSheet(def); } catch(_) {} }
}

// ── 라우팅 ──
function doGet() {
  ClassCore.registerAppUrl('관찰신고', _serviceUrl_());
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('관찰·신고 관리')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function _serviceUrl_() {
  try { return ScriptApp.getService().getUrl() || ''; } catch(_) { return ''; }
}

// ── 교사 인증 (ClassCore 위임) ──
function teacherLogin(pw)            { return ClassCore.teacherLogin(pw); }
function verifyTeacher(token)        { return ClassCore.verifyTeacher(token); }
function teacherLogout(token)        { return ClassCore.teacherLogout(token); }
function validateTeacherSession(token) { return ClassCore.verifyTeacher(token); }

// 학생 명부 (반별·관찰대상 선택용) — ClassCore 위임
function getRoster() {
  try {
    return { success: true, roster: ClassCore.getRosterValues(), homeroom: ClassCore.getConfig('담임반', '') };
  } catch(e) { return { success: false, message: e.toString() }; }
}

// ── 관찰 기록 (교사용) ──
function _observeSheet() {
  return _obsSs().getSheetByName('학생관찰기록');
}

function getObserveRecords() {
  try {
    var sheet = _observeSheet();
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: true, records: [] };
    var data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
    var records = [];
    for (var i = 0; i < data.length; i++) {
      var r = data[i];
      if (!r[1] && !r[4]) continue;
      records.push({
        rowIdx:      i + 2,
        date:        r[0] instanceof Date ? Utilities.formatDate(r[0], 'Asia/Seoul', 'yyyy-MM-dd') : String(r[0] || '').trim(),
        studentId:   String(r[1] || '').trim(),
        studentName: String(r[2] || '').trim(),
        category:    String(r[3] || '').trim(),
        content:     String(r[4] || '').trim(),
        type:        String(r[5] || '담임학급').trim()
      });
    }
    records.sort(function(a, b) { return b.date.localeCompare(a.date); });
    return { success: true, records: records };
  } catch(e) { return { success: false, message: e.toString() }; }
}

function saveObserveRecord(data) {
  try {
    var sheet = _observeSheet();
    var row = [
      data.date        || '',
      data.studentId   || '',
      data.studentName || '',
      data.category    || '',
      data.content     || '',
      data.type        || '담임학급'
    ];
    if (data.rowIdx && data.rowIdx > 1) {
      sheet.getRange(data.rowIdx, 1, 1, 6).setValues([row]);
    } else {
      sheet.getRange(Math.max(sheet.getLastRow(), 1) + 1, 1, 1, 6).setValues([row]);
    }
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}

// 여러 명의 관찰 기록을 한 번에 저장 (다중 선택)
function saveMultipleObserveRecords(dataArray) {
  try {
    var sheet = _observeSheet();
    var rows = dataArray.map(function(data) {
      return [
        data.date || '', data.studentId || '', data.studentName || '',
        data.category || '', data.content || '', data.type || '담임학급'
      ];
    });
    var lastRow = Math.max(sheet.getLastRow(), 1);
    sheet.getRange(lastRow + 1, 1, rows.length, 6).setValues(rows);
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}

function deleteObserveRecord(rowIdx) {
  try {
    if (rowIdx > 1) { _observeSheet().deleteRow(rowIdx); return { success: true }; }
    return { success: false, message: '삭제할 기록을 찾을 수 없습니다.' };
  } catch(e) { return { success: false, message: e.toString() }; }
}

// ── 신고 접수함 (교사용) ──
function getReportInbox() {
  try {
    var sheet = _obsSs().getSheetByName('신고접수');
    if (!sheet || sheet.getLastRow() < 2) return { success: true, reports: [] };
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
    var reports = [];
    for (var i = 0; i < data.length; i++) {
      var r = data[i];
      if (!r[4]) continue;
      reports.push({
        rowIdx:      i + 2,
        date:        r[0] instanceof Date ? Utilities.formatDate(r[0], 'Asia/Seoul', 'MM/dd HH:mm') : String(r[0] || '').trim(),
        studentId:   String(r[1] || '').trim(),
        studentName: String(r[2] || '').trim(),
        category:    String(r[3] || '').trim(),
        content:     String(r[4] || '').trim(),
        status:      String(r[5] || '미처리').trim()
      });
    }
    reports.sort(function(a, b) { return b.date.localeCompare(a.date); });
    return { success: true, reports: reports };
  } catch(e) { return { success: false, message: e.toString() }; }
}

function markReportDone(rowIdx) {
  try {
    _obsSs().getSheetByName('신고접수').getRange(rowIdx, 6).setValue('처리완료');
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}


// ── PWA(Pages) 프론트 → GAS 백엔드 (google.script.run 어댑터) + 토큰 게이트 ──
var RPC_WHITELIST = ["deleteObserveRecord", "getObserveRecords", "getReportInbox", "getRoster", "markReportDone", "saveMultipleObserveRecords", "saveObserveRecord", "teacherLogin", "teacherLogout", "validateTeacherSession"];
var RPC_NOAUTH = ['teacherLogin','teacherLogout'];
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

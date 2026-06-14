// 과제·채점 앱 — 교사가 수학 과제를 출제·채점하는 도구
// 데이터: 전용 "과제 시트"(과제설정·제출현황·AI채점기준)
// 공통(학생 인증·설정·키·FCM·교사 인증)은 ClassCore 라이브러리에 위임 — 하드코딩 없음
//
// 컬럼 구조(홍쌤 교실 시스템 호환, 데이터 복붙 가능):
//   과제설정   : 날짜·과제명·설명·마감일JSON·채점유형·공개여부·사진수·선택지
//   제출현황   : 날짜·학번·이름·과제명·난이도·메모·파일URL·피드백·해시·읽음·상태·첨삭URL·점수·
//                공개여부·메모2·답글·우수작유형·익명·우수작멘트·우수작키·부정행위·문항별JSON·AI채점JSON·상태변경일시
//   AI채점기준 : 과제명·채점유형·만점·기준·파일JSON·문항JSON

// ── 과제 시트 (전용) ───────────────────────────────────────
var _cachedTaskSs = null;

function _taskSs() {
  if (_cachedTaskSs) return _cachedTaskSs;
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('TASK_SHEET_ID');
  var ss;
  if (id) { try { ss = SpreadsheetApp.openById(id); } catch(_) {} }
  if (!ss) {
    ss = SpreadsheetApp.create('과제·채점 데이터');
    props.setProperty('TASK_SHEET_ID', ss.getId());
  }
  _ensureTaskSheets(ss);
  _cachedTaskSs = ss;
  return ss;
}

function _ensureTaskSheets(ss) {
  function mk(name, headers, color) {
    if (ss.getSheetByName(name)) return;
    var sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold').setBackground(color).setFontColor('white');
    sh.setFrozenRows(1);
  }
  mk('과제설정', ['날짜','과제명','설명','마감일','채점유형','공개여부','사진수','선택지'], '#1e3a8a');
  mk('제출현황', ['날짜','학번','이름','과제명','난이도','메모','파일URL','피드백','해시','읽음여부','상태','첨삭URL','점수','공개여부','메모2','답글','우수작유형','익명여부','우수작멘트','우수작키','부정행위','문항별JSON','AI채점JSON','상태변경일시'], '#1e3a8a');
  mk('AI채점기준', ['과제명','채점유형','만점','기준','파일JSON','문항JSON'], '#1e3a8a');
  var def = ss.getSheetByName('시트1') || ss.getSheetByName('Sheet1');
  if (def && ss.getSheets().length > 1) { try { ss.deleteSheet(def); } catch(_) {} }
}

// 과제 데이터용 드라이브 폴더 (제출 파일·첨삭 이미지) — 폴더ID는 ClassCore 공통 설정 사용
function _parentFolder() {
  var id = ClassCore.getConfig('드라이브폴더ID', '');
  if (id) { try { return DriveApp.getFolderById(id); } catch(_) {} }
  var it = DriveApp.getFoldersByName('홍쌤 교실 시스템');
  var folder = it.hasNext() ? it.next() : DriveApp.createFolder('홍쌤 교실 시스템');
  ClassCore.setConfig('드라이브폴더ID', folder.getId());
  return folder;
}

// ── 라우팅 ──
function doGet() {
  ClassCore.registerAppUrl('과제채점', _serviceUrl_());
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('과제·채점')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function _serviceUrl_() {
  try { return ScriptApp.getService().getUrl() || ''; } catch(_) { return ''; }
}

// ── 교사 인증 (ClassCore 위임) ──
function teacherLogin(pw)        { return ClassCore.teacherLogin(pw); }
function verifyTeacher(token)    { return ClassCore.verifyTeacher(token); }
function teacherLogout(token)    { return ClassCore.teacherLogout(token); }

// 최초 1회: 과제 시트 생성 확인 (편집기에서 실행)
function setup() {
  var ss = _taskSs();
  Logger.log('✅ 과제 시트 준비 완료: ' + ss.getUrl());
  return ss.getUrl();
}

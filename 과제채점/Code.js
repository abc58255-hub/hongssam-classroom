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
// grading.js 호환 — 폴더 ID 반환
function _getParentFolderId_() { return _parentFolder().getId(); }

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
function validateTeacherSession(token) { return ClassCore.verifyTeacher(token); }

// ── 초기 데이터 로드 (과제·제출·명부 — 비과제 필드는 빈 배열) ──
function getDashboardData() {
  try {
    var ss = _taskSs();
    var cache = CacheService.getScriptCache();
    var hrStr = ClassCore.getConfig('담임반', '');
    var defaultSlideUrl = ClassCore.getConfig('기본슬라이드URL', '');

    var roster, classList;
    var cr = cache.get('roster_v1');
    if (cr) { var p = JSON.parse(cr); roster = p.roster; classList = p.classList; }
    else {
      var r = parseRosterAndClasses(ss);
      roster = r.roster; classList = r.classList;
      try { cache.put('roster_v1', JSON.stringify({ roster: roster, classList: classList }), 60); } catch(e) {}
    }

    var tasks;
    var ct = cache.get('tasks_v1');
    if (ct) { try { tasks = JSON.parse(ct); } catch(e) { tasks = parseTasks(ss); } }
    else { tasks = parseTasks(ss); try { cache.put('tasks_v1', JSON.stringify(tasks), 30); } catch(e) {} }

    var submissions = parseSubmissions(ss);

    return {
      roster: roster, classList: classList, tasks: tasks, submissions: submissions,
      homeroom: hrStr, defaultSlideUrl: defaultSlideUrl,
      alarmRegisterUrl: ClassCore.getConfig('FCM_REGISTER_URL', ''),
      hrActivities: [], hrSubmissions: [], surveys: [], surveyRes: []
    };
  } catch(e) { throw new Error('전체 데이터 로딩 오류: ' + e.message); }
}

// ── 명부·FCM 호환 헬퍼 (grading.js의 알림 함수가 ClassCore를 통해 동작) ──
// 학생명부 시트 객체 (ClassCore 공유 시트) — parseRosterAndClasses 등이 행 단위로 읽음
function _authRoster_() {
  return SpreadsheetApp.openById(ClassCore.getAuthSheetId()).getSheetByName('학생명부');
}

// 명부 E열 토큰 유무 판단용
function _parseTokens_(raw) {
  var s = String(raw || '').trim();
  if (!s) return [];
  try {
    var a = JSON.parse(s);
    if (!Array.isArray(a)) a = [s];
    return a.map(function(e) { return typeof e === 'string' ? e : (e && e.t ? e.t : null); }).filter(Boolean);
  } catch(_) { return [s]; }
}

// 한 학생 행에 푸시 (명부 B열=학번) → ClassCore.sendPush 위임 (토큰 정리는 ClassCore가 처리)
function _sendAndPrune_(sheet, rowIdx, title, body, clickUrl, tag) {
  try {
    var sid = String(sheet.getRange(rowIdx, 2).getValue() || '').trim();
    if (!sid) return false;
    var r = ClassCore.sendPush([sid], title, body, tag, null, clickUrl);
    return !!(r && r.success && r.sent > 0);
  } catch(_) { return false; }
}

// 최초 1회: 과제 시트 생성 확인 (편집기에서 실행)
function setup() {
  var ss = _taskSs();
  Logger.log('✅ 과제 시트 준비 완료: ' + ss.getUrl());
  return ss.getUrl();
}

// ── 데이터 이전 (기존 홍쌤 공유시트 → 과제 시트, 1회 실행) ──
// 과제설정·제출현황·AI채점기준을 그대로 복사. 과제 시트의 기존 데이터는 덮어쓴다.
function importFromBoard() {
  try {
    var it = DriveApp.getFilesByName('홍쌤교실시스템_SHEET_ID');
    if (!it.hasNext()) return { success: false, message: '홍쌤 연결 파일(마커)을 찾을 수 없습니다. importFromSheet("시트ID")로 직접 지정하세요.' };
    return importFromSheet(String(it.next().getBlob().getDataAsString() || '').trim());
  } catch(e) { return { success: false, message: e.toString() }; }
}

function importFromSheet(srcId) {
  try {
    var src = SpreadsheetApp.openById(srcId);
    var dst = _taskSs();
    var names = ['과제설정', '제출현황', 'AI채점기준'];
    var report = [];
    names.forEach(function(name) {
      var s = src.getSheetByName(name);
      var d = dst.getSheetByName(name);
      if (!s || s.getLastRow() < 2 || !d) { report.push(name + ': 0행(원본 없음/비어있음)'); return; }
      var cols = Math.min(s.getLastColumn(), d.getLastColumn()); // 대상 컬럼 초과 방지
      var data = s.getRange(2, 1, s.getLastRow() - 1, cols).getValues();
      if (d.getLastRow() > 1) d.getRange(2, 1, d.getLastRow() - 1, d.getLastColumn()).clearContent();
      d.getRange(2, 1, data.length, cols).setValues(data);
      report.push(name + ': ' + data.length + '행');
    });
    Logger.log('✅ 이전 완료 — ' + report.join(' / ') + '  (원본: ' + src.getName() + ')');
    return { success: true, report: report.join(' / ') };
  } catch(e) { return { success: false, message: e.toString() }; }
}

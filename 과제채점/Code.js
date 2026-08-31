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
  // 과제 시트 ID를 ClassCore에 공유 → 학생-수학 등 다른 앱이 같은 시트를 참조
  try { if (ClassCore.getConfig('과제시트ID', '') !== ss.getId()) ClassCore.setConfig('과제시트ID', ss.getId()); } catch(_) {}
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
  mk('과제설정', ['날짜','과제명','설명','마감일','채점유형','공개여부','사진수','선택지','만점','반려기한일','피드백기한일'], '#1e3a8a');
  mk('제출현황', ['날짜','학번','이름','과제명','난이도','메모','파일URL','피드백','해시','읽음여부','상태','첨삭URL','점수','공개여부','메모2','답글','우수작유형','익명여부','우수작멘트','우수작키','부정행위','문항별JSON','AI채점JSON','상태변경일시','재제출개별마감','되돌림유형','되돌림횟수'], '#1e3a8a');
  mk('AI채점기준', ['과제명','채점유형','만점','기준','파일JSON','문항JSON'], '#1e3a8a');
  var def = ss.getSheetByName('시트1') || ss.getSheetByName('Sheet1');
  if (def && ss.getSheets().length > 1) { try { ss.deleteSheet(def); } catch(_) {} }
}

// ── [1회 실행] 피드백/재제출 루프용 컬럼 추가 (기존 시트에 헤더·기본값 보강) ──
// 과제설정: J=반려기한일, K=피드백기한일 (비어있으면 7일 기본) / 제출현황: Y·Z·AA 헤더
function setupFeedbackLoop() {
  var ss = _taskSs();
  var ts = ss.getSheetByName('과제설정');
  if (ts) {
    ts.getRange(1, 10).setValue('반려기한일');
    ts.getRange(1, 11).setValue('피드백기한일');
    var last = ts.getLastRow();
    if (last > 1) {
      var rng = ts.getRange(2, 10, last - 1, 2);
      var vals = rng.getValues();
      var changed = false;
      for (var i = 0; i < vals.length; i++) {
        if (vals[i][0] === '' || vals[i][0] === null) { vals[i][0] = 7; changed = true; }
        if (vals[i][1] === '' || vals[i][1] === null) { vals[i][1] = 7; changed = true; }
      }
      if (changed) rng.setValues(vals);
    }
  }
  var sub = ss.getSheetByName('제출현황');
  if (sub) {
    sub.getRange(1, 25).setValue('재제출개별마감');
    sub.getRange(1, 26).setValue('되돌림유형');
    sub.getRange(1, 27).setValue('되돌림횟수');
  }
  clearCache();
  return { success: true, message: '피드백 루프 컬럼 설정 완료 (기본 기한 7일)' };
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

// ── PWA(Pages) 프론트 → GAS 백엔드 호출 (google.script.run 어댑터) ──
// text/plain {fn,args,token} → {ok,result|error}. teacherLogin 외 전부 토큰 검사(공개 웹앱 보안).
var RPC_WHITELIST = ['aiGradeSubmission','approveResubmitRequest','bulkPublishTasks','cancelBestWork','changeBestScope','clearStudentReply','deleteRubric','deleteTask','getAchievements','getApiSettings','getAppLinkUrl','getDashboardData','getNotifyLog','getNotifySettings','getRubrics','getSecureFileBase64','getTasksOnly','installAutoNotifyTriggers','resetStudentPassword','runAICheatCheck','saveAiGradeResult','saveAiGradeTempResult','saveApiSetting','saveAppLinkUrl','saveFeedback','saveMultiAnnotatedImages','saveNewTask','saveRubric','sendPushToUnsubmitted','setNotifySettings','teacherLogin','toggleTaskVisibility','updateTaskSettings','uploadRubricFile','uploadTaskProblemImage','saveTaskProblems','syncAchievementSheet','validateTeacherSession'];
var RPC_NOAUTH = ['teacherLogin'];   // 로그인만 토큰 불필요
function _atOk_(token) {
  try { var v = ClassCore.verifyTeacher(token); return (v === true) || !!(v && (v.success || v.valid || v.ok)); } catch(_) { return false; }
}
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
  } catch (err) {
    out = { ok: false, error: String(err && err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function _serviceUrl_() {
  try { return ScriptApp.getService().getUrl() || ''; } catch(_) { return ''; }
}

// ── 교사 인증 (ClassCore 위임) ──
function teacherLogin(pw)        { return ClassCore.teacherLogin(pw); }
function verifyTeacher(token)    { return ClassCore.verifyTeacher(token); }
function teacherLogout(token)    { return ClassCore.teacherLogout(token); }
function validateTeacherSession(token) { return ClassCore.verifyTeacher(token); }

// ── 학생 비밀번호 초기화 (ClassCore 위임) — 빈 비번으로 리셋 → 학생이 다음 로그인 때 재설정 ──
function resetStudentPassword(studentId) {
  try { return ClassCore.resetStudentPw(studentId, ''); }
  catch (e) { return { success: false, message: e.toString() }; }
}

// ── 앱 바로가기 URL (ClassCore 앱URL 레지스트리 위임) — 클라이언트는 '바로가기_*' 키로 호출 ──
function getAppLinkUrl(cell) {
  try {
    var key = String(cell || '').replace('바로가기_', '');
    var urls = ClassCore.getAppUrls() || {};
    return { url: urls[key] || '' };
  } catch (e) { return { url: '' }; }
}
function saveAppLinkUrl(cell, url) {
  try {
    ClassCore.registerAppUrl(String(cell || '').replace('바로가기_', ''), String(url || '').trim());
    return { success: true };
  } catch (e) { return { success: false, message: e.toString() }; }
}

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


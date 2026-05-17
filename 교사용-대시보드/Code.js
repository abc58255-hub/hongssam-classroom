const SHEET_ID = "1jK7gYGFXCe3FULLs5mKttP959Aa9vp8-WNOGdJy7cZQ";
const PARENT_FOLDER_ID = "1nmo4ZtQYK3-0PFjMKO8yzlkNOVoLn9_H";

// ── 시스템설정 키-값 헬퍼 ──────────────────────────────
function _getSys(ss, key) {
  var sh = ss.getSheetByName('시스템설정');
  if (!sh || sh.getLastRow() < 2) return '';
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === key) return String(rows[i][1] || '').trim();
  }
  return '';
}

function _setSys(ss, key, value) {
  var sh = ss.getSheetByName('시스템설정');
  if (!sh || sh.getLastRow() < 2) return;
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === key) {
      sh.getRange(i + 2, 2).setValue(value);
      return;
    }
  }
}

// =====================================================
// ✅ FCM 푸시 알림 (HTTP v1 API + 서비스 계정)
// 최초 1회: GAS 스크립트 편집기에서 setupFCMCredentials() 실행
// =====================================================

function debugFCM() {
  var props = PropertiesService.getScriptProperties();
  var projectId = props.getProperty('FCM_PROJECT_ID');
  var email = props.getProperty('FCM_CLIENT_EMAIL');
  var key = props.getProperty('FCM_PRIVATE_KEY');
  Logger.log('PROJECT_ID: ' + projectId);
  Logger.log('KEY_LENGTH: ' + (key || '').length);
  try {
    var accessToken = getFCMAccessToken_();
    Logger.log('ACCESS_TOKEN OK: ' + accessToken.substring(0, 30));
    // 스프레드시트에서 첫 번째 FCM 토큰 찾아서 테스트 발송
    var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('학생명부');
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var token = String(data[i][6] || '').trim();
      if (token) {
        Logger.log('첫 번째 토큰 발견 (학번 ' + data[i][1] + '): ' + token.substring(0, 30) + '...');
        var res = sendFcmToToken_(token, '테스트 알림', '테스트입니다', '', 'test');
        Logger.log('발송 결과: ' + res);
        break;
      }
    }
  } catch(e) {
    Logger.log('오류: ' + e.toString());
  }
}

function setupFCMCredentials() {
  var props = PropertiesService.getScriptProperties();
  props.setProperty('FCM_PROJECT_ID',   'send-alarm-220c3');
  props.setProperty('FCM_CLIENT_EMAIL', 'firebase-adminsdk-fbsvc@send-alarm-220c3.iam.gserviceaccount.com');
  props.setProperty('FCM_PRIVATE_KEY',  '-----BEGIN PRIVATE KEY-----\nMIIEvgIBADANBgkqhkiG9w0BAQEFAASCBKgwggSkAgEAAoIBAQC5xXXlzic2Ykf5\naMW9BoMD1VFLvO4ShawVjtyUFwnf5Swufq6fLNxGsXdVfpt2sXDBGgXD97EwoUTQ\nMmJ/UPWw2ulG23cbbA4mK9Xu+Jh3Z9bYC3wutW3NNOpb0G/ivdcbeMu6XoDtuAMt\ndgKESakK3Ta05KuU3JFSBqTAlwIIRrtEeiXFERV6fv7jUvua/0ncTKDqOymSRuph\nYOK5TAHlTBHitqojoJ0KGK6TWAgGUGnshtXwlcyxyqAcdN6Mq75G8/LRbc6hMMIY\n5BMzsrhJt5MLcXV7Kr+iPb1yJmb1Px7vjmnsT5IAQZHkGYRDEWvKJQNSHsTJ0MW6\nEDtcknQ9AgMBAAECggEAA2cg4v+cdCKjAN+wxW9uGb3O6CXJVuDZr+h44gZb4IvW\n/BNPthRyN/drB5PQuZqZRYrpN/3vt25kRuVbxkK4Hnicm5CqPru++QD1OPqF91/U\n9ht54h9IQa31DuSOEbaYYC3ZRX9oMZSuC9IikLvGSJYx9KzYKrCgbEtbQmuXCb15\nYJ00GD1R0WUkCjoJGCWKoZmuuqOj/o/LvkTOywHw3Wx0w50V/SldpgGT0aJ1ZRjj\n2+MC7Yd+06RQ2aDcLxqIn4YM73MKd0kkUGWZ84JyfHdVxOA4kQoz52b7VIjDLx9C\n6qji9+Dddy+ULmgQFgHLsPE9I9sdszianbIK2O2evQKBgQDx083+eShG9H4iZAIj\npNpWSpj/hTzx50D+W1nuYiYM+v7AoiZgAqflsDqFcyeXcTE1WCH1Jt9HqMka3G0k\n3BomzFKXoQqPTLxEbLu7sPONzuGETUqfEC53Rm7SlZW7Zf7nSi+qHEuGNryr+IKp\nVC0Ry/yjGzU8aVT6LJ9xFOh7xwKBgQDEqKIqFolxb7+y4LZlZJm0IefgrOv4Y5m9\nLhJivnX7MT72DSyAhqUQI+jyHmjqmX8px6ieFOZs0sRtDQ3V/Rm7b16CzckoDwE8\neTy5tdA9uCJO9Q8Hn4XQDk8cCnN45tSIKGBSRCUVWBpkx2mgXBZUqRmWTZNufMuV\n5EclvgDn2wKBgQCl5GSlszucIVD+CpklFovpMlduwloioD+HvecdjxsHQI/OWe31\nYx0GhjQ/I9X/H9lf/Muev0HgiLscwCXnaU5PW081UXZLA5sLXYQTp7oMh+VQuyz7\nnUAi9qBDufXzjm6k/9Fe4vY0Zgxb9Ki9vE8GrGbtBVcp0CBJVp8yFeO+NQKBgEur\nW2J4c4A7cHaismwHLoE6Pp+bydw0btZ1IMvv3zO9Oi2w2fvGU0MHnj1zaAlE8MIh\nugbBofwjiUMwr680CS+u5Z3NEuagB2i+eZg3lh35ePIKpzLWtcVdjCENAGt33jVZ\n294rrF0vHlDCzijO5iTDQD4uMVllGWzefmXOW0jbAoGBALcUbVtH59FaHCZThuKS\nom659DWn4LePqDFal8bfEcgTko+zBZv7lZIZ9unps3GdtccR5XZCKfn9MchNKfT1\nVzqB3CPuh21vHJcMegn9WzzYwP2ezCKgjNBqbUJve7YUYuvQpZ/QFEGJEbte56pC\np2Grlq/5dEOE81GOHW8TeRAv\n-----END PRIVATE KEY-----\n');
  return '✅ FCM 자격증명 저장 완료';
}

function getFCMAccessToken_() {
  var props = PropertiesService.getScriptProperties();
  var privateKey  = props.getProperty('FCM_PRIVATE_KEY').replace(/\\n/g, '\n');
  var clientEmail = props.getProperty('FCM_CLIENT_EMAIL');
  var now = Math.floor(Date.now() / 1000);
  var header  = Utilities.base64EncodeWebSafe(JSON.stringify({alg:'RS256',typ:'JWT'})).replace(/=+$/,'');
  var payload = Utilities.base64EncodeWebSafe(JSON.stringify({
    iss: clientEmail,
    scope: 'https://www.googleapis.com/auth/firebase.messaging',
    aud: 'https://oauth2.googleapis.com/token',
    exp: now + 3600,
    iat: now
  })).replace(/=+$/,'');
  var toSign    = header + '.' + payload;
  var signature = Utilities.base64EncodeWebSafe(
    Utilities.computeRsaSha256Signature(toSign, privateKey)
  ).replace(/=+$/,'');
  var jwt = toSign + '.' + signature;
  var res = UrlFetchApp.fetch('https://oauth2.googleapis.com/token', {
    method: 'POST',
    contentType: 'application/x-www-form-urlencoded',
    payload: 'grant_type=urn%3Aietf%3Aparams%3Aoauth%3Agrant-type%3Ajwt-bearer&assertion=' + jwt
  });
  return JSON.parse(res.getContentText()).access_token;
}

function sendFcmToToken_(token, title, body, clickUrl, tag) {
  try {
    var projectId   = PropertiesService.getScriptProperties().getProperty('FCM_PROJECT_ID');
    var accessToken = getFCMAccessToken_();
    var message = {
      message: {
        token: token,
        notification: { title: title, body: body },
        data: { url: clickUrl || '', tag: tag || 'default' },
        webpush: {
          fcm_options: { link: clickUrl || 'https://script.google.com/macros/s/AKfycbyR1whn6f90-kJEAaJg_O34uP8v-KvyEqsRky58idjoxVDS5cWj80p2ScJp6V2dnz_0hA/exec' }
        }
      }
    };
    var res = UrlFetchApp.fetch(
      'https://fcm.googleapis.com/v1/projects/' + projectId + '/messages:send',
      { method:'POST', contentType:'application/json',
        headers:{'Authorization':'Bearer ' + accessToken},
        payload: JSON.stringify(message), muteHttpExceptions: true }
    );
    return res.getResponseCode() === 200;
  } catch(e) { return false; }
}

// 선택한 학생들에게 알림 발송 (클라이언트에서 호출)
function sendPushToStudents(studentIds, title, body, tag, filterClass) {
  try {
    var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('학생명부');
    var data  = sheet.getDataRange().getValues();
    var clickUrl = 'https://script.google.com/macros/s/AKfycbyR1whn6f90-kJEAaJg_O34uP8v-KvyEqsRky58idjoxVDS5cWj80p2ScJp6V2dnz_0hA/exec';
    var sent = 0, skipped = 0;
    var idSet = studentIds ? new Set(studentIds.map(String)) : null;
    for (var i = 1; i < data.length; i++) {
      var sid   = String(data[i][1] || '').trim();
      var token = String(data[i][6] || '').trim();
      if (!sid || !token) { skipped++; continue; }
      if (idSet && !idSet.has(sid)) continue;
      if (filterClass) {
        var cls = sid.length >= 2 ? (sid.substring(0,1) + '학년 ' + sid.substring(1,2) + '반') : '';
        if (cls !== filterClass) continue;
      }
      if (sendFcmToToken_(token, title, body, clickUrl, tag || 'default')) sent++;
      else skipped++;
    }
    return { success: true, sent: sent, skipped: skipped };
  } catch(e) { return { success: false, message: e.toString() }; }
}

// 전체 학생 알림
function sendPushToAll(title, body, tag) {
  return sendPushToStudents(null, title, body, tag);
}

// 특정 과제 미제출자에게 알림
function sendPushToUnsubmitted(taskName, title, body) {
  try {
    var roster  = SpreadsheetApp.openById(SHEET_ID).getSheetByName('학생명부').getDataRange().getValues();
    var history = SpreadsheetApp.openById(SHEET_ID).getSheetByName('제출현황').getDataRange().getValues();
    var submittedIds = new Set();
    for (var i = 1; i < history.length; i++) {
      var tn = String(history[i][3] || '').split(' (')[0];
      if (tn === taskName) submittedIds.add(String(history[i][1] || '').trim());
    }
    var unsubIds = [];
    for (var j = 1; j < roster.length; j++) {
      var sid = String(roster[j][1] || '').trim();
      if (sid && !submittedIds.has(sid)) unsubIds.push(sid);
    }
    if (unsubIds.length === 0) return { success: true, sent: 0, skipped: 0, message: '미제출자 없음' };
    return sendPushToStudents(unsubIds, title || ('📢 [' + taskName + '] 아직 제출 안 했어요!'), body || '지금 바로 제출해주세요.', 'unsub');
  } catch(e) { return { success: false, message: e.toString() }; }
}

// =====================================================
// ✅ 피드백 템플릿 (시스템설정 시트 '피드백_N' 키 행에 저장)
// =====================================================
function getFeedbackTemplates() {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sh = ss.getSheetByName('시스템설정');
    if (!sh || sh.getLastRow() < 2) return { success: true, templates: [] };
    var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
    var templates = [];
    rows.forEach(function(r) {
      if (String(r[0]).trim().indexOf('피드백_') === 0 && String(r[1] || '').trim()) {
        templates.push(String(r[1]).trim());
      }
    });
    return { success: true, templates: templates };
  } catch(e) { return { success: false, templates: [], message: e.toString() }; }
}

function saveFeedbackTemplates(templates) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sh = ss.getSheetByName('시스템설정');
    if (!sh) return { success: false, message: '시스템설정 시트 없음' };
    var lastRow = sh.getLastRow();
    if (lastRow < 2) return { success: false };
    var rows = sh.getRange(2, 1, lastRow - 1, 2).getValues();
    // 기존 피드백_ 행 B열 클리어
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][0]).trim().indexOf('피드백_') === 0) {
        sh.getRange(i + 2, 2).clearContent();
      }
    }
    // 다시 저장 (피드백_1, 피드백_2, ... 순서로)
    var idx = 1;
    templates.forEach(function(t) {
      if (t && t.trim()) {
        _setSys(ss, '피드백_' + idx, t.trim());
        idx++;
      }
    });
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}

// =====================================================
// ✅ 학생 비밀번호 초기화
// =====================================================
function resetStudentPassword(studentId) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName("학생명부");
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1] || "").trim() === String(studentId).trim()) {
        sheet.getRange(i + 1, 4).clearContent(); // D열 비밀번호 초기화
        clearCache();
        return { success: true };
      }
    }
    return { success: false, message: "학생을 찾을 수 없습니다." };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function doGet(e) {
  // TickTick OAuth 콜백 처리
  if (e && e.parameter && e.parameter.code) {
    return handleTickTickCallback(e.parameter.code);
  }
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('통합 교사 대시보드')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// =====================================================
// ✅ 교사 세션 인증 (UUID 토큰 기반)
// =====================================================
var TEACHER_SESSION_KEY = 'teacher_session_token';
var TEACHER_SESSION_EXPIRY_KEY = 'teacher_session_expiry';
var TEACHER_SESSION_DURATION = 4 * 60 * 60; // 4시간

function teacherLogin(pw) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const stored = _getSys(ss, '교사비밀번호');
    if (!stored) return { success: false, message: "교사 비밀번호가 설정되지 않았습니다." };
    if (pw !== stored) return { success: false, message: "비밀번호가 일치하지 않습니다." };

    // 세션 토큰 생성 및 저장
    const token = Utilities.getUuid();
    const expiry = new Date().getTime() + (TEACHER_SESSION_DURATION * 1000);
    const props = PropertiesService.getScriptProperties();
    props.setProperty(TEACHER_SESSION_KEY, token);
    props.setProperty(TEACHER_SESSION_EXPIRY_KEY, String(expiry));

    return { success: true, token: token };
  } catch(e) {
    return { success: false, message: "로그인 중 오류: " + e.message };
  }
}

function validateTeacherSession(token) {
  if (!token || typeof token !== 'string') return false;
  const props = PropertiesService.getScriptProperties();
  const stored = props.getProperty(TEACHER_SESSION_KEY);
  const expiryStr = props.getProperty(TEACHER_SESSION_EXPIRY_KEY);
  if (!stored || stored !== token) return false;
  if (expiryStr && Number(expiryStr) < new Date().getTime()) {
    props.deleteProperty(TEACHER_SESSION_KEY);
    props.deleteProperty(TEACHER_SESSION_EXPIRY_KEY);
    return false;
  }
  return true;
}

function teacherLogout(token) {
  if (!token) return;
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty(TEACHER_SESSION_KEY) === token) {
    props.deleteProperty(TEACHER_SESSION_KEY);
    props.deleteProperty(TEACHER_SESSION_EXPIRY_KEY);
  }
}

// 하위 호환용 (구버전 호출 대비)
function verifyTeacher(pw) {
  return teacherLogin(pw);
}


// =====================================================
// ✅ 캐시 유틸 함수 (명부·과제 등 자주 안 바뀌는 데이터용)
// =====================================================
function getCached(key) {
  try {
    const cached = CacheService.getScriptCache().get(key);
    return cached ? JSON.parse(cached) : null;
  } catch(e) { return null; }
}

function setCached(key, data, ttlSeconds) {
  try {
    CacheService.getScriptCache().put(key, JSON.stringify(data), ttlSeconds || 300);
  } catch(e) {}
}

function clearCache() {
  try {
    CacheService.getScriptCache().removeAll(['roster', 'classList', 'tasks', 'hrActivities', 'homeroom']);
  } catch(e) {}
}

// =====================================================
// ✅ 공통 파서들
// =====================================================
function parseRosterAndClasses(ss) {
  let cached = getCached('roster');
  if (cached) return cached;

  const rosterData = ss.getSheetByName("학생명부").getDataRange().getValues();
  let roster = [], classSet = new Set();
  for (let i = 1; i < rosterData.length; i++) {
    if (!rosterData[i][1]) continue;
    let sid = String(rosterData[i][1] || "").trim();
    let sname = String(rosterData[i][2] || "").trim();
    let clsName = sid.length >= 2 ? `${sid.substring(0,1)}학년 ${sid.substring(1,2)}반` : "기타";
    roster.push({ id: sid, name: sname, cls: clsName });
    if (clsName !== "기타") classSet.add(clsName);
  }
  let classList = Array.from(classSet).sort();
  let result = { roster, classList };
  setCached('roster', result, 600); // 10분 캐시
  return result;
}

function parseTasks(ss) {
  let cached = getCached('tasks');
  if (cached) return cached;

  const taskData = ss.getSheetByName("과제설정").getDataRange().getValues();
  let tasks = [];
  for (let i = 1; i < taskData.length; i++) {
    if (!taskData[i][1]) continue;
    let reqPics = taskData[i][6] ? parseInt(taskData[i][6]) : 1;
    let choiceList = taskData[i][7] ? String(taskData[i][7] || "").trim() : "";
    tasks.push({
      rowIdx: i + 1,
      date: taskData[i][0] ? Utilities.formatDate(new Date(taskData[i][0]), "Asia/Seoul", "yyyy-MM-dd HH:mm") : "",
      name: String(taskData[i][1] || "").trim(),
      desc: taskData[i][2] ? String(taskData[i][2] || "").trim() : "",
      deadlines: taskData[i][3] ? String(taskData[i][3] || "").trim() : "{}",
      evalType: taskData[i][4] ? String(taskData[i][4] || "").trim() : "점수제 (직접 입력)",
      isPublic: String(taskData[i][5] || "").trim() === "일괄공개",
      reqPics: reqPics,
      choiceList: choiceList
    });
  }
  tasks.reverse();
  setCached('tasks', tasks, 300); // 5분 캐시
  return tasks;
}

function parseSubmissions(ss) {
  const sheet = ss.getSheetByName('제출현황');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  // getDataRange 대신 필요한 열만 읽기 (A~W = 1~23열)
  const subData = sheet.getRange(2, 1, lastRow - 1, 23).getValues();
  const submissions = [];

  // ✅ [추가] 과제별 실시간 등수 계산용 객체
  let rankCounters = {};

  for (let i = 0; i < subData.length; i++) {
    if (!subData[i][1]) continue;

    let rowId = String(subData[i][1] || '').trim();
    let rowTask = String(subData[i][3] || '').trim();
    let baseTaskName = rowTask.split(' (')[0];
    let rowStatus = String(subData[i][10] || '').trim();
    let rowClass = rowId.length >= 2 ? `${rowId.substring(0,1)}학년 ${rowId.substring(1,2)}반` : "기타";

    // ✅ [추가] 반려나 이전기록이 아니면 등수 카운트 증가
    let myTotalRank = 0;
    let myClassRank = 0;
    if (rowStatus !== '재제출요청' && rowStatus !== '반려검토' && rowStatus !== '이전기록채점완료') {
      if (!rankCounters[baseTaskName]) rankCounters[baseTaskName] = { total: 0, classes: {} };
      if (!rankCounters[baseTaskName].classes[rowClass]) rankCounters[baseTaskName].classes[rowClass] = 0;
      
      rankCounters[baseTaskName].total++;
      rankCounters[baseTaskName].classes[rowClass]++;
      
      myTotalRank = rankCounters[baseTaskName].total;
      myClassRank = rankCounters[baseTaskName].classes[rowClass];
    }

    // URL 파싱 (G열=6)
    const rawUrl = String(subData[i][6] || '').trim();
    let urls = {};
    if (rawUrl.startsWith('{')) { try { urls = JSON.parse(rawUrl); } catch(e) {} }
    else if (rawUrl.startsWith('[')) { try { JSON.parse(rawUrl).forEach((v,j) => urls['사진 '+(j+1)] = v); } catch(e) {} }
    else if (rawUrl) { urls['사진 1'] = rawUrl; }
    
    // 첨삭 URL 파싱 (L열=11)
    const rawAnno = String(subData[i][11] || '').trim();
    let annoUrls = {};
    if (rawAnno.startsWith('{')) { try { annoUrls = JSON.parse(rawAnno); } catch(e) {} }
    else if (rawAnno.startsWith('[')) { try { JSON.parse(rawAnno).forEach((v,j) => annoUrls['사진 '+(j+1)] = v); } catch(e) {} }
    else if (rawAnno) { annoUrls['사진 1'] = rawAnno; }
    
    // 문항별 데이터 (V열=21)
    let perQuestionData = {};
    try {
      const pqRaw = String(subData[i][21] || '').trim();
      if (pqRaw.startsWith('{')) perQuestionData = JSON.parse(pqRaw);
    } catch(e) {}
    
    // 날짜
    const dateVal = subData[i][0];
    const dateMs  = dateVal instanceof Date ? dateVal.getTime() : 0;

    submissions.push({
      rowIdx:          i + 2,
      dateMs:          dateMs,
      id:    rowId,
      name:  String(subData[i][2] || '').trim(),
      task:  rowTask,
      level: String(subData[i][4] || '').trim() || undefined,
      message: String(subData[i][5] || '').trim() || undefined,
      urls,
      feedback:    String(subData[i][7]  || '').trim() || undefined,
      status:      rowStatus,
      annoUrls:    Object.keys(annoUrls).length ? annoUrls : undefined,
      score:       String(subData[i][12] || '').trim() || undefined,
      isPublic:    String(subData[i][13] || '').trim() === '공개',
      memo:        String(subData[i][14] || '').trim() || undefined,
      reply:       String(subData[i][15] || '').trim() || undefined,
      bestType:    String(subData[i][16] || '').trim() || undefined,
      isAnon:      subData[i][17] === true || String(subData[i][17] || '').toUpperCase() === 'TRUE',
      bestComment: String(subData[i][18] || '').trim() || undefined,
      bestKey:     String(subData[i][19] || '').trim() || undefined,
      cheatFlag:   String(subData[i][20] || '').trim() || undefined,
      perQuestionData: Object.keys(perQuestionData).length ? perQuestionData : undefined,
      aiGradeTemp: String(subData[i][22] || '').trim() || undefined,
      // ✅ [추가] 계산된 등수를 데이터에 포함
      totalRank: myTotalRank,
      classRank: myClassRank
    });
  }
  submissions.reverse();
  return submissions;
}

// =====================================================
// ✅ 페이지별 데이터 로딩 함수 (분리된 버전)
// =====================================================

// 채점/제출현황 페이지 (가장 자주 사용)
function getGradeData() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const { roster, classList } = parseRosterAndClasses(ss);
    const tasks = parseTasks(ss);
    const submissions = parseSubmissions(ss);
    let taskEvalMap = {};
    tasks.forEach(t => taskEvalMap[t.name] = t.evalType);
    return { roster, classList, tasks, submissions, taskEvalMap };
  } catch(e) { throw new Error("채점 데이터 오류: " + e.message); }
}

// 과제 관리 페이지
function getTaskPageData() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const { classList } = parseRosterAndClasses(ss);
    clearCache(); // 과제 수정 후 캐시 갱신
    const tasks = parseTasks(ss);
    return { classList, tasks };
  } catch(e) { throw new Error("과제 데이터 오류: " + e.message); }
}

// 통계/미제출자 페이지
function getReportData() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const { roster, classList } = parseRosterAndClasses(ss);
    const submissions = parseSubmissions(ss);
    const tasks = parseTasks(ss);
    return { roster, classList, submissions, tasks };
  } catch(e) { throw new Error("통계 데이터 오류: " + e.message); }
}

// 설문 페이지
function getSurveyData() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let surveys = [], surveyRes = [];
    const svSheet = ss.getSheetByName("설문목록");
    if (svSheet) {
      let svData = svSheet.getDataRange().getValues();
      for (let i = 1; i < svData.length; i++) {
        if (svData[i][0]) surveys.push({
          id: svData[i][0],
          date: svData[i][1] ? Utilities.formatDate(new Date(svData[i][1]), "Asia/Seoul", "yyyy-MM-dd") : "",
          title: svData[i][2], status: svData[i][3], questions: svData[i][4],
          guide: String(svData[i][5] || '').trim()
        });
      }
    }
    const srSheet = ss.getSheetByName("설문응답");
    if (srSheet) {
      let srData = srSheet.getDataRange().getValues();
      for (let i = 1; i < srData.length; i++) {
        if (srData[i][1]) surveyRes.push({
          date: srData[i][0] ? Utilities.formatDate(new Date(srData[i][0]), "Asia/Seoul", "MM/dd HH:mm") : "",
          svId: srData[i][1], stuId: srData[i][2], stuName: srData[i][3], answers: srData[i][4]
        });
      }
    }
    const { roster } = parseRosterAndClasses(ss);
    let hrStr = "";
    try {
      hrStr = _getSys(ss, '담임반');
    } catch(e) {}
    return { surveys: surveys.reverse(), surveyRes, roster, homeroom: hrStr };
  } catch(e) { throw new Error("설문 데이터 오류: " + e.message); }
}

// ✅ 1. 활동 관리: 목록 가져오기 (교사용 - rowIdx 포함)
function getHrActivitiesFull() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName("학급활동목록");
    if (!sheet) return { success: true, activities: [] };
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: true, activities: [] };
    
    const data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
    const activities = data.map((r, i) => ({
      rowIdx: i + 2,
      category: String(r[1]||'').trim(),
      name:     String(r[2]||'').trim(),
      desc:     String(r[3]||'').trim(),
      formUrl:  String(r[4]||'').trim(),
      fields:   String(r[5]||'').trim()
    }));
    return { success: true, activities: activities.reverse() };
  } catch(e) { return { success: false, message: e.toString() }; }
}

// ✅ 2. 활동 관리: 저장 (추가 및 수정 통합)
function saveHrActivity(data) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName("학급활동목록");
    if (!sheet) {
      sheet = ss.insertSheet("학급활동목록");
      sheet.getRange(1,1,1,6).setValues([['번호','카테고리','활동명','설명','폼링크','필드설정']]).setFontWeight('bold');
    }
    
    const row = [
      data.rowIdx || sheet.getLastRow(), // 번호 (기존 번호 유지 또는 마지막 번호)
      data.category || '자율활동',
      data.name || '',
      data.desc || '',
      data.formUrl || '',
      data.fields || '[]'
    ];

    if (data.rowIdx && data.rowIdx > 1) {
      // 수정 모드
      sheet.getRange(data.rowIdx, 1, 1, 6).setValues([row]);
    } else {
      // 신규 추가
      sheet.appendRow([sheet.getLastRow(), row[1], row[2], row[3], row[4], row[5]]);
    }
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}

// ✅ 3. 활동 관리: 삭제
function deleteHrActivity(rowIdx) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName("학급활동목록");
    if (sheet && rowIdx > 1) {
      sheet.deleteRow(rowIdx);
      return { success: true };
    }
    return { success: false, message: "삭제할 활동을 찾을 수 없습니다." };
  } catch(e) { return { success: false, message: e.toString() }; }
}

// ✅ 4. 홈룸 통합 데이터 (학생별 조회를 위해 수정)
function getHomeroomTabData() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    
    // 활동 목록
    let hrActivities = [];
    const actSheet = ss.getSheetByName("학급활동목록");
    if (actSheet && actSheet.getLastRow() >= 2) {
      hrActivities = actSheet.getRange(2, 1, actSheet.getLastRow()-1, 6).getValues().map((r, i) => ({
        rowIdx: i + 2, category: r[1], name: r[2], desc: r[3]
      }));
    }

    // 제출 현황 (전체)
    let hrSubmissions = [];
    const subSheet = ss.getSheetByName("창체제출현황");
    if (subSheet && subSheet.getLastRow() >= 2) {
      hrSubmissions = subSheet.getRange(2, 1, subSheet.getLastRow()-1, 8).getValues().map(r => ({
        date: r[0] ? Utilities.formatDate(new Date(r[0]), "Asia/Seoul", "MM/dd HH:mm") : "",
        id: String(r[1]).trim(), 
        name: String(r[2]).trim(),
        activity: String(r[3]).trim(),
        role: String(r[4]).trim(),
        reflection: String(r[5]).trim(),
        url: String(r[6]).trim()
      }));
    }

    return { 
      success: true, 
      hrActivities: hrActivities.reverse(), 
      hrSubmissions: hrSubmissions.reverse() 
    };
  } catch(e) { return { success: false, message: e.toString() }; }
}

// 학급활동 페이지
function getHomeroomData(studentId) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);

    // ── 1. 학급활동목록 → activities
    const activities = [];
    const actSheet = ss.getSheetByName('학급활동목록');
    if (actSheet && actSheet.getLastRow() >= 2) {
      const actData = actSheet.getRange(2, 1, actSheet.getLastRow()-1, 6).getValues();
      actData.forEach(function(r) {
        if (!r[2]) return; // 활동명 없으면 스킵
        activities.push({
          category: String(r[1]||'').trim(),
          name:     String(r[2]||'').trim(),
          desc:     String(r[3]||'').trim(),
          formUrl:  String(r[4]||'').trim(),
          fields:   String(r[5]||'').trim()
        });
      });
    }

    // ── 2. 창체제출현황 → myRecords (이 학생 기록만)
    const myRecords = [];
    const subSheet = ss.getSheetByName('창체제출현황');
    if (subSheet && subSheet.getLastRow() >= 2) {
      const subData = subSheet.getRange(2, 1, subSheet.getLastRow()-1, 8).getValues();
      subData.forEach(function(r) {
        if (!r[1]) return;
        const rid = String(r[1]||'').trim();
        if (studentId && rid !== String(studentId).trim()) return;
        myRecords.push({
          date:       r[0] ? Utilities.formatDate(new Date(r[0]), 'Asia/Seoul', 'MM/dd HH:mm') : '',
          activity:   String(r[3]||'').trim(),
          role:       String(r[4]||'').trim(),
          reflection: String(r[5]||'').trim(),
          url:        String(r[6]||'').trim() || '첨부파일 없음',
          extra:      String(r[7]||'').trim()
        });
      });
    }
    myRecords.reverse();

    // ── 3. 설문목록 → activeSurvey (진행중인 것)
    let activeSurvey = null;
    let hasSubmittedSurvey = false;
    const svSheet = ss.getSheetByName('설문목록');
    if (svSheet && svSheet.getLastRow() >= 2) {
      const svData = svSheet.getDataRange().getValues();
      for (let i = 1; i < svData.length; i++) {
        if (String(svData[i][3]).trim() === '진행중') {
          activeSurvey = {
            id:        String(svData[i][0]).trim(),
            title:     String(svData[i][2]).trim(),
            questions: String(svData[i][4]).trim()
          };
          break;
        }
      }
    }
    // 설문 응답 여부 확인
    if (activeSurvey && studentId) {
      const srSheet = ss.getSheetByName('설문응답');
      if (srSheet && srSheet.getLastRow() >= 2) {
        const srData = srSheet.getRange(2,1,srSheet.getLastRow()-1,3).getValues();
        for (let i = 0; i < srData.length; i++) {
          if (String(srData[i][0]).trim() === activeSurvey.id &&
              String(srData[i][1]).trim() === String(studentId).trim()) {
            hasSubmittedSurvey = true; break;
          }
        }
      }
    }

    // ── 4. 학급알림 → notices (공개된 것만)
    const notices = [];
    const noticeSheet = ss.getSheetByName('학급알림');
    if (noticeSheet && noticeSheet.getLastRow() >= 2) {
      const nd = noticeSheet.getRange(2,1,noticeSheet.getLastRow()-1,6).getValues();
      nd.forEach(function(r) {
        if (!r[0]) return;
        const visible = r[4] === true || String(r[4]).toUpperCase() === 'TRUE';
        if (!visible) return;
        notices.push({
          title:     String(r[0]||'').trim(),
          content:   String(r[1]||'').trim(),
          type:      String(r[2]||'알림').trim(),
          date:      r[3] instanceof Date ? Utilities.formatDate(r[3],'Asia/Seoul','MM/dd') : String(r[3]||'').trim(),
          important: r[5] === true || String(r[5]).toUpperCase() === 'TRUE'
        });
      });
    }
    notices.reverse();

    // ── 5. 명부 + 반 목록
    const { roster, classList } = parseRosterAndClasses(ss);

    return {
      activities,
      myRecords,
      activeSurvey,
      hasSubmittedSurvey,
      notices,
      roster,
      classList
    };
  } catch(e) {
    throw new Error('홈룸 데이터 오류: ' + e.message);
  }
}


// =====================================================
// ✅ 교사용: 기존 getDashboardData 복구 (+ B4 슬라이드 URL 추가)
// =====================================================
function getDashboardData() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const cache = CacheService.getScriptCache();

    let hrStr = "";
    let defaultSlideUrl = ""; // ✅ 기본슬라이드URL
    try {
      hrStr = _getSys(ss, '담임반');
      defaultSlideUrl = _getSys(ss, '기본슬라이드URL');
    } catch(e) {}

    let roster, classList;
    const rosterCacheKey = 'roster_v1';
    const cachedRoster = cache.get(rosterCacheKey);
    if (cachedRoster) {
      const parsed = JSON.parse(cachedRoster);
      roster = parsed.roster; classList = parsed.classList;
    } else {
      const r = parseRosterAndClasses(ss);
      roster = r.roster; classList = r.classList;
      try { cache.put(rosterCacheKey, JSON.stringify({roster, classList}), 60); } catch(e) {}
    }

    let tasks;
    const taskCacheKey = 'tasks_v1';
    const cachedTasks = cache.get(taskCacheKey);
    if (cachedTasks) {
      try { tasks = JSON.parse(cachedTasks); } catch(e) { tasks = parseTasks(ss); }
    } else {
      tasks = parseTasks(ss);
      try { cache.put(taskCacheKey, JSON.stringify(tasks), 30); } catch(e) {}
    }
    
    const submissions = parseSubmissions(ss);

    return {
      roster, classList, tasks, submissions,
      homeroom: hrStr,
      defaultSlideUrl: defaultSlideUrl, // ✅ 클라이언트로 전달
      hrActivities: [], hrSubmissions: [], surveys: [], surveyRes: []
    };
  } catch(e) { throw new Error("전체 데이터 로딩 오류: " + e.message); }
}
// =====================================================
// 과제 관련 함수들
// =====================================================
function saveNewTask(taskData) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName("과제설정");
    const existing = sheet.getRange("B:B").getValues().flat();
    if (existing.includes(taskData.name)) return { success: false, message: "이미 같은 이름의 과제가 존재합니다." };
    sheet.appendRow([new Date(), taskData.name, taskData.desc, taskData.deadlines, taskData.evalType, taskData.isPublic ? "일괄공개" : "비공개", taskData.reqPics, taskData.choiceList]);
    const parentFolder = DriveApp.getFolderById(PARENT_FOLDER_ID);
    let taskFolder = parentFolder.createFolder(taskData.name);
    let deadlineObj = JSON.parse(taskData.deadlines);
    for (let cls in deadlineObj) taskFolder.createFolder(cls);
    clearCache(); // 과제 추가 시 캐시 초기화
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}

function updateTaskSettings(t) {
  try {
    const s = SpreadsheetApp.openById(SHEET_ID).getSheetByName("과제설정");
    const d = s.getDataRange().getValues();
    let r = -1;
    for (let i = 1; i < d.length; i++) {
      if (String(d[i][1]).trim() === t.originalName) { r = i + 1; break; }
    }
    if (r === -1) return { success: false, message: "과제를 찾을 수 없습니다." };
    s.getRange(r, 3).setValue(t.desc);
    s.getRange(r, 4).setValue(t.deadlines);
    s.getRange(r, 5).setValue(t.evalType);
    s.getRange(r, 6).setValue(t.isPublic ? "일괄공개" : "비공개");
    clearCache(); // 과제 수정 시 캐시 초기화
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}

function toggleTaskVisibility(n, p) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const ts = ss.getSheetByName("과제설정");
    const td = ts.getDataRange().getValues();
    // ✅ 과제설정 시트에서 해당 행만 업데이트
    for (let i = 1; i < td.length; i++) {
      if (String(td[i][1]).trim() === n) {
        ts.getRange(i+1, 6).setValue(p ? "일괄공개" : "비공개");
        break;
      }
    }
    // ✅ 제출현황 시트 배치 업데이트 — 전체 읽고 해당 행만 모아서 한번에 setValues
    const s = ss.getSheetByName("제출현황");
    const sd = s.getDataRange().getValues();
    const val = p ? "공개" : "비공개";
    const updates = [];
    for (let i = 1; i < sd.length; i++) {
      if (String(sd[i][3]).split(' (')[0] === n) updates.push(i + 1);
    }
    updates.forEach(function(row) { s.getRange(row, 14).setValue(val); });
    clearCache();
    return { success: true };
  } catch(e) { return { success: false }; }
}

function bulkPublishTasks(n, c) {
  try {
    const s = SpreadsheetApp.openById(SHEET_ID).getSheetByName("제출현황");
    const d = s.getDataRange().getValues();
    // ✅ 대상 행 번호 수집 후 배치 setValue
    const rows = [];
    for (let i = 1; i < d.length; i++) {
      if (!d[i][1]) continue;
      let sid = String(d[i][1]).trim();
      let sc = sid.length >= 2 ? `${sid.substring(0,1)}학년 ${sid.substring(1,2)}반` : "기타";
      let tn = String(d[i][3]).split(' (')[0];
      if ((n === "all" || tn === n) && (c === "all" || sc === c)) rows.push(i + 1);
    }
    rows.forEach(function(row) { s.getRange(row, 14).setValue("공개"); });
    return { success: true };
  } catch(e) { return { success: false }; }
}

// ✅ 과제 목록만 빠르게 로딩 (parseTasks는 5분 캐시 사용)
function getTasksOnly() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    clearCache(); // 과제 변경 후 캐시 무효화
    return { tasks: parseTasks(ss) };
  } catch(e) { return { tasks: [] }; }
}

// =====================================================
// 채점/피드백 함수들
// =====================================================
function saveFeedback(r, f, a, sc, p, m, bestType, bestAnon, bestComment) {
  try {
    const s = SpreadsheetApp.openById(SHEET_ID).getSheetByName("제출현황");
    // ✅ 피드백 + 메타 데이터 배치로 읽기 (getRange 1회)
    let rowData = s.getRange(r, 1, 1, 4).getValues()[0];
    let studentId = String(rowData[1] || "").trim();
    let taskName = String(rowData[3] || "").trim();
    if (a === "완료") {
      s.getRange(r, 11).setValue("채점완료");
      if (taskName.includes("(재제출)")) {
        let baseTaskName = taskName.split(" (재제출)")[0];
        let records = s.getDataRange().getValues();
        for (let i = records.length - 1; i >= 1; i--) {
          if (String(records[i][1]).trim() === studentId && String(records[i][3]).trim() === baseTaskName) {
            let oldStatus = String(records[i][10] || "").trim();
            if (oldStatus !== "이전기록채점완료") s.getRange(i+1, 11).setValue("이전기록채점완료");
          }
        }
      }
    } else if (a === "재제출") {
      s.getRange(r, 10).clearContent();
      s.getRange(r, 11).setValue("재제출요청");
    } else if (a === "반려검토") {
      s.getRange(r, 11).setValue("반려검토");
    } else if (a === "채점중") {
      let curStatus = String(s.getRange(r, 11).getValue() || "").trim();
      if (curStatus !== "채점완료" && curStatus !== "재제출요청" && curStatus !== "반려검토") {
        s.getRange(r, 11).setValue("채점중");
      }
    }
    s.getRange(r, 8).setValue(f);
    s.getRange(r, 13, 1, 3).setValues([[sc, p ? "공개" : "비공개", m]]);
    // ✅ 17~20열: 우수작 정보 — bestType 값이 있을 때만 저장 (기존값 보호)
    if (bestType !== undefined && bestType !== null) {
      // bestKey: 카드 채점에서는 제출된 사진 전체 키로 자동 설정
      let existingBestKey = String(s.getRange(r, 20).getValue() || "").trim();
      if (!existingBestKey || existingBestKey === '[]') {
        // 제출 사진 URL에서 키 추출
        let rawUrl = String(s.getRange(r, 7).getValue() || "").trim();
        let urls = {};
        if (rawUrl.startsWith('{')) { try { urls = JSON.parse(rawUrl); } catch(e) {} }
        let urlKeys = Object.keys(urls).filter(k => urls[k] && urls[k] !== "첨부파일 없음");
        existingBestKey = urlKeys.length > 0 ? JSON.stringify(urlKeys) : '[]';
      }
      s.getRange(r, 17, 1, 4).setValues([[
        bestType    || "",
        bestAnon    ? "TRUE" : "FALSE",
        bestComment || "",
        existingBestKey
      ]]);
    }
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}

function getSecureFileBase64(url) {
  try {
    const f = DriveApp.getFileById(url.match(/[-\w]{25,}/)[0]);
    const m = f.getMimeType().toLowerCase();
    if (m.includes("heic") || m.includes("heif")) return { success: false, message: "HEIC 불가" };
    if (f.getSize() > 10485760) return { success: false, message: "용량 초과 (10MB)" };
    const b = f.getBlob();
    // PDF는 mimeType 그대로 전달 (OpenRouter가 지원)
    return { success: true, mimeType: b.getContentType(), data: Utilities.base64Encode(b.getBytes()) };
  } catch(e) { return { success: false }; }
}

function saveMultiAnnotatedImages(rowIdx, payloadArray, studentId, studentName, taskName, feedbackText, statusAction, score, isPublic, memo, bestType, isAnon, bestComment, bestKey, perQuestionJson) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName("제출현황");
    const parentFolder = DriveApp.getFolderById(PARENT_FOLDER_ID);
    let baseTask = String(taskName || "").split(' (')[0];
    let safeStudentId = String(studentId || "");
    let className = safeStudentId.length >= 2
      ? `${safeStudentId.substring(0,1)}학년 ${safeStudentId.substring(1,2)}반` : "기타";

    // ✅ Drive 폴더 ID를 PropertiesService에 캐시 → getFoldersByName 반복 탐색 방지
    const props = PropertiesService.getScriptProperties();
    const taskFolderKey   = 'folderId_task_' + baseTask;
    const classFolderKey  = 'folderId_class_' + baseTask + '_' + className;
    let taskFolder, classFolder;

    const cachedTaskId = props.getProperty(taskFolderKey);
    try {
      taskFolder = cachedTaskId ? DriveApp.getFolderById(cachedTaskId) : null;
    } catch(e) { taskFolder = null; }
    if (!taskFolder) {
      taskFolder = parentFolder.getFoldersByName(baseTask).hasNext()
        ? parentFolder.getFoldersByName(baseTask).next()
        : parentFolder.createFolder(baseTask);
      props.setProperty(taskFolderKey, taskFolder.getId());
    }

    const cachedClassId = props.getProperty(classFolderKey);
    try {
      classFolder = cachedClassId ? DriveApp.getFolderById(cachedClassId) : null;
    } catch(e) { classFolder = null; }
    if (!classFolder) {
      classFolder = taskFolder.getFoldersByName(className).hasNext()
        ? taskFolder.getFoldersByName(className).next()
        : taskFolder.createFolder(className);
      props.setProperty(classFolderKey, classFolder.getId());
    }

    // 첨삭 이미지 저장
    let rawAnno = sheet.getRange(rowIdx, 12).getValue();
    let currentAnnoUrls = {};
    if (String(rawAnno).startsWith('{')) { try { currentAnnoUrls = JSON.parse(rawAnno); } catch(e) {} }
    payloadArray.forEach(item => {
      const data = item.base64Data.split(',')[1];
      const blob = Utilities.newBlob(
        Utilities.base64Decode(data), 'image/jpeg',
        `[${safeStudentId}] ${studentName}_${baseTask}_${item.key}_첨삭.jpg`
      );
      currentAnnoUrls[item.key] = classFolder.createFile(blob).getUrl();
    });
    sheet.getRange(rowIdx, 12).setValue(JSON.stringify(currentAnnoUrls));

    // =====================================================
    // ✅ 문항별 데이터(perQuestionData) 처리
    // =====================================================
    let pqData = {};
    // 기존 저장된 perQuestionData 읽기
    let rawPq = sheet.getRange(rowIdx, 22).getValue();
    if (rawPq && String(rawPq).startsWith('{')) {
      try { pqData = JSON.parse(rawPq); } catch(e) {}
    }
    // 새로 전달된 데이터로 덮어쓰기 (문항별 병합)
    if (perQuestionJson) {
      try {
        let newPq = JSON.parse(perQuestionJson);
        Object.keys(newPq).forEach(k => { pqData[k] = newPq[k]; });
      } catch(e) {}
    }

    // =====================================================
    // ✅ 2형 과제 여부 판단 (문항 수 > 1)
    // =====================================================
    let submittedUrls = {};
    let rawUrl = sheet.getRange(rowIdx, 7).getValue();
    if (rawUrl && String(rawUrl).startsWith('{')) {
      try { submittedUrls = JSON.parse(rawUrl); } catch(e) {}
    }
    let submittedKeys = Object.keys(submittedUrls).filter(k =>
      submittedUrls[k] && submittedUrls[k] !== "첨부파일 없음"
    );
    let isType2 = submittedKeys.length > 1;

    // =====================================================
    // ✅ 상태 결정
    // =====================================================
    let finalStatus = statusAction;

    if (isType2) {
      // 2형: pqData 기반으로 완료 여부 자동 판단
      // allKeys: 실제 제출된 사진 키
      let checkKeys = submittedKeys.length > 0 ? submittedKeys : Object.keys(pqData);
      let doneCount   = checkKeys.filter(k => pqData[k] && pqData[k].status === '완료').length;
      let rejectCount = checkKeys.filter(k => pqData[k] && pqData[k].status === '반려검토').length;
      let processedCount = doneCount + rejectCount;

      if (checkKeys.length > 0 && processedCount === checkKeys.length) {
        // 모든 문항 처리 완료
        finalStatus = rejectCount > 0 ? '반려검토' : '완료';
      } else {
        // 일부만 처리됨
        finalStatus = '채점중';
      }
    }
    // 1형은 statusAction 그대로 사용

    // =====================================================
    // ✅ 피드백 처리 (2형: 문항별 피드백 합산)
    // =====================================================
    let finalFeedback = feedbackText;
    if (isType2 && Object.keys(pqData).length > 0 && !feedbackText) {
      let fbParts = [];
      Object.keys(pqData).forEach(k => {
        if (pqData[k] && pqData[k].fb) fbParts.push(`[${k}] ${pqData[k].fb}`);
      });
      if (fbParts.length > 0) finalFeedback = fbParts.join('\n');
    }
    sheet.getRange(rowIdx, 8).setValue(finalFeedback);

    // =====================================================
    // ✅ 점수 처리 (2형: 문항별 점수 합산 또는 나열)
    // =====================================================
    let finalScore = score;
    if (isType2 && Object.keys(pqData).length > 0) {
      let scores = Object.keys(pqData)
        .map(k => pqData[k] && pqData[k].sc ? String(pqData[k].sc).trim() : "")
        .filter(s => s !== "");
      if (scores.length > 0) {
        let allNumeric = scores.every(s => !isNaN(parseFloat(s)));
        if (allNumeric) {
          finalScore = scores.reduce((a, b) => parseFloat(a) + parseFloat(b), 0).toString();
        } else {
          finalScore = scores.join(', ');
        }
      }
    }

    // =====================================================
    // ✅ 상태 저장
    // =====================================================
    if (finalStatus === '완료') {
      sheet.getRange(rowIdx, 11).setValue('채점완료');
      if (String(taskName).includes('(재제출)')) {
        let records = sheet.getDataRange().getValues();
        for (let i = records.length - 1; i >= 1; i--) {
          if (String(records[i][1]||'').trim() === safeStudentId &&
              String(records[i][3]||'').trim() === baseTask) {
            let oldSt = String(records[i][10]||'').trim();
            if (oldSt !== '이전기록채점완료') sheet.getRange(i+1, 11).setValue('이전기록채점완료');
          }
        }
      }
    } else if (finalStatus === '반려검토') {
      sheet.getRange(rowIdx, 11).setValue('반려검토');
    } else {
      // 채점중 — 이미 완료/반려 상태면 덮어쓰지 않음
      let curSt = String(sheet.getRange(rowIdx, 11).getValue() || '').trim();
      if (curSt !== '채점완료' && curSt !== '재제출요청' && curSt !== '반려검토') {
        sheet.getRange(rowIdx, 11).setValue('채점중');
      }
    }

    // ✅ 개별 setValue 대신 한 번에 배치 저장 (성능 최적화)
    // 13~15열, 17~20열, 22열을 한 번에 처리
    // 13: 점수, 14: 공개여부, 15: 메모
    sheet.getRange(rowIdx, 13, 1, 3).setValues([[
      finalScore,
      isPublic ? "공개" : "비공개",
      memo
    ]]);
    // 17~20열: 우수작 관련
    // ✅ bestType이 빈값이면 pqData에서 best 정보 추출 (2형 스마트 채점 보완)
    let finalBestType    = bestType    || "";
    let finalBestAnon    = isAnon;
    let finalBestComment = bestComment || "";
    let finalBestKey     = bestKey     || "";
    if (!finalBestType) {
      Object.keys(pqData).forEach(k => {
        var d = pqData[k];
        if (d && d.best && d.best.scope) {
          finalBestType    = d.best.scope;
          finalBestComment = d.best.comment || "";
          finalBestAnon    = !!d.best.anon;
          // key: pqData.best.key 있으면 사용, 없으면 해당 문항 키
          var bk = d.best.key || "";
          try {
            var bkArr = JSON.parse(bk);
            finalBestKey = (Array.isArray(bkArr) && bkArr.length > 0) ? bk : JSON.stringify([k]);
          } catch(e) { finalBestKey = JSON.stringify([k]); }
        }
      });
    }
    // finalBestKey도 빈배열이면 제출 사진 전체 키로
    if (finalBestType && (!finalBestKey || finalBestKey === '[]')) {
      finalBestKey = JSON.stringify(submittedKeys.length > 0 ? submittedKeys : Object.keys(pqData));
    }
    sheet.getRange(rowIdx, 17, 1, 4).setValues([[
      finalBestType,
      finalBestAnon ? "TRUE" : "FALSE",
      finalBestComment,
      finalBestKey
    ]]);
    // 22열(V열): 문항별 독립 데이터
    if (Object.keys(pqData).length > 0) {
      sheet.getRange(rowIdx, 22).setValue(JSON.stringify(pqData));
    }
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}

function cancelBestWork(r) {
  try { SpreadsheetApp.openById(SHEET_ID).getSheetByName("제출현황").getRange(r, 17, 1, 4).clearContent(); return { success: true }; }
  catch(e) { return { success: false }; }
}

// ✅ 공개 범위 변경 (학급공개 ↔ 학년공개)
function changeBestScope(rowIdx, newScope) {
  try {
    SpreadsheetApp.openById(SHEET_ID).getSheetByName("제출현황").getRange(rowIdx, 17).setValue(newScope);
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}

// =====================================================
// 설문 함수들
// =====================================================
function saveNewSurvey(payload) {
  try {
    SpreadsheetApp.openById(SHEET_ID).getSheetByName("설문목록").appendRow(["SV_" + new Date().getTime(), new Date(), payload.title, "대기중", payload.questions, payload.guide||""]);
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}

function updateSurveyStatus(svId, newStatus) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName("설문목록");
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === svId) {
        if (newStatus === "진행중") {
          for (let j = 1; j < data.length; j++) {
            if (data[j][3] === "진행중") sheet.getRange(j+1, 4).setValue("마감됨");
          }
        }
        sheet.getRange(i+1, 4).setValue(newStatus); return { success: true };
      }
    }
    return { success: false, message: "설문 찾기 실패" };
  } catch(e) { return { success: false, message: e.toString() }; }
}

// =====================================================
// 학급활동 함수
// =====================================================
function saveNewHomeroomActivity(data) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName("학급활동목록");
    if (!sheet) {
      sheet = ss.insertSheet("학급활동목록");
      sheet.getRange(1,1,1,6).setValues([['등록일','구분','활동명','설명','폼링크','서식항목']]).setFontWeight('bold');
    }
    // ✅ 6열에 fields JSON 저장
    sheet.appendRow([new Date(), data.category||'', data.name||'', data.desc||'', data.formUrl||'', data.fields||'']);
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}

// =====================================================
// AI 부정검사 (과제 선택 전수검사)
// =====================================================
function runAICheatCheck(targetTaskName) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName("제출현황");
    const data = sheet.getDataRange().getValues();
    let globalHashMap = {};
    
    // 1. 전체 사진 해시맵 구축 (과제명 무관 전체 비교)
    for (let i = 1; i < data.length; i++) {
      if (!data[i][1]) continue;
      let hashStr = data[i][8]; 
      let sid = String(data[i][1] || "").trim();
      if (hashStr && String(hashStr).startsWith("{")) {
        try {
          let hObj = JSON.parse(hashStr);
          for (let k in hObj) { 
            let h = hObj[k]; 
            if (!globalHashMap[h]) globalHashMap[h] = []; 
            globalHashMap[h].push({ rowIdx: i+1, sid: sid, name: data[i][2] }); 
          }
        } catch(e) {}
      }
    }
    
    // 2. 검사 대상 행 추출 (해당 과제의 '모든' 제출본 검사 - 완료된 것도 포함!)
    let targetRows = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][1]) continue;
      let rowTaskName = String(data[i][3] || "").split(' (')[0].trim();
      if (rowTaskName !== targetTaskName) continue;
      
      // ✅ 기존 조건 삭제: 채점완료/이전기록채점완료 등 상태 관계없이 무조건 검사
      let currentFlag = data[i][20] ? String(data[i][20]).trim() : "";
      if (currentFlag !== "") continue; // 이미 검사해서 플래그가 찍힌 애들만 건너뜀 (원하면 이 줄도 삭제 가능)
      
      targetRows.push(i);
    }
    
    if (targetRows.length === 0) return { success: true, count: 0, total: 0, taskName: targetTaskName, message: "검사할 대상이 없습니다." };
    
    let processedCount = 0;
    for (let idx = 0; idx < targetRows.length; idx++) {
      let i = targetRows[idx]; 
      let rowIdx = i + 1;
      let sid = String(data[i][1] || "").trim(); 
      let hashStr = data[i][8]; 
      let cheatReason = "";
      
      if (hashStr && String(hashStr).startsWith("{")) {
        try {
          let hObj = JSON.parse(hashStr);
          for (let k in hObj) {
            let h = hObj[k];
            if (globalHashMap[h] && globalHashMap[h].length > 1) {
              let others = globalHashMap[h].filter(x => x.sid !== sid);
              if (others.length > 0) cheatReason += `[🚨중복: ${others[0].sid} ${others[0].name} 도용] `;
            }
          }
        } catch(e) {}
      }
      
      if (cheatReason === "" && data[i][6] && String(data[i][6]).startsWith("{")) {
        try {
          let uObj = JSON.parse(data[i][6]);
          for (let k in uObj) {
            let url = uObj[k];
            if (!url || url === "첨부파일 없음") continue;
            let fIdMatch = url.match(/[-\w]{25,}/);
            if (fIdMatch && typeof Drive !== 'undefined' && Drive.Files) {
              let docFile = Drive.Files.copy({ title: "Temp_OCR", mimeType: "application/vnd.google-apps.document" }, fIdMatch[0], { ocr: true, ocrLanguage: 'ko' });
              let doc = DocumentApp.openById(docFile.id);
              let text = doc.getBody().getText().replace(/\s+/g, "");
              DriveApp.getFileById(docFile.id).setTrashed(true);
              if (text.includes("조회자") || text.includes("조회자:")) { cheatReason += `[📸캡처 적발: 워터마크 화면 캡처] `; break; }
            }
          }
        } catch(e) {}
      }
      
      sheet.getRange(rowIdx, 21).setValue(cheatReason !== "" ? cheatReason : "✅정상(AI 스캔완료)");
      processedCount++;
    }
    return { success: true, count: processedCount, total: targetRows.length, taskName: targetTaskName };
  } catch(e) { return { success: false, message: e.toString() }; }
}

// =====================================================
// ✅ 학급 알림 관리 (학급알림 시트 연동)
// 시트 구조: A=제목, B=내용, C=유형, D=날짜, E=표시여부, F=중요여부
// =====================================================
function getNotices() {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName("학급알림");
    if (!sheet) return { success: true, notices: [] };
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: true, notices: [] };
    const data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
    const notices = [];
    for (let i = 0; i < data.length; i++) {
      const r = data[i];
      if (!r[0] && !r[1]) continue; // 빈 행 스킵
      notices.push({
        rowIdx: i + 2,
        title:     String(r[0] || '').trim(),
        content:   String(r[1] || '').trim(),
        type:      String(r[2] || '알림').trim(),
        date:      r[3] instanceof Date ? Utilities.formatDate(r[3], 'Asia/Seoul', 'yyyy-MM-dd') : String(r[3] || '').trim(),
        visible:   r[4] === true || String(r[4]).trim() === '1' || String(r[4]).trim() === 'TRUE',
        important: r[5] === true || String(r[5]).trim() === '1' || String(r[5]).trim() === 'TRUE'
      });
    }
    return { success: true, notices: notices };
  } catch(e) { return { success: false, message: e.toString() }; }
}

function saveNoticeData(data) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName("학급알림");
    if (!sheet) {
      // 시트가 없으면 생성
      sheet = ss.insertSheet("학급알림");
      sheet.getRange(1, 1, 1, 6).setValues([['제목','내용','유형','날짜','표시여부','중요여부']]);
      sheet.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#f59e0b').setFontColor('white');
    }
    const row = [
      data.title || '',
      data.content || '',
      data.type || '알림',
      data.date || '',
      data.visible ? true : false,
      data.important ? true : false
    ];
    if (data.rowIdx && data.rowIdx > 1) {
      // 수정
      sheet.getRange(data.rowIdx, 1, 1, 6).setValues([row]);
    } else {
      // 새 행 추가
      const lastRow = Math.max(sheet.getLastRow(), 1);
      sheet.getRange(lastRow + 1, 1, 1, 6).setValues([row]);
    }
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}

function deleteNoticeData(rowIdx) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName("학급알림");
    if (!sheet) return { success: false, message: '시트 없음' };
    sheet.deleteRow(rowIdx);
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}

// =====================================================
// ✅ 교실 알림판 관리 (교실 알림판 앱 스프레드시트 공지사항 시트 연동)
// 시트 ID: 시스템설정 '알림판시트ID' 키
// 시트 구조: A=날짜, B=조회공지, C=조회전달, D=종례공지, E=종례전달, F=슬라이드URL, G=유튜브URL, H=영상모드
// 교실 알림판 앱은 13시 기준으로 조회/종례 자동 전환
// =====================================================
function getBoardSheetId() {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var val = _getSys(ss, '알림판시트ID');
    return val || null;
  } catch(e) { return null; }
}

function getBoardNotices() {
  try {
    const boardId = getBoardSheetId();
    if (!boardId) return { success: false, message: '시스템설정 시트 알림판시트ID 항목에 교실 알림판 스프레드시트 ID를 입력해주세요.' };
    const sheet = SpreadsheetApp.openById(boardId).getSheetByName('공지사항');
    if (!sheet) return { success: false, message: '교실 알림판 스프레드시트에 "공지사항" 시트가 없습니다.' };
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: true, boards: [] };
    const data = sheet.getRange(2, 1, lastRow - 1, 8).getDisplayValues();
    const boards = [];
    for (let i = 0; i < data.length; i++) {
      const r = data[i];
      if (!r[0]) continue;
      boards.push({
        rowIdx:         i + 2,
        date:           String(r[0] || '').trim(),
        announcement:   String(r[1] || '').trim(),  // 조회공지
        notice:         String(r[2] || '').trim(),  // 조회전달
        announcementPm: String(r[3] || '').trim(),  // 종례공지
        noticePm:       String(r[4] || '').trim(),  // 종례전달
        slideUrl:       String(r[5] || '').trim(),
        videoUrl:       String(r[6] || '').trim(),
        videoMode:      String(r[7] || '소리만').trim()
      });
    }
    boards.sort(function(a, b) { return b.date.localeCompare(a.date); });
    return { success: true, boards: boards };
  } catch(e) { return { success: false, message: e.toString() }; }
}

function saveBoardData(data) {
  try {
    const boardId = getBoardSheetId();
    if (!boardId) return { success: false, message: '시스템설정 알림판시트ID 항목에 스프레드시트 ID가 없습니다.' };
    const sheet = SpreadsheetApp.openById(boardId).getSheetByName('공지사항');
    if (!sheet) return { success: false, message: '"공지사항" 시트를 찾을 수 없습니다.' };
    const row = [
      data.date            || '',
      data.announcement    || '',
      data.notice          || '',
      data.announcementPm  || '',
      data.noticePm        || '',
      data.slideUrl        || '',
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
    const boardId = getBoardSheetId();
    if (!boardId) return { success: false, message: 'ID 없음' };
    SpreadsheetApp.openById(boardId).getSheetByName('공지사항').deleteRow(rowIdx);
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}

// =====================================================
// ✅ 학생 관찰기록
// 시트: 학생관찰기록 (자동 생성)
// 구조: A=날짜, B=학번, C=이름, D=카테고리, E=내용, F=구분(수학수업/담임학급)
// =====================================================
function getOrCreateObserveSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName('학생관찰기록');
  if (!sheet) {
    sheet = ss.insertSheet('학생관찰기록');
    const headers = ['날짜', '학번', '이름', '카테고리', '내용', '구분'];
    sheet.getRange(1, 1, 1, 6).setValues([headers])
      .setFontWeight('bold').setBackground('#10b981').setFontColor('white');
    sheet.setColumnWidth(5, 400);
  } else {
    // ✅ F열 헤더 없으면 추가 (기존 시트 마이그레이션)
    const h = sheet.getRange(1, 6).getValue();
    if (!h) sheet.getRange(1, 6).setValue('구분').setFontWeight('bold').setBackground('#10b981').setFontColor('white');
  }
  return sheet;
}

function getObserveRecords() {
  try {
    const sheet = getOrCreateObserveSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: true, records: [] };
    const data = sheet.getRange(2, 1, lastRow - 1, 6).getValues(); // F열까지
    const records = [];
    for (let i = 0; i < data.length; i++) {
      const r = data[i];
      if (!r[1] && !r[4]) continue;
      records.push({
        rowIdx:      i + 2,
        date:        r[0] instanceof Date ? Utilities.formatDate(r[0], 'Asia/Seoul', 'yyyy-MM-dd') : String(r[0] || '').trim(),
        studentId:   String(r[1] || '').trim(),
        studentName: String(r[2] || '').trim(),
        category:    String(r[3] || '').trim(),
        content:     String(r[4] || '').trim(),
        type:        String(r[5] || '담임학급').trim() // ✅ 빈값은 담임학급으로 처리
      });
    }
    records.sort(function(a, b) { return b.date.localeCompare(a.date); });
    return { success: true, records: records };
  } catch(e) { return { success: false, message: e.toString() }; }
}

function saveObserveRecord(data) {
  try {
    const sheet = getOrCreateObserveSheet();
    const row = [
      data.date        || '',
      data.studentId   || '',
      data.studentName || '',
      data.category    || '',
      data.content     || '',
      data.type        || '담임학급'  // ✅ F열: 수학수업 / 담임학급
    ];
    if (data.rowIdx && data.rowIdx > 1) {
      sheet.getRange(data.rowIdx, 1, 1, 6).setValues([row]);
    } else {
      sheet.getRange(Math.max(sheet.getLastRow(), 1) + 1, 1, 1, 6).setValues([row]);
    }
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}

// 여러 명의 관찰 기록을 한 번에 저장 (다중 선택 대응)
function saveMultipleObserveRecords(dataArray) {
  try {
    const sheet = getOrCreateObserveSheet();
    const rows = dataArray.map(data => [
      data.date || '',
      data.studentId || '',
      data.studentName || '',
      data.category || '',
      data.content || '',
      data.type || '담임학급'
    ]);
    
    // 배열 데이터를 2차원 배열 형태로 한꺼번에 append
    const lastRow = Math.max(sheet.getLastRow(), 1);
    sheet.getRange(lastRow + 1, 1, rows.length, 6).setValues(rows);
    
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}

function deleteObserveRecord(rowIdx) {
  try {
    getOrCreateObserveSheet().deleteRow(rowIdx);
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}

// =====================================================
// ✅ 신고 접수함 조회 / 처리 완료 표시
// =====================================================
function getReportInbox() {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('신고접수');
    if (!sheet || sheet.getLastRow() < 2) return { success: true, reports: [] };
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
    const reports = [];
    for (let i = 0; i < data.length; i++) {
      const r = data[i];
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
    SpreadsheetApp.openById(SHEET_ID).getSheetByName('신고접수').getRange(rowIdx, 6).setValue('처리완료');
    return { success: true };
  } catch(e) { return { success: false }; }
}

// =====================================================
// ✅ 바로가기 URL 저장/조회 (시스템설정 키-값 구조)
// 키: 바로가기_수학교실, 바로가기_우리반교실, 바로가기_교실알림판, 바로가기_회원가입, 바로가기_포털
// =====================================================
function getAppLinkUrl(key) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    return { url: _getSys(ss, key) };
  } catch(e) { return { url: '' }; }
}

function saveAppLinkUrl(key, url) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    _setSys(ss, key, url);
    return { success: true };
  } catch(e) { return { success: false }; }
}


// =====================================================
// ✅ 설문 데이터 별도 로딩 (탭 진입 시에만 호출)
// =====================================================
function getSurveyData() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let surveys = [], surveyRes = [];
    const svSheet = ss.getSheetByName("설문목록");
    if (svSheet && svSheet.getLastRow() >= 2) {
      let svData = svSheet.getDataRange().getValues();
      for (let i = 1; i < svData.length; i++) {
        if (svData[i][0]) surveys.push({
          id: svData[i][0],
          date: svData[i][1] ? Utilities.formatDate(new Date(svData[i][1]), "Asia/Seoul", "yyyy-MM-dd") : "",
          title: svData[i][2], status: svData[i][3], questions: svData[i][4],
          guide: String(svData[i][5] || '').trim()
        });
      }
    }
    const srSheet = ss.getSheetByName("설문응답");
    if (srSheet && srSheet.getLastRow() >= 2) {
      let srData = srSheet.getDataRange().getValues();
      for (let i = 1; i < srData.length; i++) {
        if (srData[i][1]) surveyRes.push({
          date: srData[i][0] ? Utilities.formatDate(new Date(srData[i][0]), "Asia/Seoul", "MM/dd HH:mm") : "",
          svId: srData[i][1], stuId: srData[i][2], stuName: srData[i][3], answers: srData[i][4]
        });
      }
    }
    return { success: true, surveys: surveys.reverse(), surveyRes };
  } catch(e) { return { success: false, message: e.toString() }; }
}

// =====================================================
// ✅ 학급활동 데이터 별도 로딩 (우리반활동 탭 진입 시)
// =====================================================
function getHomeroomTabData() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let hrActivities = [], hrSubmissions = [];
    const hrActSheet = ss.getSheetByName("학급활동목록");
    if (hrActSheet && hrActSheet.getLastRow() >= 2) {
      let actData = hrActSheet.getDataRange().getValues();
      for (let i = 1; i < actData.length; i++) {
        if (actData[i][2]) hrActivities.push({
          category: actData[i][1], name: actData[i][2],
          desc: actData[i][3], formUrl: actData[i][4] ? String(actData[i][4]).trim() : "",
          fields: actData[i][5] ? String(actData[i][5]).trim() : ""
        });
      }
    }
    const hrSubSheet = ss.getSheetByName("창체제출현황");
    if (hrSubSheet && hrSubSheet.getLastRow() >= 2) {
      let hrData = hrSubSheet.getDataRange().getValues();
      for (let i = 1; i < hrData.length; i++) {
        if (hrData[i][1]) hrSubmissions.push({
          rowIdx: i+1,
          date: hrData[i][0] ? Utilities.formatDate(new Date(hrData[i][0]), "Asia/Seoul", "MM/dd HH:mm") : "",
          id: String(hrData[i][1]||"").trim(), name: String(hrData[i][2]||"").trim(),
          actName: String(hrData[i][3]||"").trim(), activity: String(hrData[i][3]||"").trim(), role: String(hrData[i][4]||"").trim(),
          reflection: String(hrData[i][5]||"").trim(), content: String(hrData[i][5]||"").trim(), url: String(hrData[i][6]||"").trim()
        });
      }
    }
    return { success: true, hrActivities: hrActivities.reverse(), hrSubmissions: hrSubmissions.reverse() };
  } catch(e) { return { success: false, message: e.toString() }; }
}

// =====================================================
// ✅ 캘린더 메모 (조회/종례/수업기록)
// =====================================================
function getCalendarMemos() {
  try {
    // ✅ 30초 캐시
    const cache = CacheService.getScriptCache();
    const cached = cache.get('cal_memos_v1');
    if (cached) return JSON.parse(cached);

    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName('캘린더메모');
    if (!sheet) {
      sheet = ss.insertSheet('캘린더메모');
      sheet.getRange(1, 1, 1, 4).setValues([['날짜', '유형', '내용', '색상']]);
    }
    const last = sheet.getLastRow();
    if (last < 2) return { success: true, memos: [] };
    const data = sheet.getRange(2, 1, last - 1, 4).getValues();
    const memos = data.filter(r => r[0]).map((r, i) => ({
      rowIdx: i + 2,
      date:    r[0] instanceof Date ? Utilities.formatDate(r[0], 'Asia/Seoul', 'yyyy-MM-dd') : String(r[0]).trim().substring(0, 10),
      type:    String(r[1] || '메모').trim(),
      content: String(r[2] || '').trim(),
      color:   String(r[3] || '#6366f1').trim()
    }));
    const result = { success: true, memos };
    try { cache.put('cal_memos_v1', JSON.stringify(result), 30); } catch(e) {}
    return result;
  } catch(e) { return { success: false, message: e.toString() }; }
}

function saveCalendarMemo(data) {
  try { CacheService.getScriptCache().remove('cal_memos_v1'); } catch(e) {}
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName('캘린더메모');
    if (!sheet) {
      sheet = ss.insertSheet('캘린더메모');
      sheet.getRange(1, 1, 1, 4).setValues([['날짜', '유형', '내용', '색상']]);
    }
    const row = [data.date || '', data.type || '메모', data.content || '', data.color || '#6366f1'];
    if (data.rowIdx && data.rowIdx > 1) {
      sheet.getRange(data.rowIdx, 1, 1, 4).setValues([row]);
    } else {
      sheet.appendRow(row);
    }
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}

function deleteCalendarMemo(rowIdx) {
  try { CacheService.getScriptCache().remove('cal_memos_v1'); } catch(e) {}
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('캘린더메모');
    if (sheet && rowIdx > 1) sheet.deleteRow(rowIdx);
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}

// =====================================================
// ✅ AI 업무 분석 + TickTick 연동
// =====================================================

// 시스템설정에서 API 설정값 읽기
function getApiSettings() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var m  = _getSys(ss, 'AI모델명');
  // 구 모델명 자동 교정
  if (m === 'google/gemini-2.5-flash-preview') m = 'google/gemini-2.5-flash';
  if (m === 'google/gemini-2.5-pro-preview')   m = 'google/gemini-2.5-pro';
  return {
    openrouterKey: _getSys(ss, 'OpenRouter키'),
    model:         m || 'anthropic/claude-3.5-sonnet',
    ttClientId:    _getSys(ss, 'TickTick클라이언트ID'),
    ttSecret:      _getSys(ss, 'TickTick시크릿'),
    ttToken:       _getSys(ss, 'TickTick액세스토큰'),
    ttRefresh:     _getSys(ss, 'TickTick갱신토큰')
  };
}

// AI 모델 설정값 저장 (키 이름 기반)
function saveApiSetting(key, value) {
  try {
    // 하위 호환: 'Q2' 셀 주소로 들어오면 AI모델명 키로 변환
    var resolvedKey = key;
    if (key === 'Q2') resolvedKey = 'AI모델명';
    var ss = SpreadsheetApp.openById(SHEET_ID);
    _setSys(ss, resolvedKey, value);
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}

function getOpenRouterBalance() {
  try {
    const key = _getSys(SpreadsheetApp.openById(SHEET_ID), 'OpenRouter키');
    if (!key) return { success: false, message: 'API 키 없음' };
    const res = UrlFetchApp.fetch('https://openrouter.ai/api/v1/auth/key', {
      headers: { 'Authorization': 'Bearer ' + key },
      muteHttpExceptions: true
    });
    if (res.getResponseCode() !== 200) return { success: false, message: '조회 실패 (' + res.getResponseCode() + ')' };
    const d = JSON.parse(res.getContentText()).data;
    return { success: true, usage: d.usage, limit: d.limit, isFreeTier: d.is_free_tier };
  } catch(e) { return { success: false, message: e.toString() }; }
}

// ─── Claude via OpenRouter 텍스트 분석 ─────────────────
function analyzeTaskText(text, fileAttachment) {
  try {
    const cfg = getApiSettings();
    if (!cfg.openrouterKey) return { success: false, message: '시스템설정 OpenRouter키 항목에 API 키를 입력해주세요.' };

    const today = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy년 MM월 dd일 (E)');

    const systemPrompt = '너는 한국 교사의 업무를 도와주는 AI야. 공문, 메시지, 일정이 적힌 사진, 첨부파일을 분석해서 반드시 유효한 JSON만 반환해. 사진이 첨부되면 사진 속 텍스트를 꼼꼼히 읽어서 날짜·시간·할 일을 모두 추출해. 다른 텍스트, 마크다운 코드블록, 설명은 절대 포함하지 마.';

    const userPrompt = '오늘 날짜: ' + today + '\n\n'
      + '아래 공문/메시지를 분석하고 정확히 다음 JSON 형식으로 반환해.\n\n'
      + '⚠️ 문체 규칙 (반드시 준수):\n'
      + '- 모든 텍스트 필드는 개조식으로 작성\n'
      + '- "합니다/입니다/됩니다/있습니다" 등 문장형 종결어미 절대 금지\n'
      + '- 명사형 또는 동사 원형으로 짧게 끝낼 것 (예: "제출", "확인 필요", "양식 작성")\n\n'
      + '기타 규칙:\n'
      + '- deadline: 날짜 언급 있으면 yyyy-MM-dd, 없으면 null\n'
      + '- time: 시간 언급 있으면 HH:mm (24시간), 없으면 null\n'
      + '- priority: 5=매우긴급, 3=보통, 1=낮음\n'
      + '- project: 학교업무/연수/학급운영/행사/제출/회의/기타 중 하나\n'
      + '- tags: 핵심 키워드 최대 3개 (명사, 한글)\n'
      + '- checklist: 세부 단계 있으면 배열, 없으면 빈 배열\n'
      + '- repeat: 반복 언급 있으면 RRULE 형식, 없으면 null\n\n'
      + '{\n'
      + '  "summary": "핵심 내용 2~3줄 개조식 요약 (문장 아닌 항목 나열)",\n'
      + '  "tasks": [\n'
      + '    {\n'
      + '      "title": "할 일 제목 (10자 이내, 명사형)",\n'
      + '      "deadline": "yyyy-MM-dd 또는 null",\n'
      + '      "time": "HH:mm 또는 null",\n'
      + '      "detail": "핵심 내용만 개조식으로 (2~3항목, 각 항목 20자 이내)",\n'
      + '      "priority": 1~5,\n'
      + '      "project": "분류명",\n'
      + '      "tags": ["태그1", "태그2"],\n'
      + '      "checklist": ["단계1", "단계2"],\n'
      + '      "repeat": "RRULE 또는 null"\n'
      + '    }\n'
      + '  ],\n'
      + '  "cautions": ["주의사항 개조식"]\n'
      + '}\n\n'
      + '[분석할 내용]\n' + text;

    // 파일 첨부 처리 (OpenRouter OpenAI 호환 형식 사용)
    let userContent;
    if (fileAttachment && fileAttachment.data) {
      const ext = String(fileAttachment.name || '').split('.').pop().toLowerCase();
      const isImage = ['png','jpg','jpeg','gif','webp','heic','heif'].includes(ext);
      const isPdf   = ext === 'pdf';
      if (isImage) {
        const mimeType = fileAttachment.mimeType && fileAttachment.mimeType.startsWith('image/')
          ? fileAttachment.mimeType : 'image/jpeg';
        userContent = [
          { type: 'image_url', image_url: { url: 'data:' + mimeType + ';base64,' + fileAttachment.data } },
          { type: 'text', text: userPrompt + (text ? '\n\n[추가 텍스트]\n' + text : '') }
        ];
      } else if (isPdf) {
        // PDF는 URL 방식으로 전달 불가능 → 텍스트 추출 시도 후 실패 시 안내
        userContent = userPrompt + '\n\n[PDF 첨부: ' + fileAttachment.name + ' - 텍스트 내용을 직접 붙여넣어 주세요]';
      } else {
        try {
          const decoded = Utilities.newBlob(Utilities.base64Decode(fileAttachment.data)).getDataAsString('UTF-8');
          userContent = userPrompt + '\n\n[첨부 파일: ' + fileAttachment.name + ']\n' + decoded.substring(0, 8000);
        } catch(e) {
          userContent = userPrompt + '\n\n[파일 첨부됨: ' + fileAttachment.name + ' - 텍스트 추출 실패]';
        }
      }
    } else {
      userContent = userPrompt;
    }

    const payload = {
      model: cfg.model,
      messages: [
        { role: 'system', content: systemPrompt },
        { role: 'user',   content: userContent }
      ],
      max_tokens: 2000,
      temperature: 0.3
    };

    const res = UrlFetchApp.fetch('https://openrouter.ai/api/v1/chat/completions', {
      method: 'post',
      headers: {
        'Authorization':  'Bearer ' + cfg.openrouterKey,
        'Content-Type':   'application/json',
        'HTTP-Referer':   'https://script.google.com',
        'X-Title':        'Teacher Dashboard'
      },
      payload:           JSON.stringify(payload),
      muteHttpExceptions: true
    });

    const code = res.getResponseCode();
    if (code !== 200) {
      const errBody = JSON.parse(res.getContentText());
      const errMsg  = (errBody.error && errBody.error.message) ? errBody.error.message : res.getContentText();
      return { success: false, message: 'API 오류 (' + code + '): ' + errMsg };
    }

    const body    = JSON.parse(res.getContentText());
    let   content = body.choices[0].message.content.trim();

    // 마크다운 코드블록 제거 (```json ... ```)
    content = content.replace(/^```[a-z]*\n?/i, '').replace(/\n?```$/,'').trim();

    const parsed = JSON.parse(content);
    return { success: true, result: parsed, model: cfg.model };

  } catch(e) {
    return { success: false, message: '분석 오류: ' + e.toString() };
  }
}

// ─── TickTick OAuth ────────────────────────────────────
const TT_REDIRECT_URI = 'https://example.com';

function getTickTickAuthUrl() {
  try {
    const cfg = getApiSettings();
    if (!cfg.ttClientId) return { success: false, message: '시스템설정 TickTick클라이언트ID 항목에 Client ID를 입력해주세요.' };

    const authUrl = 'https://ticktick.com/oauth/authorize'
      + '?client_id='    + encodeURIComponent(cfg.ttClientId)
      + '&scope='        + encodeURIComponent('tasks:read tasks:write')
      + '&redirect_uri=' + encodeURIComponent(TT_REDIRECT_URI)
      + '&response_type=code';

    return { success: true, authUrl: authUrl };
  } catch(e) { return { success: false, message: e.toString() }; }
}

// 인증 코드를 받아서 토큰 교환
function exchangeTickTickCode(code) {
  try {
    const cfg        = getApiSettings();
    const credential = Utilities.base64Encode(cfg.ttClientId + ':' + cfg.ttSecret);

    const res = UrlFetchApp.fetch('https://ticktick.com/oauth/token', {
      method: 'post',
      headers: {
        'Authorization': 'Basic ' + credential,
        'Content-Type':  'application/x-www-form-urlencoded'
      },
      payload: 'code='         + encodeURIComponent(code)
             + '&grant_type=authorization_code'
             + '&redirect_uri=' + encodeURIComponent(TT_REDIRECT_URI),
      muteHttpExceptions: true
    });

    const token = JSON.parse(res.getContentText());
    if (!token.access_token) {
      return { success: false, message: '코드가 잘못됐거나 만료됐어요: ' + res.getContentText() };
    }

    const _ttSs = SpreadsheetApp.openById(SHEET_ID);
    _setSys(_ttSs, 'TickTick액세스토큰', token.access_token);
    if (token.refresh_token) _setSys(_ttSs, 'TickTick갱신토큰', token.refresh_token);

    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}

function refreshTickTickToken() {
  try {
    const cfg        = getApiSettings();
    if (!cfg.ttRefresh) return { success: false, message: 'Refresh Token 없음. 재인증 필요.' };
    const credential = Utilities.base64Encode(cfg.ttClientId + ':' + cfg.ttSecret);

    const res = UrlFetchApp.fetch('https://ticktick.com/oauth/token', {
      method: 'post',
      headers: {
        'Authorization': 'Basic ' + credential,
        'Content-Type':  'application/x-www-form-urlencoded'
      },
      payload: 'grant_type=refresh_token&refresh_token=' + encodeURIComponent(cfg.ttRefresh),
      muteHttpExceptions: true
    });

    const token = JSON.parse(res.getContentText());
    if (!token.access_token) return { success: false, message: '갱신 실패. 재인증 필요.' };

    const _ttSs2 = SpreadsheetApp.openById(SHEET_ID);
    _setSys(_ttSs2, 'TickTick액세스토큰', token.access_token);
    if (token.refresh_token) _setSys(_ttSs2, 'TickTick갱신토큰', token.refresh_token);

    return { success: true, token: token.access_token };
  } catch(e) { return { success: false, message: e.toString() }; }
}

// ─── TickTick 태스크 추가 ──────────────────────────────
function addToTickTick(taskData) {
  try {
    let cfg   = getApiSettings();
    let token = cfg.ttToken;
    if (!token)          return { success: false, message: 'TickTick 미연동. 설정에서 먼저 연동해주세요.' };
    if (!cfg.ttClientId) return { success: false, message: '시스템설정 TickTick클라이언트ID 항목에 Client ID를 입력해주세요.' };

    function postTask(accessToken) {
      const body = {
        title:    taskData.title,
        content:  taskData.detail || '',
        priority: taskData.priority || 3
      };

      // 마감일 + 시간
      if (taskData.deadline) {
        const timeStr = taskData.time || '09:00';
        body.dueDate = taskData.deadline + 'T' + timeStr + ':00+0900';
      }

      // 프로젝트
      if (taskData.projectId) {
        body.projectId = taskData.projectId;
      }

      // 태그
      if (taskData.tags && taskData.tags.length > 0) {
        body.tags = taskData.tags;
      }

      // 체크리스트
      if (taskData.checklist && taskData.checklist.length > 0) {
        body.items = taskData.checklist.map(function(item, i) {
          return { title: item, status: 0, sortOrder: i };
        });
      }

      // 반복
      if (taskData.repeat) {
        body.repeatFlag = taskData.repeat;
      }

      return UrlFetchApp.fetch('https://ticktick.com/open/v1/task', {
        method: 'post',
        headers: {
          'Authorization': 'Bearer ' + accessToken,
          'Content-Type':  'application/json'
        },
        payload:            JSON.stringify(body),
        muteHttpExceptions: true
      });
    }

    let res  = postTask(token);
    let code = res.getResponseCode();

    // 401이면 토큰 만료 → 자동 갱신 후 재시도
    if (code === 401) {
      const refreshed = refreshTickTickToken();
      if (!refreshed.success) return { success: false, message: '토큰 만료. ' + refreshed.message };
      res  = postTask(refreshed.token);
      code = res.getResponseCode();
    }

    if (code === 200 || code === 201) {
      return { success: true };
    } else {
      return { success: false, message: 'TickTick 오류 (' + code + '): ' + res.getContentText() };
    }
  } catch(e) { return { success: false, message: e.toString() }; }
}


// TickTick 프로젝트 목록 가져오기
function getTickTickProjects() {
  try {
    let cfg   = getApiSettings();
    let token = cfg.ttToken;
    if (!token) return { success: false, message: 'TickTick 미연동' };

    function fetchProjects(accessToken) {
      return UrlFetchApp.fetch('https://ticktick.com/open/v1/project', {
        method: 'get',
        headers: { 'Authorization': 'Bearer ' + accessToken },
        muteHttpExceptions: true
      });
    }

    let res  = fetchProjects(token);
    if (res.getResponseCode() === 401) {
      const refreshed = refreshTickTickToken();
      if (!refreshed.success) return { success: false, message: refreshed.message };
      res = fetchProjects(refreshed.token);
    }

    const projects = JSON.parse(res.getContentText());
    return {
      success:  true,
      projects: projects.map(function(p) {
        return { id: p.id, name: p.name, color: p.color || '#6366f1' };
      })
    };
  } catch(e) { return { success: false, message: e.toString() }; }
}

// TickTick 연동 상태 확인
function getTickTickStatus() {
  try {
    const cfg = getApiSettings();
    return {
      success:   true,
      connected: !!(cfg.ttToken),
      model:     cfg.model || 'anthropic/claude-3.5-sonnet',
      hasApiKey: !!(cfg.openrouterKey)
    };
  } catch(e) { return { success: false, connected: false }; }
}

// =====================================================
// ✅ AI 자동 채점
// =====================================================

// 채점기준 시트 생성/조회
function getOrCreateRubricSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName('AI채점기준');
  if (!sh) {
    sh = ss.insertSheet('AI채점기준');
    sh.getRange(1,1,1,6).setValues([['과제명','채점유형','총점','기준설명','파일목록JSON','문항구성JSON']]);
  }
  return sh;
}

// 채점기준 목록 조회
function getRubrics() {
  try {
    const sh   = getOrCreateRubricSheet();
    const last = sh.getLastRow();
    if (last < 2) return { success: true, rubrics: [] };
    const data = sh.getRange(2, 1, last-1, 6).getValues();
    const rubrics = data.filter(r => r[0]).map((r, i) => {
      let files = [];
      try { if (r[4]) files = JSON.parse(r[4]); } catch(e) {}
      let questions = {};
      try { if (r[5]) questions = JSON.parse(r[5]); } catch(e) {}
      return {
        rowIdx:    i + 2,
        taskName:  String(r[0]||'').trim(),
        evalType:  String(r[1]||'점수제').trim(),
        maxScore:  Number(r[2]||0),
        criteria:  String(r[3]||'').trim(),
        files:     files,
        questions: questions
      };
    });
    return { success: true, rubrics };
  } catch(e) { return { success: false, message: e.toString() }; }
}


// 채점기준 파일을 base64로 받아 Drive에 저장하고 URL 반환
function uploadRubricFile(fileName, mimeType, base64Data, taskName) {
  try {
    const bytes = Utilities.base64Decode(base64Data);
    const blob  = Utilities.newBlob(bytes, mimeType, fileName);
    
    // PARENT_FOLDER_ID 안에 'AI채점기준' 폴더 생성/조회
    let rubricFolder;
    const folders = DriveApp.getFolderById(PARENT_FOLDER_ID)
                             .getFoldersByName('AI채점기준');
    if (folders.hasNext()) {
      rubricFolder = folders.next();
    } else {
      rubricFolder = DriveApp.getFolderById(PARENT_FOLDER_ID)
                             .createFolder('AI채점기준');
    }
    
    // 기존 동명 파일 삭제
    const existing = rubricFolder.getFilesByName(taskName + '_' + fileName);
    while (existing.hasNext()) existing.next().setTrashed(true);
    
    const file = rubricFolder.createFile(blob);
    file.setName(taskName + '_' + fileName);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    return { success: true, url: file.getUrl(), fileId: file.getId(), name: fileName };
  } catch(e) { return { success: false, message: e.toString() }; }
}

// 채점기준 저장
function saveRubric(data) {
  try {
    const sh  = getOrCreateRubricSheet();
    const row = [
      data.taskName||'',
      data.evalType||'점수제',
      data.maxScore||0,
      data.criteria||'',
      data.files ? JSON.stringify(data.files) : '',
      data.questions ? JSON.stringify(data.questions) : ''
    ];
    if (data.rowIdx && data.rowIdx > 1) {
      sh.getRange(data.rowIdx, 1, 1, 6).setValues([row]);
    } else {
      // 같은 과제명 있으면 덮어쓰기
      const last = sh.getLastRow();
      if (last >= 2) {
        const names = sh.getRange(2,1,last-1,1).getValues().flat();
        const idx   = names.indexOf(data.taskName);
        if (idx >= 0) {
          sh.getRange(idx+2, 1, 1, 5).setValues([row]);
          return { success: true, rowIdx: idx+2 };
        }
      }
      sh.appendRow(row);
    }
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}


// AI 채점 확정 시 perQuestionData + 피드백 + 점수 한 번에 저장
// AI 채점 결과 W열에 임시저장 (교사 확정 전)
function saveAiGradeTempResult(rowIdx, aiResultJson) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('제출현황');
    if (!sheet || rowIdx < 2) return { success: false, message: '시트 없음' };
    sheet.getRange(rowIdx, 23).setValue(aiResultJson || ''); // W열: AI임시저장
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}

// 교사 확정 저장 (피드백+점수+상태 → 실제 열에 저장, W열 비우기)
function saveAiGradeResult(rowIdx, feedback, score, statusAction, perQuestionJson) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('제출현황');
    if (!sheet || rowIdx < 2) return { success: false, message: '시트 없음' };
    const statusMap = { 'pass': '채점완료', 'reject': '재제출요청' };
    const newStatus = statusMap[statusAction] || statusAction;
    sheet.getRange(rowIdx, 8).setValue(feedback || '');
    sheet.getRange(rowIdx, 11).setValue(newStatus);
    sheet.getRange(rowIdx, 13).setValue(score || '');
    if (perQuestionJson) {
      sheet.getRange(rowIdx, 22).setValue(perQuestionJson);
    }
    // ✅ W열 초기화 제거 — 학생도 AI 결과를 계속 볼 수 있도록 유지
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}

// 채점기준 삭제
function deleteRubric(rowIdx) {
  try {
    getOrCreateRubricSheet().deleteRow(rowIdx);
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}

// Drive 파일 → base64 (채점기준용, PDF 포함)
function getRubricFileBase64(url) {
  try {
    if (!url || !url.includes('drive.google.com')) return { success: false, message: 'Drive URL이 아닙니다.' };
    const fileId = url.match(/[-\w]{25,}/);
    if (!fileId) return { success: false, message: 'File ID 추출 실패' };
    const f    = DriveApp.getFileById(fileId[0]);
    const mime = f.getMimeType();
    if (f.getSize() > 10485760) return { success: false, message: '파일 크기 10MB 초과' };
    const b64  = Utilities.base64Encode(f.getBlob().getBytes());
    return { success: true, mimeType: mime, data: b64, name: f.getName() };
  } catch(e) { return { success: false, message: e.toString() }; }
}

// ─── 핵심: AI 채점 실행 ────────────────────────────
function aiGradeSubmission(params) {
  /*
    params: {
      rowIdx,        // 제출현황 rowIdx
      taskName,      // 과제명
      studentId,     // 학번
      studentName,   // 이름
      imageUrls,     // { "1번": "url", ... } 또는 { "0": "url" }
      rubricRowIdx,  // AI채점기준 rowIdx
      isType2        // 2형 여부
    }
  */
  try {
    const cfg = getApiSettings();
    if (!cfg.openrouterKey) return { success: false, message: 'OpenRouter API 키가 없습니다.' };

    // 채점기준 가져오기
    const rubricRes = getRubrics();
    if (!rubricRes.success) return { success: false, message: '채점기준 조회 실패' };
    const rubric = rubricRes.rubrics.find(r => r.taskName === params.taskName || r.rowIdx === params.rubricRowIdx);
    if (!rubric) return { success: false, message: '해당 과제의 채점기준이 없습니다.' };

    // 메시지 구성
    const today = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');
    const messages = [];

    // content 배열 구성
    const content = [];

    // 텍스트 프롬프트
    const isType2 = params.isType2;
    const prompt = [
      '너는 수학 교사의 채점을 돕는 AI야. 반드시 JSON만 반환해.',
      '학생: ' + params.studentId + ' ' + params.studentName,
      '과제: ' + params.taskName,
      '채점유형: ' + rubric.evalType,
      '총점: ' + rubric.maxScore + '점',
      rubric.criteria ? '채점기준 설명: ' + rubric.criteria : '',
      // 문항별 채점기준 파일이 있는 경우 안내
      (rubric.files||[]).some(f=>f.questions&&f.questions.length>0)
        ? '※ 일부 파일은 특정 문항 전용 채점기준임. 파일 레이블을 확인하여 해당 문항만 적용할 것.' : '',
      // 성취기준 파일이 있는 경우 안내
      (rubric.files||[]).some(f=>f.role==='achievement')
        ? '※ 성취기준 파일을 참고하여 A/B/C/D 판별할 것.' : '',
      rubric.imageUrl ? '(채점기준 이미지/PDF 첨부 참고)' : '',
      '반환 형식:',
      (() => {
        const et = rubric.evalType || '점수제';
        if (et.includes('A-B-C') || et.includes('등급')) {
          return '{"grade":"A|B|C|D","feedback":"피드백","confidence":"high|medium|low","needsReview":true|false,"rejectReason":"사유 또는 null"}';
        } else if (et.includes('상-중-하')) {
          return '{"grade":"상|중|하","feedback":"피드백","confidence":"high|medium|low","needsReview":true|false,"rejectReason":"사유 또는 null"}';
        } else if (et.includes('P/F') || et.includes('통과')) {
          return '{"grade":"Pass|Fail","feedback":"피드백","confidence":"high|medium|low","needsReview":true|false,"rejectReason":"사유 또는 null"}';
        } else if (isType2) {
          return '{"perQuestion":{"문항명":{"score":점수,"maxScore":만점,"feedback":"피드백"}},"totalScore":합계,"overallFeedback":"종합피드백","confidence":"high|medium|low","needsReview":true|false,"rejectReason":"사유 또는 null"}';
        } else {
          return '{"score":점수,"maxScore":' + rubric.maxScore + ',"feedback":"피드백","confidence":"high|medium|low","needsReview":true|false,"rejectReason":"사유 또는 null"}';
        }
      })()
    ].filter(Boolean).join('\n');

    content.push({ type: 'text', text: prompt });

    // 채점기준 파일 첨부 (role/scope/questions 적용)
    const rubricFiles = rubric.files || [];
    for (const rf of rubricFiles) {
      if (!rf.url) continue;
      // 문항 필터링 - 특정 문항 전용 파일이면 해당 문항만
      const targetQ = rf.questions && rf.questions.length > 0 ? rf.questions : null;
      const roleLabel = rf.role === 'achievement' ? '[성취기준 파일]' : '[채점기준 파일: ' + (targetQ ? targetQ.join(', ') : '전체') + ']';
      content.push({ type: 'text', text: roleLabel });
      const rubricFile = getRubricFileBase64(rf.url);
      if (rubricFile.success) {
        content.push({ type: 'image_url', image_url: { url: 'data:' + rubricFile.mimeType + ';base64,' + rubricFile.data } });
      } else {
        content.push({ type: 'text', text: '(파일 로드 실패: ' + rubricFile.message + ')' });
      }
    }

    // 학생 제출 사진 첨부
    const urlKeys = Object.keys(params.imageUrls || {});
    for (let i = 0; i < urlKeys.length; i++) {
      const key     = urlKeys[i];
      const fileUrl = params.imageUrls[key];
      content.push({ type: 'text', text: (isType2 ? '[' + key + ' 답안 사진]' : '[학생 답안 사진]') });
      const imgRes = getSecureFileBase64(fileUrl);
      if (imgRes.success) {
        content.push({ type: 'image_url', image_url: { url: 'data:' + imgRes.mimeType + ';base64,' + imgRes.data } });
      } else {
        content.push({ type: 'text', text: '(사진 로드 실패: ' + imgRes.message + ')' });
      }
    }

    messages.push({ role: 'user', content: content });

    // API 호출
    const res = UrlFetchApp.fetch('https://openrouter.ai/api/v1/chat/completions', {
      method: 'post',
      headers: {
        'Authorization':  'Bearer ' + cfg.openrouterKey,
        'Content-Type':   'application/json',
        'HTTP-Referer':   'https://script.google.com',
        'X-Title':        'Teacher Dashboard'
      },
      payload: JSON.stringify({
        model:       cfg.model || 'google/gemini-2.5-flash',
        messages:    messages,
        max_tokens:  1000,
        temperature: 0.2
      }),
      muteHttpExceptions: true
    });

    const code = res.getResponseCode();
    if (code !== 200) {
      const err = JSON.parse(res.getContentText());
      return { success: false, message: 'API 오류(' + code + '): ' + ((err.error && err.error.message) || res.getContentText()) };
    }

    let text = JSON.parse(res.getContentText()).choices[0].message.content.trim();
    text = text.replace(/^```[a-z]*\n?/i,'').replace(/\n?```$/,'').trim();

    const result = JSON.parse(text);
    result.rowIdx      = params.rowIdx;
    result.studentId   = params.studentId;
    result.studentName = params.studentName;
    result.taskName    = params.taskName;
    return { success: true, result };

  } catch(e) { return { success: false, message: e.toString() }; }
}

// 설문 삭제
function deleteSurvey(svId) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('설문목록');
    if (!sheet) return { success: false, message: '설문목록 시트 없음' };
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(svId).trim()) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, message: '해당 설문을 찾을 수 없습니다.' };
  } catch(e) { return { success: false, message: e.toString() }; }
}

// 교실 알림판 오늘 항목 조회
function getClassroomBoardToday() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const boardId = _getSys(ss, '알림판시트ID');
    if (!boardId) return { success: true, items: [], message: '교실 알림판 시트 ID 미설정' };
    
    const boardSS = SpreadsheetApp.openById(boardId);
    const sheet   = boardSS.getSheets()[0];
    const data    = sheet.getDataRange().getValues();
    const today   = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');
    
    const items = [];
    for (let i = 1; i < data.length; i++) {
      const rowDate = data[i][0] ? Utilities.formatDate(new Date(data[i][0]), 'Asia/Seoul', 'yyyy-MM-dd') : '';
      if (rowDate === today) {
        items.push({
          type:    String(data[i][1] || '알림'),
          title:   String(data[i][2] || ''),
          content: String(data[i][3] || ''),
          time:    data[i][0] ? Utilities.formatDate(new Date(data[i][0]), 'Asia/Seoul', 'HH:mm') : ''
        });
      }
    }
    return { success: true, items: items };
  } catch(e) {
    return { success: false, message: e.toString(), items: [] };
  }
}

// 교실 알림판 저장/수정
function saveBoardNotice(data) {
  try {
    const boardId = getBoardSheetId();
    if (!boardId) return { success: false, message: '시스템설정 알림판시트ID 항목에 교실 알림판 스프레드시트 ID를 입력해주세요.' };
    const sheet = SpreadsheetApp.openById(boardId).getSheetByName('공지사항');
    if (!sheet) return { success: false, message: '교실 알림판 스프레드시트에 "공지사항" 시트가 없습니다.' };

    // 슬라이드URL: 비어있으면 기존 값 유지
    let keepSlideUrl = data.slideUrl || '';
    if (!keepSlideUrl && data.rowIdx && data.rowIdx > 1) {
      try { keepSlideUrl = String(sheet.getRange(data.rowIdx, 6).getValue() || ''); } catch(e) {}
    }
    const row = [
      data.date           || '',
      data.announcement   || '',
      data.notice         || '',
      data.announcementPm || '',
      data.noticePm       || '',
      keepSlideUrl,              // 비어있으면 기존 슬라이드 유지
      data.videoUrl       || '',
      data.videoMode      || '소리만'
    ];

    if (data.rowIdx && data.rowIdx > 1) {
      sheet.getRange(data.rowIdx, 1, 1, 8).setValues([row]);
    } else {
      // 같은 날짜 있으면 업데이트, 없으면 추가
      const lastRow = sheet.getLastRow();
      let found = -1;
      if (lastRow >= 2) {
        const dates = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
        for (let i = 0; i < dates.length; i++) {
          if (String(dates[i][0]).trim() === String(data.date).trim()) {
            found = i + 2; break;
          }
        }
      }
      if (found > 0) {
        sheet.getRange(found, 1, 1, 8).setValues([row]);
      } else {
        sheet.getRange(Math.max(lastRow, 1) + 1, 1, 1, 8).setValues([row]);
      }
    }
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}

// 교실 알림판 항목 삭제
function deleteBoardNotice(rowIdx) {
  try {
    const boardId = getBoardSheetId();
    if (!boardId) return { success: false, message: '시트 ID 없음' };
    const sheet = SpreadsheetApp.openById(boardId).getSheetByName('공지사항');
    if (!sheet) return { success: false, message: '시트 없음' };
    sheet.deleteRow(rowIdx);
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}

// 홈룸 활동 등록
function saveHrActivity(data) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    // ✅ '학급활동목록' 시트에 저장 (getHomeroomData/getHomeroomTabData와 동일한 시트)
    let sheet = ss.getSheetByName('학급활동목록');
    if (!sheet) {
      sheet = ss.insertSheet('학급활동목록');
      sheet.getRange(1,1,1,6).setValues([['번호','카테고리','활동명','설명','폼링크','필드설정']]);
      sheet.getRange(1,1,1,6).setFontWeight('bold').setBackground('#10b981').setFontColor('white');
    }
    const lastRow = Math.max(sheet.getLastRow(), 1);
    sheet.getRange(lastRow+1, 1, 1, 6).setValues([[
      lastRow,                                   // 번호
      data.category || '자율활동',               // 카테고리
      data.name     || '',                        // 활동명
      data.desc     || '',                        // 설명
      data.formUrl  || '',                        // 폼링크
      data.fields   || '[]'                       // 필드설정
    ]]);
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}

// 설문 수정
function updateSurvey(payload) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('설문목록');
    if (!sheet) return { success: false, message: '설문목록 시트 없음' };
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(payload.svId).trim()) {
        sheet.getRange(i+1, 3).setValue(payload.title);      // 제목
        sheet.getRange(i+1, 5).setValue(payload.questions);  // 질문 JSON
        if (payload.guide !== undefined) sheet.getRange(i+1, 6).setValue(payload.guide || '');
        return { success: true };
      }
    }
    return { success: false, message: '해당 설문을 찾을 수 없습니다.' };
  } catch(e) { return { success: false, message: e.toString() }; }
}

// 교실 앱 활동 제출 처리
function processHomeroomForm(p) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName('창체제출현황');
    if (!sheet) {
      sheet = ss.insertSheet('창체제출현황');
      sheet.getRange(1,1,1,8).setValues([['제출일시','학번','이름','활동명','역할','성찰','사진URL','추가응답']]);
      sheet.getRange(1,1,1,8).setFontWeight('bold').setBackground('#10b981').setFontColor('white');
    }
    // 사진 업로드
    let fileUrl = '';
    if (p.fileData && p.fileName) {
      try {
        const blob    = Utilities.newBlob(Utilities.base64Decode(p.fileData), p.fileMimeType || 'image/jpeg', p.fileName);
        const folder  = DriveApp.getFolderById(PARENT_FOLDER_ID);
        let   hrFolder;
        const folderIt = folder.getFoldersByName('창체활동사진');
        hrFolder = folderIt.hasNext() ? folderIt.next() : folder.createFolder('창체활동사진');
        const file = hrFolder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        fileUrl = 'https://drive.google.com/uc?id=' + file.getId();
      } catch(e) { fileUrl = ''; }
    }
    // 추가 응답 JSON
    const extraAnswers = p.extraAnswers ? JSON.stringify(p.extraAnswers) : '';
    sheet.appendRow([
      new Date(),
      String(p.studentId   || '').trim(),
      String(p.studentName || '').trim(),
      String(p.activity    || '').trim(),
      String(p.role        || '').trim(),
      String(p.reflection  || '').trim(),
      fileUrl,
      extraAnswers
    ]);
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}

// ✅ 학생 답장 확인 (읽음 처리 - 시트에서 비우기)
function clearStudentReply(rowIdx) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    
    // 과제 제출 데이터가 있는 시트 찾기
    const sheetNames = ["제출현황", "과제제출", "제출기록", "수학제출"];
    let sheet = null;
    for (let i = 0; i < sheetNames.length; i++) {
      sheet = ss.getSheetByName(sheetNames[i]);
      if (sheet) break;
    }
    if (!sheet) return { success: false, message: "시트를 찾을 수 없습니다." };

    // 1행(헤더)에서 '답장' 이라는 단어가 들어간 열 찾기
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    let replyCol = -1;
    for (let i = 0; i < headers.length; i++) {
      if (String(headers[i]).replace(/\s/g, '').includes('답장')) {
        replyCol = i + 1;
        break;
      }
    }

    // 해당 열이 있으면 그 칸을 깔끔하게 비움
    if (replyCol > 0 && rowIdx > 1) {
      sheet.getRange(rowIdx, replyCol).setValue('');
      return { success: true };
    }
    
    return { success: false, message: "답장 열을 찾을 수 없습니다." };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// 학생 답장 읽음 처리 (reply 필드 초기화)
function clearStudentReply(rowIdx) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('제출현황');
    if (!sheet) return { success: false };
    // P열(16번) = 학생답장
    sheet.getRange(rowIdx, 16).setValue('');
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}

// 학생 재제출 요청 승인: 상태를 재제출요청으로 변경 + P열 초기화
function approveResubmitRequest(rowIdx) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('제출현황');
    if (!sheet) return { success: false };
    sheet.getRange(rowIdx, 11).setValue('재제출요청'); // K열 = 상태
    sheet.getRange(rowIdx, 16).setValue('');           // P열 = 답글 초기화
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}

// =====================================================
// ✅ 시스템설정 시트 마이그레이션 (GAS 에디터에서 1회 직접 실행)
// 기존 셀 주소 기반 → 행 기반 키-값 구조로 변환
// =====================================================
function migrateSysSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName('시스템설정');
  if (!sh) { Logger.log('시스템설정 시트 없음'); return; }

  // 기존 데이터 읽기
  var oldData = sh.getDataRange().getValues();
  var get = function(r, c) {
    if (r < 1 || r > oldData.length) return '';
    var row = oldData[r - 1];
    if (c < 1 || c > row.length) return '';
    return String(row[c - 1] || '').trim();
  };

  var homeroomClass = get(3, 2);   // B3 담임반
  var defaultSlide  = get(4, 2);   // B4 기본슬라이드URL
  var boardSheetId  = get(2, 7);   // G2 알림판시트ID
  var shortcutH     = get(2, 8);   // H2 수학교실
  var shortcutI     = get(2, 9);   // I2 우리반교실
  var shortcutJ     = get(2, 10);  // J2 교실알림판
  var shortcutK     = get(2, 11);  // K2 회원가입
  var shortcutL     = get(2, 12);  // L2 포털
  var openrouterKey = get(2, 16);  // P2 OpenRouter키
  var aiModel       = get(2, 17);  // Q2 AI모델명
  var ttClientId    = get(2, 18);  // R2 TickTick클라이언트ID
  var ttSecret      = get(2, 19);  // S2 TickTick시크릿
  var ttToken       = get(2, 20);  // T2 액세스토큰
  var ttRefresh     = get(2, 21);  // U2 갱신토큰

  // 교사 비밀번호 (학생명부 F2)
  var teacherPw = '';
  try {
    teacherPw = String(ss.getSheetByName('학생명부').getRange('F2').getValue() || '').trim();
  } catch(e) {}

  // 피드백 템플릿 (C2:C50)
  var feedbacks = [];
  for (var r = 2; r <= Math.min(oldData.length, 51); r++) {
    var v = get(r, 3);
    if (v) feedbacks.push(v);
  }

  // 새 구조 데이터
  var rows = [
    ['항목', '값', '안내 (이 열은 참고용입니다)'],
    ['담임반',           homeroomClass, '학생-담임 앱 접근 가능 반 (예: 2학년 1반)'],
    ['교사비밀번호',     teacherPw,     '교사 대시보드 로그인 비밀번호'],
    ['기본슬라이드URL',  defaultSlide,  '교실 알림판 기본 슬라이드 주소'],
    ['알림판시트ID',     boardSheetId,  '교실 알림판 스프레드시트 ID'],
    ['── 바로가기 ──',  '',            '아래 항목에 각 앱의 URL을 붙여넣으세요'],
    ['바로가기_수학교실',   shortcutH, '수학 과제 제출 앱 URL'],
    ['바로가기_우리반교실', shortcutI, '학생-담임 앱 URL'],
    ['바로가기_교실알림판', shortcutJ, '교실 알림판 URL'],
    ['바로가기_회원가입',   shortcutK, '회원가입 앱 URL'],
    ['바로가기_포털',       shortcutL, '스마트 포털 URL'],
    ['── AI / API ──',   '',           '아래 항목에 API 키를 입력하세요'],
    ['OpenRouter키',         openrouterKey, 'AI 기능용 API 키 (openrouter.ai에서 발급)'],
    ['AI모델명',             aiModel,       '사용 모델 (예: anthropic/claude-3.5-sonnet)'],
    ['TickTick클라이언트ID', ttClientId,    'TickTick 앱 연동 (설정 화면에서 입력)'],
    ['TickTick시크릿',       ttSecret,      'TickTick 앱 연동 (설정 화면에서 입력)'],
    ['TickTick액세스토큰',   ttToken,       '자동 저장됨 — 건드리지 마세요'],
    ['TickTick갱신토큰',     ttRefresh,     '자동 저장됨 — 건드리지 마세요'],
    ['── 피드백 템플릿 ──', '', '채점 시 사용할 피드백 문구 (B열에 입력)'],
  ];
  for (var i = 0; i < 30; i++) {
    rows.push(['피드백_' + (i + 1), feedbacks[i] || '', '']);
  }

  // 시트 초기화 후 새 데이터 쓰기
  sh.clearContents();
  sh.clearFormats();
  sh.getRange(1, 1, rows.length, 3).setValues(rows);

  // 헤더 행 스타일
  sh.getRange(1, 1, 1, 3)
    .setBackground('#1e3a8a').setFontColor('white').setFontWeight('bold').setFontSize(12);

  // 구분선 행 스타일 (회색)
  [6, 12, 19].forEach(function(r) {
    sh.getRange(r, 1, 1, 3)
      .setBackground('#f1f5f9').setFontColor('#94a3b8').setFontWeight('bold').setFontStyle('italic');
  });

  // 자동저장 행 (건드리면 안됨) 연노랑 배경
  [17, 18].forEach(function(r) {
    sh.getRange(r, 2).setBackground('#fef9c3');
  });

  // 열 너비
  sh.setColumnWidth(1, 200);
  sh.setColumnWidth(2, 320);
  sh.setColumnWidth(3, 260);

  // 헤더 고정
  sh.setFrozenRows(1);

  // 학생명부 F2 비우기 (교사비번 이전 완료)
  try {
    ss.getSheetByName('학생명부').getRange('F2').clearContent();
  } catch(e) {}

  Logger.log('migrateSysSheet 완료: ' + rows.length + '행 작성');
  SpreadsheetApp.getUi().alert('시스템설정 시트 재구성 완료!\n\n' + rows.length + '개 항목이 설정됐습니다.\n교사 비밀번호도 시스템설정으로 이전했습니다.');
}
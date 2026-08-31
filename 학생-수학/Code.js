const SHEET_ID = _resolveSheetId_();
// SHEET_ID 자동 연결: 스크립트 속성 → 대시보드가 드라이브에 남긴 연결 파일 → 없으면 setup 화면
function _resolveSheetId_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('SHEET_ID');
  if (id) return id;
  try {
    var it = DriveApp.getFilesByName('홍쌤교실시스템_SHEET_ID');
    if (it.hasNext()) {
      id = String(it.next().getBlob().getDataAsString() || '').trim();
      if (/^[a-zA-Z0-9_-]{20,}$/.test(id)) { props.setProperty('SHEET_ID', id); return id; }
    }
  } catch (e) {}
  return '';
}

// 과제 데이터는 과제채점 앱의 '과제 시트'(ClassCore 공유 ID '과제시트ID')를 사용
function _taskSs_() {
  var id = StudentAuth.getConfig('과제시트ID', '');
  if (!id) throw new Error('과제 시트가 설정되지 않았습니다. 과제채점 앱을 먼저 실행하세요.');
  return SpreadsheetApp.openById(id);
}

// FCM 토큰 저장 — 인증 라이브러리(인증 시트)로 위임
function saveFcmToken(studentId, token, pwHash, deviceId) {
  return StudentAuth.saveFcmToken(studentId, token, pwHash, deviceId);
}

// 교사 푸시 — 학생 제출/재제출/답글 시 선생님 폰으로 알림 (실패해도 본 기능에 영향 없음)
// kind별 on/off: 설정 시트 '교사알림_<kind>' Y/N ('제출'만 기본 N)
function _notifyTeacher_(kind, title, body) {
  try {
    var url = '';
    try { url = (StudentAuth.getAppUrls() || {})['과제채점'] || ''; } catch(_) {}
    StudentAuth.sendTeacherPush(kind, title, body, url);
  } catch(_) {}
}


// 포털 카드·알림 클릭이 열 주소를 ClassCore 앱 URL 레지스트리에 자동 기록
// ✅ 2026-07부터 새 PWA 주소를 등록 (기존: 이 앱의 exec 주소) — 포털·알림이 새 앱을 열게 됨
var PWA_APP_URL = 'https://abc58255-hub.github.io/hongssam-classroom/math/';
function _registerAppUrl_(key) {
  try {
    var cache = CacheService.getScriptCache();
    var ck = 'appUrlReg_pwa_' + key; // 캐시 키 변경 → 배포 직후 1회는 반드시 새로 등록
    if (cache && cache.get(ck)) return;
    StudentAuth.registerAppUrl(key.replace('바로가기_', ''), PWA_APP_URL);
    if (cache) cache.put(ck, '1', 21600);
  } catch (e) {}
}

function doGet(e) {
  if (!SHEET_ID) {
    if (e && e.parameter && e.parameter.action) return ContentService.createTextOutput(JSON.stringify({ success: false, message: "앱이 아직 설정되지 않았어요." })).setMimeType(ContentService.MimeType.JSON);
    return _setupPage_();
  }
  // 수업활동 플랫폼(Supabase) 명부 동기화 — 시스템설정 '수업활동동기화키' 일치 시에만
  if (e && e.parameter && e.parameter.action === "syncRoster") {
    return ContentService.createTextOutput(JSON.stringify(_syncRosterToSupabase_(e.parameter.key)))
      .setMimeType(ContentService.MimeType.JSON);
  }
  if (e && e.parameter && (e.parameter.action === "saveFcmToken" || e.parameter.action === "saveTeacherFcmToken")) {
    var result = (e.parameter.action === "saveTeacherFcmToken")
      ? StudentAuth.saveTeacherFcmToken(e.parameter.token, e.parameter.pwHash || null, e.parameter.deviceId || '')
      : saveFcmToken(e.parameter.studentId, e.parameter.token, e.parameter.pwHash || null, e.parameter.deviceId || '');
    var output = JSON.stringify(result);
    if (e.parameter.callback) {
      return ContentService.createTextOutput(e.parameter.callback + '(' + output + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(output).setMimeType(ContentService.MimeType.JSON);
  }
  _registerAppUrl_('바로가기_수학교실');
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('수학 과제 제출기')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ✅ GitHub Pages 프론트(PWA)용 RPC 엔드포인트
// - 본문: text/plain JSON {fn:'함수명', args:[...]} (text/plain이라 CORS preflight 없음)
// - GAS exec 응답에는 Access-Control-Allow-Origin:* 이 자동으로 붙어 크로스오리진 fetch 가능
// - 화이트리스트에 있는 함수만 실행. 반환값은 {ok:true, result:...} / 오류는 {ok:false, error:...}
var RPC_WHITELIST = [
  'verifyLogin', 'getDashboardData', 'getMyGrades', 'getMyDojang',
  'getSecureFileBase64', 'processForm', 'getSubmitRank', 'autoGradeNewSubmission',
  'markBestSeen', 'markFeedbacksAsSeen', 'saveStudentReply', 'requestResubmission',
  'logFeatureUse', 'setStudentPassword', 'reviseSubmission', 'saveFcmToken'
];

function doPost(e) {
  var out;
  try {
    var req = JSON.parse(e.postData.contents);
    if (RPC_WHITELIST.indexOf(req.fn) < 0) throw new Error('허용되지 않은 함수: ' + req.fn);
    var fn = globalThis[req.fn];
    if (typeof fn !== 'function') throw new Error('함수를 찾을 수 없음: ' + req.fn);
    out = { ok: true, result: fn.apply(null, req.args || []) };
  } catch (err) {
    out = { ok: false, error: String(err && err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

// ── 수업활동 플랫폼(Supabase) 명부 동기화 ──────────────────────
// 학번·이름·비번해시를 Supabase students 테이블로 푸시 (원문 비밀번호는 어디에도 없음)
// 호출: <exec>?action=syncRoster&key=<시스템설정 '수업활동동기화키' 값>
var SUPABASE_SYNC_URL = 'https://lqjrrqhrnxctyrqccmch.supabase.co/functions/v1/sync-roster';
function _syncRosterToSupabase_(key) {
  try {
    var expect = StudentAuth.getConfig('수업활동동기화키', '');
    if (!expect || String(key || '') !== expect) return { success: false, message: '키가 올바르지 않습니다' };
    var rows = StudentAuth.getRosterValues(); // B=학번, C=이름, D=비번해시
    var students = [];
    for (var i = 1; i < rows.length; i++) {
      var sid = String(rows[i][1] || '').trim();
      if (!sid) continue;
      students.push({ sid: sid, name: String(rows[i][2] || '').trim(), pwHash: String(rows[i][3] || '').trim() });
    }
    if (students.length === 0) return { success: false, message: '명부가 비어 있습니다' };
    var res = UrlFetchApp.fetch(SUPABASE_SYNC_URL, {
      method: 'post', contentType: 'application/json',
      headers: { 'X-Sync-Key': expect },
      payload: JSON.stringify({ students: students }),
      muteHttpExceptions: true
    });
    var out;
    try { out = JSON.parse(res.getContentText()); } catch (_) { out = { success: false, message: 'HTTP ' + res.getResponseCode() }; }
    out.sent = students.length;
    return out;
  } catch (err) { return { success: false, message: String(err) }; }
}

function getHash(text) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text)
    .map(e => (e < 0 ? e + 256 : e).toString(16).padStart(2, '0'))
    .join(''); 
}

// ✅ 로그인 성공 시 실패 횟수 초기화 + 빈 비밀번호 즉시 반환
// 로그인 — 인증 라이브러리(인증 시트)로 위임. 5회 잠금·해시검증은 라이브러리가 처리.
function verifyLogin(studentId, studentName, password) {
  return StudentAuth.login(studentId, studentName, password);
}

// AI 응답 JSON 느슨하게 파싱 — 토큰 한도로 잘린 응답도 닫는 괄호({,[) 보정해서 복구
function _parseAiJsonLoose_(text) {
  if (!text) return null;
  var start = text.indexOf('{');
  if (start < 0) return null;
  var body = text.substring(start);
  // 1차: 그대로
  try { return JSON.parse(body); } catch(_) {}
  // 2차: 마지막 } 까지만
  var lastBrace = body.lastIndexOf('}');
  if (lastBrace > 0) {
    try { return JSON.parse(body.substring(0, lastBrace + 1)); } catch(_) {}
  }
  // 3차: 미완성 문자열·괄호({,[) 보정 — 스택으로 순서 추적
  var stack = [], inStr = false, esc = false;
  for (var i = 0; i < body.length; i++) {
    var c = body[i];
    if (esc) { esc = false; continue; }
    if (c === '\\') { esc = true; continue; }
    if (c === '"') { inStr = !inStr; continue; }
    if (inStr) continue;
    if (c === '{' || c === '[') stack.push(c);
    else if (c === '}' || c === ']') stack.pop();
  }
  var patched = body;
  if (inStr) patched += '"';
  patched = patched.replace(/[\s]*,[\s]*$/, '').replace(/[\s]*:[\s]*$/, ':null');
  for (var j = stack.length - 1; j >= 0; j--) {
    patched += (stack[j] === '{') ? '}' : ']';
  }
  try { return JSON.parse(patched); } catch(_) {}
  try { return JSON.parse(patched.replace(/,(\s*[}\]])/g, '$1')); } catch(_) { return null; }
}

function _normalizeImageMime(mime) {
  if (!mime) return 'image/jpeg';
  const m = mime.toLowerCase();
  if (m === 'image/jpg' || m === 'image/jpe' || m === 'image/jfif') return 'image/jpeg';
  if (['image/jpeg','image/png','image/gif','image/webp'].indexOf(m) >= 0) return m;
  return null;
}
function getSecureFileBase64(url) {
  try {
    const f = DriveApp.getFileById(url.match(/[-\w]{25,}/)[0]);
    const mime = _normalizeImageMime(f.getMimeType());
    if (!mime) return { success: false, message: '지원하지 않는 형식: ' + f.getMimeType() };
    return { success: true, mimeType: mime, data: Utilities.base64Encode(f.getBlob().getBytes()) };
  } catch(e) {
    return { success: false, message: "파일 열기 실패" };
  }
}

// =====================================================
// ✅ 학생용: getDashboardData 전체 덮어쓰기 (재제출 마감일 기능 포함)
// =====================================================
// ⚡ 제출현황(점점 커지는 로그) 캐시 — 수업 중 동시접속 시 시트 읽기 1회로 공유. TTL 45초.
//   학생 본인 제출/상태변경 시 _clearHistoryCache()로 즉시 무효화 → 방금 한 건 바로 반영.
var _HIST_TTL_ = 45;
function _getHistoryValues() {
  var cache = CacheService.getScriptCache();
  try {
    var meta = cache.get('histN');
    if (meta) {
      var n = parseInt(meta, 10), keys = [];
      for (var i = 0; i < n; i++) keys.push('hist' + i);
      var got = cache.getAll(keys), parts = [], ok = true;
      for (var j = 0; j < n; j++) { var p = got['hist' + j]; if (p == null) { ok = false; break; } parts.push(p); }
      if (ok) return JSON.parse(parts.join(''));
    }
  } catch (_) {}
  var vals = _taskSs_().getSheetByName('제출현황').getDataRange().getValues();
  try {
    var s = JSON.stringify(vals), CH = 90000, n2 = Math.ceil(s.length / CH);
    if (n2 <= 20) {
      var obj = {};
      for (var k = 0; k < n2; k++) obj['hist' + k] = s.substr(k * CH, CH);
      cache.putAll(obj, _HIST_TTL_);
      cache.put('histN', String(n2), _HIST_TTL_);
    }
  } catch (_) {}
  return vals;
}
function _clearHistoryCache() {
  try {
    var cache = CacheService.getScriptCache();
    var meta = cache.get('histN'), keys = ['histN'];
    if (meta) { var n = parseInt(meta, 10); for (var i = 0; i < n; i++) keys.push('hist' + i); }
    cache.removeAll(keys);
  } catch (_) {}
}

function getDashboardData(studentId, studentName) {
  try {
    const now = new Date();
    let safeId = String(studentId || "").trim();
    let className = safeId.length >= 2 ? `${safeId.substring(0, 1)}학년 ${safeId.substring(1, 2)}반` : "기타";
    
    const taskData = _taskSs_().getSheetByName("과제설정").getDataRange().getValues();
    let allBaseTasks = [];
    let assignedBaseTasks = []; // 이 학생 반에 배정된 과제만 (갤러리/노출용)
    let validMissingTasksSet = new Set();
    let taskSettingsMap = {};
    let taskDeadlineMap = {};

    for (let i = 1; i < taskData.length; i++) {
      let tName = String(taskData[i][1] || "").trim();
      if (!tName) continue;

      let dStr = String(taskData[i][3] || "").trim();

      // ✅ 반 배정 필터링: _classes 키가 있으면 해당 반에만 과제 표시
      if (dStr && dStr.startsWith("{")) {
        try {
          let dlCheck = JSON.parse(dStr);
          if (Array.isArray(dlCheck['_classes']) && dlCheck['_classes'].length > 0) {
            if (dlCheck['_classes'].indexOf(className) < 0) continue;
          }
        } catch(e) {}
      }

      allBaseTasks.push(tName);
      let choices = String(taskData[i][7] || "").trim();
      let choiceArray = choices ? choices.split(',').map(s => s.trim()).filter(s => s) : [];
      var maxScore = taskData[i][8] ? Number(taskData[i][8]) : 0; // col9 = 만점
      var taskPub = (function(){ var v = String(taskData[i][5] || '').trim(); return v === '일괄공개' || v === '공개'; })(); // col6 = 과제 단위 공개
      var _tProblems = (function(){ var v = taskData[i][13]; if (!v) return []; try { var a = JSON.parse(v); return Array.isArray(a) ? a : []; } catch(_) { return []; } })(); // N열 문제(JSON)
      taskSettingsMap[tName] = { reqPics: taskData[i][6] ? parseInt(taskData[i][6]) : 1, choiceArray: choiceArray, maxScore: maxScore, isPublicTask: taskPub, allowResubmit: String(taskData[i][12] || '').trim() !== 'N', desc: String(taskData[i][2] || '').trim(), problems: _tProblems }; // M(13)=재제출허용, N(14)=문제

      let isExpired = false;
      let hasDeadline = false;
      let myDeadline = null;
      let resubDeadline = null;
      let openDeadline = null;
      let assignedToMe = true; // 기본: 마감일 미설정(전체 공통) 과제는 모두에게 배정

      if (dStr && dStr.startsWith("{")) {
        try {
          let deadlines = JSON.parse(dStr);
          let dl = deadlines[className] || deadlines["all"];
          const hasClassSpecificKeys = Object.keys(deadlines).some(k => k !== "all" && !k.startsWith("resub_") && !k.startsWith("open_") && k !== "_classes");
          if (dl) {
            hasDeadline = true;
            myDeadline = dl;
            if (new Date(dl) < now) isExpired = true;
          } else if (hasClassSpecificKeys) {
            // 다른 반에만 마감일이 설정된 경우 → 이 학생에겐 과제 숨김(미배정)
            hasDeadline = true;
            isExpired = true;
            assignedToMe = false;
          }
          resubDeadline = deadlines["resub_" + className] || deadlines["resub_all"] || myDeadline;
          openDeadline = deadlines["open_" + className] || deadlines["open_all"] || null;
        } catch(e) { Logger.log('getDashboardData 마감일 파싱 실패: ' + e.message); }
      }

      taskDeadlineMap[tName] = {
        main: myDeadline,
        resub: resubDeadline,
        open: openDeadline
      };

      if (assignedToMe) assignedBaseTasks.push(tName);
      if (!hasDeadline || !isExpired) { validMissingTasksSet.add(tName); }
    }

    const historyData = _getHistoryValues();
    let history = []; 
    let taskStatusMap = {}; 
    let unreadFeedbacks = []; 
    let bestWorksMap = {};
    let rankCounters = {};

    for (let i = 1; i < historyData.length; i++) {
      let rowId = String(historyData[i][1] || "").trim();
      if (!rowId) continue;

      let rawTaskName = String(historyData[i][3] || ""); 
      let baseName = rawTaskName.split(' (')[0];
      let status = historyData[i][10] ? String(historyData[i][10]).trim() : ""; 
      let rowClass = rowId.length >= 2 ? `${rowId.substring(0,1)}학년 ${rowId.substring(1,2)}반` : "기타";

      let myTotalRank = 0;
      let myClassRank = 0;
      if (status !== '재제출요청' && status !== '피드백요청' && status !== '반려검토' && status !== '이전기록채점완료') {
        if (!rankCounters[baseName]) rankCounters[baseName] = { total: 0, classes: {} };
        if (!rankCounters[baseName].classes[rowClass]) rankCounters[baseName].classes[rowClass] = 0;
        
        rankCounters[baseName].total++;
        rankCounters[baseName].classes[rowClass]++;
        
        myTotalRank = rankCounters[baseName].total;
        myClassRank = rankCounters[baseName].classes[rowClass];
      }

      let rawUrl = String(historyData[i][6] || "").trim(); let urls = {}; 
      if (rawUrl.startsWith('{')) { try { urls = JSON.parse(rawUrl); } catch(e) {} } 
      else if (rawUrl.startsWith('[')) { try { let arr=JSON.parse(rawUrl); arr.forEach((v,idx)=>urls["사진 "+(idx+1)]=v); } catch(e){} }
      else if (rawUrl) { urls["사진 1"] = rawUrl; }

      let rawAnno = String(historyData[i][11] || "").trim(); let annoUrls = {}; 
      if (rawAnno.startsWith('{')) { try { annoUrls = JSON.parse(rawAnno); } catch(e) {} }
      else if (rawAnno.startsWith('[')) { try { let arr=JSON.parse(rawAnno); arr.forEach((v,idx)=>annoUrls["사진 "+(idx+1)]=v); } catch(e){} }
      else if (rawAnno) { annoUrls["사진 1"] = rawAnno; }
      
      let bestKey = historyData[i][19] ? String(historyData[i][19]).trim() : "";
      let bestType = historyData[i][16] ? String(historyData[i][16]).trim() : "";
      
      if (bestType !== "") {
        let isAnon = historyData[i][17] === true || String(historyData[i][17]).toUpperCase() === "TRUE";
        let authorCls = rowId.length >= 2 ? `${rowId.substring(0,1)}학년 ${rowId.substring(1,2)}반` : "기타";
        // 이 학생 반에 마감일이 없으면(= 과제 대상 아님) 우수작 팝업 대상에서 제외
        const taskDl = taskDeadlineMap[baseName];
        const isTaskAssigned = taskDl && taskDl.main;
        if (isTaskAssigned && (bestType === "학년공개" || (bestType === "학급공개" && authorCls === className))) {
          if (!bestWorksMap[baseName]) bestWorksMap[baseName] = [];
          bestWorksMap[baseName].push({
            rowIdx: i + 1, id: rowId, name: String(historyData[i][2] || ""),
            urls: urls, annoUrls: annoUrls, bestKey: bestKey,
            comment: historyData[i][18] ? String(historyData[i][18]) : "",
            isAnon: isAnon,
            bestType: bestType
          });
        }
      }

      if (rowId === safeId && String(historyData[i][2] || "").trim() === String(studentName || "").trim()) {
        let fb = historyData[i][7] ? String(historyData[i][7]).trim() : ""; 
        let isSeen = historyData[i][9] ? String(historyData[i][9]).trim() : ""; 
        let isPublic = String(historyData[i][13]).trim() === "공개" || String(historyData[i][13]).trim() === "일괄공개" || !!(taskSettingsMap[baseName] && taskSettingsMap[baseName].isPublicTask); // 과제 단위 공개도 인정
        let score = isPublic && historyData[i][12] ? String(historyData[i][12]).trim() : "";
        let reply = historyData[i][15] ? String(historyData[i][15]).trim() : ""; 
        let isMyBest = (historyData[i][16] ? String(historyData[i][16]).trim() : "") !== "";
        let ts = taskSettingsMap[baseName] || {reqPics:1, choiceArray:[]};

        let perQuestionData = {};
        try {
          let pqRaw = historyData[i][21] ? String(historyData[i][21] || "").trim() : "";
          if (pqRaw.startsWith('{')) {
            let parsed = JSON.parse(pqRaw);
            Object.keys(parsed).forEach(k => {
              let d = parsed[k];
              let qPub = d.pub !== undefined ? d.pub : isPublic;
              if (qPub || isPublic) {
                perQuestionData[k] = { fb: d.fb || '', sc: d.sc || '' };
                if (d.best) perQuestionData[k].best = d.best;
              }
            });
          }
        } catch(e) {}

        let aiGradeTemp = null;
        try {
          const aiRaw = String(historyData[i][22] || "").trim();
          if (aiRaw.startsWith('{')) aiGradeTemp = JSON.parse(aiRaw);
        } catch(e) {}

        let item = {
          rowIdx: i + 1,
          date: historyData[i][0] ? Utilities.formatDate(new Date(historyData[i][0]), "Asia/Seoul", "MM/dd HH:mm") : "",
          task: rawTaskName, baseName: baseName, level: historyData[i][4] || "",
          urls: urls, feedback: fb, status: status, annoUrls: annoUrls,
          score: score, reply: reply, isMyBest: isMyBest,
          maxScore: ts.maxScore || 0,
          allowResubmit: ts.allowResubmit !== false,
          reqPics: ts.reqPics, choices: ts.choiceArray,
          perQuestionData: perQuestionData,
          isUnread: (fb !== "" || isMyBest) && isSeen === "",
          aiGradeTemp: aiGradeTemp,
          totalRank: myTotalRank,
          classRank: myClassRank,
          deadline: taskDeadlineMap[baseName] ? taskDeadlineMap[baseName].main : null,
          resubDeadline: historyData[i][24] ? (historyData[i][24] instanceof Date ? historyData[i][24].toISOString() : String(historyData[i][24])) : null,
          returnType: historyData[i][25] ? String(historyData[i][25]).trim() : ""
        };
        history.push(item); 
        if (status !== "이전기록채점완료") {
          taskStatusMap[baseName] = item;
        } else if (!taskStatusMap[baseName]) {
          taskStatusMap[baseName] = item; 
        }
        if ((fb !== "" && isSeen === "") || (isMyBest && isSeen === "")) unreadFeedbacks.push(item);
      }
    }
    
    let missingTasks = [];
    let resubmitTasks = [];
    let voluntaryTasks = []; // 제출 완료 + 마감 전 → 학생이 자발적으로 다시 제출 가능

    allBaseTasks.forEach(t => {
      let ts = taskSettingsMap[t] || {reqPics:1, choiceArray:[]};
      let dMap = taskDeadlineMap[t] || {};

      if (!taskStatusMap[t]) {
        if (validMissingTasksSet.has(t)) {
          missingTasks.push({
            name: t, reqPics: ts.reqPics, choices: ts.choiceArray,
            submittedUrls: {}, deadline: dMap.main, openDate: dMap.open
          });
        }
      } else {
        let st = taskStatusMap[t].status;
        let submittedCount = 0;
        let currentUrls = taskStatusMap[t].urls;
        for (let k in currentUrls) {
          if (currentUrls[k] && currentUrls[k] !== "" && currentUrls[k] !== "첨부파일 없음") submittedCount++;
        }

        if (st === "재제출요청" || st === "피드백요청") {
          let isFeedback = (st === "피드백요청");
          // 개별 재제출마감(교사가 반려/피드백 누른 날 + 과제별 기한일)만 기준.
          // 없으면(레거시) 마감 없이 열어둠 — 원래 과제 마감일로 폴백하면 마감 후 피드백이 사라짐
          let resubDl = taskStatusMap[t].resubDeadline || null;
          let resubExpired = resubDl && new Date(resubDl) < now;
          if (!resubExpired) {
            let rejectionFeedback = taskStatusMap[t].feedback || "";
            let rejectionPqData = taskStatusMap[t].perQuestionData || {};
            let completedKeys = Object.keys(rejectionPqData).filter(k =>
              rejectionPqData[k] && rejectionPqData[k].status === "완료"
            );
            resubmitTasks.push({
              name: t, reqPics: ts.reqPics, choices: ts.choiceArray,
              submittedUrls: {}, isResubmit: true,
              isFeedback: isFeedback,
              returnType: taskStatusMap[t].returnType || (isFeedback ? "피드백" : "반려"),
              prevScore: (isFeedback && taskStatusMap[t].score) ? taskStatusMap[t].score : "",
              deadline: resubDl,
              rejectionFeedback: rejectionFeedback, completedKeys: completedKeys
            });
          }
        } else if (submittedCount < ts.reqPics && st === "" && validMissingTasksSet.has(t)) {
          missingTasks.push({
            name: t, reqPics: ts.reqPics, choices: ts.choiceArray,
            submittedUrls: currentUrls, deadline: dMap.main, openDate: dMap.open
          });
        } else if (st !== "완료") {
          // 제출 완료(채점완료/미채점 모두, 단 최종완료 잠금 제외) → 본 마감 전이면 자발적 다시 제출 허용
          let mdl = dMap.main;
          if (mdl && new Date(mdl) > now) {
            voluntaryTasks.push({
              name: t, reqPics: ts.reqPics, choices: ts.choiceArray,
              submittedUrls: {}, voluntary: true, deadline: mdl, openDate: dMap.open
            });
          }
        }
      }
    });

    // 과제설정에서 삭제됐거나 목록에서 누락돼도, 되돌린(반려/피드백) 상태면 학생이 다시 제출할 수 있게 보강
    let resubNames = {};
    resubmitTasks.forEach(t => { resubNames[t.name] = true; });
    Object.keys(taskStatusMap).forEach(t => {
      let it = taskStatusMap[t];
      if (!it || resubNames[t]) return;
      if (it.status !== '재제출요청' && it.status !== '피드백요청') return;
      let resubDl = it.resubDeadline || null;
      if (resubDl && new Date(resubDl) < now) return; // 기한 지남
      let isFeedback = (it.status === '피드백요청');
      let ts2 = taskSettingsMap[t] || { reqPics: 1, choiceArray: [] };
      resubmitTasks.push({
        name: t, reqPics: ts2.reqPics, choices: ts2.choiceArray,
        submittedUrls: {}, isResubmit: true, isFeedback: isFeedback,
        returnType: it.returnType || (isFeedback ? '피드백' : '반려'),
        prevScore: (isFeedback && it.score) ? it.score : '',
        deadline: resubDl,
        rejectionFeedback: it.feedback || '', completedKeys: []
      });
      resubNames[t] = true;
    });

    // 🏆 내 업적 (학급·전체 등수 포함) + 🎯 도전과제 (접속기록 포함)
    let myStats = null;
    try {
      const activity = _logVisitAndGetActivity_(safeId); // 오늘 첫 접속이면 +1 기록
      const rosterData = StudentAuth.getRosterValues();
      myStats = _computeAchievements(historyData, taskData, rosterData, safeId, className, activity);
    } catch(e) { Logger.log('업적 계산 오류: ' + e.message); }

    // AI 채점기준이 등록된 과제명 목록 (제출 후 화면 분기용 — 있으면 'AI 채점중', 없으면 '제출완료+등수')
    var rubricTaskNames = [];
    try {
      var _rsh = _taskSs_().getSheetByName('AI채점기준');
      if (_rsh && _rsh.getLastRow() >= 2) {
        var _rvals = _rsh.getRange(2, 1, _rsh.getLastRow() - 1, 1).getValues();
        for (var _ri = 0; _ri < _rvals.length; _ri++) {
          var _rn = String(_rvals[_ri][0] || '').trim();
          if (_rn && rubricTaskNames.indexOf(_rn) < 0) rubricTaskNames.push(_rn);
        }
      }
    } catch(_) {}

    // 과제별 문제(사진·글) 맵 — 학생이 과제 열면 표시
    var taskProblemsMap = {};
    Object.keys(taskSettingsMap).forEach(function(tn){
      var s = taskSettingsMap[tn];
      if (s && ((s.problems && s.problems.length) || s.desc)) taskProblemsMap[tn] = { desc: s.desc || '', problems: s.problems || [] };
    });

    return {
      history: history.reverse(),
      missingTasks: missingTasks,
      resubmitTasks: resubmitTasks,
      voluntaryTasks: voluntaryTasks,
      rubricTasks: rubricTaskNames,
      taskProblems: taskProblemsMap,
      unreadFeedbacks: unreadFeedbacks,
      bestWorksMap: bestWorksMap,
      taskOrder: assignedBaseTasks, // 이 학생 반에 배정된 과제만 (갤러리 정렬용)
      myStats: myStats,
      seenBestKeys: _getSeenBestKeys(safeId), // 서버 저장 "본 우수작" (localStorage 불안정 대비)
      fcmRegisterUrl: _getSysStudent('FCM_REGISTER_URL'),
      mathAppUrl: _getSysStudent('바로가기_수학교실')
    };

  } catch(e) {
    throw new Error("데이터 로딩 중 오류: " + e.toString());
  }
}

// ── 본 우수작 추적 (서버 저장) ───────────────────────────
function _bestSeenSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = _taskSs_().getSheetByName('우수작읽음');
  if (!sh) {
    sh = ss.insertSheet('우수작읽음');
    sh.getRange(1,1,1,2).setValues([['학번','본키JSON']]);
    sh.getRange(1,1,1,2).setFontWeight('bold').setBackground('#f59e0b').setFontColor('white');
    sh.setFrozenRows(1);
  }
  return sh;
}
function _getSeenBestKeys(studentId) {
  try {
    var sh = _bestSeenSheet();
    if (sh.getLastRow() < 2) return [];
    var rows = sh.getRange(2, 1, sh.getLastRow()-1, 2).getValues();
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === String(studentId).trim()) {
        try { var a = JSON.parse(rows[i][1] || '[]'); return Array.isArray(a) ? a : []; } catch(_) { return []; }
      }
    }
    return [];
  } catch(e) { return []; }
}
function markBestSeen(studentId, keys) {
  try {
    if (!studentId || !Array.isArray(keys)) return { success: false };
    var sh = _bestSeenSheet();
    var existing = _getSeenBestKeys(studentId);
    keys.forEach(function(k){ if (existing.indexOf(k) < 0) existing.push(String(k)); });
    if (existing.length > 300) existing = existing.slice(existing.length - 300);
    var rows = sh.getLastRow() >= 2 ? sh.getRange(2, 1, sh.getLastRow()-1, 1).getValues() : [];
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === String(studentId).trim()) {
        sh.getRange(i+2, 2).setValue(JSON.stringify(existing));
        return { success: true };
      }
    }
    sh.appendRow([String(studentId), JSON.stringify(existing)]);
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}

// ── 🎯 도전과제: 활동기록 (접속일수·기능사용) ───────────────────
// 시트: 활동기록 (학번|접속일수|마지막접속일|사용기능CSV) — 과제 스프레드시트에 자동 생성
var _FEATURE_KEYS_ = ['과제제출','재제출','피드백답글','우수작갤러리','도장확인','업적확인'];

function _activitySheet_() {
  var ss = _taskSs_();
  var sh = ss.getSheetByName('활동기록');
  if (!sh) {
    sh = ss.insertSheet('활동기록');
    sh.getRange(1,1,1,4).setValues([['학번','접속일수','마지막접속일','사용기능']]);
    sh.getRange(1,1,1,4).setFontWeight('bold').setBackground('#6366f1').setFontColor('white');
    sh.setFrozenRows(1);
  }
  return sh;
}

// 대시보드 열 때 호출 — 오늘 첫 접속이면 접속일수 +1 (학생당 하루 최대 1회 쓰기). 실패해도 무시.
function _logVisitAndGetActivity_(studentId) {
  var res = { visits: 0, features: [] };
  try {
    var sid = String(studentId || '').trim();
    if (!sid) return res;
    var sh = _activitySheet_();
    var today = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');
    var last = sh.getLastRow();
    var rows = last >= 2 ? sh.getRange(2, 1, last - 1, 4).getValues() : [];
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === sid) {
        var visits = Number(rows[i][1] || 0);
        res.features = String(rows[i][3] || '').split(',').map(function(s){ return s.trim(); }).filter(Boolean);
        // 마지막접속일이 Date로 저장돼 있을 수 있어(시트 자동변환) → yyyy-MM-dd로 정규화 후 비교
        var lastRaw = rows[i][2];
        var lastStr = (lastRaw instanceof Date)
          ? Utilities.formatDate(lastRaw, 'Asia/Seoul', 'yyyy-MM-dd')
          : String(lastRaw || '').trim();
        if (lastStr !== today) {
          visits++;
          sh.getRange(i + 2, 2, 1, 2).setValues([[visits, today]]);
        }
        res.visits = visits;
        return res;
      }
    }
    sh.appendRow([sid, 1, today, '']);
    res.visits = 1;
    return res;
  } catch (e) { return res; }
}

// 기능 첫 사용 기록 (클라이언트에서 localStorage로 중복 호출 차단, 서버도 CSV 중복 방지)
function logFeatureUse(studentId, featureKey) {
  try {
    var sid = String(studentId || '').trim();
    var key = String(featureKey || '').trim();
    if (!sid || _FEATURE_KEYS_.indexOf(key) < 0) return { success: false };
    var sh = _activitySheet_();
    var last = sh.getLastRow();
    var rows = last >= 2 ? sh.getRange(2, 1, last - 1, 4).getValues() : [];
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === sid) {
        var feats = String(rows[i][3] || '').split(',').map(function(s){ return s.trim(); }).filter(Boolean);
        if (feats.indexOf(key) < 0) { feats.push(key); sh.getRange(i + 2, 4).setValue(feats.join(',')); }
        return { success: true, features: feats };
      }
    }
    sh.appendRow([sid, 0, '', key]);
    return { success: true, features: [key] };
  } catch (e) { return { success: false }; }
}

// 학생 업적 계산 — 평균제출순위(제출률≥50%만)·등급평균(채점된 등급과제)·점수평균(채점된 5점과제)
//                 + 자동 배지(연속제출·완주·만점왕·우수작) + 도전과제 카운터. 공개된 채점만 반영.
function _computeAchievements(historyData, taskData, rosterData, myId, myClass, activity) {
  var nowTime = new Date();
  var gpaMap = { 'A':4, 'B':3, 'C':2, 'D':1 };

  // 1) 과제 분류 (생성 순서 보존)
  var taskInfo = {};          // name → { isGrade, isScore5, isPublicTask, deadlines, restrict }
  var taskOrder = [];         // 생성 순서 (연속제출 streak용)
  for (var i = 1; i < taskData.length; i++) {
    var tName = String(taskData[i][1] || '').trim();
    if (!tName) continue;
    var evalType = String(taskData[i][4] || '').trim();
    var pub = String(taskData[i][5] || '').trim();
    var maxv = Number(taskData[i][8] || 0);
    var dl = {}; var restrict = null;
    try {
      dl = JSON.parse(taskData[i][3] || '{}');
      if (Array.isArray(dl['_classes']) && dl['_classes'].length > 0) restrict = dl['_classes'];
    } catch(_) { dl = {}; }
    taskInfo[tName] = {
      isGrade: evalType.indexOf('A-B-C-D') >= 0,
      isScore5: (evalType.indexOf('점수') >= 0 && maxv === 5),
      isPublicTask: (pub === '일괄공개' || pub === '공개'),
      deadlines: dl, restrict: restrict
    };
    taskOrder.push(tName);
  }
  function assignedToClass(tName, cls) {
    var ti = taskInfo[tName]; if (!ti) return false;
    if (ti.restrict && ti.restrict.indexOf(cls) < 0) return false;
    var d = ti.deadlines || {};
    if (!(d[cls] || d['all'])) return false;
    var openDl = d['open_' + cls] || d['open_all'];
    if (openDl && new Date(openDl) > nowTime) return false;
    return true;
  }

  // 2) 학생 목록
  var students = {}, classStudents = {};
  for (var r = 1; r < rosterData.length; r++) {
    var sid = String(rosterData[r][1] || '').trim();
    if (!sid) continue;
    var cls = sid.length >= 2 ? (sid.substring(0,1) + '학년 ' + sid.substring(1,2) + '반') : '기타';
    students[sid] = { cls: cls };
    if (!classStudents[cls]) classStudents[cls] = [];
    classStudents[cls].push(sid);
  }

  // 3) 제출 집계 (최종/최초). 공개여부(col14)도 기록. + 내 답글/재제출 흔적(도전과제용)
  var latest = {}, earliest = {}, allRows = {};
  var myReplies = 0, myResubmitted = false;
  for (var h = 1; h < historyData.length; h++) {
    var hid = String(historyData[h][1] || '').trim();
    if (!hid || !students[hid]) continue;
    var hTask = String(historyData[h][3] || '').split(' (')[0];
    if (!hTask) continue;
    var hStatus = String(historyData[h][10] || '').trim();
    if (hid === myId) {
      var rpl = String(historyData[h][15] || '').trim();
      if (rpl && rpl !== '[재제출요청]') myReplies++;
      if (String(historyData[h][3] || '').indexOf('(재제출)') >= 0) myResubmitted = true;
    }
    var pubv = String(historyData[h][13] || '').trim();
    var rowObj = {
      rowIdx: h,
      ts: historyData[h][0],   // 제출 시각(Date) — 일찍 제출러 판정용
      score: String(historyData[h][12] || '').trim(),
      bestType: String(historyData[h][16] || '').trim(),
      pub: (pubv === '공개' || pubv === '일괄공개')
    };
    // 모든 차수(이전기록 포함) → 최고 점수 산출용
    if (!allRows[hTask]) allRows[hTask] = {};
    if (!allRows[hTask][hid]) allRows[hTask][hid] = [];
    allRows[hTask][hid].push(rowObj);
    if (hStatus === '이전기록채점완료') continue; // 최신/최초·제출집계에서는 제외
    if (!latest[hTask]) latest[hTask] = {};
    if (!latest[hTask][hid] || h > latest[hTask][hid].rowIdx) latest[hTask][hid] = rowObj;
    if (!earliest[hTask]) earliest[hTask] = {};
    if (!earliest[hTask][hid] || h < earliest[hTask][hid].rowIdx) earliest[hTask][hid] = rowObj;
  }
  // 공개 채점 판정: 제출이 공개거나(col14) 과제가 공개(col6)
  function isOpen(tName, lt) { return !!(lt && (lt.pub || (taskInfo[tName] && taskInfo[tName].isPublicTask))); }

  // 4) 학생별 지표
  var stat = {};
  Object.keys(students).forEach(function(sid){
    stat[sid] = { cls: students[sid].cls, assigned: 0, submitted: 0, posSum: 0, posN: 0,
                  best: 0, gradeSum: 0, gradeN: 0, scoreSum: 0, scoreN: 0 };
  });

  Object.keys(taskInfo).forEach(function(tName){
    var ti = taskInfo[tName];
    Object.keys(classStudents).forEach(function(cls){
      if (!assignedToClass(tName, cls)) return;
      var clsSids = classStudents[cls];
      // 제출 순위 (최초 제출 순서) — 제출자끼리만
      var submitters = clsSids.filter(function(sid){ return earliest[tName] && earliest[tName][sid]; })
        .sort(function(a,b){ return earliest[tName][a].rowIdx - earliest[tName][b].rowIdx; });
      var posOf = {};
      submitters.forEach(function(sid, idx){ posOf[sid] = idx + 1; });

      clsSids.forEach(function(sid){
        var st = stat[sid]; if (!st) return;
        st.assigned++;
        var lt = latest[tName] ? latest[tName][sid] : null;
        if (lt) st.submitted++;
        // 제출순위: 제출한 과제만 누적
        if (posOf[sid]) { st.posSum += posOf[sid]; st.posN++; }
        // 우수작
        if (lt && lt.bestType) st.best++;
        var rowsAll = (allRows[tName] && allRows[tName][sid]) ? allRows[tName][sid] : [];
        // 등급 평균: 공개 채점된 등급(A-D) 중 최고(A>B>C>D)
        if (ti.isGrade) {
          var bestG = null;
          rowsAll.forEach(function(ro){
            if (gpaMap[ro.score] !== undefined && isOpen(tName, ro))
              if (bestG === null || gpaMap[ro.score] > bestG) bestG = gpaMap[ro.score];
          });
          if (bestG !== null) { st.gradeSum += bestG; st.gradeN++; }
        }
        // 점수 평균: 5점만점 공개 채점 중 최고 점수
        if (ti.isScore5) {
          var bestS = null;
          rowsAll.forEach(function(ro){
            var v = parseFloat(ro.score);
            if (ro.score !== '' && !isNaN(v) && isOpen(tName, ro))
              if (bestS === null || v > bestS) bestS = v;
          });
          if (bestS !== null) { st.scoreSum += bestS; st.scoreN++; }
        }
      });
    });
  });

  // 5) 지표값 + 제출률(50% 게이트)
  var arr = Object.keys(stat).map(function(sid){
    var st = stat[sid];
    var rate = st.assigned ? (st.submitted / st.assigned) : 0;
    var eligible = (st.assigned > 0 && rate >= 0.5 && st.posN > 0);
    return {
      id: sid, cls: st.cls, rate: rate, assigned: st.assigned, submitted: st.submitted,
      submitRank: eligible ? (st.posSum / st.posN) : null,  // 제출률 50% 미만은 순위 제외
      best: st.best,
      grade: st.gradeN ? (st.gradeSum / st.gradeN) : null,
      score: st.scoreN ? (st.scoreSum / st.scoreN) : null
    };
  });

  function ranks(metric, lowerBetter, excludeZero) {
    var pool = arr.filter(function(x){
      if (x[metric] === null || x[metric] === undefined) return false;
      if (excludeZero && x[metric] === 0) return false;
      return true;
    });
    return { pool: pool, rankIn: function(list, val){
      return list.filter(function(x){ return lowerBetter ? x[metric] < val : x[metric] > val; }).length + 1;
    }};
  }
  var me = arr.filter(function(x){ return x.id === myId; })[0];
  if (!me) return null;

  function build(metric, lowerBetter, excludeZero, decimals) {
    if (me[metric] === null || me[metric] === undefined) return null;
    if (excludeZero && me[metric] === 0) return { value: 0, none: true };
    var R = ranks(metric, lowerBetter, excludeZero);
    var classPool = R.pool.filter(function(x){ return x.cls === myClass; });
    return {
      value: Math.round(me[metric] * Math.pow(10, decimals)) / Math.pow(10, decimals),
      classRank: R.rankIn(classPool, me[metric]), classTotal: classPool.length,
      overallRank: R.rankIn(R.pool, me[metric]), overallTotal: R.pool.length
    };
  }

  // 6) 내 배지 (연속제출·완주·만점왕·우수작) + 도전과제 카운터(스피드러너·불사조)
  // 점수 서열화: A-D 등급 > 통과/미통과 > 숫자 (불사조 비교용)
  function scoreVal(s) {
    if (gpaMap[s] !== undefined) return gpaMap[s];
    if (s === '통과') return 1;
    if (s === '미통과') return 0;
    var f = parseFloat(s);
    return (s !== '' && !isNaN(f)) ? f : null;
  }
  var streak = 0, cur = 0, perfect = 0, fast = 0, phoenix = 0, early = 0;
  taskOrder.forEach(function(tName){
    if (!assignedToClass(tName, myClass)) return;
    var lt = latest[tName] ? latest[tName][myId] : null;
    if (lt) { cur++; if (cur > streak) streak = cur; } else { cur = 0; }
    if (taskInfo[tName].isScore5 && lt && parseFloat(lt.score) === 5 && isOpen(tName, lt)) perfect++;
    // 🌅 일찍 제출러: 마감 24시간+ 전에 (최초) 제출한 과제
    var eRow = earliest[tName] ? earliest[tName][myId] : null;
    if (eRow && eRow.ts) {
      var dlv = taskInfo[tName].deadlines[myClass] || taskInfo[tName].deadlines['all'];
      if (dlv) {
        var dlTime = new Date(dlv), subTime = new Date(eRow.ts);
        if (!isNaN(dlTime.getTime()) && !isNaN(subTime.getTime()) && (dlTime.getTime() - subTime.getTime()) >= 24*3600*1000) early++;
      }
    }
    // ⚡ 스피드러너: 우리 반에서 1~3번째로 제출한 과제
    if (earliest[tName] && earliest[tName][myId]) {
      var myRow = earliest[tName][myId].rowIdx, fasterCnt = 0;
      (classStudents[myClass] || []).forEach(function(sid){
        if (sid !== myId && earliest[tName][sid] && earliest[tName][sid].rowIdx < myRow) fasterCnt++;
      });
      if (fasterCnt < 3) fast++;
    }
    // 🔄 불사조: 재제출(2차 이상 채점)로 첫 공개 점수보다 오른 과제 (과제당 1회)
    var rows = (allRows[tName] && allRows[tName][myId]) ? allRows[tName][myId] : [];
    if (rows.length >= 2) {
      var vals = [];
      rows.slice().sort(function(a,b){ return a.rowIdx - b.rowIdx; }).forEach(function(ro){
        if (!isOpen(tName, ro)) return;
        var v = scoreVal(ro.score);
        if (v !== null) vals.push(v);
      });
      if (vals.length >= 2 && Math.max.apply(null, vals.slice(1)) > vals[0]) phoenix++;
    }
  });
  var submitRankResult = build('submitRank', true, false, 1);
  if (!submitRankResult && me.assigned > 0) {
    submitRankResult = { lowParticipation: true, rate: Math.round(me.rate * 100) };
  }

  // 🧭 탐험가: 저장된 기능사용 + 시트 흔적으로 소급 (과제제출/재제출/피드백답글)
  var act = activity || { visits: 0, features: [] };
  var feats = (act.features || []).slice();
  function addFeat(k){ if (feats.indexOf(k) < 0) feats.push(k); }
  if (me.submitted > 0) addFeat('과제제출');
  if (myResubmitted) addFeat('재제출');
  if (myReplies > 0) addFeat('피드백답글');

  return {
    submitRank: submitRankResult,
    best:       build('best', false, true, 0),
    grade:      build('grade', false, false, 2),
    score:      build('score', false, false, 1),
    badges: {
      submissionRate: Math.round(me.rate * 100),
      complete: (me.assigned > 0 && me.submitted === me.assigned),
      streak: streak,
      perfect: perfect,
      bestCount: me.best
    },
    challenges: {
      submitted: me.submitted,
      streak: streak,
      perfect: perfect,
      best: me.best,
      fast: fast,
      phoenix: phoenix,
      early: early,
      replies: myReplies,
      visits: act.visits || 0,
      features: feats
    }
  };
}

function markFeedbacksAsSeen(rowIndices) { 
  const sheet = _taskSs_().getSheetByName("제출현황"); 
  rowIndices.forEach(idx => sheet.getRange(idx, 10).setValue("확인"));
  _clearHistoryCache();
  return true;
}

// 행 소유자 검증: 제출현황 시트와 rowIdx를 받아 학번(B열) 일치 확인 (sheet 재활용)
function _verifyRowOwner_(sheet, rowIdx, studentId) {
  if (!rowIdx || !studentId) return false;
  if (rowIdx < 2 || rowIdx > sheet.getLastRow()) return false;
  var rowSid = String(sheet.getRange(rowIdx, 2).getValue() || '').trim();
  return rowSid === String(studentId).trim();
}

// 제출현황 행에서 "X학년 Y반 이름 — 과제명" 요약 (교사 푸시 본문용)
function _rowSummary_(sheet, rowIdx) {
  try {
    var r = sheet.getRange(rowIdx, 2, 1, 3).getValues()[0]; // 학번·이름·과제명
    var sid = String(r[0] || '').trim();
    var cls = sid.length >= 2 ? (sid.substring(0, 1) + '학년 ' + sid.substring(1, 2) + '반') : '';
    return (cls ? cls + ' ' : '') + String(r[1] || '').trim() + ' — ' + String(r[2] || '').split(' (')[0].trim();
  } catch(_) { return ''; }
}

function saveStudentReply(rowIdx, replyText, studentId) {
  try {
    var sheet = _taskSs_().getSheetByName("제출현황");
    if (!_verifyRowOwner_(sheet, rowIdx, studentId)) return false;
    sheet.getRange(rowIdx, 16).setValue(replyText);
    var snip = String(replyText || '').replace(/\s+/g, ' ').trim();
    if (snip.length > 40) snip = snip.substring(0, 40) + '…';
    _notifyTeacher_('답글', '💬 학생 답글', _rowSummary_(sheet, rowIdx) + (snip ? ' : ' + snip : ''));
    return true;
  } catch(e) {
    return false;
  }
}

function requestResubmission(rowIdx, studentId) {
  try {
    var sheet = _taskSs_().getSheetByName("제출현황");
    if (!_verifyRowOwner_(sheet, rowIdx, studentId)) return { success: false };
    sheet.getRange(rowIdx, 16).setValue("[재제출요청]");
    _clearHistoryCache();
    _notifyTeacher_('답글', '🙋 재제출 요청', _rowSummary_(sheet, rowIdx) + ' — 다시 풀고 싶어해요');
    return { success: true };
  } catch(e) {
    return { success: false };
  }
}

function processForm(formData) {
  if (String(formData.voluntary) === "true") return _voluntaryResubmit(formData);
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = _taskSs_().getSheetByName("제출현황");
    const now = new Date();
    let inputId = String(formData.studentId || "").trim();
    let inputName = String(formData.studentName || "").trim();
    let isResubmit = formData.isResubmit === "true";
    let baseTaskName = formData.taskName;

    let incomingFiles = formData.filesData;
    let fileHashObj = {};
    const records = sheet.getDataRange().getValues();
    try {
      incomingFiles.forEach(f => { fileHashObj[f.key] = getHash(Utilities.base64Decode(f.b64)); });
    } catch(e) {
      return { success: false, message: '이미지 파일 처리 중 오류가 발생했습니다. (' + e.message + ')' };
    }

    // ✅ 마감일 서버측 검증
    const taskSheet = _taskSs_().getSheetByName("과제설정");
    if (taskSheet) {
      const taskRows = taskSheet.getDataRange().getValues();
      const className = inputId.length >= 2 ? `${inputId.substring(0,1)}학년 ${inputId.substring(1,2)}반` : "기타";
      for (let i = 1; i < taskRows.length; i++) {
        if (String(taskRows[i][1] || "").trim() === baseTaskName) {
          const dStr = String(taskRows[i][3] || "").trim();
          if (dStr && dStr.startsWith("{")) {
            try {
              const dl = JSON.parse(dStr);
              if (isResubmit) {
                // 개별 재제출마감(col25, 교사가 반려/피드백 누른 날+과제별 기한) 우선
                let perRowDl = null;
                for (let k = records.length - 1; k >= 1; k--) {
                  if (String(records[k][1] || "").trim() === inputId &&
                      String(records[k][3] || "").split(' (')[0].trim() === baseTaskName &&
                      (String(records[k][10] || "").trim() === "재제출요청" || String(records[k][10] || "").trim() === "피드백요청")) {
                    if (records[k][24]) perRowDl = new Date(records[k][24]);
                    break;
                  }
                }
                const resubDl = perRowDl || dl["resub_" + className] || dl["resub_all"] || dl[className] || dl["all"];
                if (resubDl && new Date(resubDl) < now) {
                  return { success: false, message: "⏰ 재제출 기한이 지났습니다. 선생님께 문의하세요." };
                }
              } else {
                const mainDl = dl[className] || dl["all"];
                if (mainDl && new Date(mainDl) < now) {
                  return { success: false, message: "⏰ 제출 기한이 지났습니다." };
                }
                // 다른 반에만 마감일이 설정된 경우 이 반은 제출 불가
                const hasClassSpecificKeys = Object.keys(dl).some(k => k !== "all" && !k.startsWith("resub_") && !k.startsWith("open_"));
                if (!mainDl && hasClassSpecificKeys) {
                  return { success: false, message: "🚫 이 과제는 해당 반에 제공되지 않습니다." };
                }
                // 공개일 이전 제출 불가
                const openDl = dl["open_" + className] || dl["open_all"];
                if (openDl && new Date(openDl) > now) {
                  return { success: false, message: "⏳ 아직 제출 가능 시간이 아닙니다." };
                }
              }
            } catch(e) { Logger.log('processForm 마감일 파싱 실패: ' + e.message + ' / dStr=' + dStr); }
          }
          break;
        }
      }
    }

    // ✅ 채점 완료 또는 채점중인 과제는 제출 불가 (중복 방지)
    for (let i = records.length - 1; i >= 1; i--) {
      if (String(records[i][1] || "").trim() === inputId && 
          String(records[i][3] || "").startsWith(baseTaskName)) {
        let currentStatus = String(records[i][10] || "").trim();
        if (currentStatus === "채점완료") {
          return { success: false, message: "🚨 채점이 끝난 과제입니다." };
        }
        // 이전기록채점완료는 건너뜀 (재제출 가능)
        if (currentStatus !== "이전기록채점완료" && currentStatus !== "재제출완료") break;
      }
    }

    // ✅ 재제출인 경우 기존 재제출요청/피드백요청 행을 재제출완료로 변경
    if (isResubmit) {
      for (let i = records.length - 1; i >= 1; i--) {
        if (String(records[i][1] || "").trim() === inputId &&
            String(records[i][3] || "").startsWith(baseTaskName) &&
            (String(records[i][10] || "").trim() === "재제출요청" || String(records[i][10] || "").trim() === "피드백요청")) {
          sheet.getRange(i + 1, 11).setValue("재제출완료");
          sheet.getRange(i + 1, 24).setValue(new Date()); // X열 = 상태변경일시
          break;
        }
      }
    }

    let finalTaskName = isResubmit ? `${baseTaskName} (재제출)` : baseTaskName; 
    let className = inputId.length >= 2 ? `${inputId.substring(0, 1)}학년 ${inputId.substring(1, 2)}반` : "기타"; 
    
    const parentFolder = DriveApp.getFolderById(_getParentFolderId_());
    let taskFolder = parentFolder.getFoldersByName(baseTaskName).hasNext() 
      ? parentFolder.getFoldersByName(baseTaskName).next() 
      : parentFolder.createFolder(baseTaskName);
    let classFolder = taskFolder.getFoldersByName(className).hasNext() 
      ? taskFolder.getFoldersByName(className).next() 
      : taskFolder.createFolder(className);

    // ✅ 임시저장 없이 항상 새 행으로 append
    let finalUrls = {};
    incomingFiles.forEach(f => {
      let suffix = isResubmit ? `_재제출_${f.key}` : `_${f.key}`;
      const blob = Utilities.newBlob(
        Utilities.base64Decode(f.b64), f.mime || 'image/jpeg',
        `[${inputId}] ${inputName}_${baseTaskName}${suffix}.jpg`
      );
      finalUrls[f.key] = classFolder.createFile(blob).getUrl(); 
    });

    sheet.appendRow([
      now, inputId, inputName, finalTaskName, formData.level, formData.message,
      JSON.stringify(finalUrls), "", JSON.stringify(fileHashObj), "", "", "", "", "", "", ""
    ]);
    _clearHistoryCache();

    // 📣 교사 푸시 — 재제출은 기본 켜짐(교사가 재채점해야 함), 첫 제출은 기본 꺼짐
    if (isResubmit) _notifyTeacher_('재제출', '🔄 재제출 도착', className + ' ' + inputName + ' — ' + baseTaskName);
    else _notifyTeacher_('제출', '📥 과제 제출', className + ' ' + inputName + ' — ' + baseTaskName);

    // ✅ 방금 추가한 행 번호를 반환 (AI 자동 채점 호출용)
    const newRowIdx = sheet.getLastRow();
    return { success: true, rowIdx: newRowIdx };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// ── 자발적 다시 제출 (마감 전, 본인) — 기존 최신 행을 새 제출로 덮어쓰고 채점 초기화 → 재채점 ──
function _voluntaryResubmit(formData) {
  try {
    var sheet = _taskSs_().getSheetByName("제출현황");
    var now = new Date();
    var inputId = String(formData.studentId || "").trim();
    var inputName = String(formData.studentName || "").trim();
    var baseTaskName = String(formData.taskName || "").split(' (')[0].trim();
    var className = inputId.length >= 2 ? (inputId.substring(0,1) + '학년 ' + inputId.substring(1,2) + '반') : "기타";

    // 1) 마감 검증 — 본 마감 전이어야 함
    var taskSheet = _taskSs_().getSheetByName("과제설정");
    var td = taskSheet.getDataRange().getValues();
    for (var k = 1; k < td.length; k++) {
      if (String(td[k][1] || "").trim() === baseTaskName) {
        var dStr = String(td[k][3] || "").trim();
        if (dStr && dStr.startsWith("{")) {
          try {
            var dl = JSON.parse(dStr);
            var mainDl = dl[className] || dl["all"];
            if (mainDl && new Date(mainDl) < now) return { success: false, message: "⏰ 마감이 지나 다시 제출할 수 없어요." };
          } catch(_) {}
        }
        break;
      }
    }

    // 2) 내 최신 제출 행 찾기
    var data = sheet.getDataRange().getValues();
    var targetRow = -1;
    for (var i = data.length - 1; i >= 1; i--) {
      if (String(data[i][1] || "").trim() === inputId &&
          String(data[i][3] || "").split(' (')[0].trim() === baseTaskName) { targetRow = i + 1; break; }
    }
    if (targetRow < 0) return { success: false, message: "제출 기록이 없어요." };

    // 3) 새 파일 업로드
    var incomingFiles = formData.filesData || [];
    var fileHashObj = {};
    try { incomingFiles.forEach(function(f){ fileHashObj[f.key] = getHash(Utilities.base64Decode(f.b64)); }); }
    catch(e) { return { success: false, message: '이미지 처리 오류: ' + e.message }; }

    var parentFolder = DriveApp.getFolderById(_getParentFolderId_());
    var taskFolder = parentFolder.getFoldersByName(baseTaskName).hasNext() ? parentFolder.getFoldersByName(baseTaskName).next() : parentFolder.createFolder(baseTaskName);
    var classFolder = taskFolder.getFoldersByName(className).hasNext() ? taskFolder.getFoldersByName(className).next() : taskFolder.createFolder(className);
    var finalUrls = {};
    incomingFiles.forEach(function(f){
      var blob = Utilities.newBlob(Utilities.base64Decode(f.b64), f.mime || 'image/jpeg',
        '[' + inputId + '] ' + inputName + '_' + baseTaskName + '_다시제출_' + f.key + '.jpg');
      finalUrls[f.key] = classFolder.createFile(blob).getUrl();
    });

    // 4) 기존 행 갱신 + 채점 관련 필드 초기화 (재채점 대기)
    sheet.getRange(targetRow, 1).setValue(now);                       // 날짜
    sheet.getRange(targetRow, 4).setValue(baseTaskName);             // 과제명(재제출 표기 제거)
    sheet.getRange(targetRow, 5).setValue(formData.level || "");     // 난이도
    sheet.getRange(targetRow, 6).setValue(formData.message || "");   // 메모
    sheet.getRange(targetRow, 7).setValue(JSON.stringify(finalUrls));// 파일URL
    sheet.getRange(targetRow, 8).setValue("");                       // 피드백
    sheet.getRange(targetRow, 9).setValue(JSON.stringify(fileHashObj)); // 해시
    sheet.getRange(targetRow, 10).setValue("");                      // 읽음
    sheet.getRange(targetRow, 11).setValue("");                      // 상태(미채점)
    sheet.getRange(targetRow, 12).setValue("");                      // 첨삭URL
    sheet.getRange(targetRow, 13).setValue("");                      // 점수
    sheet.getRange(targetRow, 14).setValue("");                      // 공개여부
    sheet.getRange(targetRow, 16).setValue("");                      // 답글
    sheet.getRange(targetRow, 17).setValue("");                      // 우수작유형
    sheet.getRange(targetRow, 22).setValue("");                      // 문항별JSON
    sheet.getRange(targetRow, 23).setValue("");                      // AI채점JSON
    sheet.getRange(targetRow, 24).setValue(now);                     // 상태변경일시
    _clearHistoryCache();
    return { success: true, rowIdx: targetRow };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}


// =====================================================
// ✅ 제출 순위 계산 (같은 과제에서 몇 번째 제출인지) - 반려/재제출 반영 버전
// =====================================================
// 특정 학생의 특정 과제 제출 등수 — 학생별 "최초 제출" 기준(재제출 중복 제거)
function getSubmitRank(studentId, taskName) {
  try {
    const sheet = _taskSs_().getSheetByName("제출현황");
    const data = sheet.getDataRange().getValues();
    const baseTaskName = String(taskName).split(' (')[0].trim();
    var safeId = String(studentId).trim();
    var myClass = safeId.substring(0, 1) + '학년 ' + safeId.substring(1, 2) + '반';

    // 학생별 최초 제출 행 번호 (한 학생당 1번만)
    var firstRow = {};
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][3] || '').split(' (')[0].trim() !== baseTaskName) continue;
      if (String(data[i][10] || '').trim() === '이전기록채점완료') continue;
      var sid = String(data[i][1] || '').trim();
      if (!sid) continue;
      if (firstRow[sid] === undefined || i < firstRow[sid]) firstRow[sid] = i;
    }
    if (firstRow[safeId] === undefined) return { success: false, message: '제출 기록 없음' };

    var myFirst = firstRow[safeId];
    var totalRank = 1, classRank = 1; // 나보다 먼저 낸 사람 수 + 1
    Object.keys(firstRow).forEach(function(sid){
      if (sid === safeId || firstRow[sid] >= myFirst) return;
      totalRank++;
      if (sid.substring(0,1) + '학년 ' + sid.substring(1,2) + '반' === myClass) classRank++;
    });
    return { success: true, totalRank: totalRank, classRank: classRank };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

// =====================================================
// ✅ 비밀번호 설정 (미설정 학생이 최초 로그인 시)
// =====================================================
function setStudentPassword(studentId, studentName, newPw) {
  var res = StudentAuth.setPassword(studentId, studentName, newPw);
  // 수업활동 플랫폼에도 즉시 반영 — 주간 동기화만 기다리면 최대 일주일간 게임 로그인 실패
  if (res && res.success) _pushStudentToLessons_(studentId, studentName, getHash(newPw));
  return res;
}

// 학생 1명 즉시 동기화 (partial — 다른 학생 비활성화 없음). 실패해도 본 기능에 영향 없음.
function _pushStudentToLessons_(sid, name, pwHash) {
  try {
    var key = StudentAuth.getConfig('수업활동동기화키', '');
    if (!key) return;
    UrlFetchApp.fetch(SUPABASE_SYNC_URL, {
      method: 'post', contentType: 'application/json',
      headers: { 'X-Sync-Key': key },
      payload: JSON.stringify({ partial: true, students: [{ sid: String(sid), name: String(name), pwHash: pwHash }] }),
      muteHttpExceptions: true
    });
  } catch (_) {}
}

// =====================================================
// ✅ AI 자동 채점 (학생 제출 직후 자동 실행)
// =====================================================

function _getParentFolderId_() {
  var id = StudentAuth.getConfig('드라이브폴더ID', '');
  if (id) return id;
  var it = DriveApp.getFoldersByName('홍쌤 교실 시스템');
  var folder = it.hasNext() ? it.next() : DriveApp.createFolder('홍쌤 교실 시스템');
  try { StudentAuth.setConfig('드라이브폴더ID', folder.getId()); } catch (e) {}
  return folder.getId();
}

var _sysStudentCache_ = null;
function _getSysStudent(key) { return StudentAuth.getConfig(key, ''); }

function _getApiSettingsForStudent() {
  var orKey = _getSysStudent('OpenRouter키');
  var model = _getSysStudent('AI모델명');
  if (model === 'google/gemini-2.5-flash-preview') model = 'google/gemini-2.5-flash';
  if (model === 'google/gemini-2.5-pro-preview')   model = 'google/gemini-2.5-pro';
  return {
    openrouterKey: orKey,
    model: model || 'google/gemini-3.5-flash'
  };
}

// 단일 출처: 과제설정의 채점유형(col5)·만점(col9)을 우선 사용
function _taskScaleFromSettings(taskName) {
  try {
    var sh = _taskSs_().getSheetByName('과제설정');
    if (!sh || sh.getLastRow() < 2) return {};
    var d = sh.getRange(2, 1, sh.getLastRow() - 1, 9).getValues();
    for (var i = 0; i < d.length; i++) {
      if (String(d[i][1] || '').trim() === taskName)
        return { evalType: String(d[i][4] || '').trim(), maxScore: Number(d[i][8] || 0) };
    }
  } catch(_) {}
  return {};
}

function _getRubricByTaskName(taskName) {
  try {
    const sh = _taskSs_().getSheetByName('AI채점기준');
    if (!sh || sh.getLastRow() < 2) return null;
    const data = sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues();
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0] || '').trim() === taskName) {
        let files = [], questions = {};
        try { if (data[i][4]) files = JSON.parse(data[i][4]); } catch(e) {}
        try { if (data[i][5]) questions = JSON.parse(data[i][5]); } catch(e) {}
        var et = String(data[i][1] || '점수제').trim();
        var mx = Number(data[i][2] || 0);
        var ov = _taskScaleFromSettings(taskName); // 과제설정 우선
        if (ov.evalType) et = ov.evalType;
        if (ov.maxScore) mx = ov.maxScore;
        return {
          taskName: String(data[i][0]).trim(),
          evalType: et,
          maxScore: mx,
          criteria: String(data[i][3] || '').trim(),
          files: files,
          questions: questions
        };
      }
    }
    return null;
  } catch(e) { return null; }
}

function _getRubricFileBase64Student(url) {
  try {
    if (!url || url.indexOf('drive.google.com') < 0) return { success: false };
    const fileId = url.match(/[-\w]{25,}/);
    if (!fileId) return { success: false };
    const f = DriveApp.getFileById(fileId[0]);
    if (f.getSize() > 10485760) return { success: false };
    const mime = _normalizeImageMime(f.getMimeType());
    if (!mime && !f.getMimeType().includes('pdf')) return { success: false };
    return { success: true, mimeType: mime || f.getMimeType(), data: Utilities.base64Encode(f.getBlob().getBytes()) };
  } catch(e) { return { success: false }; }
}

// 학생 제출 직후 클라이언트가 호출하는 자동 채점 함수
function autoGradeNewSubmission(rowIdx, taskName, studentId, studentName, prevAiForCompare) {
  try {
    const baseTask = String(taskName).split(' (')[0];
    const rubric = _getRubricByTaskName(baseTask);
    if (!rubric) return { success: false, message: '채점기준 없음' };

    const cfg = _getApiSettingsForStudent();
    if (!cfg.openrouterKey) return { success: false, message: 'API 키 없음' };

    const sheet = _taskSs_().getSheetByName('제출현황');

    const existing = String(sheet.getRange(rowIdx, 23).getValue() || '').trim();
    if (existing.startsWith('{')) return { success: false, message: '이미 채점됨' };

    const rawUrl = String(sheet.getRange(rowIdx, 7).getValue() || '').trim();
    let urls = {};
    if (rawUrl.startsWith('{')) { try { urls = JSON.parse(rawUrl); } catch(e) {} }
    else if (rawUrl) urls['사진 1'] = rawUrl;

    const urlKeys = Object.keys(urls).filter(k => urls[k] && urls[k] !== "첨부파일 없음");
    if (urlKeys.length === 0) return { success: false, message: '사진 없음' };
    const isType2 = urlKeys.length > 1;

    const prompt = [
      '너는 수학 교사의 채점을 돕는 AI야. 반드시 JSON만 반환해.',
      '학생: ' + studentId + ' ' + studentName,
      '과제: ' + baseTask,
      '채점유형: ' + rubric.evalType,
      '총점: ' + rubric.maxScore + '점',
      rubric.criteria ? '채점기준: ' + rubric.criteria : '',
      (rubric.files || []).some(f => f.role === 'achievement') ? '※ 성취기준 파일을 참고하여 A/B/C/D 판별할 것.' : '',
      '반환 형식:',
      (function() {
        const et = rubric.evalType;
        const mx = rubric.maxScore || 10;
        // 순서 중요: '상-중-하 등급제'는 '등급'을 포함하므로 A-B-C-D보다 먼저 판별
        if (et.indexOf('상-중-하') >= 0)
          return '{"grade":"상|중|하","feedback":"피드백","confidence":"high|medium|low","needsReview":true|false}';
        if (et.indexOf('A-B-C') >= 0 || et.indexOf('ABCD') >= 0 || et.indexOf('등급') >= 0)
          return '{"grade":"A|B|C|D","feedback":"피드백","confidence":"high|medium|low","needsReview":true|false}';
        if (et.indexOf('P/F') >= 0 || et.indexOf('통과') >= 0 || et.indexOf('미통과') >= 0)
          return '{"grade":"Pass|Fail","feedback":"피드백","confidence":"high|medium|low","needsReview":true|false}';
        if (isType2)
          return '{"perQuestion":{"문항명":{"score":점수,"maxScore":만점,"feedback":"피드백"}},"totalScore":합계,"overallFeedback":"종합피드백","confidence":"high|medium|low","needsReview":true|false}';
        if (et.indexOf('10점 단위') >= 0 || et.indexOf('10점단위') >= 0)
          return 'score는 반드시 0, 10, 20 … ' + mx + ' 처럼 10점 단위 값만 사용해(그 외 숫자 금지).\n{"score":10점단위점수,"maxScore":' + mx + ',"feedback":"피드백","confidence":"high|medium|low","needsReview":true|false}';
        return '{"score":0~' + mx + '사이점수,"maxScore":' + mx + ',"feedback":"피드백","confidence":"high|medium|low","needsReview":true|false}';
      })(),
      prevAiForCompare ? ('※ 이 학생의 직전 답안 AI 채점: ' + JSON.stringify({
        score: prevAiForCompare.score, grade: prevAiForCompare.grade,
        feedback: prevAiForCompare.feedback || prevAiForCompare.overallFeedback
      }) + '\n이번은 학생이 보완해서 다시 낸 답안이야. 위 JSON에 "changeSummary" 필드를 반드시 추가해서, 직전 답안 대비 무엇이 개선/변경됐는지(또는 그대로인지) 한두 문장으로 격려하듯 알려줘.') : ''
    ].filter(Boolean).join('\n');

    const content = [{ type: 'text', text: prompt }];

    (rubric.files || []).forEach(rf => {
      if (!rf.url) return;
      const roleLabel = rf.role === 'achievement' ? '[성취기준]' : '[채점기준]';
      content.push({ type: 'text', text: roleLabel });
      const rFile = _getRubricFileBase64Student(rf.url);
      if (rFile.success) {
        if (rFile.mimeType && rFile.mimeType.includes('pdf')) {
          content.push({ type: 'text', text: '(PDF 채점기준은 텍스트 기준 참고)' });
        } else {
          content.push({ type: 'image_url', image_url: { url: 'data:' + rFile.mimeType + ';base64,' + rFile.data } });
        }
      }
    });

    urlKeys.forEach(k => {
      content.push({ type: 'text', text: isType2 ? '[' + k + ' 답안]' : '[학생 답안]' });
      const imgRes = getSecureFileBase64(urls[k]);
      if (imgRes.success) {
        content.push({ type: 'image_url', image_url: { url: 'data:' + imgRes.mimeType + ';base64,' + imgRes.data } });
      }
    });

    const sysMsg = { role: 'system', content: '당신은 수학 채점 AI입니다. 반드시 JSON 객체만 반환하세요. 설명, 마크다운, 추가 텍스트 없이 오직 { } 형태의 JSON만 출력하세요.' };
    const res = UrlFetchApp.fetch('https://openrouter.ai/api/v1/chat/completions', {
      method: 'post',
      headers: {
        'Authorization': 'Bearer ' + cfg.openrouterKey,
        'Content-Type': 'application/json',
        'HTTP-Referer': 'https://script.google.com',
        'X-Title': 'Student Auto Grade'
      },
      payload: JSON.stringify({
        model: cfg.model,
        messages: [sysMsg, { role: 'user', content: content }],
        max_tokens: 4000,
        temperature: 0.2
      }),
      muteHttpExceptions: true
    });

    if (res.getResponseCode() !== 200) {
      try {
        var errJson = JSON.parse(res.getContentText());
        var raw = (errJson.error && errJson.error.metadata && errJson.error.metadata.raw) || '';
        var msg = (errJson.error && errJson.error.message) || res.getContentText().substring(0, 200);
        return { success: false, message: 'API 오류 ' + res.getResponseCode() + ': ' + msg + (raw ? ' | ' + raw : '') };
      } catch(e) { return { success: false, message: 'API 오류 ' + res.getResponseCode() }; }
    }

    let text = JSON.parse(res.getContentText()).choices[0].message.content.trim();
    text = text.replace(/^```[a-z]*\n?/i, '').replace(/\n?```$/, '').trim();
    const result = _parseAiJsonLoose_(text);
    if (!result) return { success: false, message: 'JSON 파싱 실패: ' + text.substring(0, 150) };

    sheet.getRange(rowIdx, 23).setValue(JSON.stringify(result));

    // ✅ [추가] 실시간 등수 확인 및 칭찬 멘트 생성
    const rankData = getSubmitRank(studentId, baseTask);
    if (rankData.success) {
      result.classRank = rankData.classRank;
      result.totalRank = rankData.totalRank;

      let praiseMsg = "";
      if (rankData.classRank === 1) praiseMsg = "🎉 와우! 우리 반에서 가장 먼저 제출했어요! 1등! 🥇 훌륭합니다!";
      else if (rankData.classRank === 2) praiseMsg = "👏 엄청난 스피드! 우리 반 2등으로 제출 완료! 🥈 멋져요!";
      else if (rankData.classRank === 3) praiseMsg = "👍 빠른 제출 칭찬해요! 우리 반 3등 제출입니다! 🥉 최고!";
      else if (rankData.classRank <= 5) praiseMsg = "🏃‍♂️ 엄청 빨라요! 우리 반 TOP 5 안에 들었네요! 🏅 잘했어요!";
      else if (rankData.classRank <= 10) praiseMsg = "🏃‍♀️ 부지런하네요! 우리 반 TOP 10 안에 들었어요! 🎖️";
      
      if (rankData.totalRank === 1) praiseMsg = "🏆 세상에! 전체 학년에서 당당히 1등으로 제출했어요! 엄청난 열정 최고예요! 🔥";
      result.rankMessage = praiseMsg; 
    }

    return { success: true, result: result };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

// ── 학생 셀프 보완: 같은 제출을 덮어쓰고 AI 재채점(변경점 요약 포함). 최대 2회, 교사 알림 없음 ──
function reviseSubmission(rowIdx, taskName, studentId, studentName, filesData) {
  try {
    const sheet = _taskSs_().getSheetByName('제출현황');
    if (!_verifyRowOwner_(sheet, rowIdx, studentId)) return { success: false, message: '권한이 없어요.' };

    // 교사가 이미 손댄 상태면 보완 불가 (첫 제출 직후 상태는 비어있음)
    var status = String(sheet.getRange(rowIdx, 11).getValue() || '').trim();
    if (status && status !== '이전기록채점완료') {
      return { success: false, message: '선생님이 이미 확인 중이라 보완할 수 없어요.', locked: true };
    }

    // 보완 횟수 (col28, 최대 2)
    var used = parseInt(sheet.getRange(rowIdx, 28).getValue() || '0') || 0;
    if (used >= 2) return { success: false, message: '보완은 최대 2회까지예요.', revisionsLeft: 0 };

    // 이전 AI 결과 (변경점 비교용)
    var prevAi = null;
    try { var praw = String(sheet.getRange(rowIdx, 23).getValue() || '').trim(); if (praw.startsWith('{')) prevAi = JSON.parse(praw); } catch(e) {}

    // 사진 덮어쓰기
    var incomingFiles = filesData || [];
    if (!incomingFiles.length) return { success: false, message: '사진을 첨부해주세요.' };
    var baseTask = String(taskName).split(' (')[0];
    var className = studentId.length >= 2 ? (studentId.substring(0,1) + '학년 ' + studentId.substring(1,2) + '반') : '기타';
    var parentFolder = DriveApp.getFolderById(_getParentFolderId_());
    var taskFolder = parentFolder.getFoldersByName(baseTask).hasNext() ? parentFolder.getFoldersByName(baseTask).next() : parentFolder.createFolder(baseTask);
    var classFolder = taskFolder.getFoldersByName(className).hasNext() ? taskFolder.getFoldersByName(className).next() : taskFolder.createFolder(className);

    var newUrls = {}, hashObj = {};
    incomingFiles.forEach(function(f) {
      var bytes = Utilities.base64Decode(f.b64);
      var blob = Utilities.newBlob(bytes, f.mime || 'image/jpeg',
        '[' + studentId + '] ' + studentName + '_' + baseTask + '_보완' + (used + 1) + '_' + f.key + '.jpg');
      newUrls[f.key] = classFolder.createFile(blob).getUrl();
      hashObj[f.key] = getHash(bytes);
    });
    sheet.getRange(rowIdx, 7).setValue(JSON.stringify(newUrls));   // col7 사진 덮어쓰기
    sheet.getRange(rowIdx, 9).setValue(JSON.stringify(hashObj));   // col9 해시
    sheet.getRange(rowIdx, 23).setValue('');                        // col23 이전 AI 비우기(재채점 위해)
    sheet.getRange(rowIdx, 28).setValue(used + 1);                  // col28 보완횟수++
    _clearHistoryCache();

    // AI 재채점 (직전 결과와 비교해 changeSummary 포함)
    var aiRes = autoGradeNewSubmission(rowIdx, baseTask, studentId, studentName, prevAi);
    return {
      success: true,
      result: aiRes.success ? aiRes.result : null,
      aiError: aiRes.success ? null : (aiRes.message || ''),
      revisionsLeft: 2 - (used + 1)
    };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}


// =====================================================
// 🏅 내 도장 (항목별 누적 + 등수)
// =====================================================
function _getDojangCategories_(ss) {
  var cats = _getDojangSetting_(ss, 'categories', null);
  if (!Array.isArray(cats) || !cats.length) {
    var pp = _getDojangSetting_(ss, 'praisePresets', ['박수/격려','질문 답변']);
    cats = [ { name:'발표', emoji:'🙋', type:'sheet' }, { name:'칭찬', emoji:'👏', type:'reason', presets:pp } ];
  }
  return cats.map(function(c){ return { name:String(c.name||'').trim(), emoji:String(c.emoji||'🏅'),
           type:(c.type==='sheet'?'sheet':'reason') }; }).filter(function(c){ return c.name; });
}

// 도장 전용 시트 (ClassCore 도장시트ID). 미설정 시 마스터 폴백.
var _dojangSsCache = null;
function _dojangSs_() {
  if (_dojangSsCache) return _dojangSsCache;
  var id = '';
  try { id = StudentAuth.getConfig('도장시트ID', ''); } catch(_) {}
  var ss;
  if (id) { try { ss = SpreadsheetApp.openById(id); } catch(_) {} }
  if (!ss) ss = SpreadsheetApp.openById(SHEET_ID);
  _dojangSsCache = ss;
  return ss;
}

function getMyDojang(studentId) {
  try {
    var sid = String(studentId || '').trim();
    if (!sid) return { success: false };
    var myCls = sid.length >= 2 ? (sid.substring(0,1) + '학년 ' + sid.substring(1,2) + '반') : '';
    var ss = _dojangSs_();
    var logSh = ss.getSheetByName('도장기록');
    var carrySh = ss.getSheetByName('도장_이월');
    if (!logSh) return { success: true, categories: [], recent: [], showRank: true };

    var roster = StudentAuth.getRosterValues();
    var clsOf = function(id){ id = String(id||'').trim(); return id.length>=2 ? (id.substring(0,1)+'학년 '+id.substring(1,2)+'반') : ''; };
    var cats = _getDojangCategories_(ss);

    // counts[id][catName]
    var counts = {}, rosterIds = {};
    for (var i = 1; i < roster.length; i++) {
      var rid = String(roster[i][1] || '').trim();
      if (rid) { counts[rid] = {}; rosterIds[rid] = clsOf(rid); }
    }
    // 이월 (발표/칭찬만)
    if (carrySh && carrySh.getLastRow() > 1) {
      var carry = carrySh.getRange(2, 1, carrySh.getLastRow()-1, 4).getValues();
      carry.forEach(function(c){
        var cid = String(c[0]||'').trim();
        if (!counts[cid]) return;
        counts[cid]['칭찬'] = (counts[cid]['칭찬']||0) + Number(c[2]||0);
        counts[cid]['발표'] = (counts[cid]['발표']||0) + Number(c[3]||0);
      });
    }
    var recent = [];
    var log = logSh.getLastRow() > 1 ? logSh.getRange(2, 1, logSh.getLastRow()-1, 8).getValues() : [];
    for (var r = 0; r < log.length; r++) {
      var lid = String(log[r][1] || '').trim();
      var kind = String(log[r][3] || '').trim();
      if (counts[lid]) counts[lid][kind] = (counts[lid][kind]||0) + 1;
      if (lid === sid) {
        var dt = log[r][0] ? new Date(log[r][0]) : null;
        recent.push({
          date: dt ? Utilities.formatDate(dt,'Asia/Seoul','MM/dd') : '',
          kind: kind, info: [String(log[r][6]||''), String(log[r][7]||'')].filter(Boolean).join(' ') || String(log[r][4]||''),
          ts: dt ? dt.getTime() : 0
        });
      }
    }
    recent.sort(function(a,b){ return b.ts - a.ts; });

    var rv = _getDojangSetting_(ss, 'rankVisibility', 'show');
    var showRank = (rv !== 'teacher');

    // 항목별 카운트 + 등수
    var myCounts = counts[sid] || {};
    var result = cats.map(function(c){
      var my = myCounts[c.name] || 0;
      var rankAll = 1, rankClass = 1;
      Object.keys(counts).forEach(function(id){
        var v = counts[id][c.name] || 0;
        if (v > my) { rankAll++; if (rosterIds[id] === myCls) rankClass++; }
      });
      return { name: c.name, emoji: c.emoji, count: my, rankAll: rankAll, rankClass: rankClass };
    });

    return { success: true, categories: result, showRank: showRank, records: recent };
  } catch(e) { return { success: false, message: e.toString() }; }
}

function _getDojangSetting_(ss, key, def) {
  try {
    var sh = ss.getSheetByName('도장_설정');
    if (!sh || sh.getLastRow() < 2) return def;
    var rows = sh.getRange(2, 1, sh.getLastRow()-1, 2).getValues();
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === key) {
        try { return JSON.parse(rows[i][1]); } catch(_) { return rows[i][1]; }
      }
    }
    return def;
  } catch(_) { return def; }
}

// ══════════════════════════════════════════════════════════════
//  📋 성적 (시험·수행평가) — 전용 시트 [데이터] 성적, ClassCore 성적시트ID
//  시트 구조: 1행=평가명, 2행=만점, 3행=공개(Y/N), 4행=반평균(Y/N), 5행=학번·이름 헤더, 6행~=데이터
//  A열=학번, B열=이름, C열~=평가별 점수 (구 레이아웃: 4행=헤더 — setupGradeSheet 재실행 시 자동 마이그레이션)
// ══════════════════════════════════════════════════════════════
var _gradeSsCache = null;
function _gradeSs_() {
  if (_gradeSsCache) return _gradeSsCache;
  var id = '';
  try { id = StudentAuth.getConfig('성적시트ID', ''); } catch(_) {}
  var ss = null;
  if (id) { try { ss = SpreadsheetApp.openById(id); } catch(_) {} }
  _gradeSsCache = ss;
  return ss;
}

// [교사 1회 실행] 성적 시트 생성 + 학생명부(학번·이름) 자동 채우기 + 헤더 3행
function setupGradeSheet() {
  var id = '';
  try { id = StudentAuth.getConfig('성적시트ID', ''); } catch(_) {}
  var ss;
  if (id) { try { ss = SpreadsheetApp.openById(id); } catch(_) {} }
  if (!ss) {
    ss = SpreadsheetApp.create('성적 데이터');
    try { StudentAuth.setConfig('성적시트ID', ss.getId()); } catch(_) {}
  }
  var sh = ss.getSheetByName('성적') || ss.insertSheet('성적');
  var def = ss.getSheetByName('시트1') || ss.getSheetByName('Sheet1');
  if (def && ss.getSheets().length > 1) { try { ss.deleteSheet(def); } catch(_) {} }
  // 헤더 4행 (예시 열 포함)
  if (sh.getLastRow() < 3) {
    sh.getRange(1, 1, 4, 3).setValues([
      ['평가 →', '', '1차 지필평가'],   // 1행: A1/B1 안내, C1~ 평가명
      ['만점 →', '', 100],              // 2행: C2~ 만점
      ['공개 → (Y/N 또는 7/10~7/20)', '', 'N'],  // 3행: Y=항상, 기간=그때만, 빈칸/N=숨김
      ['반평균 → (Y/N)', '', 'N']       // 4행: Y=학생에게 반평균 표시
    ]);
    sh.getRange(5, 1, 1, 2).setValues([['학번', '이름']]);
    sh.getRange(1, 1, 5, 3).setFontWeight('bold');
    sh.setFrozenRows(5); sh.setFrozenColumns(2);
  } else if (String(sh.getRange(4, 1).getValue()).trim() === '학번') {
    // 구 레이아웃(4행=헤더) → 반평균 행 삽입 마이그레이션
    sh.insertRowBefore(4);
    sh.getRange(4, 1).setValue('반평균 → (Y/N)').setFontWeight('bold');
    sh.setFrozenRows(5);
  }
  // 학생명부 → 6행부터 학번·이름 채우기 (기존 학번 있으면 스킵)
  var roster = StudentAuth.getRosterValues();
  var existing = {};
  if (sh.getLastRow() >= 6) {
    sh.getRange(6, 1, sh.getLastRow() - 5, 1).getValues().forEach(function(r){ existing[String(r[0]).trim()] = true; });
  }
  var rows = [];
  for (var i = 1; i < roster.length; i++) {
    var sid = String(roster[i][1] || '').trim();
    if (!sid || existing[sid]) continue;
    rows.push([sid, String(roster[i][2] || '').trim()]);
  }
  if (rows.length) sh.getRange(Math.max(sh.getLastRow(), 5) + 1, 1, rows.length, 2).setValues(rows);
  Logger.log('✅ 성적 시트 준비 완료: ' + ss.getUrl() + '  (학생 ' + rows.length + '명 추가)');
  return ss.getUrl();
}

// 공개 셀 판정: Y=항상 / 빈칸·N=비공개 / 기간(7/10~7/20, 7/10~, ~7/20) = 오늘이 범위 안이면 공개
function _gradeVisible(openVal) {
  var s = String(openVal || '').trim();
  if (!s) return false;
  var up = s.toUpperCase();
  if (up === 'Y') return true;
  if (up === 'N') return false;
  var today = new Date(); today.setHours(0, 0, 0, 0);
  if (s.indexOf('~') >= 0) {
    var parts = s.split('~');
    var start = _parseGradeDate(parts[0]);
    var end   = _parseGradeDate(parts[1]);
    if (start && today < start) return false;
    if (end) { var e = new Date(end); e.setHours(23, 59, 59, 999); if (today > e) return false; }
    return (start || end) ? true : false;
  }
  // '~' 없는 단일 날짜 → 그 날부터 공개(시작일)
  var d = _parseGradeDate(s);
  return d ? (today >= d) : false;
}
function _parseGradeDate(str) {
  str = String(str || '').trim();
  if (!str) return null;
  var m = str.match(/(\d{4})[.\-\/\s]+(\d{1,2})[.\-\/\s]+(\d{1,2})/); // yyyy.mm.dd
  if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
  m = str.match(/(\d{1,2})[.\-\/\s]+(\d{1,2})/); // mm.dd (올해)
  if (m) return new Date((new Date()).getFullYear(), +m[1] - 1, +m[2]);
  return null;
}

// [학생] 내 성적 — 공개된(Y 또는 기간 내) 평가만. 반평균은 4행 플래그 Y인 평가만 계산해 함께 반환
function getMyGrades(studentId) {
  try {
    var ss = _gradeSs_();
    if (!ss) return { success: true, grades: [] };
    var sh = ss.getSheetByName('성적');
    if (!sh || sh.getLastRow() < 5 || sh.getLastColumn() < 3) return { success: true, grades: [] };
    var sid = String(studentId || '').trim();
    var lastCol = sh.getLastColumn();
    // 레이아웃 감지: 5행 A열=학번이면 신(4행=반평균 플래그), 4행 A열=학번이면 구(반평균 없음)
    var hasAvgRow = String(sh.getRange(5, 1).getValue()).trim() === '학번';
    var dataStart = hasAvgRow ? 6 : 5;
    if (sh.getLastRow() < dataStart) return { success: true, grades: [] };
    var names   = sh.getRange(1, 3, 1, lastCol - 2).getValues()[0];         // 평가명
    var maxes   = sh.getRange(2, 3, 1, lastCol - 2).getValues()[0];         // 만점
    var opens   = sh.getRange(3, 3, 1, lastCol - 2).getDisplayValues()[0];  // 공개(Y/N/기간)
    var avgOns  = hasAvgRow ? sh.getRange(4, 3, 1, lastCol - 2).getDisplayValues()[0] : []; // 반평균(Y/N)
    // 내 행 찾기
    var body = sh.getRange(dataStart, 1, sh.getLastRow() - dataStart + 1, lastCol).getValues();
    var myRow = null;
    for (var i = 0; i < body.length; i++) { if (String(body[i][0]).trim() === sid) { myRow = body[i]; break; } }
    if (!myRow) return { success: true, grades: [] };
    var myCls = sid.substring(0, 2); // 학년+반
    var grades = [];
    for (var c = 0; c < names.length; c++) {
      var nm = String(names[c] || '').trim();
      if (!nm || !_gradeVisible(opens[c])) continue; // 이름 없거나 (기간상) 비공개 → 스킵
      var val = myRow[c + 2];                      // C열 = 인덱스 2
      if (val === '' || val === null || val === undefined) continue; // 점수 없으면 스킵
      var g = { name: nm, score: val, max: (maxes[c] !== '' ? maxes[c] : '') };
      // 반평균: 플래그 Y면 같은 반(학번 앞 2자리) 학생들의 숫자 점수만 평균 (개별 점수는 노출 안 함)
      if (hasAvgRow && String(avgOns[c] || '').trim().toUpperCase() === 'Y') {
        var sum = 0, n = 0;
        for (var r = 0; r < body.length; r++) {
          if (String(body[r][0] || '').trim().substring(0, 2) !== myCls) continue;
          var v = parseFloat(body[r][c + 2]);
          if (isFinite(v)) { sum += v; n++; }
        }
        if (n > 0) g.avg = Math.round(sum / n * 10) / 10;
      }
      grades.push(g);
    }
    return { success: true, grades: grades };
  } catch(e) { return { success: false, message: e.toString(), grades: [] }; }
}

// ── 초기 설정 (SHEET_ID 미설정 시 setup 화면) ──────────
function _setupPage_() {
  return HtmlService.createHtmlOutputFromFile('setup')
    .setTitle('초기 설정')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function saveSheetId(input) {
  try {
    var props = PropertiesService.getScriptProperties();
    if (props.getProperty('SHEET_ID')) return { success: false, message: '이미 설정되어 있어요. 변경하려면 GAS 편집기 > 프로젝트 설정 > 스크립트 속성에서 SHEET_ID를 삭제한 뒤 다시 열어주세요.' };
    var id = String(input || '').trim();
    var m = id.match(/\/d\/([a-zA-Z0-9_-]{20,})/);
    if (m) id = m[1];
    if (!/^[a-zA-Z0-9_-]{20,}$/.test(id)) return { success: false, message: '스프레드시트 ID 형식이 아니에요. 주소창의 URL 전체를 붙여넣어 보세요.' };
    var ss = SpreadsheetApp.openById(id);
    var name = ss.getName();
    var warn = ss.getSheetByName('시스템설정') ? '' : '\n⚠️ 시스템설정 시트가 없는 스프레드시트예요. 교사용 대시보드에서 만든 스프레드시트가 맞는지 확인해주세요.';
    props.setProperty('SHEET_ID', id);
    return { success: true, message: '"' + name + '" 연결 완료!' + warn };
  } catch (e) {
    return { success: false, message: '스프레드시트를 열 수 없어요. ID와 접근 권한을 확인해주세요.' };
  }
}

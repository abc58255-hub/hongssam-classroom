// 홍쌤 도장-입력 (수업용 빠른 기록)
// 공용 태블릿/TV에서 학생이 자기 이름을 눌러 발표·칭찬 도장을 기록.
// 데이터는 공유 스프레드시트 도장기록 시트에 누적. 진도표에서 오늘 학습지 자동 추천.

var SHEET_ID = '1jK7gYGFXCe3FULLs5mKttP959Aa9vp8-WNOGdJy7cZQ';

// 🔐 외부 요청(푸시) 권한 승인용 — 에디터에서 한 번 실행하면 됨
function grantPermissions() {
  UrlFetchApp.fetch('https://www.google.com');
  Logger.log('✅ 외부요청 권한 승인 완료 — 이제 도장 알림이 발송됩니다.');
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('홍쌤 도장-입력')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function _ensureSheets_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  function mk(name, headers, color) {
    var sh = ss.getSheetByName(name);
    if (!sh) {
      sh = ss.insertSheet(name);
      sh.getRange(1, 1, 1, headers.length).setValues([headers]);
      sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground(color || '#1e3a8a').setFontColor('white');
      sh.setFrozenRows(1);
    }
    return sh;
  }
  mk('도장기록', ['일시','학번','이름','종류','사유','차시','학습지','문제'], '#7c3aed');
  mk('도장_학습지', ['학습지명','발표문제JSON'], '#1e3a8a');
  mk('도장_설정', ['키','값'], '#0d9488');
  mk('도장_이월', ['학번','이름','칭찬이월','발표이월'], '#0d9488');
  return ss;
}

function _clsOf_(sid) {
  sid = String(sid || '').trim();
  return sid.length >= 2 ? (sid.substring(0,1) + '학년 ' + sid.substring(1,2) + '반') : '';
}
function _todayStr_() { return Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd'); }

// 설정 읽기/쓰기 (도장_설정: 키-값, 값은 JSON 문자열)
function _getSetting_(ss, key, def) {
  var sh = ss.getSheetByName('도장_설정');
  if (!sh || sh.getLastRow() < 2) return def;
  var rows = sh.getRange(2, 1, sh.getLastRow()-1, 2).getValues();
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === key) {
      try { return JSON.parse(rows[i][1]); } catch(_) { return rows[i][1]; }
    }
  }
  return def;
}

var DEFAULT_PRAISE = ['박수/격려','질문 답변','친구 도움','적극 참여','발표 경청'];
var DEFAULT_CATS = [
  { name: '발표', emoji: '🙋', type: 'sheet' },
  { name: '칭찬', emoji: '👏', type: 'reason', presets: DEFAULT_PRAISE }
];

// 도장 종류(카테고리) 설정 — 구버전(praisePresets만 있던 때) 호환
function _getCategories_(ss) {
  var cats = _getSetting_(ss, 'categories', null);
  if (!Array.isArray(cats) || !cats.length) {
    // 구버전: 칭찬 프리셋만 있던 경우 반영
    var pp = _getSetting_(ss, 'praisePresets', DEFAULT_PRAISE);
    cats = [
      { name: '발표', emoji: '🙋', type: 'sheet' },
      { name: '칭찬', emoji: '👏', type: 'reason', presets: pp }
    ];
  }
  return cats.map(function(c){
    return { name: String(c.name||'').trim(), emoji: String(c.emoji||'🏅'),
             type: (c.type==='sheet'?'sheet':'reason'),
             presets: Array.isArray(c.presets)?c.presets:[] };
  }).filter(function(c){ return c.name; });
}

// 반 목록
function getClassList() {
  try {
    var roster = SpreadsheetApp.openById(SHEET_ID).getSheetByName('학생명부').getDataRange().getValues();
    var set = {};
    for (var i = 1; i < roster.length; i++) { var c = _clsOf_(roster[i][1]); if (c) set[c] = true; }
    return { success: true, classes: Object.keys(set).sort() };
  } catch(e) { return { success: false, classes: [], message: e.toString() }; }
}

// 입력 화면 초기 데이터
function getInputData(cls) {
  try {
    var ss = _ensureSheets_();
    var roster = ss.getSheetByName('학생명부').getDataRange().getValues();
    var students = [];
    for (var i = 1; i < roster.length; i++) {
      var sid = String(roster[i][1] || '').trim();
      if (!sid) continue;
      if (cls && _clsOf_(sid) !== cls) continue;
      students.push({ sid: sid, name: String(roster[i][2] || '').trim(), today: {} });
    }
    students.sort(function(a,b){ return a.sid.localeCompare(b.sid); });

    // 오늘 카운트 (카테고리별)
    var today = _todayStr_();
    var logSh = ss.getSheetByName('도장기록');
    var log = logSh.getLastRow() > 1 ? logSh.getRange(2, 1, logSh.getLastRow()-1, 8).getValues() : [];
    var idx = {}; students.forEach(function(s){ idx[s.sid] = s; });
    for (var r = 0; r < log.length; r++) {
      var d = log[r][0] ? Utilities.formatDate(new Date(log[r][0]), 'Asia/Seoul', 'yyyy-MM-dd') : '';
      if (d !== today) continue;
      var s = idx[String(log[r][1] || '').trim()];
      if (!s) continue;
      var k = String(log[r][3]).trim();
      s.today[k] = (s.today[k] || 0) + 1;
    }

    // 학습지 목록 (숨김 처리된 건 제외)
    var wsSh = ss.getSheetByName('도장_학습지');
    var worksheets = [];
    if (wsSh.getLastRow() > 1) {
      var wd = wsSh.getRange(2, 1, wsSh.getLastRow()-1, 3).getValues();
      wd.forEach(function(w){
        var nm = String(w[0]||'').trim(); if (!nm) return;
        if (String(w[2]||'').trim().toUpperCase() === 'TRUE') return; // 숨김
        var probs = []; try { probs = JSON.parse(w[1]||'[]'); } catch(_) {}
        worksheets.push({ name: nm, problems: Array.isArray(probs)?probs:[] });
      });
    }

    var categories = _getCategories_(ss);
    var todayWs = _getTodayWorksheet_(ss, cls);

    return { success: true, students: students, worksheets: worksheets,
             categories: categories, todayWorksheet: todayWs };
  } catch(e) { return { success: false, message: e.toString() }; }
}

// 진도표에서 오늘 이 반의 학습지 추천 (best-effort, 실패 시 빈값)
function _getTodayWorksheet_(ss, cls) {
  try {
    if (!cls) return { sheet: '', lessonNo: '' };
    var today = _todayStr_();

    // 1) 이 반이 속한 그룹(gid)·계획(planId) 찾기
    var gid = '기본', planId = '기본';
    var grpSh = ss.getSheetByName('커리큘럼그룹');
    if (grpSh && grpSh.getLastRow() >= 2) {
      var grp = grpSh.getRange(2, 1, grpSh.getLastRow()-1, 5).getValues();
      for (var g = 0; g < grp.length; g++) {
        var clsList = String(grp[g][2]||'').split(',').map(function(x){return x.trim();});
        if (clsList.indexOf(cls) >= 0) {
          gid = String(grp[g][0]).trim();
          planId = String(grp[g][4]||grp[g][0]).trim() || '기본';
          break;
        }
      }
    }

    // 2) 진도체크에서 오늘+반 → 실제차시(없으면 예상차시)
    function lessonFrom(checkName) {
      var ckSh = ss.getSheetByName(checkName);
      if (!ckSh || ckSh.getLastRow() < 2) return 0;
      var cd = ckSh.getRange(2, 1, ckSh.getLastRow()-1, 6).getValues();
      for (var i = 0; i < cd.length; i++) {
        var ds = cd[i][0] instanceof Date ? Utilities.formatDate(cd[i][0],'Asia/Seoul','yyyy-MM-dd') : String(cd[i][0]).trim();
        if (ds === today && String(cd[i][1]).trim() === cls) {
          return parseInt(cd[i][2]) || parseInt(cd[i][5]) || 0;
        }
      }
      return 0;
    }
    var lessonNo = lessonFrom('진도체크_' + gid);
    if (!lessonNo && gid !== '기본') lessonNo = lessonFrom('진도체크_기본');
    if (!lessonNo) return { sheet: '', lessonNo: '' };

    // 3) 진도계획에서 차시 → 계획내용(학습지명)
    function contentFrom(planName, ln) {
      var pSh = ss.getSheetByName(planName);
      if (!pSh || pSh.getLastRow() < 2) return '';
      var pd = pSh.getRange(2, 1, pSh.getLastRow()-1, 2).getValues();
      for (var i = 0; i < pd.length; i++) {
        if ((parseInt(pd[i][0])||0) === ln) return String(pd[i][1]||'').trim();
      }
      return '';
    }
    var sheet = contentFrom('진도계획_' + planId, lessonNo);
    if (!sheet && planId !== '기본') sheet = contentFrom('진도계획_기본', lessonNo);

    return { sheet: sheet || '', lessonNo: lessonNo || '' };
  } catch(e) {
    return { sheet: '', lessonNo: '' };
  }
}

// 도장 기록
// p: { sid, name, kind(카테고리명), type('sheet'|'reason'), emoji, reason, sheet, problem, lessonNo }
function recordStamp(p) {
  try {
    var ss = _ensureSheets_();
    var sid = String(p.sid||'').trim();
    if (!sid) return { success: false, message: '학번 없음' };
    var kind = String(p.kind||'').trim();
    if (!kind) return { success: false, message: '종류 없음' };
    var isSheet = (p.type === 'sheet');
    var reason = String(p.reason||'').trim();
    var sheet = isSheet ? String(p.sheet||'').trim() : '';
    var problem = isSheet ? String(p.problem||'').trim() : '';
    var lessonNo = isSheet ? (p.lessonNo || '') : '';
    if (isSheet && !reason) reason = kind;

    var logSh = ss.getSheetByName('도장기록');
    logSh.appendRow([new Date(), sid, String(p.name||'').trim(), kind, reason, lessonNo, sheet, problem]);

    // 학습지형은 학습지·문제를 다음에 버튼으로 쓰도록 자동 누적
    if (isSheet && sheet) { _upsertWorksheetProblem_(ss, sheet, problem); }

    // 학생 폰에 자동 푸시 (best-effort)
    try {
      var emoji = String(p.emoji||'🏅');
      var info = isSheet ? [sheet, problem].filter(Boolean).join(' ') : reason;
      _pushToStudent_(ss, sid, emoji + ' ' + kind + ' 도장 +1', info || (kind + ' 잘했어요!'));
    } catch(_) {}

    // 오늘 이 학생 카운트 다시 계산 (카테고리별)
    var today = _todayStr_();
    var log = logSh.getLastRow() > 1 ? logSh.getRange(2, 1, logSh.getLastRow()-1, 8).getValues() : [];
    var todayMap = {};
    for (var r = 0; r < log.length; r++) {
      if (String(log[r][1]||'').trim() !== sid) continue;
      var d = log[r][0] ? Utilities.formatDate(new Date(log[r][0]),'Asia/Seoul','yyyy-MM-dd') : '';
      if (d !== today) continue;
      var kk = String(log[r][3]).trim();
      todayMap[kk] = (todayMap[kk]||0) + 1;
    }
    return { success: true, today: todayMap };
  } catch(e) { return { success: false, message: e.toString() }; }
}

function _upsertWorksheetProblem_(ss, sheet, problem) {
  try {
    var wsSh = ss.getSheetByName('도장_학습지');
    var rows = wsSh.getLastRow() > 1 ? wsSh.getRange(2, 1, wsSh.getLastRow()-1, 2).getValues() : [];
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === sheet) {
        if (!problem) return;
        var probs = []; try { probs = JSON.parse(rows[i][1]||'[]'); } catch(_) {}
        if (!Array.isArray(probs)) probs = [];
        if (probs.indexOf(problem) < 0) {
          probs.push(problem);
          wsSh.getRange(i+2, 2).setValue(JSON.stringify(probs));
        }
        return;
      }
    }
    wsSh.appendRow([sheet, JSON.stringify(problem ? [problem] : [])]);
  } catch(_) {}
}

// 오늘 기록 (확인·정정용)
function getTodayRecords(cls) {
  try {
    var ss = _ensureSheets_();
    var today = _todayStr_();
    var logSh = ss.getSheetByName('도장기록');
    var log = logSh.getLastRow() > 1 ? logSh.getRange(2, 1, logSh.getLastRow()-1, 8).getValues() : [];
    var items = [];
    for (var r = 0; r < log.length; r++) {
      var dt = log[r][0] ? new Date(log[r][0]) : null;
      if (!dt || Utilities.formatDate(dt,'Asia/Seoul','yyyy-MM-dd') !== today) continue;
      var sid = String(log[r][1]||'').trim();
      if (cls && _clsOf_(sid) !== cls) continue;
      items.push({
        ts: dt.getTime(), time: Utilities.formatDate(dt,'Asia/Seoul','HH:mm'),
        sid: sid, name: String(log[r][2]||'').trim(), kind: String(log[r][3]||'').trim(),
        reason: String(log[r][4]||''), sheet: String(log[r][6]||''), problem: String(log[r][7]||'')
      });
    }
    items.sort(function(a,b){ return b.ts - a.ts; });
    return { success: true, items: items };
  } catch(e) { return { success: false, message: e.toString(), items: [] }; }
}

// =====================================================
// ✅ FCM 푸시 (대시보드 로직 복제 — HTTP v1 + 서비스계정)
// =====================================================
function _getSys_(ss, key) {
  var sh = ss.getSheetByName('시스템설정');
  if (!sh || sh.getLastRow() < 2) return '';
  var rows = sh.getRange(2, 1, sh.getLastRow()-1, 2).getValues();
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === key) return String(rows[i][1] || '').trim();
  }
  return '';
}

function _getFcmToken_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var privateKey  = _getSys_(ss, 'FCM_PRIVATE_KEY').replace(/\\n/g, '\n');
  var clientEmail = _getSys_(ss, 'FCM_CLIENT_EMAIL');
  if (!privateKey || !clientEmail) return '';
  var now = Math.floor(Date.now() / 1000);
  var header  = Utilities.base64EncodeWebSafe(JSON.stringify({alg:'RS256',typ:'JWT'})).replace(/=+$/,'');
  var payload = Utilities.base64EncodeWebSafe(JSON.stringify({
    iss: clientEmail, scope: 'https://www.googleapis.com/auth/firebase.messaging',
    aud: 'https://oauth2.googleapis.com/token', exp: now + 3600, iat: now
  })).replace(/=+$/,'');
  var toSign = header + '.' + payload;
  var sig = Utilities.base64EncodeWebSafe(Utilities.computeRsaSha256Signature(toSign, privateKey)).replace(/=+$/,'');
  var res = UrlFetchApp.fetch('https://oauth2.googleapis.com/token', {
    method: 'POST', contentType: 'application/x-www-form-urlencoded',
    payload: 'grant_type=urn%3Aietf%3Aparams%3Aoauth%3Agrant-type%3Ajwt-bearer&assertion=' + toSign + '.' + sig,
    muteHttpExceptions: true
  });
  return JSON.parse(res.getContentText()).access_token || '';
}

function _pushToStudent_(ss, sid, title, body) {
  // 학생명부 E열에서 토큰 읽기 (가장 최근 토큰 1개)
  var roster = ss.getSheetByName('학생명부').getDataRange().getValues();
  var raw = '';
  for (var i = 1; i < roster.length; i++) {
    if (String(roster[i][1] || '').trim() === String(sid).trim()) { raw = String(roster[i][4] || '').trim(); break; }
  }
  if (!raw) return;
  var objs = [];
  try { var a = JSON.parse(raw); if (!Array.isArray(a)) a = [a];
        objs = a.map(function(e){ return (typeof e==='string')?e:(e&&e.t?e.t:null); }).filter(Boolean); }
  catch(_) { objs = [raw]; }
  if (!objs.length) return;
  var token = objs[objs.length - 1];

  var projectId = _getSys_(ss, 'FCM_PROJECT_ID');
  var accessToken = _getFcmToken_();
  if (!projectId || !accessToken) return;
  var link = _getSys_(ss, '바로가기_수학교실') || '';
  var t = String(title||''), b = String(body||'');
  var message = { message: {
    token: token,
    notification: { title: t, body: b },
    data: { title: t, body: b, url: link, tag: 'dojang' },
    webpush: { notification: { title: t, body: b, tag: 'dojang', icon: 'https://abc58255-hub.github.io/hongssam-classroom/icon-192.png' }, fcm_options: { link: link } }
  } };
  UrlFetchApp.fetch('https://fcm.googleapis.com/v1/projects/' + projectId + '/messages:send', {
    method:'POST', contentType:'application/json',
    headers:{'Authorization':'Bearer ' + accessToken},
    payload: JSON.stringify(message), muteHttpExceptions: true
  });
}

// 기록 삭제 (오정정) — 일시(ms)+학번으로 1행 식별
function deleteRecord(ts, sid) {
  try {
    var ss = _ensureSheets_();
    var logSh = ss.getSheetByName('도장기록');
    if (logSh.getLastRow() < 2) return { success: false, message: '기록 없음' };
    var n = logSh.getLastRow()-1;
    var vals = logSh.getRange(2, 1, n, 2).getValues(); // 일시, 학번
    for (var i = n-1; i >= 0; i--) {
      var t = vals[i][0] ? new Date(vals[i][0]).getTime() : 0;
      if (t === Number(ts) && String(vals[i][1]||'').trim() === String(sid).trim()) {
        logSh.deleteRow(i + 2);
        return { success: true };
      }
    }
    return { success: false, message: '해당 기록을 찾지 못함' };
  } catch(e) { return { success: false, message: e.toString() }; }
}

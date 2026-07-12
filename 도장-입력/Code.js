// 홍쌤 도장-입력 (수업용 빠른 기록)
// 공용 태블릿/TV에서 학생이 자기 이름을 눌러 발표·칭찬 도장을 기록.
// 데이터는 공유 스프레드시트 도장기록 시트에 누적. 진도표에서 오늘 학습지 자동 추천.

var SHEET_ID = _resolveSheetId_();
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

// 🔐 외부 요청(푸시) 권한 승인용 — 에디터에서 한 번 실행하면 됨

// 이 앱의 배포 URL을 시스템설정 '바로가기_*' 키에 자동 기록 → 대시보드 메뉴가 읽어 자동 연결
function _registerAppUrl_(key) {
  try {
    var cache = CacheService.getScriptCache();
    var ck = 'appUrlReg_' + key;
    if (cache && cache.get(ck)) return;
    var url = '';
    try { url = ScriptApp.getService().getUrl() || ''; } catch (e) { return; }
    if (!url || url.indexOf('/dev') >= 0) return;  // 게시된 /exec 주소만 기록
    var sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName('시스템설정');
    if (!sh) return;
    var last = Math.max(sh.getLastRow(), 1);
    var data = sh.getRange(1, 1, last, 2).getValues();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim() === key) {
        if (String(data[i][1]).trim() !== url) sh.getRange(i + 1, 2).setValue(url);
        if (cache) cache.put(ck, '1', 21600);
        return;
      }
    }
    sh.appendRow([key, url]);
    if (cache) cache.put(ck, '1', 21600);
  } catch (e) {}
}

function doGet() {
  if (!SHEET_ID) return _setupPage_();
  _registerAppUrl_('바로가기_도장입력');
  try { StudentAuth.registerAppUrl('도장입력', ScriptApp.getService().getUrl()); } catch(_) {}
  var t = HtmlService.createTemplateFromFile('index');
  // QR용 배포 URL 자동 주입 — 권한 문제 등으로 실패하면 시스템설정 바로가기 폴백
  var appUrl = '';
  try { appUrl = ScriptApp.getService().getUrl() || ''; } catch (e) {}
  if (!appUrl) { try { appUrl = _getSys_(SpreadsheetApp.openById(SHEET_ID), '바로가기_도장입력') || ''; } catch (e) {} }
  t.appUrl = appUrl;
  return t.evaluate()
    .setTitle('홍쌤 도장-입력')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
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

// 마스터 시트 (진도·시스템설정·인증용). 도장 데이터는 _dojangSs() 사용.
function _ensureSheets_() {
  return SpreadsheetApp.openById(SHEET_ID);
}

// 도장 전용 시트 (ClassCore 도장시트ID). 미설정 시 마스터 폴백. 탭 보장.
var _cachedDojangSs2 = null;
function _dojangSs() {
  if (_cachedDojangSs2) return _cachedDojangSs2;
  var id = '';
  try { id = StudentAuth.getConfig('도장시트ID', ''); } catch(_) {}
  var ss;
  if (id) { try { ss = SpreadsheetApp.openById(id); } catch(_) {} }
  if (!ss) ss = SpreadsheetApp.openById(SHEET_ID);
  function mk(name, headers, color) {
    var sh = ss.getSheetByName(name);
    if (!sh) {
      sh = ss.insertSheet(name);
      sh.getRange(1, 1, 1, headers.length).setValues([headers]);
      sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground(color || '#1e3a8a').setFontColor('white');
      sh.setFrozenRows(1);
    }
  }
  mk('도장기록', ['일시','학번','이름','종류','사유','차시','학습지','문제'], '#7c3aed');
  mk('도장_학습지', ['학습지명','발표문제JSON','숨김'], '#1e3a8a');
  mk('도장_설정', ['키','값'], '#475569');
  mk('도장_이월', ['학번','이름','칭찬이월','발표이월'], '#0d9488');
  _cachedDojangSs2 = ss;
  return ss;
}

function _clsOf_(sid) {
  sid = String(sid || '').trim();
  return sid.length >= 2 ? (sid.substring(0,1) + '학년 ' + sid.substring(1,2) + '반') : '';
}

// ===== 접근 잠금 (링크 공개 배포라 학생이 들어올 수 있음) =====
// 교사 대시보드와 같은 비밀번호(시스템설정 '교사비밀번호') 사용.
// 인증 성공 시 1시간짜리 세션 토큰(UUID)을 발급해 CacheService에 저장.
// 1시간이 지나면 서버에서 자동 만료 → 화면에 잠금이 다시 뜬다.
var SESSION_TTL_SEC = 3600; // 1시간
function _storedPw_(ss) { return _getSys_(ss, '교사비밀번호'); }
function _tokOk_(ss, tok) {
  if (!tok) return false;
  return CacheService.getScriptCache().get('djs_' + String(tok)) === '1';
}
var NEED_AUTH = { success: false, needAuth: true, message: '비밀번호 인증이 필요해요.' };

function unlockDojang(pw) {
  try {
    var cache = CacheService.getScriptCache();
    if (cache.get('dj_pwlock')) return { success: false, message: '🚨 5회 오류로 10분간 잠겼어요.' };
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var stored = _storedPw_(ss);
    if (!stored) return { success: false, message: '교사 비밀번호가 설정되지 않았어요. (시스템설정 시트)' };
    if (String(pw||'') !== stored) {
      var fails = parseInt(cache.get('dj_pwfail') || '0') + 1;
      if (fails >= 5) { cache.put('dj_pwlock', '1', 600); cache.remove('dj_pwfail'); return { success: false, message: '🚨 5회 오류. 10분간 잠겼어요.' }; }
      cache.put('dj_pwfail', String(fails), 600);
      return { success: false, message: '비밀번호 오류 (' + fails + '/5)' };
    }
    cache.remove('dj_pwfail');
    var token = Utilities.getUuid();
    cache.put('djs_' + token, '1', SESSION_TTL_SEC);
    return { success: true, token: token };
  } catch(e) { return { success: false, message: e.toString() }; }
}
function _todayStr_() { return Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd'); }

// 도장기록 꼬리 읽기 — 기록은 시간순 append라 '오늘' 조회엔 마지막 N행이면 충분.
// 전체 읽기는 학기말이 되면 수천 행이라 입력 화면이 점점 느려진다.
var LOG_TAIL_ROWS = 1000; // 하루 도장 1000개 초과는 사실상 불가
function _tailLog_(logSh) {
  var last = logSh.getLastRow();
  if (last < 2) return [];
  var n = Math.min(LOG_TAIL_ROWS, last - 1);
  return logSh.getRange(last - n + 1, 1, n, 8).getValues();
}

// 명부 캐시 (10분) — 학생 추가는 드물어서 안전. [{sid,name}] 전체.
function _getRosterCached_() {
  var cache = CacheService.getScriptCache();
  var hit = cache.get('dj_roster');
  if (hit) { try { return JSON.parse(hit); } catch(_) {} }
  var roster = StudentAuth.getRosterValues();
  var list = [];
  for (var i = 1; i < roster.length; i++) {
    var sid = String(roster[i][1] || '').trim();
    if (!sid) continue;
    list.push({ sid: sid, name: String(roster[i][2] || '').trim() });
  }
  try { cache.put('dj_roster', JSON.stringify(list), 600); } catch(_) {}
  return list;
}

// 설정 읽기/쓰기 (도장_설정: 키-값, 값은 JSON 문자열)
function _getSetting_(ss, key, def) {
  var sh = _dojangSs().getSheetByName('도장_설정');
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
function getClassList(tok) {
  try {
    if (!_tokOk_(SpreadsheetApp.openById(SHEET_ID), tok)) return NEED_AUTH;
    var roster = _getRosterCached_();
    var set = {};
    roster.forEach(function(s){ var c = _clsOf_(s.sid); if (c) set[c] = true; });
    return { success: true, classes: Object.keys(set).sort() };
  } catch(e) { return { success: false, classes: [], message: e.toString() }; }
}

// 입력 화면 초기 데이터
function getInputData(cls, tok) {
  try {
    var ss = _ensureSheets_();
    if (!_tokOk_(ss, tok)) return NEED_AUTH;
    var roster = _getRosterCached_();
    var students = [];
    roster.forEach(function(s){
      if (cls && _clsOf_(s.sid) !== cls) return;
      students.push({ sid: s.sid, name: s.name, today: {} });
    });
    students.sort(function(a,b){ return a.sid.localeCompare(b.sid); });

    // 오늘 카운트 (카테고리별) — 꼬리만 읽어 속도 유지
    var today = _todayStr_();
    var logSh = _dojangSs().getSheetByName('도장기록');
    var log = _tailLog_(logSh);
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
    var wsSh = _dojangSs().getSheetByName('도장_학습지');
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
    if (!_tokOk_(ss, p && p.tok)) return NEED_AUTH;
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
    var count = Math.max(1, Math.min(20, parseInt(p.count) || 1)); // 한 학생에게 여러 개

    var logSh = _dojangSs().getSheetByName('도장기록');
    var name = String(p.name||'').trim();
    // 동시 입력(여러 태블릿) 시 행 덮어쓰기 방지
    var lock = LockService.getScriptLock();
    try {
      lock.waitLock(10000);
      if (count === 1) {
        logSh.appendRow([new Date(), sid, name, kind, reason, lessonNo, sheet, problem]);
      } else {
        var rows = [];
        for (var i = 0; i < count; i++) rows.push([new Date(), sid, name, kind, reason, lessonNo, sheet, problem]);
        logSh.getRange(logSh.getLastRow()+1, 1, count, 8).setValues(rows);
      }
    } finally { try { lock.releaseLock(); } catch(_) {} }

    // 학습지형은 학습지·문제를 다음에 버튼으로 쓰도록 자동 누적
    if (isSheet && sheet) { _upsertWorksheetProblem_(ss, sheet, problem); }

    // 학생 폰에 자동 푸시 (best-effort) — 여러 개면 +N 표시, 1회만 발송
    try {
      var emoji = String(p.emoji||'🏅');
      var info = isSheet ? [sheet, problem].filter(Boolean).join(' ') : reason;
      _pushToStudent_(ss, sid, emoji + ' ' + kind + ' 도장 +' + count, info || (kind + ' 잘했어요!'));
    } catch(_) {}

    // 전체 로그 재조회 없이 즉시 응답 (속도) — 카운트는 클라이언트가 처리
    return { success: true, kind: kind, count: count };
  } catch(e) { return { success: false, message: e.toString() }; }
}

// 여러 명에게 한 번에 (칭찬 등 사유형 일괄). 푸시 없이 빠르게 일괄 기록.
function recordStampBatch(p) {
  try {
    var ss = _ensureSheets_();
    if (!_tokOk_(ss, p && p.tok)) return NEED_AUTH;
    var sids = p.sids || [], names = p.names || [];
    if (!sids.length) return { success: false, message: '대상 없음' };
    var kind = String(p.kind||'').trim();
    var reason = String(p.reason||'').trim();
    if (!kind) return { success: false, message: '종류 없음' };
    var now = new Date();
    var rows = sids.map(function(sid, i){
      return [now, String(sid).trim(), String(names[i]||'').trim(), kind, reason, '', '', ''];
    });
    var logSh = _dojangSs().getSheetByName('도장기록');
    var lock = LockService.getScriptLock();
    try {
      lock.waitLock(10000);
      logSh.getRange(logSh.getLastRow()+1, 1, rows.length, 8).setValues(rows);
    } finally { try { lock.releaseLock(); } catch(_) {} }
    return { success: true, count: rows.length, kind: kind };
  } catch(e) { return { success: false, message: e.toString() }; }
}

function _upsertWorksheetProblem_(ss, sheet, problem) {
  try {
    var wsSh = _dojangSs().getSheetByName('도장_학습지');
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
function getTodayRecords(cls, tok) {
  try {
    var ss = _ensureSheets_();
    if (!_tokOk_(ss, tok)) return NEED_AUTH;
    var today = _todayStr_();
    var logSh = _dojangSs().getSheetByName('도장기록');
    var log = _tailLog_(logSh); // 오늘 기록은 항상 꼬리에 있음
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
    // 학번순 정렬 (같은 학생은 시간순)
    items.sort(function(a,b){ return String(a.sid).localeCompare(String(b.sid), undefined, {numeric:true}) || a.ts - b.ts; });
    return { success: true, items: items };
  } catch(e) { return { success: false, message: e.toString(), items: [] }; }
}

// =====================================================
// ✅ FCM 푸시 (대시보드 로직 복제 — HTTP v1 + 서비스계정)
// =====================================================
function _getSys_(ss, key) {
  // 도장 1개당 2회 호출(프로젝트ID·링크) — 값이 거의 안 바뀌므로 1시간 캐시
  var cache = CacheService.getScriptCache();
  var ck = 'sys_' + key;
  var hit = cache.get(ck);
  if (hit !== null) return hit;
  var val = '';
  var sh = ss.getSheetByName('시스템설정');
  if (sh && sh.getLastRow() >= 2) {
    var rows = sh.getRange(2, 1, sh.getLastRow()-1, 2).getValues();
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === key) { val = String(rows[i][1] || '').trim(); break; }
    }
  }
  try { cache.put(ck, val, 3600); } catch(_) {}
  return val;
}

function _getFcmToken_() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get('fcm_access_token');
  if (cached) return cached;
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
  var at = JSON.parse(res.getContentText()).access_token || '';
  if (at) cache.put('fcm_access_token', at, 3000); // 50분 캐시 → 매번 인증 안 함
  return at;
}

function _pushToStudent_(ss, sid, title, body) {
  // 학생명부 E열에서 토큰 읽기 — TextFinder로 해당 행만 (전체 읽기 방지, 도장 1개당 1회 호출됨)
  var sh = _authRoster_();
  var hit = sh.getRange('B:B').createTextFinder(String(sid).trim()).matchEntireCell(true).findNext();
  if (!hit) return;
  var raw = String(sh.getRange(hit.getRow(), 5).getValue() || '').trim();
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
  var webNoti = { title: t, body: b, tag: 'dojang' };
  var iconUrl = _getSys_(ss, '알림아이콘URL');
  if (iconUrl) webNoti.icon = iconUrl;
  var message = { message: {
    token: token,
    notification: { title: t, body: b },
    data: { title: t, body: b, url: link, tag: 'dojang' },
    webpush: { notification: webNoti, fcm_options: { link: link } }
  } };
  UrlFetchApp.fetch('https://fcm.googleapis.com/v1/projects/' + projectId + '/messages:send', {
    method:'POST', contentType:'application/json',
    headers:{'Authorization':'Bearer ' + accessToken},
    payload: JSON.stringify(message), muteHttpExceptions: true
  });
}

// 기록 삭제 (오정정) — 일시(ms)+학번으로 1행 식별
function deleteRecord(ts, sid, tok) {
  try {
    var ss = _ensureSheets_();
    if (!_tokOk_(ss, tok)) return NEED_AUTH;
    var logSh = _dojangSs().getSheetByName('도장기록');
    var last = logSh.getLastRow();
    if (last < 2) return { success: false, message: '기록 없음' };
    // 오늘 기록만 정정 대상이라 꼬리만 스캔
    var n = Math.min(LOG_TAIL_ROWS, last - 1);
    var startRow = last - n + 1;
    var vals = logSh.getRange(startRow, 1, n, 2).getValues(); // 일시, 학번
    for (var i = n-1; i >= 0; i--) {
      var t = vals[i][0] ? new Date(vals[i][0]).getTime() : 0;
      if (t === Number(ts) && String(vals[i][1]||'').trim() === String(sid).trim()) {
        logSh.deleteRow(startRow + i);
        return { success: true };
      }
    }
    return { success: false, message: '해당 기록을 찾지 못함' };
  } catch(e) { return { success: false, message: e.toString() }; }
}

// 학생명부는 인증 시트(StudentAuth)가 단일 출처. 명부 시트 객체 직접 접근용 헬퍼.
var _authRosterSs_ = null;
function _authRoster_() {
  if (!_authRosterSs_) _authRosterSs_ = SpreadsheetApp.openById(StudentAuth.getAuthSheetId());
  return _authRosterSs_.getSheetByName('학생명부');
}

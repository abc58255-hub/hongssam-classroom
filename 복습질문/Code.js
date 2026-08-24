// 홍쌤 복습 질문 룰렛 (독립 앱)
// 반 뽑기 → 구두 복습 질문 → 통과 시 '복습질문' 도장을 공유 도장기록 시트에 기록.
// 도장 데이터는 도장-입력 앱과 같은 시트를 공유(통과 도장은 학생 도장에 그대로 반영).

var SHEET_ID = _resolveSheetId_();
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

function doGet() {
  if (!SHEET_ID) return HtmlService.createHtmlOutput('<h3>초기 설정 필요</h3><p>스크립트 속성에 SHEET_ID를 설정하거나 도장-입력 앱을 먼저 여세요.</p>');
  try { StudentAuth.registerAppUrl('복습질문', ScriptApp.getService().getUrl()); } catch(_) {}
  return HtmlService.createTemplateFromFile('index').evaluate()
    .setTitle('홍쌤 복습 질문')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── 도장 전용 시트 (도장-입력과 동일: ClassCore 도장시트ID, 폴백 SHEET_ID) ──
var _cachedDojangSs = null;
function _dojangSs() {
  if (_cachedDojangSs) return _cachedDojangSs;
  var id = '';
  try { id = StudentAuth.getConfig('도장시트ID', ''); } catch(_) {}
  var ss;
  if (id) { try { ss = SpreadsheetApp.openById(id); } catch(_) {} }
  if (!ss) ss = SpreadsheetApp.openById(SHEET_ID);
  var mk = function(name, headers, color) {
    var sh = ss.getSheetByName(name);
    if (!sh) {
      sh = ss.insertSheet(name);
      sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground(color||'#1e3a8a').setFontColor('white');
      sh.setFrozenRows(1);
    }
  };
  mk('도장기록', ['일시','학번','이름','종류','사유','차시','학습지','문제'], '#7c3aed');
  mk('도장_설정', ['키','값'], '#475569');
  _cachedDojangSs = ss;
  return ss;
}

function _clsOf_(sid) {
  sid = String(sid || '').trim();
  return sid.length >= 2 ? (sid.substring(0,1) + '학년 ' + sid.substring(1,2) + '반') : '';
}

// ── 접근 잠금 (교사 비밀번호 = 시스템설정 '교사비밀번호', 도장앱과 동일) ──
var SESSION_TTL_SEC = 3600;
var NEED_AUTH = { success: false, needAuth: true, message: '비밀번호 인증이 필요해요.' };
function _getSys_(ss, key) {
  var sh = ss.getSheetByName('시스템설정');
  if (sh && sh.getLastRow() >= 2) {
    var rows = sh.getRange(2, 1, sh.getLastRow()-1, 2).getValues();
    for (var i = 0; i < rows.length; i++) if (String(rows[i][0]).trim() === key) return String(rows[i][1]||'').trim();
  }
  return '';
}
// 토큰은 ScriptProperties에 만료시각과 함께 저장 (CacheService보다 라운드트립 안정적)
function _tokOk_(tok) {
  if (!tok) return false;
  try {
    var v = PropertiesService.getScriptProperties().getProperty('rvtok_' + String(tok));
    if (!v) return false;
    if (Date.now() > parseInt(v)) { PropertiesService.getScriptProperties().deleteProperty('rvtok_' + String(tok)); return false; }
    return true;
  } catch(_) { return false; }
}

function unlockReview(pw) {
  try {
    var cache = CacheService.getScriptCache();
    if (cache.get('rv_pwlock')) return { success: false, message: '🚨 5회 오류로 10분간 잠겼어요.' };
    var stored = _getSys_(SpreadsheetApp.openById(SHEET_ID), '교사비밀번호');
    if (!stored) return { success: false, message: '교사 비밀번호가 설정되지 않았어요. (시스템설정 시트)' };
    if (String(pw||'') !== stored) {
      var fails = parseInt(cache.get('rv_pwfail') || '0') + 1;
      if (fails >= 5) { cache.put('rv_pwlock', '1', 600); cache.remove('rv_pwfail'); return { success: false, message: '🚨 5회 오류. 10분간 잠겼어요.' }; }
      cache.put('rv_pwfail', String(fails), 600);
      return { success: false, message: '비밀번호 오류 (' + fails + '/5)' };
    }
    cache.remove('rv_pwfail');
    var token = Utilities.getUuid();
    PropertiesService.getScriptProperties().setProperty('rvtok_' + token, String(Date.now() + SESSION_TTL_SEC * 1000));
    return { success: true, token: token };
  } catch(e) { return { success: false, message: e.toString() }; }
}

// ── 명부 ──
function _getRosterCached_() {
  var cache = CacheService.getScriptCache();
  var hit = cache.get('rv_roster');
  if (hit) { try { return JSON.parse(hit); } catch(_) {} }
  var roster = StudentAuth.getRosterValues();
  var list = [];
  for (var i = 1; i < roster.length; i++) {
    var sid = String(roster[i][1] || '').trim();
    if (!sid) continue;
    list.push({ sid: sid, name: String(roster[i][2] || '').trim() });
  }
  try { cache.put('rv_roster', JSON.stringify(list), 600); } catch(_) {}
  return list;
}

function getClassList(tok) {
  try {
    if (!_tokOk_(tok)) return NEED_AUTH;
    var set = {};
    _getRosterCached_().forEach(function(s){ var c = _clsOf_(s.sid); if (c) set[c] = true; });
    return { success: true, classes: Object.keys(set).sort() };
  } catch(e) { return { success: false, classes: [], message: e.toString() }; }
}

function getReviewData(cls, tok) {
  try {
    if (!_tokOk_(tok)) return NEED_AUTH;
    var students = [];
    _getRosterCached_().forEach(function(s){ if (!cls || _clsOf_(s.sid) === cls) students.push({ sid: s.sid, name: s.name }); });
    students.sort(function(a,b){ return a.sid.localeCompare(b.sid); });
    // 누적 복습질문 도장 수 (공평 가중치)
    var counts = {};
    var sh = _dojangSs().getSheetByName('도장기록');
    if (sh && sh.getLastRow() > 1) {
      var rows = sh.getRange(2, 1, sh.getLastRow()-1, 4).getValues();
      for (var i = 0; i < rows.length; i++) {
        if (String(rows[i][3]||'').trim() !== '복습질문') continue;
        var sid = String(rows[i][1]||'').trim();
        if (!sid) continue;
        if (cls && _clsOf_(sid) !== cls) continue;
        counts[sid] = (counts[sid]||0) + 1;
      }
    }
    return { success: true, students: students, counts: counts };
  } catch(e) { return { success: false, message: e.toString() }; }
}

// 통과 → 복습질문 도장 1개 기록 (+ 학생 폰 푸시 best-effort)
function recordReviewPass(sid, name, tok) {
  try {
    if (!_tokOk_(tok)) return NEED_AUTH;
    sid = String(sid||'').trim();
    if (!sid) return { success: false, message: '학번 없음' };
    _ensureReviewCategory_();
    var logSh = _dojangSs().getSheetByName('도장기록');
    var lock = LockService.getScriptLock();
    try { lock.waitLock(10000); logSh.appendRow([new Date(), sid, String(name||'').trim(), '복습질문', '복습질문 통과', '', '', '']); }
    finally { try { lock.releaseLock(); } catch(_){} }
    try { _pushToStudent_(sid, '🔁 복습질문 도장 +1', '복습 질문 통과! 잘했어요 👍'); } catch(_) {}
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}

// 복습질문 카테고리 보장 (학생 대시보드 타일 노출용)
function _ensureReviewCategory_() {
  try {
    var ss = _dojangSs();
    var sh = ss.getSheetByName('도장_설정');
    if (!sh) return;
    var cats = null;
    if (sh.getLastRow() >= 2) {
      var rows = sh.getRange(2, 1, sh.getLastRow()-1, 2).getValues();
      for (var i = 0; i < rows.length; i++) {
        if (String(rows[i][0]).trim() === 'categories') { try { cats = JSON.parse(rows[i][1]); } catch(_){}
          if (Array.isArray(cats) && cats.some(function(c){ return c && c.name === '복습질문'; })) return;
          if (!Array.isArray(cats)) cats = [];
          cats.push({ name:'복습질문', emoji:'🔁', type:'reason', presets:[] });
          sh.getRange(i+2, 2).setValue(JSON.stringify(cats)); return;
        }
      }
    }
    // categories 키 자체가 없으면 신규 (기존 발표/칭찬 기본 + 복습질문)
    cats = [ {name:'발표',emoji:'🙋',type:'sheet'}, {name:'칭찬',emoji:'👏',type:'reason',presets:['박수/격려','질문 답변']}, {name:'복습질문',emoji:'🔁',type:'reason',presets:[]} ];
    sh.appendRow(['categories', JSON.stringify(cats)]);
  } catch(_) {}
}

// 학생 폰 푸시 (도장-입력과 동일 방식 — StudentAuth 경유)
function _pushToStudent_(sid, title, body) {
  try {
    if (StudentAuth.pushToStudent) { StudentAuth.pushToStudent(sid, title, body); }
  } catch(_) {}
}

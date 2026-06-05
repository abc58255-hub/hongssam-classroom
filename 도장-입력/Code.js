// 홍쌤 도장-입력 (수업용 빠른 기록)
// 공용 태블릿/TV에서 학생이 자기 이름을 눌러 발표·칭찬 도장을 기록.
// 데이터는 공유 스프레드시트 도장기록 시트에 누적. 진도표에서 오늘 학습지 자동 추천.

var SHEET_ID = '1jK7gYGFXCe3FULLs5mKttP959Aa9vp8-WNOGdJy7cZQ';

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
      students.push({ sid: sid, name: String(roster[i][2] || '').trim(), todayPresent: 0, todayPraise: 0 });
    }
    students.sort(function(a,b){ return a.sid.localeCompare(b.sid); });

    // 오늘 카운트
    var today = _todayStr_();
    var logSh = ss.getSheetByName('도장기록');
    var log = logSh.getLastRow() > 1 ? logSh.getRange(2, 1, logSh.getLastRow()-1, 8).getValues() : [];
    var idx = {}; students.forEach(function(s){ idx[s.sid] = s; });
    for (var r = 0; r < log.length; r++) {
      var d = log[r][0] ? Utilities.formatDate(new Date(log[r][0]), 'Asia/Seoul', 'yyyy-MM-dd') : '';
      if (d !== today) continue;
      var s = idx[String(log[r][1] || '').trim()];
      if (!s) continue;
      if (String(log[r][3]).trim() === '발표') s.todayPresent++;
      else if (String(log[r][3]).trim() === '칭찬') s.todayPraise++;
    }

    // 학습지 목록
    var wsSh = ss.getSheetByName('도장_학습지');
    var worksheets = [];
    if (wsSh.getLastRow() > 1) {
      var wd = wsSh.getRange(2, 1, wsSh.getLastRow()-1, 2).getValues();
      wd.forEach(function(w){
        var nm = String(w[0]||'').trim(); if (!nm) return;
        var probs = []; try { probs = JSON.parse(w[1]||'[]'); } catch(_) {}
        worksheets.push({ name: nm, problems: Array.isArray(probs)?probs:[] });
      });
    }

    var praisePresets = _getSetting_(ss, 'praisePresets', DEFAULT_PRAISE);
    var todayWs = _getTodayWorksheet_(ss, cls);

    return { success: true, students: students, worksheets: worksheets,
             praisePresets: praisePresets, todayWorksheet: todayWs };
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
// p: { sid, name, kind('발표'|'칭찬'), reason, sheet, problem, lessonNo }
function recordStamp(p) {
  try {
    var ss = _ensureSheets_();
    var sid = String(p.sid||'').trim();
    if (!sid) return { success: false, message: '학번 없음' };
    var kind = p.kind === '칭찬' ? '칭찬' : '발표';
    var reason = String(p.reason||'').trim();
    var sheet = String(p.sheet||'').trim();
    var problem = String(p.problem||'').trim();
    var lessonNo = p.lessonNo || '';
    if (kind === '발표' && !reason) reason = '발표';

    var logSh = ss.getSheetByName('도장기록');
    logSh.appendRow([new Date(), sid, String(p.name||'').trim(), kind, reason, lessonNo, sheet, problem]);

    // 발표 학습지·문제는 다음에 버튼으로 쓰도록 도장_학습지에 자동 누적
    if (kind === '발표' && sheet) { _upsertWorksheetProblem_(ss, sheet, problem); }

    // 오늘 이 학생 카운트 다시 계산
    var today = _todayStr_();
    var log = logSh.getLastRow() > 1 ? logSh.getRange(2, 1, logSh.getLastRow()-1, 8).getValues() : [];
    var tp = 0, tc = 0;
    for (var r = 0; r < log.length; r++) {
      if (String(log[r][1]||'').trim() !== sid) continue;
      var d = log[r][0] ? Utilities.formatDate(new Date(log[r][0]),'Asia/Seoul','yyyy-MM-dd') : '';
      if (d !== today) continue;
      if (String(log[r][3]).trim() === '발표') tp++; else if (String(log[r][3]).trim() === '칭찬') tc++;
    }
    return { success: true, todayPresent: tp, todayPraise: tc };
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

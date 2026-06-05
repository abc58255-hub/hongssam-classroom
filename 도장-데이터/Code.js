// 홍쌤 도장-데이터 (교사 전용 데이터 뷰)
// 공유 스프레드시트의 도장 데이터를 집계해 그리드로 보여준다.
// ⚠️ 교사 전용 화면 — 소외 학생·등수 등 민감정보 포함. 학생에게 노출 금지.

var SHEET_ID = '1jK7gYGFXCe3FULLs5mKttP959Aa9vp8-WNOGdJy7cZQ';

// 소외 기준 (마지막 발표 이후 경과일)
var STALE_DAYS = 14;

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('홍쌤 도장-데이터')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 필요한 시트가 없으면 생성
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
  mk('도장_이월', ['학번','이름','칭찬이월','발표이월'], '#0d9488');
  mk('도장_학습지', ['학습지명','발표문제JSON'], '#1e3a8a');
  return ss;
}

function _clsOf_(sid) {
  sid = String(sid || '').trim();
  return sid.length >= 2 ? (sid.substring(0,1) + '학년 ' + sid.substring(1,2) + '반') : '';
}

function _daysBetween_(from, to) {
  return Math.floor((to.getTime() - from.getTime()) / 86400000);
}

function _fmtDate_(d) {
  try { return Utilities.formatDate(new Date(d), 'Asia/Seoul', 'yyyy-MM-dd'); }
  catch(_) { return ''; }
}

// 반 목록 (학생명부 기반)
function getClassList() {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var roster = ss.getSheetByName('학생명부').getDataRange().getValues();
    var set = {};
    for (var i = 1; i < roster.length; i++) {
      var c = _clsOf_(roster[i][1]);
      if (c) set[c] = true;
    }
    var list = Object.keys(set).sort();
    return { success: true, classes: list };
  } catch(e) { return { success: false, message: e.toString(), classes: [] }; }
}

// 메인 집계 (카테고리 동적)
// filterClass: '' 또는 'all' = 전체, 그 외 '2학년 1반'
// period: 'all' | '7' | '30'
function getDojangData(filterClass, period, semesterStart) {
  try {
    var ss = _ensureSheets_();
    var cats = _getCategories_(ss);
    var catNames = cats.map(function(c){ return c.name; });
    var staleCat = (cats.filter(function(c){ return c.type === 'sheet'; })[0] || cats[0] || {name:''}).name;

    var roster = ss.getSheetByName('학생명부').getDataRange().getValues();
    var logSh = ss.getSheetByName('도장기록');
    var carrySh = ss.getSheetByName('도장_이월');
    var log = logSh.getLastRow() > 1 ? logSh.getRange(2, 1, logSh.getLastRow()-1, 8).getValues() : [];
    var carry = carrySh.getLastRow() > 1 ? carrySh.getRange(2, 1, carrySh.getLastRow()-1, 4).getValues() : [];

    var now = new Date();
    var cutoff = null;
    var includeCarry = (period === 'all' || !period);
    if (period === '7') cutoff = new Date(now.getTime() - 7*86400000);
    else if (period === '30') cutoff = new Date(now.getTime() - 30*86400000);

    var onlyClass = (filterClass && filterClass !== 'all') ? filterClass : null;

    var students = {}, order = [];
    for (var i = 1; i < roster.length; i++) {
      var sid = String(roster[i][1] || '').trim();
      if (!sid) continue;
      var cls = _clsOf_(sid);
      if (onlyClass && cls !== onlyClass) continue;
      students[sid] = { sid: sid, name: String(roster[i][2] || '').trim(), cls: cls,
                        counts: {}, last: {}, recent: [] };
      order.push(sid);
    }

    // 이월 합산 (이월은 발표/칭찬 두 항목만 존재)
    if (includeCarry) {
      for (var c = 0; c < carry.length; c++) {
        var csid = String(carry[c][0] || '').trim();
        if (!csid || !students[csid]) continue;
        students[csid].counts['칭찬'] = (students[csid].counts['칭찬']||0) + Number(carry[c][2] || 0);
        students[csid].counts['발표'] = (students[csid].counts['발표']||0) + Number(carry[c][3] || 0);
      }
    }

    // 도장기록 집계
    for (var r = 0; r < log.length; r++) {
      var lsid = String(log[r][1] || '').trim();
      if (!lsid || !students[lsid]) continue;
      var dt = log[r][0] ? new Date(log[r][0]) : null;
      if (cutoff && (!dt || dt < cutoff)) continue;
      var kind = String(log[r][3] || '').trim();
      var reason = String(log[r][4] || '');
      var sheetName = String(log[r][6] || '');
      var problem = String(log[r][7] || '');
      var st = students[lsid];
      st.counts[kind] = (st.counts[kind]||0) + 1;
      var info = [sheetName, problem].filter(Boolean).join(' · ') || reason;
      if (dt && (!st.last[kind] || dt > st.last[kind].date)) st.last[kind] = { date: dt, info: info };
      st.recent.push({ date: dt ? _fmtDate_(dt) : '', kind: kind, reason: reason,
                       sheet: sheetName, problem: problem, ts: dt ? dt.getTime() : 0 });
    }

    // 행 구성
    var rows = order.map(function(sid){
      var s = students[sid];
      var lastMap = {};
      catNames.forEach(function(cn){
        if (s.last[cn]) lastMap[cn] = { date: _fmtDate_(s.last[cn].date), elapsed: _daysBetween_(s.last[cn].date, now), info: s.last[cn].info };
      });
      var scnt = s.counts[staleCat] || 0;
      var sElapsed = (lastMap[staleCat] ? lastMap[staleCat].elapsed : null);
      var stale = staleCat ? ((scnt === 0) || (sElapsed !== null && sElapsed >= STALE_DAYS)) : false;
      var total = 0; catNames.forEach(function(cn){ total += (s.counts[cn]||0); });
      s.recent.sort(function(a,b){ return b.ts - a.ts; });
      return { sid: s.sid, name: s.name, cls: s.cls, counts: s.counts, last: lastMap,
               total: total, stale: stale, recent: s.recent.slice(0, 8),
               rankAll: {}, rankClass: {} };
    });

    // 항목별 등수 (전체 + 반)
    catNames.forEach(function(cn){
      var all = rows.slice().sort(function(a,b){ return (b.counts[cn]||0) - (a.counts[cn]||0); });
      _assignRankByCount_(all, cn, 'rankAll');
      var byClass = {};
      rows.forEach(function(r){ (byClass[r.cls] = byClass[r.cls] || []).push(r); });
      Object.keys(byClass).forEach(function(k){
        var arr = byClass[k].slice().sort(function(a,b){ return (b.counts[cn]||0) - (a.counts[cn]||0); });
        _assignRankByCount_(arr, cn, 'rankClass');
      });
    });

    // 기본 정렬: 소외 항목(발표) 많은 순, 그다음 총합
    rows.sort(function(a,b){
      var av = a.counts[staleCat]||0, bv = b.counts[staleCat]||0;
      if (bv !== av) return bv - av;
      if (b.total !== a.total) return b.total - a.total;
      return a.name.localeCompare(b.name, 'ko');
    });

    var summary = { students: rows.length, perCat: {},
                    staleCount: rows.filter(function(r){ return r.stale; }).length };
    catNames.forEach(function(cn){
      summary.perCat[cn] = rows.reduce(function(s,r){ return s + (r.counts[cn]||0); }, 0);
    });

    return { success: true, rows: rows, summary: summary, staleDays: STALE_DAYS,
             categories: cats, staleCat: staleCat };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function _assignRankByCount_(sortedArr, catName, rankField) {
  var rank = 0, prev = null, count = 0;
  for (var i = 0; i < sortedArr.length; i++) {
    count++;
    var v = sortedArr[i].counts[catName] || 0;
    if (prev === null || v !== prev) { rank = count; prev = v; }
    sortedArr[i][rankField][catName] = rank;
  }
}

function _assignRank_(sortedArr, key, rankField) {
  var rank = 0, prev = null, count = 0;
  for (var i = 0; i < sortedArr.length; i++) {
    count++;
    var v = sortedArr[i][key];
    if (prev === null || v !== prev) { rank = count; prev = v; }
    sortedArr[i][rankField] = rank;
  }
}

var DEFAULT_PRAISE = ['박수/격려','질문 답변','친구 도움','적극 참여','발표 경청'];

function _getCategories_(ss) {
  var cats = _getSettingRaw_(ss, 'categories', null);
  if (!Array.isArray(cats) || !cats.length) {
    var pp = _getSettingRaw_(ss, 'praisePresets', DEFAULT_PRAISE);
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

function _getSettingRaw_(ss, key, def) {
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
function _setSetting_(ss, key, value) {
  var sh = ss.getSheetByName('도장_설정');
  if (!sh) { sh = ss.insertSheet('도장_설정'); sh.getRange(1,1,1,2).setValues([['키','값']]); }
  var rows = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow()-1, 1).getValues() : [];
  var v = JSON.stringify(value);
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === key) { sh.getRange(i+2, 2).setValue(v); return; }
  }
  sh.appendRow([key, v]);
}

function getDojangSettings() {
  try {
    var ss = _ensureSheets_();
    return { success: true,
      categories: _getCategories_(ss),
      rankVisibility: _getSettingRaw_(ss, 'rankVisibility', 'show') };
  } catch(e) { return { success: false, message: e.toString() }; }
}
function saveDojangSettings(categories, rankVisibility) {
  try {
    var ss = _ensureSheets_();
    var clean = (categories || []).map(function(c){
      return { name: String(c.name||'').trim(), emoji: String(c.emoji||'🏅').trim() || '🏅',
               type: (c.type==='sheet'?'sheet':'reason'),
               presets: (Array.isArray(c.presets)?c.presets:[]).map(function(p){return String(p).trim();}).filter(Boolean) };
    }).filter(function(c){ return c.name; });
    if (!clean.length) return { success: false, message: '최소 1개 항목이 필요해요.' };
    _setSetting_(ss, 'categories', clean);
    _setSetting_(ss, 'rankVisibility', rankVisibility === 'teacher' ? 'teacher' : 'show');
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}

// ===== 학습지·문제 관리 =====
function getWorksheetsList() {
  try {
    var ss = _ensureSheets_();
    var sh = ss.getSheetByName('도장_학습지');
    var list = [];
    if (sh.getLastRow() > 1) {
      var wd = sh.getRange(2, 1, sh.getLastRow()-1, 2).getValues();
      wd.forEach(function(w){
        var nm = String(w[0]||'').trim(); if (!nm) return;
        var probs = []; try { probs = JSON.parse(w[1]||'[]'); } catch(_) {}
        list.push({ name: nm, problems: Array.isArray(probs)?probs:[] });
      });
    }
    return { success: true, worksheets: list };
  } catch(e) { return { success: false, message: e.toString(), worksheets: [] }; }
}
function saveWorksheet(origName, name, problems) {
  try {
    var ss = _ensureSheets_();
    var sh = ss.getSheetByName('도장_학습지');
    name = String(name||'').trim();
    if (!name) return { success: false, message: '이름 필요' };
    var probs = (Array.isArray(problems)?problems:[]).map(function(p){return String(p).trim();}).filter(Boolean);
    var rows = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow()-1, 1).getValues() : [];
    var key = String(origName||name).trim();
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === key) {
        sh.getRange(i+2, 1, 1, 2).setValues([[name, JSON.stringify(probs)]]);
        return { success: true };
      }
    }
    sh.appendRow([name, JSON.stringify(probs)]);
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}
function deleteWorksheet(name) {
  try {
    var ss = _ensureSheets_();
    var sh = ss.getSheetByName('도장_학습지');
    var rows = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow()-1, 1).getValues() : [];
    for (var i = rows.length-1; i >= 0; i--) {
      if (String(rows[i][0]).trim() === String(name).trim()) { sh.deleteRow(i+2); return { success: true }; }
    }
    return { success: false, message: '없음' };
  } catch(e) { return { success: false, message: e.toString() }; }
}

// 이월 입력용: 반 학생 + 현재 이월값
function getCarryList(cls) {
  try {
    var ss = _ensureSheets_();
    var roster = ss.getSheetByName('학생명부').getDataRange().getValues();
    var carrySh = ss.getSheetByName('도장_이월');
    var carry = carrySh.getLastRow() > 1 ? carrySh.getRange(2, 1, carrySh.getLastRow()-1, 4).getValues() : [];
    var cmap = {};
    carry.forEach(function(c){ cmap[String(c[0]||'').trim()] = { praise: Number(c[2]||0), present: Number(c[3]||0) }; });
    var students = [];
    for (var i = 1; i < roster.length; i++) {
      var sid = String(roster[i][1] || '').trim();
      if (!sid) continue;
      if (cls && cls !== 'all' && _clsOf_(sid) !== cls) continue;
      var c = cmap[sid] || { praise: 0, present: 0 };
      students.push({ sid: sid, name: String(roster[i][2] || '').trim(), praise: c.praise, present: c.present });
    }
    students.sort(function(a,b){ return a.sid.localeCompare(b.sid); });
    return { success: true, students: students };
  } catch(e) { return { success: false, message: e.toString() }; }
}

// 이월 저장 (upsert)
function saveCarry(sid, name, praise, present) {
  try {
    var ss = _ensureSheets_();
    sid = String(sid||'').trim();
    if (!sid) return { success: false };
    praise = Math.max(0, parseInt(praise)||0);
    present = Math.max(0, parseInt(present)||0);
    var sh = ss.getSheetByName('도장_이월');
    var rows = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow()-1, 1).getValues() : [];
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === sid) {
        sh.getRange(i+2, 1, 1, 4).setValues([[sid, String(name||'').trim(), praise, present]]);
        return { success: true };
      }
    }
    sh.appendRow([sid, String(name||'').trim(), praise, present]);
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}

// 한 학생의 전체 도장 내역 (상세)
function getStudentDetail(studentId) {
  try {
    var ss = _ensureSheets_();
    var sid = String(studentId || '').trim();
    var logSh = ss.getSheetByName('도장기록');
    var log = logSh.getLastRow() > 1 ? logSh.getRange(2, 1, logSh.getLastRow()-1, 8).getValues() : [];
    var carrySh = ss.getSheetByName('도장_이월');
    var carry = carrySh.getLastRow() > 1 ? carrySh.getRange(2, 1, carrySh.getLastRow()-1, 4).getValues() : [];

    var items = [];
    var name = '';
    for (var r = 0; r < log.length; r++) {
      if (String(log[r][1] || '').trim() !== sid) continue;
      name = String(log[r][2] || '').trim() || name;
      var dt = log[r][0] ? new Date(log[r][0]) : null;
      items.push({
        date: dt ? _fmtDate_(dt) : '',
        time: dt ? Utilities.formatDate(dt, 'Asia/Seoul', 'HH:mm') : '',
        kind: String(log[r][3] || ''), reason: String(log[r][4] || ''),
        period: String(log[r][5] || ''), sheet: String(log[r][6] || ''),
        problem: String(log[r][7] || ''), ts: dt ? dt.getTime() : 0
      });
    }
    items.sort(function(a,b){ return b.ts - a.ts; });

    var carryRow = { praise: 0, present: 0 };
    for (var c = 0; c < carry.length; c++) {
      if (String(carry[c][0] || '').trim() === sid) {
        carryRow.praise = Number(carry[c][2] || 0);
        carryRow.present = Number(carry[c][3] || 0);
        name = String(carry[c][1] || '').trim() || name;
        break;
      }
    }
    return { success: true, name: name, items: items, carry: carryRow };
  } catch(e) { return { success: false, message: e.toString() }; }
}

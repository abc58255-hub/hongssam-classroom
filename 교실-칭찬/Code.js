// =====================================================
// 교실 칭찬 — 교사용 관리 앱
// 공유 스프레드시트 사용 (학생-담임 앱과 시트 공유)
// =====================================================
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
  _registerAppUrl_('바로가기_칭찬');
  try { StudentAuth.registerAppUrl('칭찬', ScriptApp.getService().getUrl()); } catch(_) {}
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('교실 칭찬 관리')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── 시트 헬퍼 ──────────────────────────────────────────
function _ss() { return SpreadsheetApp.openById(SHEET_ID); }

function _ensureSheet(name, headers) {
  var ss = _ss();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#7c3aed').setFontColor('white');
    sh.setFrozenRows(1);
  }
  return sh;
}
function _evSheet()  { return _ensureSheet('칭찬이벤트', ['ID','제목','모드','반','시작일','종료일','참여자JSON','옵션JSON','활성','생성일']); }
function _asSheet()  { return _ensureSheet('칭찬배정', ['날짜','이벤트ID','배정JSON']); }
function _recSheet() { return _ensureSheet('칭찬기록', ['일시','이벤트ID','보낸학번','보낸이름','받은학번','받은이름','내용','익명','숨김']); }

function _fmtDate(d) { return Utilities.formatDate(new Date(d), 'Asia/Seoul', 'yyyy-MM-dd'); }
function _clsOfId(sid) {
  sid = String(sid || '').trim();
  return sid.length >= 2 ? (sid.substring(0,1) + '학년 ' + sid.substring(1,2) + '반') : '기타';
}

// ── 반 목록 / 학생 명부 ─────────────────────────────────
function getClassList() {
  try {
    var data = StudentAuth.getRosterValues();
    var set = {};
    for (var i = 1; i < data.length; i++) {
      var sid = String(data[i][1] || '').trim();
      if (!sid) continue;
      var c = _clsOfId(sid);
      if (c !== '기타') set[c] = true;
    }
    return { success: true, classes: Object.keys(set).sort() };
  } catch(e) { return { success: false, message: e.toString(), classes: [] }; }
}

function getRoster(cls) {
  try {
    var data = StudentAuth.getRosterValues();
    var list = [];
    for (var i = 1; i < data.length; i++) {
      var sid = String(data[i][1] || '').trim();
      if (!sid) continue;
      if (cls && _clsOfId(sid) !== cls) continue;
      list.push({ id: sid, name: String(data[i][2] || '').trim() });
    }
    list.sort(function(a,b){ return String(a.id).localeCompare(String(b.id)); });
    return { success: true, roster: list };
  } catch(e) { return { success: false, message: e.toString(), roster: [] }; }
}

// ── 이벤트 CRUD ────────────────────────────────────────
function getEvents() {
  try {
    var sh = _evSheet();
    if (sh.getLastRow() < 2) return { success: true, events: [] };
    var rows = sh.getRange(2, 1, sh.getLastRow()-1, 10).getValues();
    var events = rows.map(function(r, i) {
      var opts = {}; try { opts = JSON.parse(r[7] || '{}'); } catch(_) {}
      var parts = []; try { parts = JSON.parse(r[6] || '[]'); } catch(_) {}
      return {
        rowIdx: i + 2,
        id: String(r[0]), title: String(r[1]), mode: String(r[2]), cls: String(r[3]),
        start: r[4] ? _fmtDate(r[4]) : '', end: r[5] ? _fmtDate(r[5]) : '',
        participants: parts, options: opts, active: String(r[8]).toUpperCase() === 'Y',
        created: r[9] ? _fmtDate(r[9]) : ''
      };
    }).reverse();
    return { success: true, events: events };
  } catch(e) { return { success: false, message: e.toString(), events: [] }; }
}

function saveEvent(data) {
  try {
    var sh = _evSheet();
    var roster = getRoster(data.cls).roster;
    var participants = (data.participants && data.participants.length) ? data.participants : roster.map(function(s){ return s.id; });
    var id = data.id || ('EV' + Date.now());
    var row = [ id, data.title, data.mode, data.cls, data.start, data.end,
                JSON.stringify(participants), JSON.stringify(data.options || {}),
                data.active ? 'Y' : 'N', new Date() ];
    if (data.rowIdx) {
      sh.getRange(data.rowIdx, 1, 1, 10).setValues([row]);
    } else {
      sh.appendRow(row);
    }
    return { success: true, id: id };
  } catch(e) { return { success: false, message: e.toString() }; }
}

function toggleEventActive(rowIdx, active) {
  try { _evSheet().getRange(rowIdx, 9).setValue(active ? 'Y' : 'N'); return { success: true }; }
  catch(e) { return { success: false, message: e.toString() }; }
}

function deleteEvent(rowIdx, eventId) {
  try {
    _evSheet().deleteRow(rowIdx);
    // 관련 배정·기록도 삭제
    [_asSheet(), _recSheet()].forEach(function(sh) {
      if (sh.getLastRow() < 2) return;
      var idCol = (sh.getName() === '칭찬배정') ? 2 : 2;
      var data = sh.getRange(2, 1, sh.getLastRow()-1, sh.getLastColumn()).getValues();
      for (var i = data.length - 1; i >= 0; i--) {
        if (String(data[i][idCol-1]) === String(eventId)) sh.deleteRow(i + 2);
      }
    });
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}

// ── 배정 미리보기/생성 (relay/spotlight) ────────────────
function _shuffle(arr) {
  var a = arr.slice();
  for (var i = a.length - 1; i > 0; i--) {
    var j = Math.floor(Math.random() * (i + 1));
    var t = a[i]; a[i] = a[j]; a[j] = t;
  }
  return a;
}

// 랜덤 derangement (셔플 후 원형 연결) → { 보낸:받은 }
function _makeRelay(ids) {
  if (ids.length < 2) return {};
  var s = _shuffle(ids);
  var map = {};
  for (var i = 0; i < s.length; i++) map[s[i]] = s[(i + 1) % s.length];
  return map;
}

function previewAssignment(eventId) {
  try {
    var ev = getEvents().events.filter(function(e){ return e.id === eventId; })[0];
    if (!ev) return { success: false, message: '이벤트 없음' };
    var today = _fmtDate(new Date());
    var asSh = _asSheet();
    var existing = null, existRow = -1;
    if (asSh.getLastRow() >= 2) {
      var rows = asSh.getRange(2, 1, asSh.getLastRow()-1, 3).getValues();
      for (var i = 0; i < rows.length; i++) {
        if (_fmtDate(rows[i][0]) === today && String(rows[i][1]) === String(eventId)) {
          try { existing = JSON.parse(rows[i][2]); existRow = i + 2; } catch(_) {}
          break;
        }
      }
    }
    var roster = getRoster(ev.cls).roster;
    var nameOf = {}; roster.forEach(function(s){ nameOf[s.id] = s.name; });
    return { success: true, mode: ev.mode, assignment: existing, hasToday: !!existing, nameMap: nameOf };
  } catch(e) { return { success: false, message: e.toString() }; }
}

// ── 결과 조회 ──────────────────────────────────────────
function getEventResults(eventId) {
  try {
    var sh = _recSheet();
    if (sh.getLastRow() < 2) return { success: true, records: [] };
    var rows = sh.getRange(2, 1, sh.getLastRow()-1, 10).getValues(); // 10열=하트
    var records = [];
    rows.forEach(function(r, i) {
      if (String(r[1]) !== String(eventId)) return;
      records.push({
        rowIdx: i + 2,
        time: r[0] ? Utilities.formatDate(new Date(r[0]), 'Asia/Seoul', 'MM/dd HH:mm') : '',
        fromId: String(r[2]), fromName: String(r[3]),
        toId: String(r[4]), toName: String(r[5]),
        content: String(r[6]), anon: String(r[7]).toUpperCase() === 'Y',
        hidden: String(r[8]).toUpperCase() === 'Y',
        hearted: String(r[9]).toUpperCase() === 'Y'
      });
    });
    records.reverse();
    return { success: true, records: records };
  } catch(e) { return { success: false, message: e.toString(), records: [] }; }
}

function deletePraise(rowIdx) {
  try { _recSheet().deleteRow(rowIdx); return { success: true }; }
  catch(e) { return { success: false, message: e.toString() }; }
}

function togglePraiseHidden(rowIdx, hidden) {
  try { _recSheet().getRange(rowIdx, 9).setValue(hidden ? 'Y' : 'N'); return { success: true }; }
  catch(e) { return { success: false, message: e.toString() }; }
}

// ── 통계 ───────────────────────────────────────────────
function getEventStats(eventId) {
  try {
    var recs = getEventResults(eventId).records;
    var roster = [];
    var ev = getEvents().events.filter(function(e){ return e.id === eventId; })[0];
    if (ev) roster = getRoster(ev.cls).roster;
    var nameOf = {}; roster.forEach(function(s){ nameOf[s.id] = s.name; });

    var givenCount = {}, recvCount = {}, heartGivenCount = {}, heartRecvCount = {};
    var gaveSet = {}; // 숨김 포함, 한 번이라도 보낸 학생(실시자)
    recs.forEach(function(r) {
      if (r.fromId) gaveSet[r.fromId] = true;
      if (r.hidden) return;
      givenCount[r.fromId] = (givenCount[r.fromId] || 0) + 1;
      recvCount[r.toId]   = (recvCount[r.toId]   || 0) + 1;
      if (r.hearted) {
        heartGivenCount[r.toId]  = (heartGivenCount[r.toId]  || 0) + 1; // 하트 누른 사람(받은 학생) = 선물한 하트
        heartRecvCount[r.fromId] = (heartRecvCount[r.fromId] || 0) + 1; // 하트 달린 칭찬의 작성자 = 받은 하트
      }
    });
    function toRank(countMap) {
      return Object.keys(countMap).map(function(id){
        return { id: id, name: nameOf[id] || id, count: countMap[id] };
      }).sort(function(a,b){ return b.count - a.count; });
    }
    // 한 번도 못 받은 학생
    var noReceive = roster.filter(function(s){ return !recvCount[s.id]; }).map(function(s){ return s.name; });

    // 미실시자 — 참여 대상 중 한 번도 보내지 않은 학생 (교사 전용, 학생 익명과 무관)
    var partIds = (ev && ev.participants && ev.participants.length) ? ev.participants.map(String) : roster.map(function(s){ return s.id; });
    var notGiven = partIds.filter(function(id){ return !gaveSet[id]; })
      .map(function(id){ return { id: id, name: nameOf[id] || id }; })
      .sort(function(a,b){ return String(a.id).localeCompare(String(b.id), undefined, {numeric:true}); });

    return {
      success: true,
      total: recs.filter(function(r){ return !r.hidden; }).length,
      topGivers: toRank(givenCount),
      topReceivers: toRank(recvCount),
      topHeartsGiven: toRank(heartGivenCount),
      topHeartsRecv: toRank(heartRecvCount),
      topHearts: toRank(heartGivenCount), // 하위호환
      heartTotal: Object.keys(heartGivenCount).reduce(function(s,k){ return s + heartGivenCount[k]; }, 0),
      noReceive: noReceive,
      notGiven: notGiven,
      doneCount: partIds.length - notGiven.length,
      participantCount: partIds.length,
      rosterCount: roster.length
    };
  } catch(e) { return { success: false, message: e.toString() }; }
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

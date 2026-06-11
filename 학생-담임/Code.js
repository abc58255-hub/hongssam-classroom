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

function doGet() { if (!SHEET_ID) return _setupPage_(); return HtmlService.createHtmlOutputFromFile('index').setTitle('우리 반 교실').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL).addMetaTag('viewport', 'width=device-width, initial-scale=1'); }
function getHash(text) { return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text).map(e => (e < 0 ? e + 256 : e).toString(16).padStart(2, '0')).join(''); }

var _sysHrCache_ = null;
function _getSysHrKey_(key) {
  if (!_sysHrCache_) {
    _sysHrCache_ = {};
    try {
      var sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName('시스템설정');
      if (sh && sh.getLastRow() >= 2) {
        var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
        for (var i = 0; i < rows.length; i++) {
          _sysHrCache_[String(rows[i][0]).trim()] = String(rows[i][1] || '').trim();
        }
      }
    } catch(e) {}
  }
  return _sysHrCache_[key] || '';
}

function _getParentFolderId_() {
  var id = _getSysHrKey_('드라이브폴더ID');
  if (id) return id;
  // 미설정 시 자동 생성(같은 이름 폴더가 있으면 재사용) 후 시스템설정에 저장
  var it = DriveApp.getFoldersByName('홍쌤 교실 시스템');
  var folder = it.hasNext() ? it.next() : DriveApp.createFolder('홍쌤 교실 시스템');
  try {
    var sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName('시스템설정');
    if (sh) sh.appendRow(['드라이브폴더ID', folder.getId()]);
  } catch (e) {}
  return folder.getId();
}

function _getSysHr(ss) {
  var sh = ss.getSheetByName('시스템설정');
  if (!sh || sh.getLastRow() < 2) return '';
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === '담임반') return String(rows[i][1] || '').trim();
  }
  // 구 구조 폴백 (B3)
  try { return String(sh.getRange('B3').getValue() || '').trim(); } catch(e) { return ''; }
}

function verifyHomeroomLogin(studentId, studentName, password) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let hrStr = _getSysHr(ss);
  let studentHr = `${String(studentId).substring(0,1)}학년 ${String(studentId).substring(1,2)}반`;
  if (hrStr !== "" && hrStr !== studentHr) return { success: false, message: `🚨 ${hrStr} 학생 전용입니다.` };

  const cache = CacheService.getScriptCache(); const lockKey = "lock_" + studentId;
  if (cache.get(lockKey)) return { success: false, message: "🚨 5회 오류. 10분간 잠금" };
  const rosterData = ss.getSheetByName("학생명부").getDataRange().getValues();
  const inputHash = getHash(password); 
  
  for (let i = 1; i < rosterData.length; i++) {
    if (String(rosterData[i][1]).trim() === String(studentId).trim() && String(rosterData[i][2]).trim() === String(studentName).trim()) {
      if (String(rosterData[i][3]).trim() === inputHash) { cache.remove("fail_" + studentId); return { success: true, homeroom: hrStr }; } 
      else {
        let fails = parseInt(cache.get("fail_" + studentId) || "0") + 1;
        if (fails >= 5) { cache.put(lockKey, "locked", 600); cache.remove("fail_" + studentId); return { success: false, message: "🚨 10분 잠금" }; }
        cache.put("fail_" + studentId, fails.toString(), 600); return { success: false, message: `비번 오류 (${fails}/5)` };
      }
    }
  }
  return { success: false, message: "정보 확인 요망" };
}

function getHomeroomData(studentId) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  // ── 활동목록 (학급활동목록 시트): [0]=번호, [1]=카테고리, [2]=이름, [3]=설명, [4]=폼링크, [5]=필드
  let activities = []; const actSheet = ss.getSheetByName('학급활동목록');
  if (actSheet && actSheet.getLastRow() >= 2) {
    const actData = actSheet.getRange(2, 1, actSheet.getLastRow()-1, 6).getValues();
    actData.forEach(function(r) {
      if (!r[2]) return;
      activities.push({
        category: String(r[1]||'').trim(),
        name:     String(r[2]||'').trim(),
        desc:     String(r[3]||'').trim(),
        formUrl:  String(r[4]||'').trim(),
        fields:   String(r[5]||'').trim()  // ✅ 추가 입력항목 JSON
      });
    });
    activities.reverse();
  }
  let myRecords = [];
  const subSheet = ss.getSheetByName("창체제출현황");
  if (subSheet) {
    const subData = subSheet.getDataRange().getValues();
    const sidStr = String(studentId).trim();
    for (let i = 1; i < subData.length; i++) {
      if (String(subData[i][1]).trim() !== sidStr) continue;
      myRecords.push({
        date: Utilities.formatDate(new Date(subData[i][0]), "Asia/Seoul", "MM/dd HH:mm"),
        activity: subData[i][3],
        role: subData[i][4],
        reflection: subData[i][5],
        url: subData[i][6]
      });
    }
  }
  
  let roster = []; let hrStr = _getSysHr(ss);
  const rosterData = ss.getSheetByName("학생명부").getDataRange().getValues();
  for(let i=1; i<rosterData.length; i++) {
    let sid = String(rosterData[i][1]).trim();
    if (`${sid.substring(0,1)}학년 ${sid.substring(1,2)}반` === hrStr) roster.push({ id: sid, name: String(rosterData[i][2]).trim() });
  }

  let activeSurvey = null; const svSheet = ss.getSheetByName("설문목록");
  if(svSheet) { let svData = svSheet.getDataRange().getValues(); for(let i=1; i<svData.length; i++) { if(svData[i][3] === "진행중") { activeSurvey = { id: svData[i][0], title: svData[i][2], questions: svData[i][4] }; break; } } }
  
  let hasSubmittedSurvey = false;
  if(activeSurvey) {
    const srSheet = ss.getSheetByName("설문응답");
    if(srSheet) {
      let srData = srSheet.getDataRange().getValues();
      for(let i=1; i<srData.length; i++) { if(srData[i][1] === activeSurvey.id && String(srData[i][2]).trim() === String(studentId).trim()) { hasSubmittedSurvey = true; break; } }
    }
  }

  // ✅ 학급 알림 읽기 (학급알림 시트 — 교사 대시보드에서 작성)
  // 시트 구조: A=제목, B=내용, C=유형, D=날짜, E=표시여부, F=중요여부
  let notices = [];
  try {
    const noticeSheet = ss.getSheetByName("학급알림");
    if (noticeSheet && noticeSheet.getLastRow() >= 2) {
      const noticeData = noticeSheet.getRange(2, 1, noticeSheet.getLastRow() - 1, 6).getValues();
      const today = new Date(); today.setHours(0,0,0,0);
      for (let i = 0; i < noticeData.length; i++) {
        const r = noticeData[i];
        if (!r[0]) continue; // 제목 없으면 스킵
        const visible = r[4] === true || String(r[4]).trim() === 'TRUE' || String(r[4]).trim() === '1';
        if (!visible) continue; // 숨김 처리된 알림은 학생에게 안 보임
        // D-day 계산
        let dday = null;
        if (r[3]) {
          const eventDate = new Date(r[3]); eventDate.setHours(0,0,0,0);
          const diff = Math.round((eventDate - today) / (1000 * 60 * 60 * 24));
          if (diff === 0) dday = 'D-Day!';
          else if (diff > 0) dday = 'D-' + diff;
          else dday = 'D+' + Math.abs(diff);
        }
        notices.push({
          title:     String(r[0] || '').trim(),
          content:   String(r[1] || '').trim(),
          type:      String(r[2] || '알림').trim(),
          date:      r[3] instanceof Date ? Utilities.formatDate(r[3], 'Asia/Seoul', 'yyyy-MM-dd') : String(r[3] || '').trim(),
          important: r[5] === true || String(r[5]).trim() === 'TRUE' || String(r[5]).trim() === '1',
          dday:      dday
        });
      }
      // 중요 알림 먼저, 그 다음 날짜순
      notices.sort(function(a, b) {
        if (a.important && !b.important) return -1;
        if (!a.important && b.important) return 1;
        return (b.date || '').localeCompare(a.date || '');
      });
    }
  } catch(e) {}

  return { activities: activities.reverse(), myRecords: myRecords.reverse(), roster: roster.sort((a,b)=>a.id.localeCompare(b.id)), activeSurvey: activeSurvey, hasSubmittedSurvey: hasSubmittedSurvey, notices: notices, fcmRegisterUrl: _getSysHrKey_('FCM_REGISTER_URL') };
}

// 설문 중복 제출 방지
function submitSurveyResponse(payload) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName("설문응답");
    const data = sheet.getDataRange().getValues();
    for(let i = 1; i < data.length; i++) {
      if(data[i][1] === payload.svId && String(data[i][2]).trim() === String(payload.stuId).trim()) {
        return { success: false, message: "🚨 이미 이 설문을 제출하셨습니다! 한 번만 참여 가능합니다." };
      }
    }
    sheet.appendRow([new Date(), payload.svId, payload.stuId, payload.stuName, payload.answers]);
    return { success: true };
  } catch(e) { 
    return { success: false, message: e.toString() }; 
  }
}

function processHomeroomForm(data) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName("창체제출현황"); const parentFolder = DriveApp.getFolderById(_getParentFolderId_());
    const records = sheet.getDataRange().getValues(); for (let i = records.length - 1; i >= 1; i--) { if (String(records[i][1]).trim() === String(data.studentId).trim() && String(records[i][3]).trim() === String(data.activity).trim()) { return { success: false, message: "이미 제출 완료!" }; } }
    let fileUrl = "첨부파일 없음";
    if (data.fileData) { let hrFolder = parentFolder.getFoldersByName("학급기록").hasNext() ? parentFolder.getFoldersByName("학급기록").next() : parentFolder.createFolder("학급기록"); let actFolder = hrFolder.getFoldersByName(data.activity).hasNext() ? hrFolder.getFoldersByName(data.activity).next() : hrFolder.createFolder(data.activity); const blob = Utilities.newBlob(Utilities.base64Decode(data.fileData), data.fileMimeType, `[${data.studentId}] ${data.studentName}_${data.activity}.${data.fileName.split('.').pop()}`); fileUrl = actFolder.createFile(blob).getUrl(); }
    const extra = data.extraAnswers ? (typeof data.extraAnswers === 'object' ? JSON.stringify(data.extraAnswers) : String(data.extraAnswers)) : '';
    sheet.appendRow([new Date(), data.studentId, data.studentName, data.activity, data.role, data.reflection, fileUrl, extra]);
    return { success: true };
  } catch (e) { return { success: false, message: e.toString() }; }
}

// =====================================================
// ✅ 신고 접수
// 시트: 신고접수 (자동 생성)
// 구조: A=접수일시, B=학번, C=이름, D=유형, E=내용, F=처리여부
// =====================================================
function submitReport(data) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName('신고접수');
    if (!sheet) {
      sheet = ss.insertSheet('신고접수');
      sheet.getRange(1, 1, 1, 6).setValues([['접수일시','학번','이름','유형','내용','처리여부']]);
      sheet.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#ef4444').setFontColor('white');
      sheet.setColumnWidth(5, 400);
    }
    sheet.appendRow([new Date(), data.studentId || '', data.studentName || '', data.category || '', data.content || '', '미처리']);
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}

// =====================================================
// 💝 칭찬 기능 (교사용 칭찬앱과 시트 공유)
// 시트: 칭찬이벤트, 칭찬배정, 칭찬기록
// =====================================================
function _prClsOf(sid) {
  sid = String(sid || '').trim();
  return sid.length >= 2 ? (sid.substring(0,1) + '학년 ' + sid.substring(1,2) + '반') : '기타';
}
function _prFmtDate(d) { return Utilities.formatDate(new Date(d), 'Asia/Seoul', 'yyyy-MM-dd'); }
function _prTime(d) { return d ? Utilities.formatDate(new Date(d), 'Asia/Seoul', 'MM/dd HH:mm') : ''; }

function _prRoster(cls) {
  var data = SpreadsheetApp.openById(SHEET_ID).getSheetByName('학생명부').getDataRange().getValues();
  var list = [];
  for (var i = 1; i < data.length; i++) {
    var sid = String(data[i][1] || '').trim();
    if (!sid) continue;
    if (cls && _prClsOf(sid) !== cls) continue;
    list.push({ id: sid, name: String(data[i][2] || '').trim() });
  }
  return list;
}

// 활성 이벤트 중 이 학생 반에 해당하는 것
function getActivePraiseEvents(studentId) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sh = ss.getSheetByName('칭찬이벤트');
    if (!sh || sh.getLastRow() < 2) return { events: [] };
    var myCls = _prClsOf(studentId);
    var today = _prFmtDate(new Date());
    var rows = sh.getRange(2, 1, sh.getLastRow()-1, 10).getValues();
    var events = [];
    rows.forEach(function(r) {
      if (String(r[8]).toUpperCase() !== 'Y') return;       // 활성만
      if (String(r[3]).trim() !== myCls) return;            // 우리 반만
      var start = r[4] ? _prFmtDate(r[4]) : '';
      var end = r[5] ? _prFmtDate(r[5]) : '';
      if (start && today < start) return;
      if (end && today > end) return;
      events.push({ id: String(r[0]), title: String(r[1]), mode: String(r[2]) });
    });
    return { events: events };
  } catch(e) { return { events: [], message: e.toString() }; }
}

function _prGetEventRow(eventId) {
  var sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName('칭찬이벤트');
  if (!sh || sh.getLastRow() < 2) return null;
  var rows = sh.getRange(2, 1, sh.getLastRow()-1, 10).getValues();
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][0]) === String(eventId)) {
      var opts = {}; try { opts = JSON.parse(rows[i][7] || '{}'); } catch(_) {}
      var parts = []; try { parts = JSON.parse(rows[i][6] || '[]'); } catch(_) {}
      return { id:String(rows[i][0]), title:String(rows[i][1]), mode:String(rows[i][2]), cls:String(rows[i][3]),
               participants: parts, options: opts };
    }
  }
  return null;
}

function _prShuffle(arr){ var a=arr.slice(); for(var i=a.length-1;i>0;i--){var j=Math.floor(Math.random()*(i+1));var t=a[i];a[i]=a[j];a[j]=t;} return a; }

// 오늘 배정 가져오기/생성 (relay, spotlight) — LockService로 동시 생성 방지
function _prEnsureAssignment(ev, today) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName('칭찬배정') || ss.insertSheet('칭찬배정');
  if (sh.getLastRow() === 0) sh.appendRow(['날짜','이벤트ID','배정JSON']);
  // 기존 조회
  if (sh.getLastRow() >= 2) {
    var rows = sh.getRange(2, 1, sh.getLastRow()-1, 3).getValues();
    for (var i = 0; i < rows.length; i++) {
      if (_prFmtDate(rows[i][0]) === today && String(rows[i][1]) === String(ev.id)) {
        try { return JSON.parse(rows[i][2]); } catch(_) { return null; }
      }
    }
  }
  // 생성 (Lock)
  var lock = LockService.getScriptLock();
  try { lock.waitLock(8000); } catch(e) { return null; }
  try {
    // Lock 후 재확인
    if (sh.getLastRow() >= 2) {
      var rows2 = sh.getRange(2, 1, sh.getLastRow()-1, 3).getValues();
      for (var j = 0; j < rows2.length; j++) {
        if (_prFmtDate(rows2[j][0]) === today && String(rows2[j][1]) === String(ev.id)) {
          try { return JSON.parse(rows2[j][2]); } catch(_) { return null; }
        }
      }
    }
    var ids = ev.participants && ev.participants.length ? ev.participants : _prRoster(ev.cls).map(function(s){return s.id;});
    var assignment;
    if (ev.mode === 'spotlight') {
      // 최근 주인공 안 겹치게: 과거 주인공 목록 제외 후 랜덤, 다 돌면 리셋
      var past = _prPastSpotlights(ev.id);
      var pool = ids.filter(function(x){ return past.indexOf(x) < 0; });
      if (pool.length === 0) pool = ids;
      var star = pool[Math.floor(Math.random()*pool.length)];
      assignment = { spotlight: star };
    } else { // relay: 랜덤 derangement
      var map = {};
      if (ids.length >= 2) { var s = _prShuffle(ids); for (var k=0;k<s.length;k++) map[s[k]] = s[(k+1)%s.length]; }
      assignment = { map: map };
    }
    sh.appendRow([today, ev.id, JSON.stringify(assignment)]);
    return assignment;
  } finally { lock.releaseLock(); }
}

function _prPastSpotlights(eventId) {
  var sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName('칭찬배정');
  var out = [];
  if (!sh || sh.getLastRow() < 2) return out;
  var rows = sh.getRange(2, 1, sh.getLastRow()-1, 3).getValues();
  rows.forEach(function(r){
    if (String(r[1]) !== String(eventId)) return;
    try { var a = JSON.parse(r[2]); if (a.spotlight) out.push(a.spotlight); } catch(_) {}
  });
  return out;
}

// 칭찬기록 시트
function _prRecSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName('칭찬기록');
  if (!sh) {
    sh = ss.insertSheet('칭찬기록');
    sh.getRange(1,1,1,9).setValues([['일시','이벤트ID','보낸학번','보낸이름','받은학번','받은이름','내용','익명','숨김']]);
    sh.getRange(1,1,1,9).setFontWeight('bold').setBackground('#7c3aed').setFontColor('white');
    sh.setFrozenRows(1);
  }
  return sh;
}

function _prRecords(eventId) {
  var sh = _prRecSheet();
  if (sh.getLastRow() < 2) return [];
  var rows = sh.getRange(2, 1, sh.getLastRow()-1, 9).getValues();
  var out = [];
  rows.forEach(function(r){
    if (String(r[1]) !== String(eventId)) return;
    out.push({ time:_prTime(r[0]), dateStr:_prFmtDate(r[0]), fromId:String(r[2]), fromName:String(r[3]),
               toId:String(r[4]), toName:String(r[5]), content:String(r[6]),
               anon:String(r[7]).toUpperCase()==='Y', hidden:String(r[8]).toUpperCase()==='Y' });
  });
  return out;
}

// 이벤트 상세 (학생용)
function getPraiseEventDetail(eventId, studentId, studentName) {
  try {
    var ev = _prGetEventRow(eventId);
    if (!ev) return { success: false, message: '이벤트를 찾을 수 없어요.' };
    var today = _prFmtDate(new Date());
    var roster = _prRoster(ev.cls);
    var nameOf = {}; roster.forEach(function(s){ nameOf[s.id] = s.name; });
    var recs = _prRecords(eventId);
    var opts = ev.options || {};

    // 내가 받은 / 보낸
    var received = recs.filter(function(r){ return r.toId === studentId && !r.hidden; })
      .map(function(r){ return { content:r.content, fromName:r.fromName, anon:r.anon, time:r.time }; }).reverse();
    var given = recs.filter(function(r){ return r.fromId === studentId; })
      .map(function(r){ return { content:r.content, toName:r.toName, time:r.time }; }).reverse();

    // 반 게시판 (옵션 class일 때만)
    var board = [];
    if (opts.board === 'class') {
      board = recs.filter(function(r){ return !r.hidden; })
        .map(function(r){ return { fromName:r.fromName, toName:r.toName, content:r.content, anon:r.anon }; }).reverse().slice(0, 50);
    }

    // 오늘 작성 횟수
    var todayCount = recs.filter(function(r){ return r.fromId === studentId && r.dateStr === today; }).length;
    var dailyLimit = opts.dailyLimit || 0;

    // 모드별 대상/후보
    var result = {
      success: true, title: ev.title, mode: ev.mode,
      anonChoice: opts.anon === 'choice',
      received: received, given: given, board: board,
      doneToday: false, canWrite: true
    };

    if (ev.mode === 'relay') {
      var asg = _prEnsureAssignment(ev, today);
      var tid = asg && asg.map ? asg.map[studentId] : null;
      if (!tid) { result.canWrite = false; result.doneToday = false; result.message = '오늘 배정이 없어요.'; return result; }
      // 오늘 이미 그 대상에게 보냈으면 완료
      var didToday = recs.some(function(r){ return r.fromId===studentId && r.toId===tid && r.dateStr===today; });
      result.targetId = tid; result.targetName = nameOf[tid] || tid;
      result.canWrite = !didToday; result.doneToday = didToday;
    } else if (ev.mode === 'spotlight') {
      var asg2 = _prEnsureAssignment(ev, today);
      var star = asg2 ? asg2.spotlight : null;
      result.targetId = star; result.targetName = nameOf[star] || star;
      if (star === studentId) { result.canWrite = false; result.doneToday = false; result.message = '오늘은 내가 주인공이에요! 🎉'; result.iAmStar = true; }
      else {
        var didStar = recs.some(function(r){ return r.fromId===studentId && r.toId===star && r.dateStr===today; });
        result.canWrite = !didStar; result.doneToday = didStar;
      }
    } else {
      // random / pick / chain → 후보 = 나 제외 전원
      result.candidates = roster.filter(function(s){ return s.id !== studentId; });
      if (dailyLimit > 0 && todayCount >= dailyLimit) { result.canWrite = false; result.doneToday = true; }
    }
    return result;
  } catch(e) { return { success: false, message: e.toString() }; }
}

// 칭찬 제출
function submitPraise(data) {
  try {
    var ev = _prGetEventRow(data.eventId);
    if (!ev) return { success: false, message: '이벤트를 찾을 수 없어요.' };
    var opts = ev.options || {};
    var content = String(data.content || '').trim();
    var minLen = opts.minLen || 0;
    if (content.length < minLen) return { success: false, message: '칭찬은 최소 ' + minLen + '자 이상 적어주세요.' };
    if (!data.toId) return { success: false, message: '칭찬할 친구를 선택해주세요.' };
    if (data.toId === data.fromId) return { success: false, message: '자기 자신은 칭찬할 수 없어요.' };

    var today = _prFmtDate(new Date());
    var recs = _prRecords(data.eventId);

    // 하루 횟수 제한
    var dailyLimit = opts.dailyLimit || 0;
    var todayCount = recs.filter(function(r){ return r.fromId===data.fromId && r.dateStr===today; }).length;
    if (dailyLimit > 0 && todayCount >= dailyLimit) return { success: false, message: '오늘 작성 가능한 칭찬을 다 썼어요.' };

    // relay/spotlight: 오늘 같은 대상 중복 방지
    if (ev.mode === 'relay' || ev.mode === 'spotlight') {
      var dup = recs.some(function(r){ return r.fromId===data.fromId && r.toId===data.toId && r.dateStr===today; });
      if (dup) return { success: false, message: '오늘 이미 이 친구를 칭찬했어요.' };
    }

    // 익명 결정
    var anon = false;
    if (opts.anon === 'anon') anon = true;
    else if (opts.anon === 'choice') anon = !!data.anon;

    var roster = _prRoster(ev.cls);
    var nameOf = {}; roster.forEach(function(s){ nameOf[s.id] = s.name; });
    var toName = nameOf[data.toId] || data.toId;

    _prRecSheet().appendRow([ new Date(), data.eventId, data.fromId, data.fromName || '', data.toId, toName, content, anon ? 'Y' : 'N', 'N' ]);
    return { success: true };
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

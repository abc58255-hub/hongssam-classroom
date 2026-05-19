const SHEET_ID = PropertiesService.getScriptProperties().getProperty('SHEET_ID')
  || '1jK7gYGFXCe3FULLs5mKttP959Aa9vp8-WNOGdJy7cZQ';

function doGet() { return HtmlService.createHtmlOutputFromFile('index').setTitle('우리 반 교실').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL).addMetaTag('viewport', 'width=device-width, initial-scale=1'); }
function getHash(text) { return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text).map(e => (e < 0 ? e + 256 : e).toString(16).padStart(2, '0')).join(''); }

function _getSysHrKey_(key) {
  var sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName('시스템설정');
  if (!sh || sh.getLastRow() < 2) return '';
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === key) return String(rows[i][1] || '').trim();
  }
  return '';
}

function _getParentFolderId_() {
  return _getSysHrKey_('드라이브폴더ID') || '1nmo4ZtQYK3-0PFjMKO8yzlkNOVoLn9_H';
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
  let myRecords = []; const subSheet = ss.getSheetByName("창체제출현황");
  if (subSheet) { let subData = subSheet.getDataRange().getValues(); for (let i = 1; i < subData.length; i++) { if (String(subData[i][1]).trim() === String(studentId).trim()) { myRecords.push({ date: Utilities.formatDate(new Date(subData[i][0]), "Asia/Seoul", "MM/dd HH:mm"), activity: subData[i][3], role: subData[i][4], reflection: subData[i][5], url: subData[i][6] }); } } }
  
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
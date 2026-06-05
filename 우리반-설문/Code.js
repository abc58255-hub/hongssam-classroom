// =====================================================
// 우리 반 설문 — 교사용 통합 앱 (제작 + 통계 분석)
// 공유 스프레드시트 사용 (설문목록 / 설문응답 시트)
// =====================================================
const SHEET_ID = PropertiesService.getScriptProperties().getProperty('SHEET_ID')
  || '1jK7gYGFXCe3FULLs5mKttP959Aa9vp8-WNOGdJy7cZQ';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('우리 반 설문')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function _getSys(ss, key) {
  var sh = ss.getSheetByName('시스템설정');
  if (!sh || sh.getLastRow() < 2) return '';
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === key) return String(rows[i][1] || '').trim();
  }
  return '';
}

// 초기 데이터: 담임반 + 명부 + 설문 목록/응답
function getInitData() {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var hrStr = _getSys(ss, '담임반') || '';
    // 명부
    var roster = [];
    var rd = ss.getSheetByName('학생명부').getDataRange().getValues();
    for (var i = 1; i < rd.length; i++) {
      var sid = String(rd[i][1] || '').trim();
      if (!sid) continue;
      var cls = sid.length >= 2 ? (sid.substring(0,1) + '학년 ' + sid.substring(1,2) + '반') : '기타';
      roster.push({ id: sid, name: String(rd[i][2] || '').trim(), cls: cls });
    }
    var sv = getSurveyData();
    return { success: true, homeroom: hrStr, roster: roster,
             surveys: sv.surveys || [], surveyRes: sv.surveyRes || [] };
  } catch(e) { return { success: false, message: e.toString() }; }
}

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

function saveNewSurvey(payload) {
  try {
    SpreadsheetApp.openById(SHEET_ID).getSheetByName("설문목록").appendRow(["SV_" + new Date().getTime(), new Date(), payload.title, "대기중", payload.questions, payload.guide||""]);
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}

function updateSurvey(payload) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('설문목록');
    if (!sheet) return { success: false, message: '설문목록 시트 없음' };
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(payload.svId).trim()) {
        sheet.getRange(i+1, 3).setValue(payload.title);
        sheet.getRange(i+1, 5).setValue(payload.questions);
        if (payload.guide !== undefined) sheet.getRange(i+1, 6).setValue(payload.guide || '');
        return { success: true };
      }
    }
    return { success: false, message: '해당 설문을 찾을 수 없습니다.' };
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

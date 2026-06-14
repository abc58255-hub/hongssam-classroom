// =====================================================
// 우리 반 설문 — 교사용 (제작 + 통계 분석)
// 데이터: 전용 '설문 시트'(설문목록·설문응답)
// 공통(명부·설정·앱URL)은 ClassCore 라이브러리(식별자 StudentAuth)에 위임
// =====================================================

// ── 설문 시트 (전용) ───────────────────────────────────────
var _cachedSurveySs = null;
function _surveySs() {
  if (_cachedSurveySs) return _cachedSurveySs;
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('SURVEY_SHEET_ID');
  var ss;
  if (id) { try { ss = SpreadsheetApp.openById(id); } catch(_) {} }
  if (!ss) { ss = SpreadsheetApp.create('설문 데이터'); props.setProperty('SURVEY_SHEET_ID', ss.getId()); }
  _ensureSurveySheets(ss);
  // 설문 시트 ID를 ClassCore에 공유 (다른 앱·포털 참조용)
  try { if (StudentAuth.getConfig('설문시트ID', '') !== ss.getId()) StudentAuth.setConfig('설문시트ID', ss.getId()); } catch(_) {}
  _cachedSurveySs = ss;
  return ss;
}

function _ensureSurveySheets(ss) {
  function mk(name, headers) {
    if (ss.getSheetByName(name)) return;
    var sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold').setBackground('#0d9488').setFontColor('white');
    sh.setFrozenRows(1);
  }
  mk('설문목록', ['ID', '날짜', '제목', '상태', '질문JSON', '안내문']);
  mk('설문응답', ['날짜', '설문ID', '학번', '이름', '답변']);
  var def = ss.getSheetByName('시트1') || ss.getSheetByName('Sheet1');
  if (def && ss.getSheets().length > 1) { try { ss.deleteSheet(def); } catch(_) {} }
}

// 이 앱의 배포 URL을 ClassCore 앱 URL 레지스트리에 기록 (포털 카드용)
function _registerAppUrl_(key) {
  try {
    var cache = CacheService.getScriptCache();
    var ck = 'appUrlReg_' + key;
    if (cache && cache.get(ck)) return;
    var url = ScriptApp.getService().getUrl() || '';
    if (!url || url.indexOf('/dev') >= 0) return;
    StudentAuth.registerAppUrl(key.replace('바로가기_', ''), url);
    if (cache) cache.put(ck, '1', 21600);
  } catch (e) {}
}

function doGet() {
  _registerAppUrl_('바로가기_설문');
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('우리 반 설문')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 초기 데이터: 담임반 + 명부 + 설문 목록/응답
function getInitData() {
  try {
    var hrStr = StudentAuth.getConfig('담임반', '');
    var roster = [];
    var rd = StudentAuth.getRosterValues();
    for (var i = 1; i < rd.length; i++) {
      var sid = String(rd[i][1] || '').trim();
      if (!sid) continue;
      var cls = sid.length >= 2 ? (sid.substring(0, 1) + '학년 ' + sid.substring(1, 2) + '반') : '기타';
      roster.push({ id: sid, name: String(rd[i][2] || '').trim(), cls: cls });
    }
    var sv = getSurveyData();
    return { success: true, homeroom: hrStr, roster: roster, surveys: sv.surveys || [], surveyRes: sv.surveyRes || [] };
  } catch(e) { return { success: false, message: e.toString() }; }
}

function getSurveyData() {
  try {
    var ss = _surveySs();
    var surveys = [], surveyRes = [];
    var svSheet = ss.getSheetByName("설문목록");
    if (svSheet && svSheet.getLastRow() >= 2) {
      var svData = svSheet.getDataRange().getValues();
      for (var i = 1; i < svData.length; i++) {
        if (svData[i][0]) surveys.push({
          id: svData[i][0],
          date: svData[i][1] ? Utilities.formatDate(new Date(svData[i][1]), "Asia/Seoul", "yyyy-MM-dd") : "",
          title: svData[i][2], status: svData[i][3], questions: svData[i][4],
          guide: String(svData[i][5] || '').trim()
        });
      }
    }
    var srSheet = ss.getSheetByName("설문응답");
    if (srSheet && srSheet.getLastRow() >= 2) {
      var srData = srSheet.getDataRange().getValues();
      for (var i = 1; i < srData.length; i++) {
        if (srData[i][1]) surveyRes.push({
          date: srData[i][0] ? Utilities.formatDate(new Date(srData[i][0]), "Asia/Seoul", "MM/dd HH:mm") : "",
          svId: srData[i][1], stuId: srData[i][2], stuName: srData[i][3], answers: srData[i][4]
        });
      }
    }
    return { success: true, surveys: surveys.reverse(), surveyRes: surveyRes };
  } catch(e) { return { success: false, message: e.toString() }; }
}

function saveNewSurvey(payload) {
  try {
    _surveySs().getSheetByName("설문목록").appendRow(["SV_" + new Date().getTime(), new Date(), payload.title, "대기중", payload.questions, payload.guide || ""]);
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}

function updateSurvey(payload) {
  try {
    var sheet = _surveySs().getSheetByName('설문목록');
    if (!sheet) return { success: false, message: '설문목록 시트 없음' };
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(payload.svId).trim()) {
        sheet.getRange(i + 1, 3).setValue(payload.title);
        sheet.getRange(i + 1, 5).setValue(payload.questions);
        if (payload.guide !== undefined) sheet.getRange(i + 1, 6).setValue(payload.guide || '');
        return { success: true };
      }
    }
    return { success: false, message: '해당 설문을 찾을 수 없습니다.' };
  } catch(e) { return { success: false, message: e.toString() }; }
}

function updateSurveyStatus(svId, newStatus) {
  try {
    var sheet = _surveySs().getSheetByName("설문목록");
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === svId) {
        if (newStatus === "진행중") {
          for (var j = 1; j < data.length; j++) {
            if (data[j][3] === "진행중") sheet.getRange(j + 1, 4).setValue("마감됨");
          }
        }
        sheet.getRange(i + 1, 4).setValue(newStatus); return { success: true };
      }
    }
    return { success: false, message: "설문 찾기 실패" };
  } catch(e) { return { success: false, message: e.toString() }; }
}

function deleteSurvey(svId) {
  try {
    var sheet = _surveySs().getSheetByName('설문목록');
    if (!sheet) return { success: false, message: '설문목록 시트 없음' };
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(svId).trim()) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, message: '해당 설문을 찾을 수 없습니다.' };
  } catch(e) { return { success: false, message: e.toString() }; }
}

// 최초 1회: 설문 시트 생성 확인 (편집기에서 실행)
function setup() {
  var ss = _surveySs();
  Logger.log('✅ 설문 시트 준비 완료: ' + ss.getUrl());
  return ss.getUrl();
}

// 데이터 이전: 기존 홍쌤 공유시트의 설문목록·설문응답 → 설문 시트 (1회 실행)
function importFromBoard() {
  try {
    var it = DriveApp.getFilesByName('홍쌤교실시스템_SHEET_ID');
    if (!it.hasNext()) return { success: false, message: '홍쌤 연결 파일(마커)을 찾을 수 없습니다. importFromSheet("시트ID")로 직접 지정하세요.' };
    return importFromSheet(String(it.next().getBlob().getDataAsString() || '').trim());
  } catch(e) { return { success: false, message: e.toString() }; }
}

function importFromSheet(srcId) {
  try {
    var src = SpreadsheetApp.openById(srcId);
    var dst = _surveySs();
    var names = ['설문목록', '설문응답'];
    var report = [];
    names.forEach(function(name) {
      var s = src.getSheetByName(name);
      var d = dst.getSheetByName(name);
      if (!s || s.getLastRow() < 2 || !d) { report.push(name + ': 0행'); return; }
      var cols = Math.min(s.getLastColumn(), d.getLastColumn());
      var data = s.getRange(2, 1, s.getLastRow() - 1, cols).getValues();
      if (d.getLastRow() > 1) d.getRange(2, 1, d.getLastRow() - 1, d.getLastColumn()).clearContent();
      d.getRange(2, 1, data.length, cols).setValues(data);
      report.push(name + ': ' + data.length + '행');
    });
    Logger.log('✅ 이전 완료 — ' + report.join(' / '));
    return { success: true, report: report.join(' / ') };
  } catch(e) { return { success: false, message: e.toString() }; }
}

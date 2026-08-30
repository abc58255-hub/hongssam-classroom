// 창체(학급활동) 관리 앱 — 교사가 창의적 체험활동을 출제·관리하는 도구
// 데이터: 전용 "창체 시트"(학급활동목록·창체제출현황)
// 공통(학생 인증·설정·키·교사 인증·드라이브폴더)은 ClassCore 라이브러리에 위임 — 하드코딩 없음
//
// 컬럼 구조(홍쌤 교실 시스템 호환, 데이터 복붙 가능):
//   학급활동목록 : 번호·카테고리·활동명·설명·폼링크·필드JSON
//   창체제출현황 : 날짜·학번·이름·활동명·역할·소감·파일URL·추가답변

// ── 창체 시트 (전용) ───────────────────────────────────────
var _cachedSs = null;

function _changcheSs() {
  if (_cachedSs) return _cachedSs;
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('CHANGCHE_SHEET_ID');
  var ss;
  if (id) { try { ss = SpreadsheetApp.openById(id); } catch(_) {} }
  if (!ss) {
    ss = SpreadsheetApp.create('창체 데이터');
    props.setProperty('CHANGCHE_SHEET_ID', ss.getId());
  }
  _ensureSheets(ss);
  // 창체 시트 ID를 ClassCore에 공유 → 학생-담임 등 다른 앱이 같은 시트를 참조
  try { if (ClassCore.getConfig('창체시트ID', '') !== ss.getId()) ClassCore.setConfig('창체시트ID', ss.getId()); } catch(_) {}
  _cachedSs = ss;
  return ss;
}

function _ensureSheets(ss) {
  function mk(name, headers, color) {
    if (ss.getSheetByName(name)) return;
    var sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold').setBackground(color).setFontColor('white');
    sh.setFrozenRows(1);
  }
  mk('학급활동목록', ['번호','카테고리','활동명','설명','폼링크','필드JSON'], '#0d9488');
  mk('창체제출현황', ['날짜','학번','이름','활동명','역할','소감','파일URL','추가답변'], '#0d9488');
  var def = ss.getSheetByName('시트1') || ss.getSheetByName('Sheet1');
  if (def && ss.getSheets().length > 1) { try { ss.deleteSheet(def); } catch(_) {} }
}

// 창체 첨부파일용 드라이브 폴더 — 폴더ID는 ClassCore 공통 설정 사용
function _parentFolder() {
  var id = ClassCore.getConfig('드라이브폴더ID', '');
  if (id) { try { return DriveApp.getFolderById(id); } catch(_) {} }
  var it = DriveApp.getFoldersByName('홍쌤 교실 시스템');
  var folder = it.hasNext() ? it.next() : DriveApp.createFolder('홍쌤 교실 시스템');
  ClassCore.setConfig('드라이브폴더ID', folder.getId());
  return folder;
}
function _getParentFolderId_() { return _parentFolder().getId(); }

// ── 라우팅 ──
function doGet() {
  ClassCore.registerAppUrl('창체관리', _serviceUrl_());
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('창체 관리')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function _serviceUrl_() {
  try { return ScriptApp.getService().getUrl() || ''; } catch(_) { return ''; }
}

// ── 교사 인증 (ClassCore 위임) ──
function teacherLogin(pw)            { return ClassCore.teacherLogin(pw); }
function verifyTeacher(token)        { return ClassCore.verifyTeacher(token); }
function teacherLogout(token)        { return ClassCore.teacherLogout(token); }
function validateTeacherSession(token) { return ClassCore.verifyTeacher(token); }

// ── 활동 관리 (교사용) ──

// 활동 목록 (rowIdx 포함)
function getActivities() {
  try {
    var sheet = _changcheSs().getSheetByName('학급활동목록');
    if (!sheet || sheet.getLastRow() < 2) return { success: true, activities: [] };
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
    var activities = data.map(function(r, i) {
      return {
        rowIdx:   i + 2,
        category: String(r[1] || '').trim(),
        name:     String(r[2] || '').trim(),
        desc:     String(r[3] || '').trim(),
        formUrl:  String(r[4] || '').trim(),
        fields:   String(r[5] || '').trim()
      };
    });
    return { success: true, activities: activities.reverse() };
  } catch(e) { return { success: false, message: e.toString() }; }
}

// 활동 저장 (추가/수정 통합)
function saveActivity(data) {
  try {
    var sheet = _changcheSs().getSheetByName('학급활동목록');
    var row = [
      data.rowIdx || sheet.getLastRow(),
      data.category || '자율활동',
      data.name || '',
      data.desc || '',
      data.formUrl || '',
      data.fields || '[]'
    ];
    if (data.rowIdx && data.rowIdx > 1) {
      sheet.getRange(data.rowIdx, 1, 1, 6).setValues([row]);
    } else {
      sheet.appendRow([sheet.getLastRow(), row[1], row[2], row[3], row[4], row[5]]);
    }
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}

// 활동 삭제
function deleteActivity(rowIdx) {
  try {
    var sheet = _changcheSs().getSheetByName('학급활동목록');
    if (sheet && rowIdx > 1) { sheet.deleteRow(rowIdx); return { success: true }; }
    return { success: false, message: '삭제할 활동을 찾을 수 없습니다.' };
  } catch(e) { return { success: false, message: e.toString() }; }
}

// 제출 현황 (전체) + 활동 목록 (학생별 조회용)
function getSubmissions() {
  try {
    var ss = _changcheSs();

    var activities = [];
    var actSheet = ss.getSheetByName('학급활동목록');
    if (actSheet && actSheet.getLastRow() >= 2) {
      activities = actSheet.getRange(2, 1, actSheet.getLastRow() - 1, 6).getValues().map(function(r, i) {
        return { rowIdx: i + 2, category: r[1], name: r[2], desc: r[3] };
      });
    }

    var submissions = [];
    var subSheet = ss.getSheetByName('창체제출현황');
    if (subSheet && subSheet.getLastRow() >= 2) {
      submissions = subSheet.getRange(2, 1, subSheet.getLastRow() - 1, 8).getValues().map(function(r) {
        return {
          date:       r[0] ? Utilities.formatDate(new Date(r[0]), 'Asia/Seoul', 'MM/dd HH:mm') : '',
          id:         String(r[1] || '').trim(),
          name:       String(r[2] || '').trim(),
          activity:   String(r[3] || '').trim(),
          role:       String(r[4] || '').trim(),
          reflection: String(r[5] || '').trim(),
          url:        String(r[6] || '').trim()
        };
      });
    }

    return {
      success: true,
      activities: activities.reverse(),
      submissions: submissions.reverse()
    };
  } catch(e) { return { success: false, message: e.toString() }; }
}

// 학생 명부 (반별 필터용) — ClassCore 위임
function getRoster() {
  try {
    return { success: true, roster: ClassCore.getRosterValues(), homeroom: ClassCore.getConfig('담임반', '') };
  } catch(e) { return { success: false, message: e.toString() }; }
}


// ── PWA(Pages) 프론트 → GAS 백엔드 (google.script.run 어댑터) + 토큰 게이트 ──
var RPC_WHITELIST = ["deleteActivity", "getActivities", "getRoster", "getSubmissions", "saveActivity", "teacherLogin", "teacherLogout", "validateTeacherSession"];
var RPC_NOAUTH = ['teacherLogin','teacherLogout'];
function _atOk_(token) { try { var v = ClassCore.verifyTeacher(token); return (v === true) || !!(v && (v.success || v.valid || v.ok)); } catch(_) { return false; } }
function doPost(e) {
  var out;
  try {
    var req = JSON.parse(e.postData.contents);
    if (RPC_WHITELIST.indexOf(req.fn) < 0) throw new Error('허용되지 않은 함수: ' + req.fn);
    if (RPC_NOAUTH.indexOf(req.fn) < 0 && !_atOk_(req.token)) {
      return ContentService.createTextOutput(JSON.stringify({ ok: false, error: '로그인이 필요해요', needAuth: true })).setMimeType(ContentService.MimeType.JSON);
    }
    var fn = globalThis[req.fn];
    if (typeof fn !== 'function') throw new Error('함수를 찾을 수 없음: ' + req.fn);
    out = { ok: true, result: fn.apply(null, req.args || []) };
  } catch (err) { out = { ok: false, error: String(err && err.message || err) }; }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

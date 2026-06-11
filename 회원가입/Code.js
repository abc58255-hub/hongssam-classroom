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
  _registerAppUrl_('바로가기_회원가입');
  return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('수학 교실 회원가입')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// 비밀번호 해시
function getHash(text) {
  const rawBytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text);
  return rawBytes.map(function(e) {
    return (e < 0 ? e + 256 : e).toString(16).padStart(2, '0');
  }).join('');
}

// ✅ QR 이미지는 register_index.html에 직접 embed됨 — 서버 함수 불필요

// 회원가입 처리
function registerStudent(studentId, studentName, password, inputCode) {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  // 가입코드: 시스템설정 '가입코드' 키 우선, 없으면 학생명부 E2 폴백
  let currentSecretCode = '';
  const sysSh = ss.getSheetByName('시스템설정');
  if (sysSh && sysSh.getLastRow() >= 2) {
    const rows = sysSh.getRange(2, 1, sysSh.getLastRow() - 1, 2).getValues();
    for (let i = 0; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === '가입코드') { currentSecretCode = String(rows[i][1] || '').trim(); break; }
    }
  }

  if (!currentSecretCode || String(inputCode).trim() !== currentSecretCode) {
    return { success: false, message: "🚨 가입 코드가 틀렸거나, 현재 가입 기간이 아닙니다." };
  }

  const rosterSheet = ss.getSheetByName('학생명부');
  if (!rosterSheet) return { success: false, message: "학생 명부를 찾을 수 없습니다." };
  const data = rosterSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim() === String(studentId).trim() &&
        String(data[i][2]).trim() === String(studentName).trim()) {
      if (data[i][3]) {
        return { success: false, message: "⚠️ 이미 가입이 완료된 계정입니다. 비밀번호를 잊었다면 선생님께 초기화를 요청하세요." };
      }
      rosterSheet.getRange(i + 1, 4).setValue(getHash(password));
      return { success: true };
    }
  }
  return { success: false, message: "❌ 명부에 없는 학번/이름입니다. 이름을 정확히 띄어쓰기 없이 입력했는지 확인하세요." };
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

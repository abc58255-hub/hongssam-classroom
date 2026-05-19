const SHEET_ID = PropertiesService.getScriptProperties().getProperty('SHEET_ID')
  || '1jK7gYGFXCe3FULLs5mKttP959Aa9vp8-WNOGdJy7cZQ';

function doGet() {
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
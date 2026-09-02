function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('우리반 알림판')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// =====================================================
// ✅ 시트 구조 (공지사항 시트)
// A=날짜, B=조회공지, C=조회전달, D=종례공지, E=종례전달, F=슬라이드URL, G=유튜브URL, H=영상모드
// 13시 이전 → 조회(B,C) 표시 / 13시 이후 → 종례(D,E) 표시
// 종례 항목이 비어있으면 조회 내용 그대로 표시
// =====================================================
function getTodayData() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("공지사항");
    if (!sheet) return { error: "시트 탭 이름을 '공지사항'으로 확인해주세요." };

    var data = sheet.getDataRange().getDisplayValues();
    var todayStr = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd");
    var nowHour = parseInt(Utilities.formatDate(new Date(), "Asia/Seoul", "HH"));
    var isAfternoon = nowHour >= 13; // 13시 이후 = 종례

    for (var i = 1; i < data.length; i++) {
      var rowDateStr = data[i][0].replace(/\s/g, '');
      if (rowDateStr !== todayStr) continue;

      var announcementAm = data[i][1] || "등록된 공지사항이 없습니다.";
      var noticeAm       = data[i][2] || "등록된 전달사항이 없습니다.";
      var announcementPm = data[i][3] || ""; // 비어있으면 조회 내용 사용
      var noticePm       = data[i][4] || "";
      var slideUrl       = data[i][5] ? data[i][5].trim() : "";
      var videoUrl       = data[i][6] ? data[i][6].trim() : "";
      var videoMode      = data[i][7] ? data[i][7].trim() : "소리만";

      // iframe 태그 통째로 붙여넣은 경우 URL만 추출
      if (slideUrl.indexOf('<iframe') !== -1) {
        var match = slideUrl.match(/src="([^"]+)"/);
        if (match) slideUrl = match[1];
      }

      // 13시 기준으로 조회/종례 전환
      var announcement, notice;
      if (isAfternoon) {
        announcement = announcementPm || announcementAm; // 종례 비어있으면 조회 내용
        notice       = noticePm       || noticeAm;
      } else {
        announcement = announcementAm;
        notice       = noticeAm;
      }

      return {
        announcement: announcement,
        notice:       notice,
        slideUrl:     slideUrl,
        videoUrl:     videoUrl,
        videoMode:    videoMode,
        timeMode:     isAfternoon ? '종례' : '조회' // 디버그용
      };
    }

    return { error: "오늘 날짜(" + todayStr + ") 데이터가 없습니다." };
  } catch(e) {
    return { error: "시스템 오류: " + e.message };
  }
}

// ── 📺 TV 실시간 알림 — 같은 스프레드시트 '_TV_FLASH' 탭에서 읽음 (알림관리가 씀) ──
function getTvFlash() {
  try {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('_TV_FLASH');
    if (!sh) return { id: 0, msg: '' };
    var v = sh.getRange('A2:B2').getValues()[0];
    return { id: Number(v[0]) || 0, msg: String(v[1] || '') };
  } catch (_) { return { id: 0, msg: '' }; }
}

// ── PWA(Pages) 프론트 → GAS 백엔드 호출 (google.script.run 어댑터) ──
var RPC_WHITELIST = ["getTodayData", "getTvFlash"];
function doPost(e) {
  var out;
  try {
    var req = JSON.parse(e.postData.contents);
    if (RPC_WHITELIST.indexOf(req.fn) < 0) throw new Error('허용되지 않은 함수: ' + req.fn);
    var fn = globalThis[req.fn];
    if (typeof fn !== 'function') throw new Error('함수를 찾을 수 없음: ' + req.fn);
    out = { ok: true, result: fn.apply(null, req.args || []) };
  } catch (err) { out = { ok: false, error: String(err && err.message || err) }; }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

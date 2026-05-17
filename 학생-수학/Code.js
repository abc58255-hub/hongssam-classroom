const SHEET_ID = "1jK7gYGFXCe3FULLs5mKttP959Aa9vp8-WNOGdJy7cZQ"; 
const PARENT_FOLDER_ID = "1nmo4ZtQYK3-0PFjMKO8yzlkNOVoLn9_H";

// ✅ Firebase 서버 키 — Firebase 콘솔 → 프로젝트 설정 → 클라우드 메시징 → 서버 키
const FCM_SERVER_KEY = "여기에_FCM_SERVER_KEY_입력";

// =====================================================
// FCM 푸시 알림 발송 함수들
// =====================================================

// FCM 토큰 저장 (학생 앱에서 로그인 시 호출)
// pwHash 파라미터가 있으면 비밀번호 검증 후 저장
function saveFcmToken(studentId, token, pwHash) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName("학생명부");
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1] || "").trim() === String(studentId || "").trim()) {
        if (pwHash) {
          const storedHash = String(data[i][3] || "").trim(); // D열 = 비밀번호해시
          if (storedHash && storedHash !== pwHash) {
            return { success: false, message: "비밀번호가 올바르지 않습니다." };
          }
          if (!storedHash) {
            return { success: false, message: "미가입 학생입니다. 먼저 회원가입을 해주세요." };
          }
        }
        sheet.getRange(i + 1, 5).setValue(token); // E열 (FCM토큰)
        return { success: true };
      }
    }
    return { success: false, message: "학번을 찾을 수 없습니다." };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

// 특정 학생에게 푸시 알림 발송
function sendPushToStudent(studentId, title, body, tag) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName("학생명부");
    const data = sheet.getDataRange().getValues();
    let token = "";
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1] || "").trim() === String(studentId || "").trim()) {
        token = String(data[i][4] || "").trim(); // E열에서 FCM 토큰 읽기
        break;
      }
    }
    if (!token) return { success: false, message: "FCM 토큰 없음 (앱 미설치)" };
    return sendFcmMessage(token, title, body, tag);
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

// 실제 FCM API 호출
function sendFcmMessage(token, title, body, tag) {
  try {
    const payload = {
      to: token,
      notification: {
        title: title,
        body: body,
        icon: "icon-192x192.png",
        click_action: "FLUTTER_NOTIFICATION_CLICK"
      },
      data: {
        tag: tag || "default"
      }
    };

    const response = UrlFetchApp.fetch("https://fcm.googleapis.com/fcm/send", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Authorization": "key=" + FCM_SERVER_KEY
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    const result = JSON.parse(response.getContentText());
    if (result.success === 1) {
      return { success: true };
    } else {
      return { success: false, message: JSON.stringify(result) };
    }
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}


function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'saveFcmToken') {
    var result = saveFcmToken(e.parameter.studentId, e.parameter.token, e.parameter.pwHash || null);
    var output = JSON.stringify(result);
    if (e.parameter.callback) {
      return ContentService.createTextOutput(e.parameter.callback + '(' + output + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(output).setMimeType(ContentService.MimeType.JSON);
  }
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('수학 과제 제출기')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getHash(text) { 
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text)
    .map(e => (e < 0 ? e + 256 : e).toString(16).padStart(2, '0'))
    .join(''); 
}

// ✅ 로그인 성공 시 실패 횟수 초기화 + 빈 비밀번호 즉시 반환
function verifyLogin(studentId, studentName, password) { 
  // 비밀번호가 비어있으면 카운트 건드리지 않고 즉시 반환 (자동로그인 시도 방지)
  if (!password || password.trim() === "") {
    return { success: false, message: "" };
  }

  const cache = CacheService.getScriptCache(); 
  const lockKey = "lock_" + studentId; 
  if (cache.get(lockKey)) return { success: false, message: "🚨 5회 오류. 10분 정지됨." }; 

  const rosterData = SpreadsheetApp.openById(SHEET_ID).getSheetByName("학생명부").getDataRange().getValues(); 
  const inputHash = getHash(password); 

  for (let i = 1; i < rosterData.length; i++) { 
    if (String(rosterData[i][1] || "").trim() === String(studentId || "").trim() && 
        String(rosterData[i][2] || "").trim() === String(studentName || "").trim()) { 
      if (String(rosterData[i][3] || "").trim() === inputHash) { 
        // ✅ 로그인 성공 시 실패 횟수 + 잠금 모두 초기화
        cache.remove("fail_" + studentId); 
        cache.remove(lockKey);
        return { success: true }; 
      } else { 
        let fails = parseInt(cache.get("fail_" + studentId) || "0") + 1; 
        if (fails >= 5) { 
          cache.put(lockKey, "locked", 600); 
          cache.remove("fail_" + studentId); 
          return { success: false, message: "🚨 10분 정지됨" }; 
        } 
        cache.put("fail_" + studentId, fails.toString(), 600); 
        return { success: false, message: "비번 오류 (" + fails + "/5)" }; 
      } 
    } 
  } 
  return { success: false, message: "정보 확인 요망" }; 
}

function getSecureFileBase64(url) { 
  try { 
    return { 
      success: true, 
      mimeType: DriveApp.getFileById(url.match(/[-\w]{25,}/)[0]).getBlob().getContentType(), 
      data: Utilities.base64Encode(DriveApp.getFileById(url.match(/[-\w]{25,}/)[0]).getBlob().getBytes()) 
    }; 
  } catch(e) { 
    return { success: false, message: "파일 열기 실패" }; 
  } 
}

// =====================================================
// ✅ 학생용: getDashboardData 전체 덮어쓰기 (재제출 마감일 기능 포함)
// =====================================================
function getDashboardData(studentId, studentName) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID); 
    const now = new Date(); 
    let safeId = String(studentId || "").trim();
    let className = safeId.length >= 2 ? `${safeId.substring(0, 1)}학년 ${safeId.substring(1, 2)}반` : "기타";
    
    const taskData = ss.getSheetByName("과제설정").getDataRange().getValues();
    let allBaseTasks = []; 
    let validMissingTasksSet = new Set(); 
    let taskSettingsMap = {};
    let taskDeadlineMap = {};

    for (let i = 1; i < taskData.length; i++) {
      let tName = String(taskData[i][1] || "").trim(); 
      if (!tName) continue;
      allBaseTasks.push(tName); 
      let choices = String(taskData[i][7] || "").trim();
      let choiceArray = choices ? choices.split(',').map(s => s.trim()).filter(s => s) : [];
      taskSettingsMap[tName] = { reqPics: taskData[i][6] ? parseInt(taskData[i][6]) : 1, choiceArray: choiceArray };

      let dStr = String(taskData[i][3] || "").trim(); 
      let isExpired = false; 
      let hasDeadline = false;
      let myDeadline = null;
      let resubDeadline = null;
      let openDeadline = null;

      if (dStr && dStr.startsWith("{")) {
        try {
          let deadlines = JSON.parse(dStr);
          let dl = deadlines[className] || deadlines["all"];
          const hasClassSpecificKeys = Object.keys(deadlines).some(k => k !== "all" && !k.startsWith("resub_"));
          if (dl) {
            hasDeadline = true;
            myDeadline = dl;
            if (new Date(dl) < now) isExpired = true;
          } else if (hasClassSpecificKeys) {
            // 다른 반에만 마감일이 설정된 경우 → 이 학생에겐 과제 숨김
            hasDeadline = true;
            isExpired = true;
          }
          resubDeadline = deadlines["resub_" + className] || deadlines["resub_all"] || myDeadline;
          openDeadline = deadlines["open_" + className] || deadlines["open_all"] || null;
        } catch(e) {}
      }

      taskDeadlineMap[tName] = {
        main: myDeadline,
        resub: resubDeadline,
        open: openDeadline
      };

      if (!hasDeadline || !isExpired) { validMissingTasksSet.add(tName); }
    }

    const historyData = ss.getSheetByName("제출현황").getDataRange().getValues();
    let history = []; 
    let taskStatusMap = {}; 
    let unreadFeedbacks = []; 
    let bestWorksMap = {};
    let rankCounters = {};

    for (let i = 1; i < historyData.length; i++) {
      let rowId = String(historyData[i][1] || "").trim();
      if (!rowId) continue;

      let rawTaskName = String(historyData[i][3] || ""); 
      let baseName = rawTaskName.split(' (')[0];
      let status = historyData[i][10] ? String(historyData[i][10]).trim() : ""; 
      let rowClass = rowId.length >= 2 ? `${rowId.substring(0,1)}학년 ${rowId.substring(1,2)}반` : "기타";

      let myTotalRank = 0;
      let myClassRank = 0;
      if (status !== '재제출요청' && status !== '반려검토' && status !== '이전기록채점완료') {
        if (!rankCounters[baseName]) rankCounters[baseName] = { total: 0, classes: {} };
        if (!rankCounters[baseName].classes[rowClass]) rankCounters[baseName].classes[rowClass] = 0;
        
        rankCounters[baseName].total++;
        rankCounters[baseName].classes[rowClass]++;
        
        myTotalRank = rankCounters[baseName].total;
        myClassRank = rankCounters[baseName].classes[rowClass];
      }

      let rawUrl = String(historyData[i][6] || "").trim(); let urls = {}; 
      if (rawUrl.startsWith('{')) { try { urls = JSON.parse(rawUrl); } catch(e) {} } 
      else if (rawUrl.startsWith('[')) { try { let arr=JSON.parse(rawUrl); arr.forEach((v,idx)=>urls["사진 "+(idx+1)]=v); } catch(e){} }
      else if (rawUrl) { urls["사진 1"] = rawUrl; }

      let rawAnno = String(historyData[i][11] || "").trim(); let annoUrls = {}; 
      if (rawAnno.startsWith('{')) { try { annoUrls = JSON.parse(rawAnno); } catch(e) {} }
      else if (rawAnno.startsWith('[')) { try { let arr=JSON.parse(rawAnno); arr.forEach((v,idx)=>annoUrls["사진 "+(idx+1)]=v); } catch(e){} }
      else if (rawAnno) { annoUrls["사진 1"] = rawAnno; }
      
      let bestKey = historyData[i][19] ? String(historyData[i][19]).trim() : "";
      let bestType = historyData[i][16] ? String(historyData[i][16]).trim() : "";
      
      if (bestType !== "") {
        let isAnon = historyData[i][17] === true || String(historyData[i][17]).toUpperCase() === "TRUE"; 
        let authorCls = rowId.length >= 2 ? `${rowId.substring(0,1)}학년 ${rowId.substring(1,2)}반` : "기타";
        
        if (bestType === "학년공개" || (bestType === "학급공개" && authorCls === className)) {
          if (!bestWorksMap[baseName]) bestWorksMap[baseName] = [];
          bestWorksMap[baseName].push({ 
            rowIdx: i + 1, id: rowId, name: String(historyData[i][2] || ""), 
            urls: urls, annoUrls: annoUrls, bestKey: bestKey, 
            comment: historyData[i][18] ? String(historyData[i][18]) : "", 
            isAnon: isAnon,
            bestType: bestType
          });
        }
      }

      if (rowId === safeId && String(historyData[i][2] || "").trim() === String(studentName || "").trim()) {
        let fb = historyData[i][7] ? String(historyData[i][7]).trim() : ""; 
        let isSeen = historyData[i][9] ? String(historyData[i][9]).trim() : ""; 
        let isPublic = String(historyData[i][13]).trim() === "공개" || String(historyData[i][13]).trim() === "일괄공개"; 
        let score = isPublic && historyData[i][12] ? String(historyData[i][12]).trim() : ""; 
        let reply = historyData[i][15] ? String(historyData[i][15]).trim() : ""; 
        let isMyBest = (historyData[i][16] ? String(historyData[i][16]).trim() : "") !== "";
        let ts = taskSettingsMap[baseName] || {reqPics:1, choiceArray:[]};

        let perQuestionData = {};
        try {
          let pqRaw = historyData[i][21] ? String(historyData[i][21] || "").trim() : "";
          if (pqRaw.startsWith('{')) {
            let parsed = JSON.parse(pqRaw);
            Object.keys(parsed).forEach(k => {
              let d = parsed[k];
              let qPub = d.pub !== undefined ? d.pub : isPublic;
              if (qPub || isPublic) {
                perQuestionData[k] = { fb: d.fb || '', sc: d.sc || '' };
                if (d.best) perQuestionData[k].best = d.best;
              }
            });
          }
        } catch(e) {}

        let aiGradeTemp = null;
        try {
          const aiRaw = String(historyData[i][22] || "").trim();
          if (aiRaw.startsWith('{')) aiGradeTemp = JSON.parse(aiRaw);
        } catch(e) {}

        let item = {
          rowIdx: i + 1,
          date: historyData[i][0] ? Utilities.formatDate(new Date(historyData[i][0]), "Asia/Seoul", "MM/dd HH:mm") : "",
          task: rawTaskName, baseName: baseName, level: historyData[i][4] || "",
          urls: urls, feedback: fb, status: status, annoUrls: annoUrls,
          score: score, reply: reply, isMyBest: isMyBest,
          reqPics: ts.reqPics, choices: ts.choiceArray,
          perQuestionData: perQuestionData,
          isUnread: (fb !== "" || isMyBest) && isSeen === "",
          aiGradeTemp: aiGradeTemp,
          totalRank: myTotalRank,
          classRank: myClassRank,
          deadline: taskDeadlineMap[baseName] ? taskDeadlineMap[baseName].main : null
        };
        history.push(item); 
        if (status !== "이전기록채점완료") {
          taskStatusMap[baseName] = item;
        } else if (!taskStatusMap[baseName]) {
          taskStatusMap[baseName] = item; 
        }
        if ((fb !== "" && isSeen === "") || (isMyBest && isSeen === "")) unreadFeedbacks.push(item);
      }
    }
    
    let missingTasks = []; 
    let resubmitTasks = [];

    allBaseTasks.forEach(t => {
      let ts = taskSettingsMap[t] || {reqPics:1, choiceArray:[]};
      let dMap = taskDeadlineMap[t] || {};

      if (!taskStatusMap[t]) {
        if (validMissingTasksSet.has(t)) {
          missingTasks.push({
            name: t, reqPics: ts.reqPics, choices: ts.choiceArray,
            submittedUrls: {}, deadline: dMap.main, openDate: dMap.open
          });
        }
      } else {
        let st = taskStatusMap[t].status;
        let submittedCount = 0; 
        let currentUrls = taskStatusMap[t].urls;
        for (let k in currentUrls) { 
          if (currentUrls[k] && currentUrls[k] !== "" && currentUrls[k] !== "첨부파일 없음") submittedCount++; 
        }

        if (st === "재제출요청") {
          let resubDl = dMap.resub || dMap.main;
          let resubExpired = resubDl && new Date(resubDl) < now;
          if (!resubExpired) {
            let rejectionFeedback = taskStatusMap[t].feedback || "";
            let rejectionPqData = taskStatusMap[t].perQuestionData || {};
            let completedKeys = Object.keys(rejectionPqData).filter(k =>
              rejectionPqData[k] && rejectionPqData[k].status === "완료"
            );
            resubmitTasks.push({
              name: t, reqPics: ts.reqPics, choices: ts.choiceArray,
              submittedUrls: {}, isResubmit: true,
              deadline: dMap.resub,
              rejectionFeedback: rejectionFeedback, completedKeys: completedKeys
            });
          }
        } else if (submittedCount < ts.reqPics && st === "" && validMissingTasksSet.has(t)) {
          missingTasks.push({
            name: t, reqPics: ts.reqPics, choices: ts.choiceArray,
            submittedUrls: currentUrls, deadline: dMap.main, openDate: dMap.open
          });
        }
      }
    });

    return {
      history: history.reverse(),
      missingTasks: missingTasks,
      resubmitTasks: resubmitTasks,
      unreadFeedbacks: unreadFeedbacks,
      bestWorksMap: bestWorksMap,
      fcmRegisterUrl: _getSysStudent('FCM_REGISTER_URL')
    };

  } catch(e) {
    throw new Error("데이터 로딩 중 오류: " + e.toString());
  }
}

function markFeedbacksAsSeen(rowIndices) { 
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName("제출현황"); 
  rowIndices.forEach(idx => sheet.getRange(idx, 10).setValue("확인")); 
  return true; 
}

function saveStudentReply(rowIdx, replyText) {
  try {
    SpreadsheetApp.openById(SHEET_ID).getSheetByName("제출현황").getRange(rowIdx, 16).setValue(replyText);
    return true;
  } catch(e) {
    return false;
  }
}

function requestResubmission(rowIdx) {
  try {
    SpreadsheetApp.openById(SHEET_ID).getSheetByName("제출현황").getRange(rowIdx, 16).setValue("[재제출요청]");
    return { success: true };
  } catch(e) {
    return { success: false };
  }
}

function processForm(formData) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName("제출현황"); 
    const now = new Date();
    let inputId = String(formData.studentId || "").trim(); 
    let inputName = String(formData.studentName || "").trim();
    let isResubmit = formData.isResubmit === "true"; 
    let baseTaskName = formData.taskName;
    
    let incomingFiles = formData.filesData; 
    let fileHashObj = {}; 
    const records = sheet.getDataRange().getValues();
    incomingFiles.forEach(f => { fileHashObj[f.key] = getHash(Utilities.base64Decode(f.b64)); });

    // ✅ 마감일 서버측 검증
    const taskSheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName("과제설정");
    if (taskSheet) {
      const taskRows = taskSheet.getDataRange().getValues();
      const className = inputId.length >= 2 ? `${inputId.substring(0,1)}학년 ${inputId.substring(1,2)}반` : "기타";
      for (let i = 1; i < taskRows.length; i++) {
        if (String(taskRows[i][1] || "").trim() === baseTaskName) {
          const dStr = String(taskRows[i][3] || "").trim();
          if (dStr && dStr.startsWith("{")) {
            try {
              const dl = JSON.parse(dStr);
              if (isResubmit) {
                const resubDl = dl["resub_" + className] || dl["resub_all"] || dl[className] || dl["all"];
                if (resubDl && new Date(resubDl) < now) {
                  return { success: false, message: "⏰ 재제출 기한이 지났습니다. 선생님께 문의하세요." };
                }
              } else {
                const mainDl = dl[className] || dl["all"];
                if (mainDl && new Date(mainDl) < now) {
                  return { success: false, message: "⏰ 제출 기한이 지났습니다." };
                }
                // 다른 반에만 마감일이 설정된 경우 이 반은 제출 불가
                const hasClassSpecificKeys = Object.keys(dl).some(k => k !== "all" && !k.startsWith("resub_") && !k.startsWith("open_"));
                if (!mainDl && hasClassSpecificKeys) {
                  return { success: false, message: "🚫 이 과제는 해당 반에 제공되지 않습니다." };
                }
                // 공개일 이전 제출 불가
                const openDl = dl["open_" + className] || dl["open_all"];
                if (openDl && new Date(openDl) > now) {
                  return { success: false, message: "⏳ 아직 제출 가능 시간이 아닙니다." };
                }
              }
            } catch(e) {}
          }
          break;
        }
      }
    }

    // ✅ 채점 완료 또는 채점중인 과제는 제출 불가 (중복 방지)
    for (let i = records.length - 1; i >= 1; i--) {
      if (String(records[i][1] || "").trim() === inputId && 
          String(records[i][3] || "").startsWith(baseTaskName)) {
        let currentStatus = String(records[i][10] || "").trim();
        if (currentStatus === "채점완료") {
          return { success: false, message: "🚨 채점이 끝난 과제입니다." };
        }
        // 이전기록채점완료는 건너뜀 (재제출 가능)
        if (currentStatus !== "이전기록채점완료" && currentStatus !== "재제출완료") break;
      }
    }

    // ✅ 재제출인 경우 기존 재제출요청 행을 재제출완료로 변경
    if (isResubmit) {
      for (let i = records.length - 1; i >= 1; i--) {
        if (String(records[i][1] || "").trim() === inputId && 
            String(records[i][3] || "").startsWith(baseTaskName) && 
            String(records[i][10] || "").trim() === "재제출요청") {
          sheet.getRange(i + 1, 11).setValue("재제출완료"); 
          break;
        }
      }
    }

    let finalTaskName = isResubmit ? `${baseTaskName} (재제출)` : baseTaskName; 
    let className = inputId.length >= 2 ? `${inputId.substring(0, 1)}학년 ${inputId.substring(1, 2)}반` : "기타"; 
    
    const parentFolder = DriveApp.getFolderById(PARENT_FOLDER_ID);
    let taskFolder = parentFolder.getFoldersByName(baseTaskName).hasNext() 
      ? parentFolder.getFoldersByName(baseTaskName).next() 
      : parentFolder.createFolder(baseTaskName);
    let classFolder = taskFolder.getFoldersByName(className).hasNext() 
      ? taskFolder.getFoldersByName(className).next() 
      : taskFolder.createFolder(className);

    // ✅ 임시저장 없이 항상 새 행으로 append
    let finalUrls = {};
    incomingFiles.forEach(f => {
      let suffix = isResubmit ? `_재제출_${f.key}` : `_${f.key}`;
      const blob = Utilities.newBlob(
        Utilities.base64Decode(f.b64), 'image/jpeg', 
        `[${inputId}] ${inputName}_${baseTaskName}${suffix}.jpg`
      );
      finalUrls[f.key] = classFolder.createFile(blob).getUrl(); 
    });

    sheet.appendRow([
      now, inputId, inputName, finalTaskName, formData.level, formData.message, 
      JSON.stringify(finalUrls), "", JSON.stringify(fileHashObj), "", "", "", "", "", "", ""
    ]);
    
    // ✅ 방금 추가한 행 번호를 반환 (AI 자동 채점 호출용)
    const newRowIdx = sheet.getLastRow();
    return { success: true, rowIdx: newRowIdx };
  } catch (error) { 
    return { success: false, message: error.toString() }; 
  }
}


// =====================================================
// ✅ 제출 순위 계산 (같은 과제에서 몇 번째 제출인지) - 반려/재제출 반영 버전
// =====================================================
// 특정 학생의 특정 과제 제출 등수 계산
function getSubmitRank(studentId, taskName) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName("제출현황");
    const data = sheet.getDataRange().getValues();
    const baseTaskName = String(taskName).split(' (')[0].trim(); // (재제출) 등 제외한 기본 과제명
    
    let safeId = String(studentId).trim();
    let myClass = safeId.substring(0, 1) + '학년 ' + safeId.substring(1, 2) + '반';

    let totalRank = 0;
    let classRank = 0;
    let found = false;

    // 1행(헤더) 제외하고 전수 조사
    for (let i = 1; i < data.length; i++) {
      let rowTask = String(data[i][3] || '').split(' (')[0].trim();
      let rowStatus = String(data[i][10] || '').trim();
      let rowId = String(data[i][1] || '').trim();
      let rowClass = rowId.substring(0, 1) + '학년 ' + rowId.substring(1, 2) + '반';

      if (rowTask === baseTaskName) {
        // 반려되거나 이전 기록인 경우는 등수에서 제외 (선택 사항)
        if (rowStatus === '재제출요청' || rowStatus === '이전기록채점완료') continue;

        totalRank++;
        if (rowClass === myClass) classRank++;

        // 현재 학생의 행을 찾으면 카운트 중단 (해당 시점까지의 제출 순서가 등수)
        if (rowId === safeId) {
          found = true;
          break;
        }
      }
    }
    
    return { success: true, totalRank: totalRank, classRank: classRank };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

// =====================================================
// ✅ 비밀번호 설정 (미설정 학생이 최초 로그인 시)
// =====================================================
function setStudentPassword(studentId, studentName, newPw) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName("학생명부");
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1] || '').trim() === String(studentId).trim() &&
          String(data[i][2] || '').trim() === String(studentName).trim()) {
        sheet.getRange(i + 1, 4).setValue(getHash(newPw));
        return { success: true };
      }
    }
    return { success: false, message: '학생 정보를 찾을 수 없습니다.' };
  } catch(e) { return { success: false, message: e.toString() }; }
}

// verifyLogin 수정: 비밀번호 미설정 학생 감지
function checkNeedsPwSetup(studentId, studentName) {
  try {
    const data = SpreadsheetApp.openById(SHEET_ID).getSheetByName("학생명부").getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1] || '').trim() === String(studentId).trim() &&
          String(data[i][2] || '').trim() === String(studentName).trim()) {
        return { found: true, needsSetup: !String(data[i][3] || '').trim() };
      }
    }
    return { found: false, needsSetup: false };
  } catch(e) { return { found: false, needsSetup: false }; }
}

// =====================================================
// ✅ AI 자동 채점 (학생 제출 직후 자동 실행)
// =====================================================

function _getSysStudent(key) {
  var sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName('시스템설정');
  if (!sh || sh.getLastRow() < 2) return '';
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === key) return String(rows[i][1] || '').trim();
  }
  return '';
}

function _getApiSettingsForStudent() {
  var orKey = _getSysStudent('OpenRouter키');
  var model = _getSysStudent('AI모델명');
  // 구 구조 폴백 (P2:Q2)
  if (!orKey) {
    try {
      var sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName('시스템설정');
      if (sh) {
        var old = sh.getRange('P2:Q2').getValues()[0];
        if (!orKey) orKey = String(old[0] || '').trim();
        if (!model) model = String(old[1] || '').trim();
      }
    } catch(e) {}
  }
  if (model === 'google/gemini-2.5-flash-preview') model = 'google/gemini-2.5-flash';
  if (model === 'google/gemini-2.5-pro-preview')   model = 'google/gemini-2.5-pro';
  return {
    openrouterKey: orKey,
    model: model || 'google/gemini-2.5-flash'
  };
}

function _getRubricByTaskName(taskName) {
  try {
    const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName('AI채점기준');
    if (!sh || sh.getLastRow() < 2) return null;
    const data = sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues();
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0] || '').trim() === taskName) {
        let files = [], questions = {};
        try { if (data[i][4]) files = JSON.parse(data[i][4]); } catch(e) {}
        try { if (data[i][5]) questions = JSON.parse(data[i][5]); } catch(e) {}
        return {
          taskName: String(data[i][0]).trim(),
          evalType: String(data[i][1] || '점수제').trim(),
          maxScore: Number(data[i][2] || 0),
          criteria: String(data[i][3] || '').trim(),
          files: files,
          questions: questions
        };
      }
    }
    return null;
  } catch(e) { return null; }
}

function _getRubricFileBase64Student(url) {
  try {
    if (!url || url.indexOf('drive.google.com') < 0) return { success: false };
    const fileId = url.match(/[-\w]{25,}/);
    if (!fileId) return { success: false };
    const f = DriveApp.getFileById(fileId[0]);
    if (f.getSize() > 10485760) return { success: false };
    return {
      success: true,
      mimeType: f.getMimeType(),
      data: Utilities.base64Encode(f.getBlob().getBytes())
    };
  } catch(e) { return { success: false }; }
}

// 학생 제출 직후 클라이언트가 호출하는 자동 채점 함수
function autoGradeNewSubmission(rowIdx, taskName, studentId, studentName) {
  try {
    const baseTask = String(taskName).split(' (')[0];
    const rubric = _getRubricByTaskName(baseTask);
    if (!rubric) return { success: false, message: '채점기준 없음' };

    const cfg = _getApiSettingsForStudent();
    if (!cfg.openrouterKey) return { success: false, message: 'API 키 없음' };

    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('제출현황');

    const existing = String(sheet.getRange(rowIdx, 23).getValue() || '').trim();
    if (existing.startsWith('{')) return { success: false, message: '이미 채점됨' };

    const rawUrl = String(sheet.getRange(rowIdx, 7).getValue() || '').trim();
    let urls = {};
    if (rawUrl.startsWith('{')) { try { urls = JSON.parse(rawUrl); } catch(e) {} }
    else if (rawUrl) urls['사진 1'] = rawUrl;

    const urlKeys = Object.keys(urls).filter(k => urls[k] && urls[k] !== "첨부파일 없음");
    if (urlKeys.length === 0) return { success: false, message: '사진 없음' };
    const isType2 = urlKeys.length > 1;

    const prompt = [
      '너는 수학 교사의 채점을 돕는 AI야. 반드시 JSON만 반환해.',
      '학생: ' + studentId + ' ' + studentName,
      '과제: ' + baseTask,
      '채점유형: ' + rubric.evalType,
      '총점: ' + rubric.maxScore + '점',
      rubric.criteria ? '채점기준: ' + rubric.criteria : '',
      (rubric.files || []).some(f => f.role === 'achievement') ? '※ 성취기준 파일을 참고하여 A/B/C/D 판별할 것.' : '',
      '반환 형식:',
      (function() {
        const et = rubric.evalType;
        if (et.indexOf('A-B-C') >= 0 || et.indexOf('등급') >= 0)
          return '{"grade":"A|B|C|D","feedback":"피드백","confidence":"high|medium|low","needsReview":true|false}';
        if (et.indexOf('상-중-하') >= 0)
          return '{"grade":"상|중|하","feedback":"피드백","confidence":"high|medium|low","needsReview":true|false}';
        if (et.indexOf('P/F') >= 0 || et.indexOf('통과') >= 0)
          return '{"grade":"Pass|Fail","feedback":"피드백","confidence":"high|medium|low","needsReview":true|false}';
        if (isType2)
          return '{"perQuestion":{"문항명":{"score":점수,"maxScore":만점,"feedback":"피드백"}},"totalScore":합계,"overallFeedback":"종합피드백","confidence":"high|medium|low","needsReview":true|false}';
        return '{"score":점수,"maxScore":' + rubric.maxScore + ',"feedback":"피드백","confidence":"high|medium|low","needsReview":true|false}';
      })()
    ].filter(Boolean).join('\n');

    const content = [{ type: 'text', text: prompt }];

    (rubric.files || []).forEach(rf => {
      if (!rf.url) return;
      const roleLabel = rf.role === 'achievement' ? '[성취기준]' : '[채점기준]';
      content.push({ type: 'text', text: roleLabel });
      const rFile = _getRubricFileBase64Student(rf.url);
      if (rFile.success) {
        content.push({ type: 'image_url', image_url: { url: 'data:' + rFile.mimeType + ';base64,' + rFile.data } });
      }
    });

    urlKeys.forEach(k => {
      content.push({ type: 'text', text: isType2 ? '[' + k + ' 답안]' : '[학생 답안]' });
      const imgRes = getSecureFileBase64(urls[k]);
      if (imgRes.success) {
        content.push({ type: 'image_url', image_url: { url: 'data:' + imgRes.mimeType + ';base64,' + imgRes.data } });
      }
    });

    const res = UrlFetchApp.fetch('https://openrouter.ai/api/v1/chat/completions', {
      method: 'post',
      headers: {
        'Authorization': 'Bearer ' + cfg.openrouterKey,
        'Content-Type': 'application/json',
        'HTTP-Referer': 'https://script.google.com',
        'X-Title': 'Student Auto Grade'
      },
      payload: JSON.stringify({
        model: cfg.model,
        messages: [{ role: 'user', content: content }],
        max_tokens: 1000,
        temperature: 0.2
      }),
      muteHttpExceptions: true
    });

    if (res.getResponseCode() !== 200) {
      return { success: false, message: 'API 오류 ' + res.getResponseCode() };
    }

    let text = JSON.parse(res.getContentText()).choices[0].message.content.trim();
    text = text.replace(/^```[a-z]*\n?/i, '').replace(/\n?```$/, '').trim();
    const result = JSON.parse(text);

    sheet.getRange(rowIdx, 23).setValue(JSON.stringify(result));

    // ✅ [추가] 실시간 등수 확인 및 칭찬 멘트 생성
    const rankData = getSubmitRank(studentId, baseTask);
    if (rankData.success) {
      result.classRank = rankData.classRank;
      result.totalRank = rankData.rank;

      let praiseMsg = "";
      if (rankData.classRank === 1) praiseMsg = "🎉 와우! 우리 반에서 가장 먼저 제출했어요! 1등! 🥇 훌륭합니다!";
      else if (rankData.classRank === 2) praiseMsg = "👏 엄청난 스피드! 우리 반 2등으로 제출 완료! 🥈 멋져요!";
      else if (rankData.classRank === 3) praiseMsg = "👍 빠른 제출 칭찬해요! 우리 반 3등 제출입니다! 🥉 최고!";
      else if (rankData.classRank <= 5) praiseMsg = "🏃‍♂️ 엄청 빨라요! 우리 반 TOP 5 안에 들었네요! 🏅 잘했어요!";
      else if (rankData.classRank <= 10) praiseMsg = "🏃‍♀️ 부지런하네요! 우리 반 TOP 10 안에 들었어요! 🎖️";
      
      if (rankData.rank === 1) praiseMsg = "🏆 세상에! 전체 학년에서 당당히 1등으로 제출했어요! 엄청난 열정 최고예요! 🔥";
      result.rankMessage = praiseMsg; 
    }

    return { success: true, result: result };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

// =====================================================
// 🧪 AI 자동 채점 디버그 테스트 (직접 실행해보기)
// =====================================================
function testAutoGrade() {
  // 1. 설정 확인
  const cfg = _getApiSettingsForStudent();
  Logger.log('📌 API 키 앞 10자: ' + cfg.openrouterKey.substring(0, 10));
  Logger.log('📌 모델: ' + cfg.model);
  
  if (!cfg.openrouterKey) {
    Logger.log('❌ OpenRouter API 키가 없습니다! 시스템설정 OpenRouter키 항목 확인!');
    return;
  }
  
  // 2. 채점기준 시트 확인
  const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName('AI채점기준');
  if (!sh) {
    Logger.log('❌ "AI채점기준" 시트가 없습니다!');
    return;
  }
  const rubricCount = sh.getLastRow() - 1;
  Logger.log('📌 등록된 채점기준 개수: ' + rubricCount);
  
  if (rubricCount < 1) {
    Logger.log('❌ 등록된 채점기준이 없습니다. 교사 대시보드에서 먼저 등록하세요!');
    return;
  }
  
  // 3. 첫 번째 채점기준 정보 출력
  const firstRubric = sh.getRange(2, 1, 1, 6).getValues()[0];
  Logger.log('📌 첫 번째 채점기준 과제명: ' + firstRubric[0]);
  Logger.log('📌 채점유형: ' + firstRubric[1]);
  Logger.log('📌 파일 개수: ' + (firstRubric[4] ? JSON.parse(firstRubric[4]).length : 0));
  
  // 4. 최근 제출된 행을 기준으로 실제 채점 시도
  const sub = SpreadsheetApp.openById(SHEET_ID).getSheetByName('제출현황');
  const lastRow = sub.getLastRow();
  Logger.log('📌 제출현황 마지막 행: ' + lastRow);
  
  if (lastRow < 2) {
    Logger.log('❌ 제출된 과제가 하나도 없습니다.');
    return;
  }
  
  const rowData = sub.getRange(lastRow, 1, 1, 8).getValues()[0];
  const sid = String(rowData[1] || '').trim();
  const sname = String(rowData[2] || '').trim();
  const tname = String(rowData[3] || '').trim();
  
  Logger.log('📌 마지막 제출 학생: ' + sid + ' ' + sname);
  Logger.log('📌 과제명: ' + tname);
  
  // 5. 해당 과제의 채점기준 찾기
  const baseTask = tname.split(' (')[0];
  const rubric = _getRubricByTaskName(baseTask);
  if (!rubric) {
    Logger.log('❌ 해당 과제(' + baseTask + ')의 채점기준이 없습니다!');
    Logger.log('👉 교사 대시보드 → AI 채점기준 관리에서 "' + baseTask + '" 과제 추가하세요.');
    return;
  }
  Logger.log('✅ 채점기준 발견!');
  
  // 6. 실제 채점 실행
  Logger.log('🤖 AI 채점 시작...');
  const result = autoGradeNewSubmission(lastRow, tname, sid, sname);
  Logger.log('📌 결과: ' + JSON.stringify(result));
}

// 🔐 외부 API 권한 승인용 (한 번만 실행)
function grantPermissions() {
  // 이 함수 실행 시 구글이 권한 요청 창을 띄워줘요
  UrlFetchApp.fetch('https://www.google.com');
  Logger.log('✅ 권한 승인 완료! 이제 AI 자동 채점이 작동합니다.');
}
function sendPushToUnsubmitted(taskName, title, body) {
  try {
    var ss = _taskSs();
    var roster  = _authRoster_().getDataRange().getValues();
    var history = ss.getSheetByName('제출현황').getDataRange().getValues();
    var tasks   = ss.getSheetByName('과제설정').getDataRange().getValues();

    // 해당 과제의 마감일 설정 로드
    var deadlines = {};
    for (var k = 1; k < tasks.length; k++) {
      if (String(tasks[k][1] || '').trim() === taskName) {
        try { deadlines = JSON.parse(tasks[k][3] || '{}'); } catch(_) {}
        break;
      }
    }

    var submittedIds = new Set();
    for (var i = 1; i < history.length; i++) {
      var tn = String(history[i][3] || '').split(' (')[0];
      if (tn === taskName) submittedIds.add(String(history[i][1] || '').trim());
    }

    var unsubIds = [];
    for (var j = 1; j < roster.length; j++) {
      var sid = String(roster[j][1] || '').trim();
      if (!sid) continue;
      // 이 학생 반에 마감일이 설정된 경우에만 대상
      var cls = sid.length >= 2 ? (sid.substring(0,1) + '학년 ' + sid.substring(1,2) + '반') : '';
      var dl = deadlines[cls] || deadlines['all'];
      if (!dl) continue;
      if (!submittedIds.has(sid)) unsubIds.push(sid);
    }
    if (unsubIds.length === 0) return { success: true, sent: 0, skipped: 0, message: '미제출자 없음' };
    return sendPushToStudents(unsubIds, title || ('📢 [' + taskName + '] 아직 제출 안 했어요!'), body || '지금 바로 제출해주세요.', 'unsub');
  } catch(e) { return { success: false, message: e.toString() }; }
}
function _logNotify_(type, content, sent) {
  try {
    var ss = _taskSs();
    var sh = ss.getSheetByName('알림발송로그');
    if (!sh) {
      sh = ss.insertSheet('알림발송로그');
      sh.getRange(1, 1, 1, 4).setValues([['발송시각','유형','내용','발송수']]);
      sh.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#1e3a8a').setFontColor('white');
      sh.setFrozenRows(1);
    }
    sh.appendRow([new Date(), type, content, sent]);
    // 최대 500행 유지
    if (sh.getLastRow() > 501) sh.deleteRow(2);
  } catch(e) { Logger.log('_logNotify_ 오류: ' + e.message); }
}
function getNotifyLog() {
  try {
    var sh = _taskSs().getSheetByName('알림발송로그');
    if (!sh || sh.getLastRow() < 2) return { success: true, logs: [] };
    var lastRow = sh.getLastRow();
    var startRow = Math.max(2, lastRow - 29);
    var rows = sh.getRange(startRow, 1, lastRow - startRow + 1, 4).getValues();
    var logs = rows.reverse().map(function(r) {
      return {
        time: r[0] ? Utilities.formatDate(new Date(r[0]), 'Asia/Seoul', 'MM/dd HH:mm') : '',
        type: String(r[1] || ''),
        content: String(r[2] || ''),
        sent: Number(r[3] || 0)
      };
    });
    return { success: true, logs: logs };
  } catch(e) { return { success: false, logs: [] }; }
}
function _setSubmissionStatus_(sheet, rowIdx, status) {
  sheet.getRange(rowIdx, 11).setValue(status);
  sheet.getRange(rowIdx, 24).setValue(new Date());
}
function _formatDeadline_(dl) {
  try { return Utilities.formatDate(new Date(dl), 'Asia/Seoul', 'MM/dd HH:mm'); }
  catch(_) { return String(dl); }
}
function _dDayText_(deadline, now) {
  var diffMs = deadline - now;
  if (diffMs < 0) return '마감 지남';
  var diffH = Math.floor(diffMs / 3600000);
  if (diffH < 24) {
    var hh = Utilities.formatDate(deadline, 'Asia/Seoul', 'HH:mm');
    return '오늘 ' + hh + ' 마감';
  }
  var diffD = Math.floor(diffH / 24);
  return 'D-' + diffD;
}
function _notifyNewTask_(taskName, deadlinesJson) {
  try {
    var deadlines;
    try { deadlines = JSON.parse(deadlinesJson || '{}'); } catch(_) { return 0; }
    var hasAny = Object.keys(deadlines).some(function(k){
      return !k.startsWith('resub_') && !k.startsWith('open_') && deadlines[k];
    });
    if (!hasAny) return 0;

    var ss = _taskSs();
    var sheet = _authRoster_();
    var roster = sheet.getDataRange().getValues();
    var clickUrl = ClassCore.getConfig('바로가기_수학교실') || '';
    var sent = 0;

    for (var i = 1; i < roster.length; i++) {
      var sid = String(roster[i][1] || '').trim();
      if (!sid) continue;
      var cls = sid.length >= 2 ? (sid.substring(0,1) + '학년 ' + sid.substring(1,2) + '반') : '';
      var dl = deadlines[cls] || deadlines['all'];
      if (!dl) continue;
      var tokens = _parseTokens_(roster[i][4]);
      if (!tokens.length) continue;
      var title = '📝 새 과제: ' + taskName;
      var body  = '마감: ' + _formatDeadline_(dl);
      if (_sendAndPrune_(sheet, i + 1, title, body, clickUrl, 'newTask')) sent++;
    }
    return sent;
  } catch(e) {
    Logger.log('_notifyNewTask_ 오류: ' + e.message);
    return 0;
  }
}
function _notifyTaskDeadlineChange_(taskName, oldJson, newJson) {
  try {
    var oldDl, newDl;
    try { oldDl = JSON.parse(oldJson || '{}'); } catch(_) { oldDl = {}; }
    try { newDl = JSON.parse(newJson || '{}'); } catch(_) { return 0; }

    var ss = _taskSs();
    var sheet = _authRoster_();
    var roster = sheet.getDataRange().getValues();
    var clickUrl = ClassCore.getConfig('바로가기_수학교실') || '';
    var sent = 0;

    function eff(map, cls) {
      var v = map[cls] || map['all'];
      return v ? String(v) : '';
    }

    for (var i = 1; i < roster.length; i++) {
      var sid = String(roster[i][1] || '').trim();
      if (!sid) continue;
      var cls = sid.length >= 2 ? (sid.substring(0,1) + '학년 ' + sid.substring(1,2) + '반') : '';
      var oldEff = eff(oldDl, cls);
      var newEff = eff(newDl, cls);
      if (newEff === oldEff) continue;      // 변경 없음
      if (!newEff) continue;                 // 마감일이 사라진 경우는 알리지 않음
      var tokens = _parseTokens_(roster[i][4]);
      if (!tokens.length) continue;

      var title = oldEff
        ? '📅 마감일 변경: ' + taskName
        : '📝 마감일 추가: ' + taskName;
      var body = '새 마감: ' + _formatDeadline_(newEff);
      if (_sendAndPrune_(sheet, i + 1, title, body, clickUrl, 'taskUpdate')) sent++;
    }
    return sent;
  } catch(e) {
    Logger.log('_notifyTaskDeadlineChange_ 오류: ' + e.message);
    return 0;
  }
}
function notifyGradedHourly() {
  // 동시 실행 방지 (트리거 중복 설치 시 한 번만 실행되도록)
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    Logger.log('notifyGradedHourly: 이미 실행 중, 건너뜀');
    return { skipped: true };
  }
  try {
    var props = PropertiesService.getScriptProperties();
    var lastTs = parseInt(props.getProperty('last_grade_notify_ts') || '0');
    var now = Date.now();
    if (!lastTs) lastTs = now - 3600 * 1000;

    var ss = _taskSs();
    var sheet = ss.getSheetByName('제출현황');
    var data = sheet.getDataRange().getValues();
    var rosterSheet = _authRoster_();
    var roster = rosterSheet.getDataRange().getValues();
    var clickUrl = ClassCore.getConfig('바로가기_수학교실') || '';

    var tokenMap = {}, rowMap = {};
    for (var r = 1; r < roster.length; r++) {
      var rsid = String(roster[r][1] || '').trim();
      if (rsid) { tokenMap[rsid] = _parseTokens_(roster[r][4]); rowMap[rsid] = r + 1; }
    }

    var sent = 0;
    for (var i = 1; i < data.length; i++) {
      var changeTs = data[i][23]; // X열
      if (!changeTs) continue;
      var changeTime = new Date(changeTs).getTime();
      if (changeTime <= lastTs || changeTime > now) continue;

      var status = String(data[i][10] || '').trim();
      var sid = String(data[i][1] || '').trim();
      var taskName = String(data[i][3] || '').split(' (')[0];
      var tokens = tokenMap[sid] || [];
      if (!tokens.length) continue;

      var title, body;
      if (status === '채점완료') {
        title = '✅ 채점 완료: ' + taskName;
        body  = '선생님의 피드백을 확인하세요';
      } else if (status === '재제출요청') {
        title = '🔄 재제출 요청: ' + taskName;
        body  = '과제를 다시 제출해주세요';
      } else if (status === '피드백요청') {
        title = '📩 피드백 도착: ' + taskName;
        body  = '선생님 피드백을 보고 보완해서 다시 제출해주세요';
      } else {
        continue;
      }
      if (_sendAndPrune_(rosterSheet, rowMap[sid], title, body, clickUrl, 'graded')) sent++;
    }
    props.setProperty('last_grade_notify_ts', String(now));
    Logger.log('notifyGradedHourly: ' + sent + '건 발송');
    if (sent > 0) _logNotify_('✅ 채점/반려 알림', '지난 1시간 채점·반려 학생', sent);
    return { sent: sent };
  } catch(e) {
    Logger.log('notifyGradedHourly 오류: ' + e.message);
    return { error: e.message };
  } finally {
    lock.releaseLock();
  }
}
function notifyUnsubmittedDaily() {
  // 동시/중복 실행 방지 (트리거 중복 설치 대비)
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) { Logger.log('notifyUnsubmittedDaily: 이미 실행 중, 건너뜀'); return { skipped: true }; }
  try {
    var props0 = PropertiesService.getScriptProperties();
    var today = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');
    if (props0.getProperty('last_unsub_notify_day') === today) {
      Logger.log('notifyUnsubmittedDaily: 오늘 이미 발송됨, 건너뜀');
      return { skipped: true };
    }
    var ss = _taskSs();
    var tasks = ss.getSheetByName('과제설정').getDataRange().getValues();
    var subs  = ss.getSheetByName('제출현황').getDataRange().getValues();
    var rosterSheet = _authRoster_();
    var roster = rosterSheet.getDataRange().getValues();
    var clickUrl = ClassCore.getConfig('바로가기_수학교실') || '';
    var now = new Date();

    var tokenMap = {}, rowMap = {};
    for (var r = 1; r < roster.length; r++) {
      var rsid = String(roster[r][1] || '').trim();
      if (rsid) { tokenMap[rsid] = _parseTokens_(roster[r][4]); rowMap[rsid] = r + 1; }
    }

    // 학번 → {과제명 → true}
    var submittedMap = {};
    for (var s = 1; s < subs.length; s++) {
      var ssid = String(subs[s][1] || '').trim();
      var stask = String(subs[s][3] || '').split(' (')[0];
      if (!ssid || !stask) continue;
      if (!submittedMap[ssid]) submittedMap[ssid] = {};
      submittedMap[ssid][stask] = true;
    }

    // 학생별 미제출 과제 모음
    var studentMissing = {}; // sid → [{name, deadline}, ...]

    for (var i = 1; i < tasks.length; i++) {
      var tName = String(tasks[i][1] || '').trim();
      if (!tName) continue;
      var deadlines;
      try { deadlines = JSON.parse(tasks[i][3] || '{}'); } catch(_) { continue; }

      for (var r2 = 1; r2 < roster.length; r2++) {
        var rsid2 = String(roster[r2][1] || '').trim();
        if (!rsid2) continue;
        var cls = rsid2.length >= 2 ? (rsid2.substring(0,1) + '학년 ' + rsid2.substring(1,2) + '반') : '';
        var dl = deadlines[cls] || deadlines['all'];
        if (!dl) continue;
        var dlDate = new Date(dl);
        if (dlDate < now) continue;
        if ((submittedMap[rsid2] || {})[tName]) continue;

        if (!studentMissing[rsid2]) studentMissing[rsid2] = [];
        studentMissing[rsid2].push({ name: tName, deadline: dlDate });
      }
    }

    var sent = 0, studentCount = 0;
    Object.keys(studentMissing).forEach(function(sid) {
      var tokens = tokenMap[sid] || [];
      if (!tokens.length) return;
      var missing = studentMissing[sid];
      missing.sort(function(a, b) { return a.deadline - b.deadline; });
      studentCount++;

      var title, body;
      if (missing.length === 1) {
        var m = missing[0];
        title = '📌 미제출 과제: ' + m.name;
        body  = _dDayText_(m.deadline, now);
      } else {
        title = '📌 미제출 과제 ' + missing.length + '개';
        body  = missing.slice(0, 4).map(function(m){
          return m.name + ' (' + _dDayText_(m.deadline, now) + ')';
        }).join(', ');
        if (missing.length > 4) body += ' 외 ' + (missing.length - 4) + '개';
      }

      if (_sendAndPrune_(rosterSheet, rowMap[sid], title, body, clickUrl, 'unsub')) sent++;
    });

    props0.setProperty('last_unsub_notify_day', today);
    Logger.log('notifyUnsubmittedDaily: 학생 ' + studentCount + '명, ' + sent + '건 발송');
    if (sent > 0) _logNotify_('📌 미제출 알림', '미제출자 ' + studentCount + '명', sent);
    return { sent: sent, students: studentCount };
  } catch(e) {
    Logger.log('notifyUnsubmittedDaily 오류: ' + e.message);
    return { error: e.message };
  } finally {
    lock.releaseLock();
  }
}
function getTriggerStatus() {
  try {
    var triggers = ScriptApp.getProjectTriggers();
    var counts = {};
    triggers.forEach(function(t) {
      var fn = t.getHandlerFunction();
      counts[fn] = (counts[fn] || 0) + 1;
    });
    return { success: true, counts: counts, total: triggers.length };
  } catch(e) { return { success: false, message: e.toString() }; }
}
function installAutoNotifyTriggers() {
  var deleted = 0;
  ScriptApp.getProjectTriggers().forEach(function(t) {
    var fn = t.getHandlerFunction();
    if (fn === 'notifyGradedHourly' || fn === 'notifyUnsubmittedDaily') {
      ScriptApp.deleteTrigger(t);
      deleted++;
    }
  });

  ScriptApp.newTrigger('notifyGradedHourly')
    .timeBased().everyHours(1).nearMinute(0).create();

  ScriptApp.newTrigger('notifyUnsubmittedDaily')
    .timeBased().atHour(8).nearMinute(20).everyDays(1).create();

  return { success: true, message: '기존 ' + deleted + '개 정리 후 트리거 2개 설치 완료' };
}
function getCached(key) {
  try {
    const cached = CacheService.getScriptCache().get(key);
    return cached ? JSON.parse(cached) : null;
  } catch(e) { return null; }
}
function setCached(key, data, ttlSeconds) {
  try {
    CacheService.getScriptCache().put(key, JSON.stringify(data), ttlSeconds || 300);
  } catch(e) {}
}
function clearCache() {
  try {
    CacheService.getScriptCache().removeAll(['roster', 'classList', 'tasks', 'hrActivities', 'homeroom']);
  } catch(e) {}
}
function parseRosterAndClasses(ss) {
  let cached = getCached('roster');
  if (cached) return cached;

  const rosterData = _authRoster_().getDataRange().getValues();
  let roster = [], classSet = new Set();
  for (let i = 1; i < rosterData.length; i++) {
    if (!rosterData[i][1]) continue;
    let sid = String(rosterData[i][1] || "").trim();
    let sname = String(rosterData[i][2] || "").trim();
    let clsName = sid.length >= 2 ? `${sid.substring(0,1)}학년 ${sid.substring(1,2)}반` : "기타";
    roster.push({ id: sid, name: sname, cls: clsName });
    if (clsName !== "기타") classSet.add(clsName);
  }
  let classList = Array.from(classSet).sort();
  let result = { roster, classList };
  setCached('roster', result, 600); // 10분 캐시
  return result;
}
function parseTasks(ss) {
  let cached = getCached('tasks');
  if (cached) return cached;

  const taskData = ss.getSheetByName("과제설정").getDataRange().getValues();
  let tasks = [];
  for (let i = 1; i < taskData.length; i++) {
    if (!taskData[i][1]) continue;
    let reqPics = taskData[i][6] ? parseInt(taskData[i][6]) : 1;
    let choiceList = taskData[i][7] ? String(taskData[i][7] || "").trim() : "";
    tasks.push({
      rowIdx: i + 1,
      date: taskData[i][0] ? Utilities.formatDate(new Date(taskData[i][0]), "Asia/Seoul", "yyyy-MM-dd HH:mm") : "",
      name: String(taskData[i][1] || "").trim(),
      desc: taskData[i][2] ? String(taskData[i][2] || "").trim() : "",
      deadlines: taskData[i][3] ? String(taskData[i][3] || "").trim() : "{}",
      evalType: taskData[i][4] ? String(taskData[i][4] || "").trim() : "점수제 (직접 입력)",
      isPublic: String(taskData[i][5] || "").trim() === "일괄공개",
      reqPics: reqPics,
      choiceList: choiceList,
      rejectDays:   taskData[i][9]  !== '' && taskData[i][9]  != null ? parseInt(taskData[i][9])  : 7,
      feedbackDays: taskData[i][10] !== '' && taskData[i][10] != null ? parseInt(taskData[i][10]) : 7
    });
  }
  tasks.reverse();
  setCached('tasks', tasks, 300); // 5분 캐시
  return tasks;
}
function parseSubmissions(ss) {
  const sheet = ss.getSheetByName('제출현황');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  // getDataRange 대신 필요한 열만 읽기 (A~AA = 1~27열)
  const subData = sheet.getRange(2, 1, lastRow - 1, 27).getValues();
  const submissions = [];

  // ✅ [추가] 과제별 실시간 등수 계산용 객체
  let rankCounters = {};

  for (let i = 0; i < subData.length; i++) {
    if (!subData[i][1]) continue;

    let rowId = String(subData[i][1] || '').trim();
    let rowTask = String(subData[i][3] || '').trim();
    let baseTaskName = rowTask.split(' (')[0];
    let rowStatus = String(subData[i][10] || '').trim();
    let rowClass = rowId.length >= 2 ? `${rowId.substring(0,1)}학년 ${rowId.substring(1,2)}반` : "기타";

    // ✅ [추가] 반려나 이전기록이 아니면 등수 카운트 증가
    let myTotalRank = 0;
    let myClassRank = 0;
    if (rowStatus !== '재제출요청' && rowStatus !== '피드백요청' && rowStatus !== '반려검토' && rowStatus !== '이전기록채점완료') {
      if (!rankCounters[baseTaskName]) rankCounters[baseTaskName] = { total: 0, classes: {} };
      if (!rankCounters[baseTaskName].classes[rowClass]) rankCounters[baseTaskName].classes[rowClass] = 0;
      
      rankCounters[baseTaskName].total++;
      rankCounters[baseTaskName].classes[rowClass]++;
      
      myTotalRank = rankCounters[baseTaskName].total;
      myClassRank = rankCounters[baseTaskName].classes[rowClass];
    }

    // URL 파싱 (G열=6)
    const rawUrl = String(subData[i][6] || '').trim();
    let urls = {};
    if (rawUrl.startsWith('{')) { try { urls = JSON.parse(rawUrl); } catch(e) {} }
    else if (rawUrl.startsWith('[')) { try { JSON.parse(rawUrl).forEach((v,j) => urls['사진 '+(j+1)] = v); } catch(e) {} }
    else if (rawUrl) { urls['사진 1'] = rawUrl; }
    
    // 첨삭 URL 파싱 (L열=11)
    const rawAnno = String(subData[i][11] || '').trim();
    let annoUrls = {};
    if (rawAnno.startsWith('{')) { try { annoUrls = JSON.parse(rawAnno); } catch(e) {} }
    else if (rawAnno.startsWith('[')) { try { JSON.parse(rawAnno).forEach((v,j) => annoUrls['사진 '+(j+1)] = v); } catch(e) {} }
    else if (rawAnno) { annoUrls['사진 1'] = rawAnno; }
    
    // 문항별 데이터 (V열=21)
    let perQuestionData = {};
    try {
      const pqRaw = String(subData[i][21] || '').trim();
      if (pqRaw.startsWith('{')) perQuestionData = JSON.parse(pqRaw);
    } catch(e) {}
    
    // 날짜
    const dateVal = subData[i][0];
    const dateMs  = dateVal instanceof Date ? dateVal.getTime() : 0;

    submissions.push({
      rowIdx:          i + 2,
      dateMs:          dateMs,
      id:    rowId,
      name:  String(subData[i][2] || '').trim(),
      task:  rowTask,
      level: String(subData[i][4] || '').trim() || undefined,
      message: String(subData[i][5] || '').trim() || undefined,
      urls,
      feedback:    String(subData[i][7]  || '').trim() || undefined,
      status:      rowStatus,
      annoUrls:    Object.keys(annoUrls).length ? annoUrls : undefined,
      score:       String(subData[i][12] || '').trim() || undefined,
      isPublic:    String(subData[i][13] || '').trim() === '공개',
      memo:        String(subData[i][14] || '').trim() || undefined,
      reply:       String(subData[i][15] || '').trim() || undefined,
      bestType:    String(subData[i][16] || '').trim() || undefined,
      isAnon:      subData[i][17] === true || String(subData[i][17] || '').toUpperCase() === 'TRUE',
      bestComment: String(subData[i][18] || '').trim() || undefined,
      bestKey:     String(subData[i][19] || '').trim() || undefined,
      cheatFlag:   String(subData[i][20] || '').trim() || undefined,
      perQuestionData: Object.keys(perQuestionData).length ? perQuestionData : undefined,
      aiGradeTemp: String(subData[i][22] || '').trim() || undefined,
      // 재제출 루프
      resubDeadline: subData[i][24] instanceof Date ? Utilities.formatDate(subData[i][24], 'Asia/Seoul', 'yyyy-MM-dd') : (String(subData[i][24] || '').trim() || undefined),
      returnType:    String(subData[i][25] || '').trim() || undefined,
      returnCount:   subData[i][26] !== '' && subData[i][26] != null ? parseInt(subData[i][26]) : 0,
      // ✅ [추가] 계산된 등수를 데이터에 포함
      totalRank: myTotalRank,
      classRank: myClassRank
    });
  }
  submissions.reverse();
  return submissions;
}
// 재제출 기한(col25)이 지났는데 학생이 안 낸 반려/피드백 행 → 직전 점수 유지한 채 자동 완료(잠금)
function _autoCompleteExpiredReturns_() {
  try {
    const s = _taskSs().getSheetByName('제출현황');
    const last = s.getLastRow();
    if (last < 2) return 0;
    const d = s.getRange(2, 11, last - 1, 15).getValues(); // K(11=상태)~Y(25=재제출마감)
    const now = new Date();
    let changed = 0;
    for (let i = 0; i < d.length; i++) {
      const st = String(d[i][0] || '').trim(); // col11
      if (st !== '재제출요청' && st !== '피드백요청') continue;
      const dl = d[i][14]; // col25
      if (!dl) continue;
      const dlDate = (dl instanceof Date) ? dl : new Date(dl);
      if (isNaN(dlDate.getTime()) || dlDate >= now) continue;
      _setSubmissionStatus_(s, i + 2, '완료'); // 직전 점수 그대로 두고 잠금
      changed++;
    }
    return changed;
  } catch(e) { return 0; }
}

function getGradeData() {
  try {
    const ss = _taskSs();
    _autoCompleteExpiredReturns_();
    const { roster, classList } = parseRosterAndClasses(ss);
    const tasks = parseTasks(ss);
    const submissions = parseSubmissions(ss);
    let taskEvalMap = {};
    tasks.forEach(t => taskEvalMap[t.name] = t.evalType);
    return { roster, classList, tasks, submissions, taskEvalMap };
  } catch(e) { throw new Error("채점 데이터 오류: " + e.message); }
}
function getTaskPageData() {
  try {
    const ss = _taskSs();
    const { classList } = parseRosterAndClasses(ss);
    clearCache(); // 과제 수정 후 캐시 갱신
    const tasks = parseTasks(ss);
    return { classList, tasks };
  } catch(e) { throw new Error("과제 데이터 오류: " + e.message); }
}
function saveNewTask(taskData) {
  try {
    const sheet = _taskSs().getSheetByName("과제설정");
    const existing = sheet.getRange("B:B").getValues().flat();
    if (existing.includes(taskData.name)) return { success: false, message: "이미 같은 이름의 과제가 존재합니다." };
    // I(9)=만점 자리 유지, J(10)=성취기준 코드 JSON 배열
    sheet.appendRow([new Date(), taskData.name, taskData.desc, taskData.deadlines, taskData.evalType, taskData.isPublic ? "일괄공개" : "비공개", taskData.reqPics, taskData.choiceList, taskData.maxScore || '', taskData.standards || '']);
    const parentFolder = DriveApp.getFolderById(_getParentFolderId_());
    let taskFolder = parentFolder.createFolder(taskData.name);
    let deadlineObj = JSON.parse(taskData.deadlines);
    for (let cls in deadlineObj) taskFolder.createFolder(cls);
    clearCache();
    // 🔔 새 과제 알림 — 마감일 설정된 반 학생만
    var pushed = _notifyNewTask_(taskData.name, taskData.deadlines);
    if (pushed > 0) _logNotify_('📝 새 과제', taskData.name, pushed);
    return { success: true, pushed: pushed };
  } catch(e) { return { success: false, message: e.toString() }; }
}
function deleteTask(taskName) {
  try {
    if (!taskName) return { success: false, message: '과제명이 없습니다.' };
    const s = _taskSs().getSheetByName('과제설정');
    const d = s.getDataRange().getValues();
    for (let i = 1; i < d.length; i++) {
      if (String(d[i][1]).trim() === taskName) {
        s.deleteRow(i + 1);
        clearCache();
        return { success: true };
      }
    }
    return { success: false, message: '과제를 찾을 수 없습니다.' };
  } catch(e) { return { success: false, message: e.toString() }; }
}
function updateTaskSettings(t) {
  try {
    const s = _taskSs().getSheetByName("과제설정");
    const d = s.getDataRange().getValues();
    let r = -1;
    for (let i = 1; i < d.length; i++) {
      if (String(d[i][1]).trim() === t.originalName) { r = i + 1; break; }
    }
    if (r === -1) return { success: false, message: "과제를 찾을 수 없습니다." };
    const oldDeadlines = String(d[r-1][3] || '');
    s.getRange(r, 3).setValue(t.desc);
    s.getRange(r, 4).setValue(t.deadlines);
    s.getRange(r, 5).setValue(t.evalType);
    s.getRange(r, 6).setValue(t.isPublic ? "일괄공개" : "비공개");
    if (t.rejectDays != null)   s.getRange(r, 10).setValue(parseInt(t.rejectDays)   || 7);
    if (t.feedbackDays != null) s.getRange(r, 11).setValue(parseInt(t.feedbackDays) || 7);
    clearCache();
    // 🔔 마감일 변경/신규 반 추가 시 해당 학생만 알림
    var pushed = _notifyTaskDeadlineChange_(t.originalName, oldDeadlines, t.deadlines);
    if (pushed > 0) _logNotify_('📅 마감일 변경', t.originalName, pushed);
    return { success: true, pushed: pushed };
  } catch(e) { return { success: false, message: e.toString() }; }
}
function toggleTaskVisibility(n, p) {
  try {
    const ss = _taskSs();
    const ts = ss.getSheetByName("과제설정");
    const td = ts.getDataRange().getValues();
    // ✅ 과제설정 시트에서 해당 행만 업데이트
    for (let i = 1; i < td.length; i++) {
      if (String(td[i][1]).trim() === n) {
        ts.getRange(i+1, 6).setValue(p ? "일괄공개" : "비공개");
        break;
      }
    }
    // ✅ 제출현황 시트 배치 업데이트 — 전체 읽고 해당 행만 모아서 한번에 setValues
    const s = ss.getSheetByName("제출현황");
    const sd = s.getDataRange().getValues();
    const val = p ? "공개" : "비공개";
    const updates = [];
    for (let i = 1; i < sd.length; i++) {
      if (String(sd[i][3]).split(' (')[0] === n) updates.push(i + 1);
    }
    // 배치 업데이트: setValue 다중 호출 → 연속 범위 묶어서 처리
    let start = null, prev = null;
    function flushRange(from, to) {
      const numRows = to - from + 1;
      s.getRange(from, 14, numRows, 1).setValues(Array(numRows).fill([val]));
    }
    updates.sort(function(a, b) { return a - b; }).forEach(function(row) {
      if (start === null) { start = row; prev = row; }
      else if (row === prev + 1) { prev = row; }
      else { flushRange(start, prev); start = row; prev = row; }
    });
    if (start !== null) flushRange(start, prev);
    clearCache();
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}
function bulkPublishTasks(n, c) {
  try {
    const s = _taskSs().getSheetByName("제출현황");
    const d = s.getDataRange().getValues();
    // ✅ 대상 행 번호 수집 후 배치 setValue
    const rows = [];
    for (let i = 1; i < d.length; i++) {
      if (!d[i][1]) continue;
      let sid = String(d[i][1]).trim();
      let sc = sid.length >= 2 ? `${sid.substring(0,1)}학년 ${sid.substring(1,2)}반` : "기타";
      let tn = String(d[i][3]).split(' (')[0];
      if ((n === "all" || tn === n) && (c === "all" || sc === c)) rows.push(i + 1);
    }
    rows.forEach(function(row) { s.getRange(row, 14).setValue("공개"); });
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}
function getTasksOnly() {
  try {
    const ss = _taskSs();
    clearCache(); // 과제 변경 후 캐시 무효화
    return { tasks: parseTasks(ss) };
  } catch(e) { return { tasks: [] }; }
}
// ── 재제출 루프 헬퍼 ──
// 과제별 반려/피드백 재제출기한(일수) 조회 (미설정 시 7일)
function _taskResubDays_(taskName) {
  try {
    var base = String(taskName || '').split(' (')[0].trim();
    var d = _taskSs().getSheetByName('과제설정').getDataRange().getValues();
    for (var i = 1; i < d.length; i++) {
      if (String(d[i][1] || '').trim() === base) {
        return {
          reject:   d[i][9]  !== '' && d[i][9]  != null ? parseInt(d[i][9])  : 7,
          feedback: d[i][10] !== '' && d[i][10] != null ? parseInt(d[i][10]) : 7
        };
      }
    }
  } catch(e) {}
  return { reject: 7, feedback: 7 };
}
// 이 학생·이 과제의 지금까지 되돌린(반려+피드백) 횟수 = 이전기록채점완료 행 수
function _returnCountFor_(sheet, studentId, baseTask) {
  var d = sheet.getDataRange().getValues();
  var cnt = 0;
  for (var i = 1; i < d.length; i++) {
    if (String(d[i][1]).trim() === studentId &&
        String(d[i][3]).split(' (')[0].trim() === baseTask &&
        String(d[i][10] || '').trim() === '이전기록채점완료') cnt++;
  }
  return cnt;
}

function saveFeedback(r, f, a, sc, p, m, bestType, bestAnon, bestComment) {
  try {
    const s = _taskSs().getSheetByName("제출현황");
    // ✅ 피드백 + 메타 데이터 배치로 읽기 (getRange 1회)
    let rowData = s.getRange(r, 1, 1, 4).getValues()[0];
    let studentId = String(rowData[1] || "").trim();
    let taskName = String(rowData[3] || "").trim();
    let baseTaskName2 = taskName.split(" (")[0].trim();
    // ── 재제출 루프: 반려/피드백/최종완료 ──
    if (a === "반려요청" || a === "피드백요청") {
      // 2회 제한: 이미 2회 되돌렸으면 차단
      if (_returnCountFor_(s, studentId, baseTaskName2) >= 2) {
        return { success: false, message: "이미 2회 되돌렸습니다. 더 이상 재제출을 요청할 수 없어요. (완료로 마무리해 주세요)" };
      }
      var days = _taskResubDays_(taskName);
      var isFb = (a === "피드백요청");
      var dl = new Date(); dl.setDate(dl.getDate() + (isFb ? days.feedback : days.reject));
      _setSubmissionStatus_(s, r, isFb ? "피드백요청" : "재제출요청");
      s.getRange(r, 8).setValue(f);                                   // 피드백 내용
      if (sc !== undefined && sc !== null && sc !== "")
        s.getRange(r, 13, 1, 3).setValues([[sc, p ? "공개" : "비공개", m || ""]]); // 점수(이전 점수 보존용)
      s.getRange(r, 25).setValue(dl);                                 // 재제출 개별마감
      s.getRange(r, 26).setValue(isFb ? "피드백" : "반려");            // 되돌림유형
      s.getRange(r, 27).setValue(_returnCountFor_(s, studentId, baseTaskName2) + 1); // 되돌림횟수
      return { success: true, deadline: Utilities.formatDate(dl, "Asia/Seoul", "yyyy-MM-dd") };
    }
    if (a === "완료해제") {
      _setSubmissionStatus_(s, r, "채점완료");
      s.getRange(r, 25, 1, 3).clearContent(); // 재제출마감/유형/횟수 초기화
      return { success: true };
    }
    if (a === "최종완료") {
      _setSubmissionStatus_(s, r, "완료");
      s.getRange(r, 8).setValue(f);
      if (sc !== undefined && sc !== null && sc !== "")
        s.getRange(r, 13, 1, 3).setValues([[sc, p ? "공개" : "비공개", m || ""]]);
      // 이전 차수 행 정리
      let records = s.getDataRange().getValues();
      for (let i = records.length - 1; i >= 1; i--) {
        if (String(records[i][1]).trim() === studentId &&
            String(records[i][3]).split(' (')[0].trim() === baseTaskName2 &&
            (i + 1) !== r) {
          let oldStatus = String(records[i][10] || "").trim();
          if (oldStatus !== "이전기록채점완료") _setSubmissionStatus_(s, i + 1, "이전기록채점완료");
        }
      }
      return { success: true };
    }
    if (a === "완료") {
      _setSubmissionStatus_(s, r, "채점완료");
      if (taskName.includes("(재제출)")) {
        let baseTaskName = taskName.split(" (재제출)")[0];
        let records = s.getDataRange().getValues();
        for (let i = records.length - 1; i >= 1; i--) {
          if (String(records[i][1]).trim() === studentId && String(records[i][3]).trim() === baseTaskName) {
            let oldStatus = String(records[i][10] || "").trim();
            if (oldStatus !== "이전기록채점완료") _setSubmissionStatus_(s, i+1, "이전기록채점완료");
          }
        }
      }
    } else if (a === "재제출") {
      s.getRange(r, 10).clearContent();
      _setSubmissionStatus_(s, r, "재제출요청");
    } else if (a === "반려검토") {
      _setSubmissionStatus_(s, r, "반려검토");
    } else if (a === "채점중") {
      let curStatus = String(s.getRange(r, 11).getValue() || "").trim();
      if (curStatus !== "채점완료" && curStatus !== "재제출요청" && curStatus !== "반려검토") {
        _setSubmissionStatus_(s, r, "채점중");
      }
    }
    s.getRange(r, 8).setValue(f);
    s.getRange(r, 13, 1, 3).setValues([[sc, p ? "공개" : "비공개", m]]);
    // ✅ 17~20열: 우수작 정보 — bestType 값이 있을 때만 저장 (기존값 보호)
    if (bestType !== undefined && bestType !== null) {
      // bestKey: 카드 채점에서는 제출된 사진 전체 키로 자동 설정
      let existingBestKey = String(s.getRange(r, 20).getValue() || "").trim();
      if (!existingBestKey || existingBestKey === '[]') {
        // 제출 사진 URL에서 키 추출
        let rawUrl = String(s.getRange(r, 7).getValue() || "").trim();
        let urls = {};
        if (rawUrl.startsWith('{')) { try { urls = JSON.parse(rawUrl); } catch(e) {} }
        let urlKeys = Object.keys(urls).filter(k => urls[k] && urls[k] !== "첨부파일 없음");
        existingBestKey = urlKeys.length > 0 ? JSON.stringify(urlKeys) : '[]';
      }
      s.getRange(r, 17, 1, 4).setValues([[
        bestType    || "",
        bestAnon    ? "TRUE" : "FALSE",
        bestComment || "",
        existingBestKey
      ]]);
    }
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}
function getSecureFileBase64(url) {
  try {
    const f = DriveApp.getFileById(url.match(/[-\w]{25,}/)[0]);
    if (f.getSize() > 10485760) return { success: false, message: "용량 초과 (10MB)" };
    const rawMime = f.getMimeType().toLowerCase();
    const mime = _normalizeImageMime(rawMime);
    if (!mime) return { success: false, message: "지원하지 않는 형식: " + rawMime };
    const b = f.getBlob();
    return { success: true, mimeType: mime, data: Utilities.base64Encode(b.getBytes()) };
  } catch(e) { return { success: false }; }
}
function saveMultiAnnotatedImages(rowIdx, payloadArray, studentId, studentName, taskName, feedbackText, statusAction, score, isPublic, memo, bestType, isAnon, bestComment, bestKey, perQuestionJson) {
  try {
    const sheet = _taskSs().getSheetByName("제출현황");
    const parentFolder = DriveApp.getFolderById(_getParentFolderId_());
    let baseTask = String(taskName || "").split(' (')[0];
    let safeStudentId = String(studentId || "");
    let className = safeStudentId.length >= 2
      ? `${safeStudentId.substring(0,1)}학년 ${safeStudentId.substring(1,2)}반` : "기타";

    // ✅ Drive 폴더 ID를 PropertiesService에 캐시 → getFoldersByName 반복 탐색 방지
    const props = PropertiesService.getScriptProperties();
    const taskFolderKey   = 'folderId_task_' + baseTask;
    const classFolderKey  = 'folderId_class_' + baseTask + '_' + className;
    let taskFolder, classFolder;

    const cachedTaskId = props.getProperty(taskFolderKey);
    try {
      taskFolder = cachedTaskId ? DriveApp.getFolderById(cachedTaskId) : null;
    } catch(e) { taskFolder = null; }
    if (!taskFolder) {
      taskFolder = parentFolder.getFoldersByName(baseTask).hasNext()
        ? parentFolder.getFoldersByName(baseTask).next()
        : parentFolder.createFolder(baseTask);
      props.setProperty(taskFolderKey, taskFolder.getId());
    }

    const cachedClassId = props.getProperty(classFolderKey);
    try {
      classFolder = cachedClassId ? DriveApp.getFolderById(cachedClassId) : null;
    } catch(e) { classFolder = null; }
    if (!classFolder) {
      classFolder = taskFolder.getFoldersByName(className).hasNext()
        ? taskFolder.getFoldersByName(className).next()
        : taskFolder.createFolder(className);
      props.setProperty(classFolderKey, classFolder.getId());
    }

    // 첨삭 이미지 저장
    let rawAnno = sheet.getRange(rowIdx, 12).getValue();
    let currentAnnoUrls = {};
    if (String(rawAnno).startsWith('{')) { try { currentAnnoUrls = JSON.parse(rawAnno); } catch(e) {} }
    payloadArray.forEach(item => {
      const data = item.base64Data.split(',')[1];
      const blob = Utilities.newBlob(
        Utilities.base64Decode(data), 'image/jpeg',
        `[${safeStudentId}] ${studentName}_${baseTask}_${item.key}_첨삭.jpg`
      );
      currentAnnoUrls[item.key] = classFolder.createFile(blob).getUrl();
    });
    sheet.getRange(rowIdx, 12).setValue(JSON.stringify(currentAnnoUrls));

    // =====================================================
    // ✅ 문항별 데이터(perQuestionData) 처리
    // =====================================================
    let pqData = {};
    // 기존 저장된 perQuestionData 읽기
    let rawPq = sheet.getRange(rowIdx, 22).getValue();
    if (rawPq && String(rawPq).startsWith('{')) {
      try { pqData = JSON.parse(rawPq); } catch(e) {}
    }
    // 새로 전달된 데이터로 덮어쓰기 (문항별 병합)
    if (perQuestionJson) {
      try {
        let newPq = JSON.parse(perQuestionJson);
        Object.keys(newPq).forEach(k => { pqData[k] = newPq[k]; });
      } catch(e) {}
    }

    // =====================================================
    // ✅ 2형 과제 여부 판단 (문항 수 > 1)
    // =====================================================
    let submittedUrls = {};
    let rawUrl = sheet.getRange(rowIdx, 7).getValue();
    if (rawUrl && String(rawUrl).startsWith('{')) {
      try { submittedUrls = JSON.parse(rawUrl); } catch(e) {}
    }
    let submittedKeys = Object.keys(submittedUrls).filter(k =>
      submittedUrls[k] && submittedUrls[k] !== "첨부파일 없음"
    );
    let isType2 = submittedKeys.length > 1;

    // =====================================================
    // ✅ 상태 결정
    // =====================================================
    let finalStatus = statusAction;

    if (isType2) {
      // 2형: pqData 기반으로 완료 여부 자동 판단
      // allKeys: 실제 제출된 사진 키
      let checkKeys = submittedKeys.length > 0 ? submittedKeys : Object.keys(pqData);
      let doneCount   = checkKeys.filter(k => pqData[k] && pqData[k].status === '완료').length;
      let rejectCount = checkKeys.filter(k => pqData[k] && pqData[k].status === '반려검토').length;
      let processedCount = doneCount + rejectCount;

      if (checkKeys.length > 0 && processedCount === checkKeys.length) {
        // 모든 문항 처리 완료
        finalStatus = rejectCount > 0 ? '반려검토' : '완료';
      } else {
        // 일부만 처리됨
        finalStatus = '채점중';
      }
    }
    // 1형은 statusAction 그대로 사용

    // =====================================================
    // ✅ 피드백 처리 (2형: 문항별 피드백 합산)
    // =====================================================
    let finalFeedback = feedbackText;
    if (isType2 && Object.keys(pqData).length > 0 && !feedbackText) {
      let fbParts = [];
      Object.keys(pqData).forEach(k => {
        if (pqData[k] && pqData[k].fb) fbParts.push(`[${k}] ${pqData[k].fb}`);
      });
      if (fbParts.length > 0) finalFeedback = fbParts.join('\n');
    }
    sheet.getRange(rowIdx, 8).setValue(finalFeedback);

    // =====================================================
    // ✅ 점수 처리 (2형: 문항별 점수 합산 또는 나열)
    // =====================================================
    let finalScore = score;
    if (isType2 && Object.keys(pqData).length > 0) {
      let scores = Object.keys(pqData)
        .map(k => pqData[k] && pqData[k].sc ? String(pqData[k].sc).trim() : "")
        .filter(s => s !== "");
      if (scores.length > 0) {
        let allNumeric = scores.every(s => !isNaN(parseFloat(s)));
        if (allNumeric) {
          finalScore = scores.reduce((a, b) => parseFloat(a) + parseFloat(b), 0).toString();
        } else {
          finalScore = scores.join(', ');
        }
      }
    }

    // =====================================================
    // ✅ 상태 저장
    // =====================================================
    if (finalStatus === '최종완료') {
      _setSubmissionStatus_(sheet, rowIdx, '완료');
      let recF = sheet.getDataRange().getValues();
      for (let i = recF.length - 1; i >= 1; i--) {
        if (String(recF[i][1]||'').trim() === safeStudentId &&
            String(recF[i][3]||'').split(' (')[0].trim() === baseTask && (i+1) !== rowIdx) {
          let oldSt = String(recF[i][10]||'').trim();
          if (oldSt !== '이전기록채점완료') _setSubmissionStatus_(sheet, i+1, '이전기록채점완료');
        }
      }
    } else if (finalStatus === '반려요청' || finalStatus === '피드백요청') {
      var isFbA = (finalStatus === '피드백요청');
      if (_returnCountFor_(sheet, safeStudentId, baseTask) >= 2) {
        _setSubmissionStatus_(sheet, rowIdx, '채점완료'); // 2회 초과 → 그냥 채점완료
      } else {
        var daysA = _taskResubDays_(taskName);
        var dlA = new Date(); dlA.setDate(dlA.getDate() + (isFbA ? daysA.feedback : daysA.reject));
        _setSubmissionStatus_(sheet, rowIdx, isFbA ? '피드백요청' : '재제출요청');
        sheet.getRange(rowIdx, 25).setValue(dlA);
        sheet.getRange(rowIdx, 26).setValue(isFbA ? '피드백' : '반려');
        sheet.getRange(rowIdx, 27).setValue(_returnCountFor_(sheet, safeStudentId, baseTask) + 1);
      }
    } else if (finalStatus === '완료') {
      _setSubmissionStatus_(sheet, rowIdx, '채점완료');
      if (String(taskName).includes('(재제출)')) {
        let records = sheet.getDataRange().getValues();
        for (let i = records.length - 1; i >= 1; i--) {
          if (String(records[i][1]||'').trim() === safeStudentId &&
              String(records[i][3]||'').trim() === baseTask) {
            let oldSt = String(records[i][10]||'').trim();
            if (oldSt !== '이전기록채점완료') _setSubmissionStatus_(sheet, i+1, '이전기록채점완료');
          }
        }
      }
    } else if (finalStatus === '반려검토') {
      _setSubmissionStatus_(sheet, rowIdx, '반려검토');
    } else {
      // 채점중 — 이미 완료/반려 상태면 덮어쓰지 않음
      let curSt = String(sheet.getRange(rowIdx, 11).getValue() || '').trim();
      if (curSt !== '채점완료' && curSt !== '재제출요청' && curSt !== '반려검토') {
        _setSubmissionStatus_(sheet, rowIdx, '채점중');
      }
    }

    // ✅ 개별 setValue 대신 한 번에 배치 저장 (성능 최적화)
    // 13~15열, 17~20열, 22열을 한 번에 처리
    // 13: 점수, 14: 공개여부, 15: 메모
    sheet.getRange(rowIdx, 13, 1, 3).setValues([[
      finalScore,
      isPublic ? "공개" : "비공개",
      memo
    ]]);
    // 17~20열: 우수작 관련
    // ✅ bestType이 빈값이면 pqData에서 best 정보 추출 (2형 스마트 채점 보완)
    let finalBestType    = bestType    || "";
    let finalBestAnon    = isAnon;
    let finalBestComment = bestComment || "";
    let finalBestKey     = bestKey     || "";
    if (!finalBestType) {
      Object.keys(pqData).forEach(k => {
        var d = pqData[k];
        if (d && d.best && d.best.scope) {
          finalBestType    = d.best.scope;
          finalBestComment = d.best.comment || "";
          finalBestAnon    = !!d.best.anon;
          // key: pqData.best.key 있으면 사용, 없으면 해당 문항 키
          var bk = d.best.key || "";
          try {
            var bkArr = JSON.parse(bk);
            finalBestKey = (Array.isArray(bkArr) && bkArr.length > 0) ? bk : JSON.stringify([k]);
          } catch(e) { finalBestKey = JSON.stringify([k]); }
        }
      });
    }
    // finalBestKey도 빈배열이면 제출 사진 전체 키로
    if (finalBestType && (!finalBestKey || finalBestKey === '[]')) {
      finalBestKey = JSON.stringify(submittedKeys.length > 0 ? submittedKeys : Object.keys(pqData));
    }
    sheet.getRange(rowIdx, 17, 1, 4).setValues([[
      finalBestType,
      finalBestAnon ? "TRUE" : "FALSE",
      finalBestComment,
      finalBestKey
    ]]);
    // 22열(V열): 문항별 독립 데이터
    if (Object.keys(pqData).length > 0) {
      sheet.getRange(rowIdx, 22).setValue(JSON.stringify(pqData));
    }
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}
function cancelBestWork(r) {
  try { _taskSs().getSheetByName("제출현황").getRange(r, 17, 1, 4).clearContent(); return { success: true }; }
  catch(e) { return { success: false, message: e.toString() }; }
}
function changeBestScope(rowIdx, newScope) {
  try {
    _taskSs().getSheetByName("제출현황").getRange(rowIdx, 17).setValue(newScope);
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}
function runAICheatCheck(targetTaskName) {
  try {
    const sheet = _taskSs().getSheetByName("제출현황");
    const data = sheet.getDataRange().getValues();
    let globalHashMap = {};
    
    // 1. 전체 사진 해시맵 구축 (과제명 무관 전체 비교)
    for (let i = 1; i < data.length; i++) {
      if (!data[i][1]) continue;
      let hashStr = data[i][8]; 
      let sid = String(data[i][1] || "").trim();
      if (hashStr && String(hashStr).startsWith("{")) {
        try {
          let hObj = JSON.parse(hashStr);
          for (let k in hObj) { 
            let h = hObj[k]; 
            if (!globalHashMap[h]) globalHashMap[h] = []; 
            globalHashMap[h].push({ rowIdx: i+1, sid: sid, name: data[i][2] }); 
          }
        } catch(e) {}
      }
    }
    
    // 2. 검사 대상 행 추출 (해당 과제의 '모든' 제출본 검사 - 완료된 것도 포함!)
    let targetRows = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][1]) continue;
      let rowTaskName = String(data[i][3] || "").split(' (')[0].trim();
      if (rowTaskName !== targetTaskName) continue;
      
      // ✅ 기존 조건 삭제: 채점완료/이전기록채점완료 등 상태 관계없이 무조건 검사
      let currentFlag = data[i][20] ? String(data[i][20]).trim() : "";
      if (currentFlag !== "") continue; // 이미 검사해서 플래그가 찍힌 애들만 건너뜀 (원하면 이 줄도 삭제 가능)
      
      targetRows.push(i);
    }
    
    if (targetRows.length === 0) return { success: true, count: 0, total: 0, taskName: targetTaskName, message: "검사할 대상이 없습니다." };
    
    let processedCount = 0;
    for (let idx = 0; idx < targetRows.length; idx++) {
      let i = targetRows[idx]; 
      let rowIdx = i + 1;
      let sid = String(data[i][1] || "").trim(); 
      let hashStr = data[i][8]; 
      let cheatReason = "";
      
      if (hashStr && String(hashStr).startsWith("{")) {
        try {
          let hObj = JSON.parse(hashStr);
          for (let k in hObj) {
            let h = hObj[k];
            if (globalHashMap[h] && globalHashMap[h].length > 1) {
              let others = globalHashMap[h].filter(x => x.sid !== sid);
              if (others.length > 0) cheatReason += `[🚨중복: ${others[0].sid} ${others[0].name} 도용] `;
            }
          }
        } catch(e) {}
      }
      
      if (cheatReason === "" && data[i][6] && String(data[i][6]).startsWith("{")) {
        try {
          let uObj = JSON.parse(data[i][6]);
          for (let k in uObj) {
            let url = uObj[k];
            if (!url || url === "첨부파일 없음") continue;
            let fIdMatch = url.match(/[-\w]{25,}/);
            if (fIdMatch && typeof Drive !== 'undefined' && Drive.Files) {
              let docFile = Drive.Files.copy({ title: "Temp_OCR", mimeType: "application/vnd.google-apps.document" }, fIdMatch[0], { ocr: true, ocrLanguage: 'ko' });
              let doc = DocumentApp.openById(docFile.id);
              let text = doc.getBody().getText().replace(/\s+/g, "");
              DriveApp.getFileById(docFile.id).setTrashed(true);
              if (text.includes("조회자") || text.includes("조회자:")) { cheatReason += `[📸캡처 적발: 워터마크 화면 캡처] `; break; }
            }
          }
        } catch(e) {}
      }
      
      sheet.getRange(rowIdx, 21).setValue(cheatReason !== "" ? cheatReason : "✅정상(AI 스캔완료)");
      processedCount++;
    }
    return { success: true, count: processedCount, total: targetRows.length, taskName: targetTaskName };
  } catch(e) { return { success: false, message: e.toString() }; }
}
function getApiSettings() {
  var ss = _taskSs();
  var sh = ss.getSheetByName('시스템설정');

  var orKey     = ClassCore.getConfig('OpenRouter키');
  var m         = ClassCore.getConfig('AI모델명');
  var ttClient  = ClassCore.getConfig('TickTick클라이언트ID');
  var ttSecret  = ClassCore.getConfig('TickTick시크릿');
  var ttToken   = ClassCore.getConfig('TickTick액세스토큰');
  var ttRefresh = ClassCore.getConfig('TickTick갱신토큰');

  // 구 구조 폴백 (P2:U2) — migrateSysSheet 실행 전이거나 이전 실패 시
  if (!orKey && sh) {
    try {
      var old = sh.getRange('P2:U2').getValues()[0];
      if (!orKey)    orKey    = String(old[0] || '').trim();
      if (!m)        m        = String(old[1] || '').trim();
      if (!ttClient) ttClient = String(old[2] || '').trim();
      if (!ttSecret) ttSecret = String(old[3] || '').trim();
      if (!ttToken)  ttToken  = String(old[4] || '').trim();
      if (!ttRefresh)ttRefresh= String(old[5] || '').trim();
    } catch(e) {}
  }

  if (m === 'google/gemini-2.5-flash-preview') m = 'google/gemini-2.5-flash';
  if (m === 'google/gemini-2.5-pro-preview')   m = 'google/gemini-2.5-pro';
  return {
    openrouterKey: orKey,
    model:         m || 'anthropic/claude-3.5-sonnet',
    ttClientId:    ttClient,
    ttSecret:      ttSecret,
    ttToken:       ttToken,
    ttRefresh:     ttRefresh
  };
}
function saveApiSetting(key, value) {
  try {
    // 하위 호환: 'Q2' 셀 주소로 들어오면 AI모델명 키로 변환
    var resolvedKey = key;
    if (key === 'Q2') resolvedKey = 'AI모델명';
    var ss = _taskSs();
    ClassCore.setConfig(resolvedKey, value);
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}
function getOpenRouterBalance() {
  try {
    const key = ClassCore.getConfig('OpenRouter키');
    if (!key) return { success: false, message: 'API 키 없음' };
    const res = UrlFetchApp.fetch('https://openrouter.ai/api/v1/auth/key', {
      headers: { 'Authorization': 'Bearer ' + key },
      muteHttpExceptions: true
    });
    if (res.getResponseCode() !== 200) return { success: false, message: '조회 실패 (' + res.getResponseCode() + ')' };
    const d = JSON.parse(res.getContentText()).data;
    return { success: true, usage: d.usage, limit: d.limit, isFreeTier: d.is_free_tier };
  } catch(e) { return { success: false, message: e.toString() }; }
}
function getOrCreateRubricSheet() {
  const ss = _taskSs();
  let sh = ss.getSheetByName('AI채점기준');
  if (!sh) {
    sh = ss.insertSheet('AI채점기준');
    sh.getRange(1,1,1,6).setValues([['과제명','채점유형','총점','기준설명','파일목록JSON','문항구성JSON']]);
  }
  return sh;
}
function getRubrics() {
  try {
    const sh   = getOrCreateRubricSheet();
    const last = sh.getLastRow();
    if (last < 2) return { success: true, rubrics: [] };
    const data = sh.getRange(2, 1, last-1, 6).getValues();
    const rubrics = data.filter(r => r[0]).map((r, i) => {
      let files = [];
      try { if (r[4]) files = JSON.parse(r[4]); } catch(e) {}
      let questions = {};
      try { if (r[5]) questions = JSON.parse(r[5]); } catch(e) {}
      return {
        rowIdx:    i + 2,
        taskName:  String(r[0]||'').trim(),
        evalType:  String(r[1]||'점수제').trim(),
        maxScore:  Number(r[2]||0),
        criteria:  String(r[3]||'').trim(),
        files:     files,
        questions: questions
      };
    });
    return { success: true, rubrics };
  } catch(e) { return { success: false, message: e.toString() }; }
}
function uploadRubricFile(fileName, mimeType, base64Data, taskName) {
  try {
    const bytes = Utilities.base64Decode(base64Data);
    const blob  = Utilities.newBlob(bytes, mimeType, fileName);
    
    // PARENT_FOLDER_ID 안에 'AI채점기준' 폴더 생성/조회
    let rubricFolder;
    const folders = DriveApp.getFolderById(_getParentFolderId_())
                             .getFoldersByName('AI채점기준');
    if (folders.hasNext()) {
      rubricFolder = folders.next();
    } else {
      rubricFolder = DriveApp.getFolderById(_getParentFolderId_())
                             .createFolder('AI채점기준');
    }
    
    // 기존 동명 파일 삭제
    const existing = rubricFolder.getFilesByName(taskName + '_' + fileName);
    while (existing.hasNext()) existing.next().setTrashed(true);
    
    const file = rubricFolder.createFile(blob);
    file.setName(taskName + '_' + fileName);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    return { success: true, url: file.getUrl(), fileId: file.getId(), name: fileName };
  } catch(e) { return { success: false, message: e.toString() }; }
}
function saveRubric(data) {
  try {
    const sh  = getOrCreateRubricSheet();
    const row = [
      data.taskName||'',
      data.evalType||'점수제',
      data.maxScore||0,
      data.criteria||'',
      data.files ? JSON.stringify(data.files) : '',
      data.questions ? JSON.stringify(data.questions) : ''
    ];
    if (data.rowIdx && data.rowIdx > 1) {
      sh.getRange(data.rowIdx, 1, 1, 6).setValues([row]);
    } else {
      // 같은 과제명 있으면 덮어쓰기
      const last = sh.getLastRow();
      if (last >= 2) {
        const names = sh.getRange(2,1,last-1,1).getValues().flat();
        const idx   = names.indexOf(data.taskName);
        if (idx >= 0) {
          sh.getRange(idx+2, 1, 1, 5).setValues([row]);
          return { success: true, rowIdx: idx+2 };
        }
      }
      sh.appendRow(row);
    }
    _syncTaskScale(data.taskName, data.evalType, data.maxScore); // 과제설정 단일출처 동기화
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}

// 채점기준 저장 시 과제설정의 채점유형(5열)·만점(9열)을 함께 맞춤 (SSOT 유지)
function _syncTaskScale(taskName, evalType, maxScore) {
  try {
    if (!taskName) return;
    var ts = _taskSs().getSheetByName('과제설정');
    var d = ts.getDataRange().getValues();
    for (var i = 1; i < d.length; i++) {
      if (String(d[i][1]).trim() === String(taskName).trim()) {
        if (evalType) ts.getRange(i + 1, 5).setValue(evalType);
        if (maxScore && Number(maxScore) > 0) ts.getRange(i + 1, 9).setValue(Number(maxScore));
        clearCache();
        return;
      }
    }
  } catch(_) {}
}
function saveAiGradeTempResult(rowIdx, aiResultJson) {
  try {
    const sheet = _taskSs().getSheetByName('제출현황');
    if (!sheet || rowIdx < 2) return { success: false, message: '시트 없음' };
    sheet.getRange(rowIdx, 23).setValue(aiResultJson || ''); // W열: AI임시저장
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}
function saveAiGradeResult(rowIdx, feedback, score, statusAction, perQuestionJson) {
  try {
    const sheet = _taskSs().getSheetByName('제출현황');
    if (!sheet || rowIdx < 2) return { success: false, message: '시트 없음' };
    const statusMap = { 'pass': '채점완료', 'reject': '재제출요청' };
    const newStatus = statusMap[statusAction] || statusAction;
    sheet.getRange(rowIdx, 8).setValue(feedback || '');
    _setSubmissionStatus_(sheet, rowIdx, newStatus);
    sheet.getRange(rowIdx, 13).setValue(score || '');
    if (perQuestionJson) {
      sheet.getRange(rowIdx, 22).setValue(perQuestionJson);
    }
    // ✅ W열 초기화 제거 — 학생도 AI 결과를 계속 볼 수 있도록 유지
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}
function deleteRubric(rowIdx) {
  try {
    getOrCreateRubricSheet().deleteRow(rowIdx);
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}
function _parseAiJsonLoose_(text) {
  if (!text) return null;
  var start = text.indexOf('{');
  if (start < 0) return null;
  var body = text.substring(start);
  // 1차: 그대로
  try { return JSON.parse(body); } catch(_) {}
  // 2차: 마지막 } 까지만
  var lastBrace = body.lastIndexOf('}');
  if (lastBrace > 0) {
    try { return JSON.parse(body.substring(0, lastBrace + 1)); } catch(_) {}
  }
  // 3차: 미완성 문자열·괄호({,[) 보정 — 스택으로 순서 추적
  var stack = [], inStr = false, esc = false;
  for (var i = 0; i < body.length; i++) {
    var c = body[i];
    if (esc) { esc = false; continue; }
    if (c === '\\') { esc = true; continue; }
    if (c === '"') { inStr = !inStr; continue; }
    if (inStr) continue;
    if (c === '{' || c === '[') stack.push(c);
    else if (c === '}' || c === ']') stack.pop();
  }
  var patched = body;
  if (inStr) patched += '"';
  // 끝에 남은 쉼표/콜론/여는따옴표 정리
  patched = patched.replace(/[\s]*,[\s]*$/, '').replace(/[\s]*:[\s]*$/, ':null');
  // 스택을 역순으로 닫기
  for (var j = stack.length - 1; j >= 0; j--) {
    patched += (stack[j] === '{') ? '}' : ']';
  }
  try { return JSON.parse(patched); } catch(_) {}
  // 4차: 끝쉼표 제거 후 재시도
  try { return JSON.parse(patched.replace(/,(\s*[}\]])/g, '$1')); } catch(_) { return null; }
}
function _normalizeImageMime(mime) {
  if (!mime) return 'image/jpeg';
  const m = mime.toLowerCase();
  if (m === 'image/jpg' || m === 'image/jpe' || m === 'image/jfif') return 'image/jpeg';
  if (['image/jpeg','image/png','image/gif','image/webp'].indexOf(m) >= 0) return m;
  return null; // 지원 불가
}
function getRubricFileBase64(url) {
  try {
    if (!url || !url.includes('drive.google.com')) return { success: false, message: 'Drive URL이 아닙니다.' };
    const fileId = url.match(/[-\w]{25,}/);
    if (!fileId) return { success: false, message: 'File ID 추출 실패' };
    const f    = DriveApp.getFileById(fileId[0]);
    if (f.getSize() > 10485760) return { success: false, message: '파일 크기 10MB 초과' };
    const mime = _normalizeImageMime(f.getMimeType());
    if (!mime && !f.getMimeType().includes('pdf')) return { success: false, message: '지원하지 않는 이미지 형식: ' + f.getMimeType() };
    const b64  = Utilities.base64Encode(f.getBlob().getBytes());
    return { success: true, mimeType: mime || f.getMimeType(), data: b64, name: f.getName() };
  } catch(e) { return { success: false, message: e.toString() }; }
}
function aiGradeSubmission(params) {
  /*
    params: {
      rowIdx,        // 제출현황 rowIdx
      taskName,      // 과제명
      studentId,     // 학번
      studentName,   // 이름
      imageUrls,     // { "1번": "url", ... } 또는 { "0": "url" }
      rubricRowIdx,  // AI채점기준 rowIdx
      isType2        // 2형 여부
    }
  */
  try {
    const cfg = getApiSettings();
    if (!cfg.openrouterKey) return { success: false, message: 'OpenRouter API 키가 없습니다.' };

    // 채점기준 가져오기
    const rubricRes = getRubrics();
    if (!rubricRes.success) return { success: false, message: '채점기준 조회 실패' };
    const rubric = rubricRes.rubrics.find(r => r.taskName === params.taskName || r.rowIdx === params.rubricRowIdx);
    if (!rubric) return { success: false, message: '해당 과제의 채점기준이 없습니다.' };

    // 단일 출처: 과제설정의 채점유형(col5)·만점(col9)이 있으면 우선 사용 (없으면 AI채점기준 폴백)
    try {
      var _ts = _taskSs().getSheetByName('과제설정');
      if (_ts && _ts.getLastRow() >= 2) {
        var _td = _ts.getRange(2, 1, _ts.getLastRow() - 1, 9).getValues();
        for (var _i = 0; _i < _td.length; _i++) {
          if (String(_td[_i][1] || '').trim() === rubric.taskName) {
            var _et = String(_td[_i][4] || '').trim(), _mx = Number(_td[_i][8] || 0);
            if (_et) rubric.evalType = _et;
            if (_mx) rubric.maxScore = _mx;
            break;
          }
        }
      }
    } catch(_) {}

    // 메시지 구성
    const today = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');
    const messages = [];

    // content 배열 구성
    const content = [];

    // 텍스트 프롬프트
    const isType2 = params.isType2;
    const prompt = [
      '너는 수학 교사의 채점을 돕는 AI야. 반드시 JSON만 반환해.',
      '학생: ' + params.studentId + ' ' + params.studentName,
      '과제: ' + params.taskName,
      '채점유형: ' + rubric.evalType,
      '총점: ' + rubric.maxScore + '점',
      rubric.criteria ? '채점기준 설명: ' + rubric.criteria : '',
      // 문항별 채점기준 파일이 있는 경우 안내
      (rubric.files||[]).some(f=>f.questions&&f.questions.length>0)
        ? '※ 일부 파일은 특정 문항 전용 채점기준임. 파일 레이블을 확인하여 해당 문항만 적용할 것.' : '',
      // 성취기준 파일이 있는 경우 안내
      (rubric.files||[]).some(f=>f.role==='achievement')
        ? '※ 성취기준 파일을 참고하여 A/B/C/D 판별할 것.' : '',
      rubric.imageUrl ? '(채점기준 이미지/PDF 첨부 참고)' : '',
      '반환 형식:',
      (() => {
        const et = rubric.evalType || '점수제';
        if (et.includes('A-B-C') || et.includes('등급')) {
          return '{"grade":"A|B|C|D","feedback":"피드백","confidence":"high|medium|low","needsReview":true|false,"rejectReason":"사유 또는 null"}';
        } else if (et.includes('상-중-하')) {
          return '{"grade":"상|중|하","feedback":"피드백","confidence":"high|medium|low","needsReview":true|false,"rejectReason":"사유 또는 null"}';
        } else if (et.includes('P/F') || et.includes('통과')) {
          return '{"grade":"Pass|Fail","feedback":"피드백","confidence":"high|medium|low","needsReview":true|false,"rejectReason":"사유 또는 null"}';
        } else if (isType2) {
          return '{"perQuestion":{"문항명":{"score":점수,"maxScore":만점,"feedback":"피드백"}},"totalScore":합계,"overallFeedback":"종합피드백","confidence":"high|medium|low","needsReview":true|false,"rejectReason":"사유 또는 null"}';
        } else {
          return '{"score":점수,"maxScore":' + rubric.maxScore + ',"feedback":"피드백","confidence":"high|medium|low","needsReview":true|false,"rejectReason":"사유 또는 null"}';
        }
      })()
    ].filter(Boolean).join('\n');

    content.push({ type: 'text', text: prompt });

    // 채점기준 파일 첨부 (role/scope/questions 적용)
    const rubricFiles = rubric.files || [];
    for (const rf of rubricFiles) {
      if (!rf.url) continue;
      // 문항 필터링 - 특정 문항 전용 파일이면 해당 문항만
      const targetQ = rf.questions && rf.questions.length > 0 ? rf.questions : null;
      const roleLabel = rf.role === 'achievement' ? '[성취기준 파일]' : '[채점기준 파일: ' + (targetQ ? targetQ.join(', ') : '전체') + ']';
      content.push({ type: 'text', text: roleLabel });
      const rubricFile = getRubricFileBase64(rf.url);
      if (rubricFile.success) {
        if (rubricFile.mimeType && rubricFile.mimeType.includes('pdf')) {
          content.push({ type: 'text', text: '(PDF 채점기준 파일은 텍스트 기준을 참고하세요: ' + (rf.name || '') + ')' });
        } else {
          content.push({ type: 'image_url', image_url: { url: 'data:' + rubricFile.mimeType + ';base64,' + rubricFile.data } });
        }
      } else {
        content.push({ type: 'text', text: '(파일 로드 실패: ' + rubricFile.message + ')' });
      }
    }

    // 학생 제출 사진 첨부
    const urlKeys = Object.keys(params.imageUrls || {});
    for (let i = 0; i < urlKeys.length; i++) {
      const key     = urlKeys[i];
      const fileUrl = params.imageUrls[key];
      content.push({ type: 'text', text: (isType2 ? '[' + key + ' 답안 사진]' : '[학생 답안 사진]') });
      const imgRes = getSecureFileBase64(fileUrl);
      if (imgRes.success) {
        content.push({ type: 'image_url', image_url: { url: 'data:' + imgRes.mimeType + ';base64,' + imgRes.data } });
      } else {
        content.push({ type: 'text', text: '(사진 로드 실패: ' + imgRes.message + ')' });
      }
    }

    messages.unshift({ role: 'system', content: '당신은 수학 채점 AI입니다. 반드시 JSON 객체만 반환하세요. 설명, 마크다운, 추가 텍스트 없이 오직 { } 형태의 JSON만 출력하세요.' });
    messages.push({ role: 'user', content: content });

    // API 호출
    const res = UrlFetchApp.fetch('https://openrouter.ai/api/v1/chat/completions', {
      method: 'post',
      headers: {
        'Authorization':  'Bearer ' + cfg.openrouterKey,
        'Content-Type':   'application/json',
        'HTTP-Referer':   'https://script.google.com',
        'X-Title':        'Teacher Dashboard'
      },
      payload: JSON.stringify({
        model:           cfg.model || 'google/gemini-3.5-flash',
        messages:        messages,
        max_tokens:      4000,
        temperature:     0.2
      }),
      muteHttpExceptions: true
    });

    const code = res.getResponseCode();
    if (code !== 200) {
      const err = JSON.parse(res.getContentText());
      const raw = (err.error && err.error.metadata && err.error.metadata.raw) || '';
      const msg = (err.error && err.error.message) || res.getContentText();
      return { success: false, message: 'API 오류(' + code + '): ' + msg + (raw ? ' | ' + raw : '') };
    }

    let text = JSON.parse(res.getContentText()).choices[0].message.content.trim();
    text = text.replace(/^```[a-z]*\n?/i,'').replace(/\n?```$/,'').trim();
    const result = _parseAiJsonLoose_(text);
    if (!result) return { success: false, message: 'JSON 파싱 실패: ' + text.substring(0, 150) };
    result.rowIdx      = params.rowIdx;
    result.studentId   = params.studentId;
    result.studentName = params.studentName;
    result.taskName    = params.taskName;
    return { success: true, result };

  } catch(e) { return { success: false, message: e.toString() }; }
}
function clearStudentReply(rowIdx) {
  try {
    const ss = _taskSs();
    
    // 과제 제출 데이터가 있는 시트 찾기
    const sheetNames = ["제출현황", "과제제출", "제출기록", "수학제출"];
    let sheet = null;
    for (let i = 0; i < sheetNames.length; i++) {
      sheet = ss.getSheetByName(sheetNames[i]);
      if (sheet) break;
    }
    if (!sheet) return { success: false, message: "시트를 찾을 수 없습니다." };

    // 1행(헤더)에서 '답장' 이라는 단어가 들어간 열 찾기
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    let replyCol = -1;
    for (let i = 0; i < headers.length; i++) {
      if (String(headers[i]).replace(/\s/g, '').includes('답장')) {
        replyCol = i + 1;
        break;
      }
    }

    // 해당 열이 있으면 그 칸을 깔끔하게 비움
    if (replyCol > 0 && rowIdx > 1) {
      sheet.getRange(rowIdx, replyCol).setValue('');
      return { success: true };
    }
    
    return { success: false, message: "답장 열을 찾을 수 없습니다." };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}
function clearStudentReply(rowIdx) {
  try {
    const sheet = _taskSs().getSheetByName('제출현황');
    if (!sheet) return { success: false };
    // P열(16번) = 학생답장
    sheet.getRange(rowIdx, 16).setValue('');
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}
function approveResubmitRequest(rowIdx) {
  try {
    const sheet = _taskSs().getSheetByName('제출현황');
    if (!sheet) return { success: false };
    _setSubmissionStatus_(sheet, rowIdx, '재제출요청'); // K열 + X열
    sheet.getRange(rowIdx, 16).setValue('');           // P열 = 답글 초기화
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}

const SHEET_ID = "1jK7gYGFXCe3FULLs5mKttP959Aa9vp8-WNOGdJy7cZQ";

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('진도표 관리')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function _getSysSetting(key) {
  try {
    var sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName('시스템설정');
    if (!sh || sh.getLastRow() < 2) return '';
    var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === key) return String(rows[i][1] || '').trim();
    }
    return '';
  } catch(e) { return ''; }
}

function verifyTeacher(password) {
  try {
    // 시스템설정 '교사비밀번호' 키 우선, 없으면 학생명부 F2 폴백
    var pw = _getSysSetting('교사비밀번호');
    if (!pw) {
      var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('학생명부');
      if (sheet) pw = String(sheet.getRange('F2').getValue() || '').trim();
    }
    if (!pw) return { success: false, message: '비밀번호가 설정되어 있지 않습니다. 시스템설정 시트를 확인하세요.' };
    if (String(password).trim() === pw) return { success: true };
    return { success: false, message: '비밀번호가 틀렸습니다.' };
  } catch(e) { return { success: false, message: e.toString() }; }
}

function getClassList() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const rosterSheet = ss.getSheetByName('학생명부');
    if (!rosterSheet || rosterSheet.getLastRow() < 2) return { success: true, classes: [] };
    const data = rosterSheet.getRange(2, 1, rosterSheet.getLastRow() - 1, 3).getValues();
    const clsSet = {};
    data.forEach(function(r) {
      const id = String(r[1] || '').trim();
      if (id.length >= 2) clsSet[id.substring(0,1) + '학년 ' + id.substring(1,2) + '반'] = true;
    });
    return { success: true, classes: Object.keys(clsSet).sort() };
  } catch(e) { return { success: false, classes: [], message: e.toString() }; }
}

// =====================================================
// 헬퍼
// =====================================================

// 그룹ID에 따라 시트명 반환. 기본 그룹은 원래 이름 그대로.
function _sn(base, groupId) {
  return (!groupId || groupId === '기본') ? base : base + '_' + groupId;
}

// 시트 없으면 생성 후 반환
function _ensureSh(ss, name, headers, bg) {
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (headers) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers])
        .setFontWeight('bold').setBackground(bg || '#1e3a8a').setFontColor('white');
    }
  }
  return sh;
}

// =====================================================
// 커리큘럼 그룹 관리
// 시트: 커리큘럼그룹 — A=그룹ID, B=그룹명, C=반목록(쉼표), D=생성일
// =====================================================

function getGroups() {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sh = _ensureSh(ss, '커리큘럼그룹', ['그룹ID','그룹명','반목록','생성일','계획ID'], '#7c3aed');
    var groups = [{ id: '기본', name: '기본', classes: _getDefaultClasses_(ss), isDefault: true, planId: '기본' }];
    if (sh.getLastRow() >= 2) {
      var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues();
      rows.forEach(function(r) {
        if (!r[0]) return;
        var cls = r[2] ? String(r[2]).split(',').map(function(c){return c.trim();}).filter(Boolean) : [];
        var pid = String(r[4]||'').trim() || String(r[0]).trim();
        groups.push({ id: String(r[0]).trim(), name: String(r[1]||r[0]).trim(), classes: cls, planId: pid });
      });
    }
    return { success: true, groups: groups };
  } catch(e) { return { success: false, groups: [], message: e.toString() }; }
}

function _getDefaultClasses_(ss) {
  var clsSh = ss.getSheetByName('진도반설정');
  if (clsSh && clsSh.getLastRow() >= 2) {
    var cv = clsSh.getRange(2, 1).getValue();
    if (cv) return String(cv).split(',').map(function(c){return c.trim();}).filter(Boolean);
  }
  return [];
}

// 그룹이 사용하는 계획ID 반환 (E열이 비어있으면 그룹ID 자체를 fallback으로 사용)
function _getGroupPlanId_(gid, ss) {
  if (gid === '기본') return '기본';
  var sh = ss.getSheetByName('커리큘럼그룹');
  if (!sh || sh.getLastRow() < 2) return gid;
  var rows = sh.getRange(2, 1, sh.getLastRow()-1, 5).getValues();
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === gid) {
      return String(rows[i][4]||'').trim() || gid;
    }
  }
  return gid;
}

// =====================================================
// 진도 계획 관리
// 시트: 진도계획목록 — A=계획ID, B=계획명, C=생성일
// =====================================================

function getPlanList() {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sh = _ensureSh(ss, '진도계획목록', ['계획ID','계획명','생성일'], '#1e3a8a');
    var plans = [{ id: '기본', name: '기본 계획', isDefault: true }];
    if (sh.getLastRow() >= 2) {
      var rows = sh.getRange(2, 1, sh.getLastRow()-1, 3).getValues();
      rows.forEach(function(r) {
        if (!r[0]) return;
        plans.push({ id: String(r[0]).trim(), name: String(r[1]||r[0]).trim() });
      });
    }
    return { success: true, plans: plans };
  } catch(e) { return { success: false, plans: [], message: e.toString() }; }
}

function createPlan(data) {
  try {
    if (!data.name) return { success: false, message: '계획 이름을 입력하세요.' };
    var id = (data.id || data.name).replace(/[\s\/\\]/g,'').substring(0, 20);
    if (!id || id === '기본') return { success: false, message: '유효하지 않은 이름입니다.' };
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sh = _ensureSh(ss, '진도계획목록', ['계획ID','계획명','생성일'], '#1e3a8a');
    if (sh.getLastRow() >= 2) {
      var rows = sh.getRange(2, 1, sh.getLastRow()-1, 1).getValues();
      for (var i = 0; i < rows.length; i++) {
        if (String(rows[i][0]).trim() === id) return { success: false, message: '이미 존재하는 계획입니다.' };
      }
    }
    var today = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');
    sh.appendRow([id, data.name, today]);
    _ensureSh(ss, '진도계획_' + id, ['차시','계획내용','수업자료URL','메모'], '#1e3a8a');
    return { success: true, id: id };
  } catch(e) { return { success: false, message: e.toString() }; }
}

function renamePlan(planId, newName) {
  try {
    if (!planId || planId === '기본') return { success: false, message: '기본 계획은 이름을 변경할 수 없습니다.' };
    if (!newName) return { success: false, message: '이름을 입력하세요.' };
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sh = ss.getSheetByName('진도계획목록');
    if (!sh || sh.getLastRow() < 2) return { success: false, message: '계획을 찾을 수 없습니다.' };
    var rows = sh.getRange(2, 1, sh.getLastRow()-1, 1).getValues();
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === planId) {
        sh.getRange(i+2, 2).setValue(newName);
        return { success: true };
      }
    }
    return { success: false, message: '계획을 찾을 수 없습니다.' };
  } catch(e) { return { success: false, message: e.toString() }; }
}

function deletePlan(planId) {
  try {
    if (!planId || planId === '기본') return { success: false, message: '기본 계획은 삭제할 수 없습니다.' };
    var ss = SpreadsheetApp.openById(SHEET_ID);
    // 사용 중인 그룹 확인
    var grpSh = ss.getSheetByName('커리큘럼그룹');
    if (grpSh && grpSh.getLastRow() >= 2) {
      var grpRows = grpSh.getRange(2, 1, grpSh.getLastRow()-1, 5).getValues();
      for (var i = 0; i < grpRows.length; i++) {
        if (String(grpRows[i][4]||'').trim() === planId) {
          return { success: false, message: '"' + planId + '" 계획을 사용하는 그룹이 있습니다.\n먼저 그룹에서 다른 계획을 선택하세요.' };
        }
      }
    }
    var sh = ss.getSheetByName('진도계획목록');
    if (sh && sh.getLastRow() >= 2) {
      var rows = sh.getRange(2, 1, sh.getLastRow()-1, 1).getValues();
      for (var j = rows.length-1; j >= 0; j--) {
        if (String(rows[j][0]).trim() === planId) sh.deleteRow(j+2);
      }
    }
    var planSh = ss.getSheetByName('진도계획_' + planId);
    if (planSh) ss.deleteSheet(planSh);
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}

function getSyllabusPlanById(planId) {
  try {
    var pid = planId || '기본';
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sh = _ensureSh(ss, _sn('진도계획', pid), ['차시','계획내용','수업자료URL','메모'], '#1e3a8a');
    var plans = [];
    if (sh.getLastRow() >= 2) {
      var pd = sh.getRange(2, 1, sh.getLastRow()-1, 4).getValues();
      pd.forEach(function(r) {
        if (!r[0]) return;
        plans.push({ lessonNo: parseInt(r[0])||0, content: String(r[1]||'').trim(), memo: String(r[3]||'').trim() });
      });
      plans.sort(function(a,b){ return a.lessonNo-b.lessonNo; });
    }
    return { success: true, plans: plans };
  } catch(e) { return { success: false, message: e.toString() }; }
}

function saveSyllabusPlanById(plans, planId) {
  try {
    var pid = planId || '기본';
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sh = _ensureSh(ss, _sn('진도계획', pid), ['차시','계획내용','수업자료URL','메모'], '#1e3a8a');
    if (sh.getLastRow() > 1) sh.deleteRows(2, sh.getLastRow()-1);
    var rows = plans.map(function(p){ return [p.lessonNo, p.content||'', '', p.memo||'']; });
    if (rows.length > 0) sh.getRange(2, 1, rows.length, 4).setValues(rows);
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}

function createGroup(data) {
  try {
    if (!data.name) return { success: false, message: '그룹명을 입력하세요.' };
    // ID: 이름에서 공백 제거한 값 (최대 20자)
    var id = (data.id || data.name).replace(/\s/g, '').substring(0, 20);
    if (!id) return { success: false, message: '유효하지 않은 그룹명입니다.' };

    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sh = _ensureSh(ss, '커리큘럼그룹', ['그룹ID','그룹명','반목록','생성일'], '#7c3aed');

    // 중복 확인
    if (sh.getLastRow() >= 2) {
      var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
      for (var i = 0; i < rows.length; i++) {
        if (String(rows[i][0]).trim() === id) return { success: false, message: '이미 존재하는 그룹입니다.' };
      }
    }

    var today = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');
    var planId = String(data.planId || '기본').trim();
    sh.appendRow([id, data.name, (data.classes || []).join(','), today, planId]);

    _ensureSh(ss, '진도체크_' + id, ['날짜','반','실제차시','메모','상태','예상차시'], '#10b981');

    return { success: true, id: id };
  } catch(e) { return { success: false, message: e.toString() }; }
}

function updateGroup(data) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sh = ss.getSheetByName('커리큘럼그룹');
    if (!sh || sh.getLastRow() < 2) return { success: false, message: '그룹을 찾을 수 없습니다.' };
    var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues();
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === String(data.id).trim()) {
        sh.getRange(i + 2, 2, 1, 2).setValues([[data.name, (data.classes || []).join(',')]]);
        if (data.planId !== undefined) sh.getRange(i + 2, 5).setValue(data.planId || '기본');
        return { success: true };
      }
    }
    return { success: false, message: '그룹을 찾을 수 없습니다.' };
  } catch(e) { return { success: false, message: e.toString() }; }
}

function deleteGroup(id) {
  try {
    if (!id || id === '기본') return { success: false, message: '기본 그룹은 삭제할 수 없습니다.' };
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var planId = _getGroupPlanId_(id, ss);
    var sh = ss.getSheetByName('커리큘럼그룹');
    if (sh && sh.getLastRow() >= 2) {
      var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
      for (var i = rows.length - 1; i >= 0; i--) {
        if (String(rows[i][0]).trim() === String(id).trim()) sh.deleteRow(i + 2);
      }
    }
    var ckSh = ss.getSheetByName('진도체크_' + id);
    if (ckSh) ss.deleteSheet(ckSh);
    // 진도계획 시트는 그룹 자체 계획인 경우만 삭제 (공유 계획은 유지)
    if (planId === id) {
      var ps = ss.getSheetByName('진도계획_' + id);
      if (ps) ss.deleteSheet(ps);
    }
    // 시간표·학기 설정 행 삭제
    _deleteRowsByGroupId_(ss, '시간표설정', id);
    _deleteRowsByGroupId_(ss, '학기설정', id);
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}

function _deleteRowsByGroupId_(ss, sheetName, id) {
  var sh = ss.getSheetByName(sheetName);
  if (!sh || sh.getLastRow() < 2) return;
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  for (var i = rows.length - 1; i >= 0; i--) {
    if (String(rows[i][0]).trim() === String(id).trim()) sh.deleteRow(i + 2);
  }
}

// =====================================================
// 학기 설정
// 시트: 학기설정 — A=그룹ID, B=학기명, C=시작일, D=종료일, E=공휴일(쉼표)
// =====================================================

function getSemesterList(groupId) {
  try {
    var gid = groupId || '기본';
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sh = _ensureSh(ss, '학기설정', ['그룹ID','학기명','시작일','종료일','공휴일(쉼표구분)'], '#0284c7');
    var semesters = [];
    if (sh.getLastRow() >= 2) {
      var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues();
      rows.forEach(function(r) {
        if (String(r[0]).trim() !== gid) return;
        semesters.push({
          name:     String(r[1] || '').trim(),
          start:    r[2] ? Utilities.formatDate(new Date(r[2]), 'Asia/Seoul', 'yyyy-MM-dd') : '',
          end:      r[3] ? Utilities.formatDate(new Date(r[3]), 'Asia/Seoul', 'yyyy-MM-dd') : '',
          holidays: r[4] ? String(r[4]).split(',').map(function(d){return d.trim();}).filter(Boolean) : []
        });
      });
    }
    return { success: true, semesters: semesters };
  } catch(e) { return { success: false, semesters: [], message: e.toString() }; }
}

function saveSemester(groupId, data) {
  try {
    var gid = groupId || '기본';
    if (!data.name) return { success: false, message: '학기명을 입력하세요.' };
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sh = _ensureSh(ss, '학기설정', ['그룹ID','학기명','시작일','종료일','공휴일(쉼표구분)'], '#0284c7');
    var found = false;
    if (sh.getLastRow() >= 2) {
      var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues();
      for (var i = 0; i < rows.length; i++) {
        if (String(rows[i][0]).trim() === gid && String(rows[i][1]).trim() === data.name) {
          sh.getRange(i + 2, 1, 1, 5).setValues([[gid, data.name, data.start, data.end, (data.holidays||[]).join(',')]]);
          found = true; break;
        }
      }
    }
    if (!found) sh.appendRow([gid, data.name, data.start, data.end, (data.holidays||[]).join(',')]);
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}

function deleteSemester(groupId, semesterName) {
  try {
    var gid = groupId || '기본';
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sh = ss.getSheetByName('학기설정');
    if (!sh || sh.getLastRow() < 2) return { success: true };
    var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
    for (var i = rows.length - 1; i >= 0; i--) {
      if (String(rows[i][0]).trim() === gid && String(rows[i][1]).trim() === semesterName) sh.deleteRow(i + 2);
    }
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}

// =====================================================
// 시간표 설정
// 시트: 시간표설정 — A=그룹ID, B=반, C=학기명, D=월, E=화, F=수, G=목, H=금
// (교시번호, 0=없음)
// =====================================================

function getTimetableList(groupId) {
  try {
    var gid = groupId || '기본';
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sh = _ensureSh(ss, '시간표설정', ['그룹ID','반','학기명','월','화','수','목','금'], '#0f172a');
    var entries = [];
    if (sh.getLastRow() >= 2) {
      var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 8).getValues();
      rows.forEach(function(r) {
        if (String(r[0]).trim() !== gid) return;
        entries.push({
          cls:      String(r[1] || '').trim(),
          semester: String(r[2] || '').trim(),
          mon: parseInt(r[3]) || 0,
          tue: parseInt(r[4]) || 0,
          wed: parseInt(r[5]) || 0,
          thu: parseInt(r[6]) || 0,
          fri: parseInt(r[7]) || 0
        });
      });
    }
    return { success: true, entries: entries };
  } catch(e) { return { success: false, entries: [], message: e.toString() }; }
}

function saveTimetable(groupId, data) {
  // data: { cls, semester, mon, tue, wed, thu, fri }
  try {
    var gid = groupId || '기본';
    if (!data.cls || !data.semester) return { success: false, message: '반과 학기를 선택하세요.' };
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sh = _ensureSh(ss, '시간표설정', ['그룹ID','반','학기명','월','화','수','목','금'], '#0f172a');
    var found = false;
    if (sh.getLastRow() >= 2) {
      var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues();
      for (var i = 0; i < rows.length; i++) {
        if (String(rows[i][0]).trim() === gid && String(rows[i][1]).trim() === data.cls && String(rows[i][2]).trim() === data.semester) {
          sh.getRange(i + 2, 1, 1, 8).setValues([[gid, data.cls, data.semester, data.mon||0, data.tue||0, data.wed||0, data.thu||0, data.fri||0]]);
          found = true; break;
        }
      }
    }
    if (!found) sh.appendRow([gid, data.cls, data.semester, data.mon||0, data.tue||0, data.wed||0, data.thu||0, data.fri||0]);
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}

function deleteTimetable(groupId, cls, semester) {
  try {
    var gid = groupId || '기본';
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sh = ss.getSheetByName('시간표설정');
    if (!sh || sh.getLastRow() < 2) return { success: true };
    var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues();
    for (var i = rows.length - 1; i >= 0; i--) {
      if (String(rows[i][0]).trim() === gid && String(rows[i][1]).trim() === cls && String(rows[i][2]).trim() === semester) sh.deleteRow(i + 2);
    }
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}

// =====================================================
// 자동 일정 생성
// =====================================================

// 내부 헬퍼: 시트에 쓰지 않고 { cls: ['yyyy-MM-dd', ...] } 반환
function _computeSchedule_(gid, ss) {
  var semSh = ss.getSheetByName('학기설정');
  var ttSh  = ss.getSheetByName('시간표설정');
  var sems  = [];
  if (semSh && semSh.getLastRow() >= 2) {
    var sr = semSh.getRange(2, 1, semSh.getLastRow()-1, 5).getValues();
    sr.forEach(function(r) {
      if (String(r[0]).trim() !== gid) return;
      if (!r[2] || !r[3]) return;
      var hols = {};
      if (r[4]) String(r[4]).split(',').forEach(function(d){ if (d.trim()) hols[d.trim()] = true; });
      sems.push({
        name:     String(r[1]||'').trim(),
        start:    Utilities.formatDate(new Date(r[2]), 'Asia/Seoul', 'yyyy-MM-dd'),
        end:      Utilities.formatDate(new Date(r[3]), 'Asia/Seoul', 'yyyy-MM-dd'),
        holidays: hols
      });
    });
  }
  var ttMap = {};
  if (ttSh && ttSh.getLastRow() >= 2) {
    var tr = ttSh.getRange(2, 1, ttSh.getLastRow()-1, 8).getValues();
    tr.forEach(function(r) {
      if (String(r[0]).trim() !== gid) return;
      var k = String(r[1]||'').trim() + '|' + String(r[2]||'').trim();
      ttMap[k] = { cls: String(r[1]||'').trim(), semester: String(r[2]||'').trim(),
        mon: parseInt(r[3])||0, tue: parseInt(r[4])||0, wed: parseInt(r[5])||0,
        thu: parseInt(r[6])||0, fri: parseInt(r[7])||0 };
    });
  }
  var DAY_MAP = { mon:1, tue:2, wed:3, thu:4, fri:5 };
  var schedule = {};
  sems.forEach(function(sem) {
    Object.keys(ttMap).forEach(function(k) {
      var entry = ttMap[k];
      if (entry.semester !== sem.name) return;
      var dow = [];
      Object.keys(DAY_MAP).forEach(function(d){ if (entry[d]>0) dow.push(DAY_MAP[d]); });
      if (dow.length === 0) return;
      var cur = new Date(sem.start + 'T09:00:00');
      var end = new Date(sem.end   + 'T09:00:00');
      while (cur <= end) {
        var ds = Utilities.formatDate(cur, 'Asia/Seoul', 'yyyy-MM-dd');
        if (dow.indexOf(cur.getDay()) >= 0 && !sem.holidays[ds]) {
          if (!schedule[entry.cls]) schedule[entry.cls] = {};
          schedule[entry.cls][ds] = true;
        }
        cur.setDate(cur.getDate()+1);
      }
    });
  });
  var result = {};
  Object.keys(schedule).forEach(function(cls){
    result[cls] = Object.keys(schedule[cls]).sort();
  });
  return result;
}

function generateSchedule(groupId, semesterName) {
  try {
    var gid = groupId || '기본';

    // 학기 설정 조회
    var semList = getSemesterList(gid);
    if (!semList.success) return { success: false, message: '학기 설정 조회 실패' };
    var sem = null;
    semList.semesters.forEach(function(s) { if (s.name === semesterName) sem = s; });
    if (!sem) return { success: false, message: '"' + semesterName + '" 학기 설정이 없습니다.' };
    if (!sem.start || !sem.end) return { success: false, message: '학기 시작일/종료일을 입력하세요.' };

    // 시간표 조회
    var ttList = getTimetableList(gid);
    if (!ttList.success) return { success: false, message: '시간표 조회 실패' };
    var ttEntries = ttList.entries.filter(function(e) { return e.semester === semesterName; });
    if (ttEntries.length === 0) return { success: false, message: '"' + semesterName + '" 학기 시간표가 없습니다.\n시간표 탭에서 먼저 등록해주세요.' };

    // 공휴일 Set
    var holidaySet = {};
    (sem.holidays || []).forEach(function(d) { holidaySet[d] = true; });

    // 요일 → getDay() 매핑 (1=월, 2=화, 3=수, 4=목, 5=금)
    var DAY_MAP = { mon: 1, tue: 2, wed: 3, thu: 4, fri: 5 };
    var DAY_KR  = { mon: '월', tue: '화', wed: '수', thu: '목', fri: '금' };

    var startDt = new Date(sem.start + 'T09:00:00');
    var endDt   = new Date(sem.end   + 'T09:00:00');

    var allItems  = []; // { lessonNo, cls, plannedDate }
    var preview   = {}; // { cls: [{ lessonNo, date, day }] }
    var warnings  = [];

    ttEntries.forEach(function(entry) {
      var cls = entry.cls;
      // 이 반의 수업 요일 목록 (getDay 값)
      var scheduledDow = [];
      Object.keys(DAY_MAP).forEach(function(k) {
        if (entry[k] > 0) scheduledDow.push(DAY_MAP[k]);
      });
      if (scheduledDow.length === 0) {
        warnings.push(cls + ': 시간표에 수업 요일이 없습니다.');
        return;
      }

      var lessonDates = [];
      var cur = new Date(startDt);
      while (cur <= endDt) {
        var dow = cur.getDay();
        var dateStr = Utilities.formatDate(cur, 'Asia/Seoul', 'yyyy-MM-dd');
        if (scheduledDow.indexOf(dow) >= 0 && !holidaySet[dateStr]) {
          lessonDates.push(dateStr);
        }
        cur.setDate(cur.getDate() + 1);
      }

      if (lessonDates.length === 0) {
        warnings.push(cls + ': 학기 기간 내 수업일이 없습니다.');
        return;
      }

      preview[cls] = [];
      lessonDates.forEach(function(dt, idx) {
        var ln = idx + 1;
        var d = new Date(dt + 'T09:00:00');
        var dowK = ['일','월','화','수','목','금','토'][d.getDay()];
        allItems.push({ lessonNo: ln, cls: cls, plannedDate: dt });
        preview[cls].push({ lessonNo: ln, date: dt, day: dowK });
      });
    });

    if (allItems.length === 0) return { success: false, message: '생성된 일정이 없습니다.\n' + warnings.join('\n') };

    return { success: true, items: allItems, preview: preview, total: allItems.length, warnings: warnings };
  } catch(e) { return { success: false, message: e.toString() }; }
}


function generateAndApplySchedule(groupId, semData, ttEntries) {
  try {
    var gid = groupId || '기본';
    if (!semData || !semData.name) return { success: false, message: '학기명을 입력하세요.' };
    if (!semData.start || !semData.end) return { success: false, message: '시작일과 종료일을 입력하세요.' };
    var entries = (ttEntries || []).filter(function(e) {
      return e.cls && (e.mon||e.tue||e.wed||e.thu||e.fri);
    });
    if (entries.length === 0) return { success: false, message: '수업 요일이 있는 반을 최소 1개 입력하세요.' };

    // 학기 저장
    var r1 = saveSemester(gid, semData);
    if (!r1.success) return r1;

    // 시간표 저장
    for (var i = 0; i < entries.length; i++) {
      var r2 = saveTimetable(gid, {
        cls: entries[i].cls, semester: semData.name,
        mon: entries[i].mon||0, tue: entries[i].tue||0,
        wed: entries[i].wed||0, thu: entries[i].thu||0, fri: entries[i].fri||0
      });
      if (!r2.success) return r2;
    }

    // 일정 미리보기 계산 (진도체크 시트에 쓰지 않음 — 일정은 매번 자동 계산됨)
    var gen = generateSchedule(gid, semData.name);
    if (!gen.success) return gen;

    var summary = {};
    Object.keys(gen.preview).forEach(function(cls) { summary[cls] = gen.preview[cls].length; });
    return { success: true, count: gen.items.length, summary: summary, warnings: gen.warnings || [] };
  } catch(e) { return { success: false, message: e.toString() }; }
}

// =====================================================
// 진도표 데이터 (그룹 지원)
// =====================================================

function getSyllabusData(groupId) {
  try {
    var gid = groupId || '기본';
    var ss = SpreadsheetApp.openById(SHEET_ID);

    // 진도계획 (그룹이 참조하는 계획 시트 사용)
    var planId = _getGroupPlanId_(gid, ss);
    var planSh = _ensureSh(ss, _sn('진도계획', planId), ['차시','계획내용','수업자료URL','메모'], '#1e3a8a');
    var plans = [];
    if (planSh.getLastRow() >= 2) {
      var pd = planSh.getRange(2, 1, planSh.getLastRow() - 1, 4).getValues();
      pd.forEach(function(r) {
        if (!r[0]) return;
        plans.push({ lessonNo: parseInt(r[0])||0, content: String(r[1]||'').trim(), material: '', memo: String(r[3]||'').trim() });
      });
      plans.sort(function(a,b){ return a.lessonNo - b.lessonNo; });
    }

    // 진도체크 (스키마: 날짜, 반, 실제차시, 메모, 상태, 예상차시)
    // key = 'yyyy-MM-dd_반이름'
    // status = '' | '취소[:사유]' | '이동:targetDate' | '받음:sourceDate'(이동받은 수업)
    var checkSh = _ensureSh(ss, _sn('진도체크', gid), ['날짜','반','실제차시','메모','상태','예상차시'], '#10b981');
    var checks = {};
    var movedIn = {}; // key → [{ lessonNo, sourceDate, memo }]
    if (checkSh.getLastRow() >= 2) {
      var cd = checkSh.getRange(2, 1, checkSh.getLastRow() - 1, 6).getValues();
      cd.forEach(function(r) {
        if (!r[0] || !r[1]) return;
        var dateStr = r[0] instanceof Date
          ? Utilities.formatDate(r[0], 'Asia/Seoul', 'yyyy-MM-dd')
          : String(r[0]).trim();
        if (!dateStr) return;
        var status = String(r[4]||'').trim();
        var ln  = parseInt(r[2]) || 0;
        var pln = parseInt(r[5]) || 0;
        var key = dateStr + '_' + String(r[1]).trim();
        // 이동받은 수업 (다른 날에서 이 날로 이동됨)
        if (status.indexOf('받음:') === 0) {
          var srcDate = status.substring(3);
          if (!movedIn[key]) movedIn[key] = [];
          movedIn[key].push({ lessonNo: ln, sourceDate: srcDate, memo: String(r[3]||'').trim() });
          return;
        }
        if (!ln && !pln && status.indexOf('취소') !== 0 && status.indexOf('이동:') !== 0) return;
        checks[key] = { lessonNo: ln, plannedLessonNo: pln, memo: String(r[3]||'').trim(), status: status };
      });
    }

    // 반 목록
    var classes = [];
    if (gid === '기본') {
      classes = _getDefaultClasses_(ss);
      if (classes.length === 0) {
        var res = getClassList();
        classes = res.classes || [];
      }
    } else {
      var grpSh = ss.getSheetByName('커리큘럼그룹');
      if (grpSh && grpSh.getLastRow() >= 2) {
        var grpRows = grpSh.getRange(2, 1, grpSh.getLastRow() - 1, 3).getValues();
        for (var gi = 0; gi < grpRows.length; gi++) {
          if (String(grpRows[gi][0]).trim() === gid) {
            var cv = String(grpRows[gi][2] || '');
            if (cv) classes = cv.split(',').map(function(c){return c.trim();}).filter(Boolean);
            break;
          }
        }
      }
    }

    // 수업 일정 계산 (시트 기록 없이 시간표+학기 설정으로 매번 계산)
    var schedule = _computeSchedule_(gid, ss);

    // 공휴일 날짜 수집 (학기설정 E열)
    var holidayDates = {};
    var holSemSh = ss.getSheetByName('학기설정');
    if (holSemSh && holSemSh.getLastRow() >= 2) {
      var holRows = holSemSh.getRange(2, 1, holSemSh.getLastRow() - 1, 5).getValues();
      holRows.forEach(function(r) {
        if (String(r[0]).trim() !== gid || !r[4]) return;
        String(r[4]).split(',').forEach(function(d) {
          var dt = d.trim();
          if (dt) holidayDates[dt] = true;
        });
      });
    }

    return { success: true, plans: plans, checks: checks, movedIn: movedIn, classes: classes, schedule: schedule, holidays: holidayDates };
  } catch(e) { return { success: false, message: e.toString() }; }
}

function saveSyllabusPlan(plans, groupId) {
  try {
    var gid = groupId || '기본';
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var planId = _getGroupPlanId_(gid, ss);
    var sh = _ensureSh(ss, _sn('진도계획', planId), ['차시','계획내용','수업자료URL','메모'], '#1e3a8a');
    if (sh.getLastRow() > 1) sh.deleteRows(2, sh.getLastRow() - 1);
    var rows = plans.map(function(p){ return [p.lessonNo, p.content||'', '', p.memo||'']; });
    if (rows.length > 0) sh.getRange(2, 1, rows.length, 4).setValues(rows);
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}

// data = { date: 'yyyy-MM-dd', cls, lessonNo, plannedLessonNo, memo, status }
// status: '' = 정상, '취소' = 수업 없음, '이동:yyyy-MM-dd' = 날짜 이동
function saveSyllabusCheck(data, groupId) {
  try {
    var gid = groupId || '기본';
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sh = _ensureSh(ss, _sn('진도체크', gid), ['날짜','반','실제차시','메모','상태','예상차시'], '#10b981');
    var status = String(data.status||'').trim();
    var ln  = data.lessonNo || 0;
    var pln = data.plannedLessonNo || 0;
    var found = false;
    if (sh.getLastRow() >= 2) {
      var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
      for (var i = 0; i < rows.length; i++) {
        var rowDate = rows[i][0] instanceof Date
          ? Utilities.formatDate(rows[i][0], 'Asia/Seoul', 'yyyy-MM-dd')
          : String(rows[i][0]).trim();
        if (rowDate === String(data.date).trim() && String(rows[i][1]).trim() === String(data.cls).trim()) {
          sh.getRange(i + 2, 1, 1, 6).setValues([[data.date, data.cls, ln, data.memo||'', status, pln]]);
          found = true; break;
        }
      }
    }
    if (!found) sh.appendRow([data.date, data.cls, ln, data.memo||'', status, pln]);
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}

function deleteSyllabusCheck(date, cls, groupId) {
  try {
    var gid = groupId || '기본';
    var sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(_sn('진도체크', gid));
    if (!sh || sh.getLastRow() < 2) return { success: true };
    var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
    for (var i = rows.length - 1; i >= 0; i--) {
      var rowDate = rows[i][0] instanceof Date
        ? Utilities.formatDate(rows[i][0], 'Asia/Seoul', 'yyyy-MM-dd')
        : String(rows[i][0]).trim();
      if (rowDate === String(date).trim() && String(rows[i][1]).trim() === String(cls).trim()) sh.deleteRow(i + 2);
    }
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}

function saveSylClassList(classes, groupId) {
  try {
    var gid = groupId || '기본';
    var ss = SpreadsheetApp.openById(SHEET_ID);
    if (gid === '기본') {
      var sh = _ensureSh(ss, '진도반설정', ['반목록'], '#334155');
      sh.getRange(2, 1).setValue(classes.join(','));
    } else {
      updateGroup({ id: gid, name: gid, classes: classes }); // classes만 업데이트
    }
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}

// 이동받은 수업 저장 — status='받음:sourceDate' 행 추가/갱신
// data: { date, cls, lessonNo, sourceDate, memo }
function saveMovedInLesson(data, groupId) {
  try {
    var gid = groupId || '기본';
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sh = _ensureSh(ss, _sn('진도체크', gid), ['날짜','반','실제차시','메모','상태','예상차시'], '#10b981');
    var status = '받음:' + String(data.sourceDate || '');
    var ln = data.lessonNo || 0;
    var found = false;
    if (sh.getLastRow() >= 2) {
      var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues();
      for (var i = 0; i < rows.length; i++) {
        var rowDate = rows[i][0] instanceof Date
          ? Utilities.formatDate(rows[i][0], 'Asia/Seoul', 'yyyy-MM-dd')
          : String(rows[i][0]).trim();
        if (rowDate === String(data.date).trim() && String(rows[i][1]).trim() === String(data.cls).trim()
            && String(rows[i][4]||'').trim() === status) {
          sh.getRange(i + 2, 1, 1, 6).setValues([[data.date, data.cls, ln, data.memo||'', status, 0]]);
          found = true; break;
        }
      }
    }
    if (!found) sh.appendRow([data.date, data.cls, ln, data.memo||'', status, 0]);
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}

// 이동받은 수업 삭제 — status='받음:sourceDate' 행 제거
function deleteMovedInLesson(date, cls, sourceDate, groupId) {
  try {
    var gid = groupId || '기본';
    var sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(_sn('진도체크', gid));
    if (!sh || sh.getLastRow() < 2) return { success: true };
    var status = '받음:' + String(sourceDate);
    var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues();
    for (var i = rows.length - 1; i >= 0; i--) {
      var rowDate = rows[i][0] instanceof Date
        ? Utilities.formatDate(rows[i][0], 'Asia/Seoul', 'yyyy-MM-dd')
        : String(rows[i][0]).trim();
      if (rowDate === String(date).trim() && String(rows[i][1]).trim() === String(cls).trim()
          && String(rows[i][4]||'').trim() === status) {
        sh.deleteRow(i + 2);
      }
    }
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}


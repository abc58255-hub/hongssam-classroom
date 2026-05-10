const SHEET_ID = "1jK7gYGFXCe3FULLs5mKttP959Aa9vp8-WNOGdJy7cZQ";

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('진도표 관리')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function verifyTeacher(password) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('학생명부');
    if (!sheet) return { success: false, message: '시트를 찾을 수 없습니다.' };
    const pw = String(sheet.getRange('F2').getValue() || '').trim();
    if (!pw) return { success: false, message: '비밀번호가 설정되어 있지 않습니다.' };
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
// 진도표 관리 v2
// 시트: 진도계획   — A=차시, B=내용, C=자료URL, D=메모
// 시트: 진도체크   — A=차시, B=반, C=날짜, D=메모, E=예상날짜
// 시트: 진도반설정 — A=반목록(쉼표구분)
// 시트: 시험범위   — A=시험명, B=차시목록(쉼표구분)
// =====================================================

function getSyllabusData() {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);

    // 진도계획
    var planSheet = ss.getSheetByName('진도계획');
    var plans = [];
    if (!planSheet) {
      planSheet = ss.insertSheet('진도계획');
      planSheet.getRange(1,1,1,4).setValues([['차시','계획내용','수업자료URL','메모']]).setFontWeight('bold').setBackground('#1e3a8a').setFontColor('white');
    } else if (planSheet.getLastRow() >= 2) {
      var pd = planSheet.getRange(2,1,planSheet.getLastRow()-1,4).getValues();
      pd.forEach(function(r) {
        if (!r[0]) return;
        plans.push({ lessonNo: parseInt(r[0])||0, content: String(r[1]||'').trim(), material: '', memo: String(r[2]||'').trim() });
      });
      plans.sort(function(a,b){ return a.lessonNo - b.lessonNo; });
    }

    // 진도체크
    var checkSheet = ss.getSheetByName('진도체크');
    var checks = {};
    if (!checkSheet) {
      checkSheet = ss.insertSheet('진도체크');
      checkSheet.getRange(1,1,1,5).setValues([['차시','반','날짜','메모','예상날짜']]).setFontWeight('bold').setBackground('#10b981').setFontColor('white');
    } else if (checkSheet.getLastRow() >= 2) {
      var cd = checkSheet.getRange(2,1,checkSheet.getLastRow()-1,5).getValues();
      cd.forEach(function(r) {
        if (!r[0] || !r[1]) return;
        var dateStr = r[2] instanceof Date ? Utilities.formatDate(r[2],'Asia/Seoul','M/d') : String(r[2]||'').trim();
        var key = String(r[0]) + '_' + String(r[1]).trim();
        checks[key] = { date: dateStr, memo: String(r[3]||'').trim(), plannedDate: (r[4] instanceof Date ? Utilities.formatDate(r[4], 'Asia/Seoul', 'M/d') : String(r[4]||'').trim()) };
      });
    }

    // 반 설정
    var clsSheet = ss.getSheetByName('진도반설정');
    var classes = [];
    if (clsSheet && clsSheet.getLastRow() >= 2) {
      var cv = clsSheet.getRange(2,1).getValue();
      if (cv) classes = String(cv).split(',').map(function(c){return c.trim();}).filter(Boolean);
    }
    if (classes.length === 0) {
      var rosterSheet = ss.getSheetByName('학생명부');
      if (rosterSheet && rosterSheet.getLastRow() >= 2) {
        var rData = rosterSheet.getRange(2,1,rosterSheet.getLastRow()-1,3).getValues();
        var clsSetFallback = {};
        rData.forEach(function(r){
          var id = String(r[1]||'').trim();
          if (id.length >= 2) clsSetFallback[id.substring(0,1)+'학년 '+id.substring(1,2)+'반'] = true;
        });
        classes = Object.keys(clsSetFallback).sort();
      }
    }

    // 시험범위
    var examSheet = ss.getSheetByName('시험범위');
    var examRanges = [];
    if (!examSheet) {
      examSheet = ss.insertSheet('시험범위');
      examSheet.getRange(1,1,1,2).setValues([['시험명','차시목록']]).setFontWeight('bold').setBackground('#f59e0b').setFontColor('white');
    } else if (examSheet.getLastRow() >= 2) {
      var ed = examSheet.getRange(2,1,examSheet.getLastRow()-1,2).getValues();
      ed.forEach(function(r) {
        if (!r[0]) return;
        examRanges.push({ name: String(r[0]).trim(), units: String(r[1]||'').trim() });
      });
    }

    return { success: true, plans: plans, checks: checks, classes: classes, examRanges: examRanges };
  } catch(e) { return { success: false, message: e.toString() }; }
}

function saveSyllabusPlan(plans) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName('진도계획');
    if (!sheet) {
      sheet = ss.insertSheet('진도계획');
      sheet.getRange(1,1,1,3).setValues([['차시','계획내용','메모']]).setFontWeight('bold').setBackground('#1e3a8a').setFontColor('white');
    }
    if (sheet.getLastRow() > 1) sheet.deleteRows(2, sheet.getLastRow()-1);
    var rows = plans.map(function(p){ return [p.lessonNo, p.content||'', p.memo||'']; });
    if (rows.length > 0) sheet.getRange(2,1,rows.length,3).setValues(rows);
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}

function saveSyllabusCheck(data) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName('진도체크');
    if (!sheet) {
      sheet = ss.insertSheet('진도체크');
      sheet.getRange(1,1,1,5).setValues([['차시','반','날짜','메모','예상날짜']]).setFontWeight('bold').setBackground('#10b981').setFontColor('white');
    }
    var found = false;
    if (sheet.getLastRow() >= 2) {
      var rows = sheet.getRange(2,1,sheet.getLastRow()-1,5).getValues();
      for (var i=0; i<rows.length; i++) {
        if (parseInt(rows[i][0])===parseInt(data.lessonNo) && String(rows[i][1]).trim()===String(data.cls).trim()) {
          sheet.getRange(i+2,1,1,5).setValues([[data.lessonNo, data.cls, data.date||'', data.memo||'', data.plannedDate||'']]);
          found = true; break;
        }
      }
    }
    if (!found) sheet.appendRow([data.lessonNo, data.cls, data.date||'', data.memo||'', data.plannedDate||'']);
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}

function deleteSyllabusCheck(lessonNo, cls) {
  try {
    var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('진도체크');
    if (!sheet || sheet.getLastRow() < 2) return { success: true };
    var rows = sheet.getRange(2,1,sheet.getLastRow()-1,2).getValues();
    for (var i=rows.length-1; i>=0; i--) {
      if (parseInt(rows[i][0])===parseInt(lessonNo) && String(rows[i][1]).trim()===String(cls).trim()) {
        sheet.deleteRow(i+2);
      }
    }
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}

function saveSylClassList(classes) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName('진도반설정');
    if (!sheet) { sheet = ss.insertSheet('진도반설정'); sheet.getRange(1,1).setValue('반목록'); }
    sheet.getRange(2,1).setValue(classes.join(','));
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}

function saveExamRange(data) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName('시험범위');
    if (!sheet) {
      sheet = ss.insertSheet('시험범위');
      sheet.getRange(1,1,1,2).setValues([['시험명','차시목록']]).setFontWeight('bold').setBackground('#f59e0b').setFontColor('white');
    }
    sheet.appendRow([data.name, data.units]);
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}

function saveBulkPlannedDates(saveData) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName('진도체크');
    if (!sheet) {
      sheet = ss.insertSheet('진도체크');
      sheet.getRange(1,1,1,5).setValues([['차시','반','날짜','메모','예상날짜']]).setFontWeight('bold').setBackground('#10b981').setFontColor('white');
    }
    var existing = {};
    if (sheet.getLastRow() >= 2) {
      var rows = sheet.getRange(2,1,sheet.getLastRow()-1,5).getValues();
      rows.forEach(function(r,i) {
        if (r[0] && r[1]) existing[r[0]+'_'+String(r[1]).trim()] = i+2;
      });
    }
    saveData.forEach(function(d) {
      var key = String(d.lessonNo) + '_' + String(d.cls).trim();
      if (existing[key]) {
        var curDate = sheet.getRange(existing[key], 3).getValue();
        if (!curDate) sheet.getRange(existing[key], 5).setValue(d.plannedDate);
      } else {
        var last = sheet.getLastRow() + 1;
        sheet.getRange(last, 1, 1, 5).setValues([[d.lessonNo, d.cls, '', '', d.plannedDate]]);
        existing[key] = last;
      }
    });
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}

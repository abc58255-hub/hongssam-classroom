// 교사 포털 — 분리된 교사 앱들을 한 곳에서 접속하는 카드형 허브
// URL 소스: ClassCore 'getAppUrls()' (바로가기_* 키) 단일 소스
// 인증·설정은 ClassCore 라이브러리에 위임 — 하드코딩 없음

// ── 라우팅 ──
function doGet() {
  try { ClassCore.registerAppUrl('교사포털', _serviceUrl_()); } catch(_) {}
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('교사 포털')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
function _serviceUrl_() {
  try { return ScriptApp.getService().getUrl() || ''; } catch(_) { return ''; }
}

// ✅ GitHub Pages 프론트(PWA)용 RPC 엔드포인트 — 수학교실 PWA와 동일 패턴
// text/plain JSON {fn, args} → 화이트리스트 함수만 실행 → {ok, result|error}
// 아래 목록은 fcm-sw/portal/index.html 어댑터의 FNS와 반드시 일치해야 함
var RPC_WHITELIST = [
  'teacherLogin', 'teacherLogout', 'validateTeacherSession',
  'getPortalData', 'setAppUrl', 'getStudentList', 'resetStudentPassword',
  'getLessonGames', 'setLessonGame', 'lessonScores', 'lessonDeleteScore', 'lessonResetGame'
];

function doPost(e) {
  var out;
  try {
    var req = JSON.parse(e.postData.contents);
    if (RPC_WHITELIST.indexOf(req.fn) < 0) throw new Error('허용되지 않은 함수: ' + req.fn);
    var fn = globalThis[req.fn];
    if (typeof fn !== 'function') throw new Error('함수를 찾을 수 없음: ' + req.fn);
    out = { ok: true, result: fn.apply(null, req.args || []) };
  } catch (err) {
    out = { ok: false, error: String(err && err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

// ── 교사 인증 (ClassCore 위임) ──
function teacherLogin(pw)            { return ClassCore.teacherLogin(pw); }
function teacherLogout(token)        { return ClassCore.teacherLogout(token); }
function validateTeacherSession(token) { return ClassCore.verifyTeacher(token); }

// ── 포털 데이터: 카드 정의 + URL 매칭 ──
// 카드별 keys 는 ClassCore 레지스트리에서 찾을 후보 이름(앞 순서 우선).
// suffix 는 URL 뒤에 붙일 쿼리(같은 앱의 다른 페이지용).
function getPortalData() {
  var urls = {};
  try { urls = ClassCore.getAppUrls() || {}; } catch(_) {}

  function resolve(keys) {
    for (var i = 0; i < keys.length; i++) {
      if (urls[keys[i]]) return urls[keys[i]];
    }
    return '';
  }

  var groups = [
    { title: '수업·평가', cards: [
      { name: '과제채점',  icon: '📝', desc: '수학 과제 출제·채점·첨삭',     keys: ['과제채점'] },
      { name: '진도표',    icon: '📋', desc: '진도 계획·기록',              keys: ['진도표'] },
      { name: '수업활동',  icon: '🎮', desc: '단원 게임·자유 도전 공개 관리', keys: [], action: 'lessonAdmin' }
    ]},
    { title: '학급운영', cards: [
      { name: '알림관리',  icon: '🔔', desc: '학급알림 작성·푸시 알림 전송',    keys: ['알림관리'] },
      { name: '창체관리',  icon: '🎯', desc: '창의적 체험활동 관리',         keys: ['창체관리'] },
      { name: '설문관리',  icon: '📊', desc: '우리 반 설문 만들기·결과',      keys: ['설문관리', '설문'] },
      { name: '칭찬',      icon: '💝', desc: '칭찬 이벤트·배정·기록',        keys: ['칭찬'] },
      { name: '학급비',    icon: '💰', desc: '학급비 영수증·예산',          keys: ['업무도구'], suffix: '?page=budget' },
      { name: 'TV알림판',  icon: '📺', desc: '교실 TV 공지 화면',           keys: ['교실알림판', 'TV알림판', '교실TV알림판'] }
    ]},
    { title: '안전·기록', cards: [
      { name: '관찰·신고', icon: '🚨', desc: '학생 관찰 기록·신고 접수함',    keys: ['관찰신고'] }
    ]},
    { title: '도구', cards: [
      { name: 'AI업무분석', icon: '🤖', desc: '공문·메시지 AI 분석 → TickTick', keys: ['AI업무분석'] },
      { name: '도장입력',  icon: '🏅', desc: '학생 도장 부여',             keys: ['도장입력'] },
      { name: '도장조회',  icon: '🔍', desc: '도장 데이터 조회',            keys: ['도장데이터', '도장조회'] },
      { name: '방탈출',    icon: '🔓', desc: '방탈출 게임 만들기',          keys: ['업무도구'], suffix: '?page=escape' }
    ]},
    { title: '학생용 (학생에게 안내)', student: true, cards: [
      { name: '스마트포털', icon: '🎓', desc: '학생용 통합 포털(수학 교실 대문)', keys: ['스마트포털', '학생포털', '포털'] },
      { name: '수학교실',   icon: '➗', desc: '수학 과제 제출·피드백',       keys: ['수학교실'] },
      { name: '우리반교실', icon: '🏫', desc: '학급알림·창체·설문·내 도장',   keys: ['우리반교실'] },
      { name: '회원가입',   icon: '✏️', desc: '학생 회원가입',             keys: ['회원가입'] }
    ]}
  ];

  groups.forEach(function(g) {
    g.cards.forEach(function(c) {
      var base = resolve(c.keys);
      c.url = base ? (base + (c.suffix || '')) : '';
    });
  });

  return { success: true, groups: groups, homeroom: _cfg('담임반', '') };
}

function _cfg(k, d) { try { return ClassCore.getConfig(k, d); } catch(_) { return d; } }

// ── 🎮 수업활동(Supabase 게임) 관리 — sync-roster 함수를 SYNC_KEY로 호출 ──
var LESSON_SYNC_URL = 'https://lqjrrqhrnxctyrqccmch.supabase.co/functions/v1/sync-roster';
function _lessonCall_(payload) {
  var key = _cfg('수업활동동기화키', '');
  if (!key) return { success: false, message: '시스템설정에 수업활동동기화키가 없습니다' };
  try {
    var res = UrlFetchApp.fetch(LESSON_SYNC_URL, {
      method: 'post', contentType: 'application/json',
      headers: { 'X-Sync-Key': key },
      payload: JSON.stringify(payload), muteHttpExceptions: true
    });
    return JSON.parse(res.getContentText());
  } catch (e) { return { success: false, message: String(e) }; }
}
// ClassCore.verifyTeacher의 반환 형식이 버전에 따라 달라 관대하게 판정
// (true / {success} / {valid} / {ok} 모두 허용 — 형식 불일치로 "로그인 만료" 오탐났던 이력)
function _teacherOk_(token) {
  try {
    var v = ClassCore.verifyTeacher(token);
    return v === true || !!(v && (v.success || v.valid || v.ok));
  } catch (_) { return false; }
}
function getLessonGames(token) {
  if (!_teacherOk_(token)) return { success: false, message: '로그인이 만료됐어요. 새로고침 후 다시 로그인해주세요.' };
  return _lessonCall_({ op: 'listGames' });
}
// g = {unit, game, title, url, scope, status, max_score?}
function setLessonGame(token, g) {
  if (!_teacherOk_(token)) return { success: false, message: '로그인이 만료됐어요. 새로고침 후 다시 로그인해주세요.' };
  return _lessonCall_({ games: [g] });
}
// 기록 관리 — 목록·1건 삭제·전체 리셋
function lessonScores(token, unit, game) {
  if (!_teacherOk_(token)) return { success: false, message: '로그인이 만료됐어요' };
  return _lessonCall_({ op: 'listScores', unit: unit, game: game });
}
function lessonDeleteScore(token, id) {
  if (!_teacherOk_(token)) return { success: false, message: '로그인이 만료됐어요' };
  return _lessonCall_({ op: 'deleteScore', id: id });
}
function lessonResetGame(token, unit, game) {
  if (!_teacherOk_(token)) return { success: false, message: '로그인이 만료됐어요' };
  return _lessonCall_({ op: 'resetGame', unit: unit, game: game });
}

// 카드 URL 수동 등록(미연결 카드 보정용) — 이름은 카드 keys[0]
function setAppUrl(name, url) {
  try { return ClassCore.registerAppUrl(name, url); } catch(e) { return { success: false, message: e.toString() }; }
}

// ── 비밀번호 초기화 (ClassCore 위임) ──
function getStudentList() {
  try { return { success: true, students: ClassCore.getStudents() }; }
  catch(e) { return { success: false, message: e.toString(), students: [] }; }
}
// newPw 비우면 비번 삭제 → 학생이 다음 로그인 때 새 비밀번호 설정
function resetStudentPassword(studentId, newPw) {
  try { return ClassCore.resetStudentPw(studentId, newPw || ''); }
  catch(e) { return { success: false, message: e.toString() }; }
}

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
  'getLessonGames', 'setLessonGame', 'lessonScores', 'lessonDeleteScore', 'lessonResetGame',
  'getLinks', 'saveLink', 'deleteLink',
  'graphCreate', 'graphUpdate', 'graphClear', 'graphClose', 'graphClasses',
  'graphPresetSave', 'graphPresetDelete', 'syncRoster'
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

// ── 교사 인증 ──
// 비밀번호 검증은 ClassCore에 위임하되, 세션 토큰은 포털 자체 ScriptProperties로 관리.
// (ClassCore가 토큰을 CacheService에 저장하는데, 캐시 라운드트립이 불안정할 때
//  로그인 직후에도 "로그인 만료"가 뜨는 문제 회피 — 복습질문 앱과 동일 방식)
var PTOK_TTL_MS = 6 * 3600 * 1000; // 6시간
function _ptokOk_(token) {
  if (!token) return false;
  try {
    var v = PropertiesService.getScriptProperties().getProperty('ptok_' + String(token));
    if (!v) return false;
    if (Date.now() > parseInt(v)) { PropertiesService.getScriptProperties().deleteProperty('ptok_' + String(token)); return false; }
    return true;
  } catch (_) { return false; }
}
function teacherLogin(pw) {
  var r;
  try { r = ClassCore.teacherLogin(pw); } catch (e) { return { success: false, message: '인증 오류: ' + e.message }; }
  var ok = (r === true) || !!(r && (r.success || r.valid || r.ok || r.token));
  if (!ok) return (r && r.message) ? r : { success: false, message: '비밀번호가 올바르지 않습니다.' };
  var token = Utilities.getUuid();
  PropertiesService.getScriptProperties().setProperty('ptok_' + token, String(Date.now() + PTOK_TTL_MS));
  return { success: true, token: token };
}
function teacherLogout(token) {
  try { PropertiesService.getScriptProperties().deleteProperty('ptok_' + String(token)); } catch (_) {}
  try { ClassCore.teacherLogout(token); } catch (_) {}
  return { success: true };
}
function validateTeacherSession(token) { return { success: _ptokOk_(token) }; }

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
      { name: '과제채점',  icon: '📝', desc: '수학 과제 출제·채점·첨삭',     keys: ['과제채점'], url: 'https://abc58255-hub.github.io/hongssam-classroom/grade/' },
      { name: '진도표',    icon: '📋', desc: '진도 계획·기록',              keys: ['진도표'], url: 'https://abc58255-hub.github.io/hongssam-classroom/progress/' },
      { name: '수업활동',  icon: '🎮', desc: '단원 게임·자유 도전 공개 관리', keys: [], action: 'lessonAdmin' },
      { name: '복습 질문', icon: '🎲', desc: '반 뽑기 → 구두 복습 → 통과 도장(교사용)', keys: ['복습질문'], url: 'https://abc58255-hub.github.io/hongssam-classroom/review/' }
    ]},
    { title: '학급운영', cards: [
      { name: '알림관리',  icon: '🔔', desc: '학급알림 작성·푸시 알림 전송',    keys: ['알림관리'], url: 'https://abc58255-hub.github.io/hongssam-classroom/notify/' },
      { name: '창체관리',  icon: '🎯', desc: '창의적 체험활동 관리',         keys: ['창체관리'], url: 'https://abc58255-hub.github.io/hongssam-classroom/activity/' },
      { name: '설문관리',  icon: '📊', desc: '우리 반 설문 만들기·결과',      keys: ['설문관리', '설문'] },
      { name: '칭찬',      icon: '💝', desc: '칭찬 이벤트·배정·기록',        keys: ['칭찬'] },
      { name: '학급비',    icon: '💰', desc: '학급비 영수증·예산',          keys: ['업무도구'], suffix: '?page=budget' },
      { name: 'TV알림판',  icon: '📺', desc: '교실 TV 공지 화면',           keys: ['교실알림판', 'TV알림판', '교실TV알림판'], url: 'https://abc58255-hub.github.io/hongssam-classroom/tvboard/' }
    ]},
    { title: '안전·기록', cards: [
      { name: '관찰·신고', icon: '🚨', desc: '학생 관찰 기록·신고 접수함',    keys: ['관찰신고'], url: 'https://abc58255-hub.github.io/hongssam-classroom/observe/' }
    ]},
    { title: '도구', cards: [
      { name: 'AI업무분석', icon: '🤖', desc: '공문·메시지 AI 분석 → TickTick', keys: ['AI업무분석'], url: 'https://abc58255-hub.github.io/hongssam-classroom/ai/' },
      { name: '도장입력',  icon: '🏅', desc: '학생 도장 부여',             keys: ['도장입력'], url: 'https://abc58255-hub.github.io/hongssam-classroom/stamp/' },
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
      if (c.url) return; // 하드코딩 URL(예: 복습질문 별도 앱)은 유지
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
// 포털 자체 토큰(ScriptProperties)으로 검증 — ClassCore 캐시 라운드트립 불안정 회피
function _teacherOk_(token) { return _ptokOk_(token); }
function getLessonGames(token) {
  if (!_teacherOk_(token)) return { success: false, message: '로그인이 만료됐어요. 새로고침 후 다시 로그인해주세요.' };
  return _lessonCall_({ op: 'listGames' });
}
// g = {unit, game, title, url, scope, status, max_score?}
function setLessonGame(token, g) {
  if (!_teacherOk_(token)) return { success: false, message: '로그인이 만료됐어요. 새로고침 후 다시 로그인해주세요.' };
  return _lessonCall_({ games: [g] });
}
// ── 협력 직선그래프 탐구 ──
function graphCreate(token, room) {
  if (!_teacherOk_(token)) return { success: false, message: '로그인이 만료됐어요. 새로고침 후 다시 로그인해주세요.' };
  return _lessonCall_({ op: 'graphCreate', room: room });
}
function graphUpdate(token, id, patch) {
  if (!_teacherOk_(token)) return { success: false, message: '로그인이 만료됐어요.' };
  return _lessonCall_({ op: 'graphUpdate', id: id, patch: patch });
}
function graphClear(token, id) {
  if (!_teacherOk_(token)) return { success: false, message: '로그인이 만료됐어요.' };
  return _lessonCall_({ op: 'graphClear', id: id });
}
function graphClose(token, id) {
  if (!_teacherOk_(token)) return { success: false, message: '로그인이 만료됐어요.' };
  return _lessonCall_({ op: 'graphClose', id: id });
}
// 실제 학생이 있는 반 목록만 반환 — 학번(sid) 앞2자리로 유도해 학생앱 className과 형식 일치 보장
function graphClasses(token) {
  if (!_teacherOk_(token)) return { success: false, message: '로그인이 만료됐어요.' };
  var list;
  try { list = ClassCore.getStudents() || []; } catch (e) { return { success: false, message: String(e) }; }
  var set = {};
  list.forEach(function (s) {
    var sid = String((s && (s.id || s.sid || s.studentId || s['학번'] || s.number)) || '').trim();
    if (sid.length >= 2) set[sid.charAt(0) + '학년 ' + sid.charAt(1) + '반'] = 1;
  });
  var classes = Object.keys(set).sort(function (a, b) { return a.localeCompare(b, 'ko', { numeric: true }); });
  return { success: true, classes: classes };
}
// 그래프 직선 프리셋 저장/삭제 (읽기는 클라이언트가 anon으로 직접)
function graphPresetSave(token, preset) {
  if (!_teacherOk_(token)) return { success: false, message: '로그인이 만료됐어요.' };
  return _lessonCall_({ op: 'graphPresetSave', preset: preset });
}
function graphPresetDelete(token, id) {
  if (!_teacherOk_(token)) return { success: false, message: '로그인이 만료됐어요.' };
  return _lessonCall_({ op: 'graphPresetDelete', id: id });
}

// ── 수업 링크 보드 (학생 대시보드 상단 카드) ──
function getLinks(token) {
  if (!_teacherOk_(token)) return { success: false, message: '로그인이 만료됐어요. 새로고침 후 다시 로그인해주세요.' };
  return _lessonCall_({ op: 'listLinks' });
}
// link = {id?, title, url, class_scope, active, sort}
function saveLink(token, link) {
  if (!_teacherOk_(token)) return { success: false, message: '로그인이 만료됐어요. 새로고침 후 다시 로그인해주세요.' };
  return _lessonCall_({ op: 'saveLink', link: link });
}
function deleteLink(token, id) {
  if (!_teacherOk_(token)) return { success: false, message: '로그인이 만료됐어요' };
  return _lessonCall_({ op: 'deleteLink', id: id });
}

// 명단 즉시 동기화 — 학생-수학 GAS의 syncRoster(전체 로스터→Supabase students/Auth) 호출.
// 시트에서 학생을 추가·수정한 뒤 이걸 누르면 수업활동·그래프 로그인이 바로 반영됨.
var STUDENT_MATH_URL = 'https://script.google.com/macros/s/AKfycbyR1whn6f90-kJEAaJg_O34uP8v-KvyEqsRky58idjoxVDS5cWj80p2ScJp6V2dnz_0hA/exec';
function syncRoster(token) {
  if (!_teacherOk_(token)) return { success: false, message: '로그인이 만료됐어요.' };
  var key = _cfg('수업활동동기화키', '');
  if (!key) return { success: false, message: '시스템설정에 수업활동동기화키가 없습니다' };
  try {
    var res = UrlFetchApp.fetch(STUDENT_MATH_URL + '?action=syncRoster&key=' + encodeURIComponent(key), { muteHttpExceptions: true, followRedirects: true });
    return JSON.parse(res.getContentText());
  } catch (e) { return { success: false, message: String(e) }; }
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

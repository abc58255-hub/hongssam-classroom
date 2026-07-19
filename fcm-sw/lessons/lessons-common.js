// ═══════════════════════════════════════════════════════════════
// 수업활동 공용 모듈 (LessonHub)
// 단원 페이지에서 이 3줄만 넣으면 로그인·점수저장·등수판이 준비됨:
//   <script src="https://cdn.jsdelivr.net/npm/@supabase/supabase-js@2"></script>
//   <script src="../lessons/supabase-config.js"></script>
//   <script src="../lessons/lessons-common.js"></script>
// 사용:
//   LessonHub.init({ unit:'일차함수', game:'graph-race', onReady:function(user){...} });
//   LessonHub.saveScore(120, {stage:3});          // 점수 저장
//   LessonHub.leaderboard().then(rows => ...);     // 등수판 (RLS가 공개범위 강제)
//   LessonHub.onNewScore(function(row){...});      // 실시간 새 기록 알림
//   LessonHub.user  → { studentId, name, className }
// ═══════════════════════════════════════════════════════════════
var LessonHub = (function () {
  var sb = supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY);
  var cfg = null;   // { unit, game, onReady }
  var user = null;  // { studentId, name, className }

  function classNameOf(sid) {
    sid = String(sid || '');
    return sid.length >= 2 ? sid.substring(0, 1) + '학년 ' + sid.substring(1, 2) + '반' : '';
  }

  async function sha256(text) {
    var buf = new TextEncoder().encode(text);
    var h = await crypto.subtle.digest('SHA-256', buf);
    return Array.from(new Uint8Array(h)).map(function (b) { return b.toString(16).padStart(2, '0'); }).join('');
  }

  function setUserFromSession(session) {
    var m = (session && session.user && session.user.user_metadata) || {};
    var sid = m.student_id || '';
    user = { studentId: sid, name: m.name || '', className: classNameOf(sid) };
  }

  // ── 로그인 오버레이 ──────────────────────────────────────────
  function showLogin() {
    var ov = document.createElement('div');
    ov.id = 'lhLogin';
    ov.style.cssText = 'position:fixed;top:0;left:0;right:0;bottom:0;background:#f3f4f6;z-index:9999;display:flex;align-items:center;justify-content:center;padding:20px;font-family:Pretendard,-apple-system,sans-serif;';
    ov.innerHTML =
      '<div style="background:white;width:100%;max-width:360px;border-radius:20px;box-shadow:0 10px 30px rgba(0,0,0,.08);padding:28px;box-sizing:border-box;">'
      + '<div style="font-size:20px;font-weight:900;color:#1e3a8a;text-align:center;margin-bottom:4px;">🎮 수업활동 로그인</div>'
      + '<div style="font-size:12px;color:#94a3b8;text-align:center;margin-bottom:20px;">수학교실과 같은 학번·비밀번호예요</div>'
      + '<input id="lhSid" type="number" placeholder="학번 4자리" style="width:100%;padding:13px;border:1.5px solid #e5e7eb;border-radius:12px;box-sizing:border-box;font-size:15px;background:#f9fafb;margin-bottom:10px;">'
      + '<input id="lhPw" type="password" placeholder="비밀번호" style="width:100%;padding:13px;border:1.5px solid #e5e7eb;border-radius:12px;box-sizing:border-box;font-size:15px;background:#f9fafb;margin-bottom:14px;">'
      + '<button id="lhBtn" style="width:100%;padding:14px;background:#2563eb;color:white;border:none;border-radius:12px;font-size:15px;font-weight:800;cursor:pointer;">입장하기</button>'
      + '<div id="lhMsg" style="text-align:center;color:#ef4444;font-size:13px;font-weight:600;margin-top:12px;min-height:18px;"></div>'
      + '</div>';
    document.body.appendChild(ov);
    var saved = localStorage.getItem('lh_sid');
    if (saved) document.getElementById('lhSid').value = saved;
    document.getElementById('lhPw').addEventListener('keydown', function (e) { if (e.key === 'Enter') doLogin(); });
    document.getElementById('lhBtn').addEventListener('click', doLogin);
  }

  async function doLogin() {
    var sid = document.getElementById('lhSid').value.trim();
    var pw = document.getElementById('lhPw').value;
    var msg = document.getElementById('lhMsg');
    var btn = document.getElementById('lhBtn');
    if (!sid || !pw) { msg.textContent = '학번과 비밀번호를 입력해요'; return; }
    btn.disabled = true; btn.textContent = '확인 중...'; msg.textContent = '';
    try {
      var pwHash = await sha256(pw);
      // 1) 서버 함수로 검증 (기존 비밀번호 그대로)
      var res = await fetch(SUPABASE_URL + '/functions/v1/login', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'apikey': SUPABASE_ANON_KEY, 'Authorization': 'Bearer ' + SUPABASE_ANON_KEY },
        body: JSON.stringify({ studentId: sid, pwHash: pwHash })
      });
      var out = await res.json();
      if (!out.success) { msg.textContent = out.message || '로그인 실패'; btn.disabled = false; btn.textContent = '입장하기'; return; }
      // 2) 세션 발급
      var auth = await sb.auth.signInWithPassword({ email: out.email, password: pwHash });
      if (auth.error) { msg.textContent = '세션 오류: ' + auth.error.message; btn.disabled = false; btn.textContent = '입장하기'; return; }
      localStorage.setItem('lh_sid', sid);
      setUserFromSession(auth.data.session);
      var ov = document.getElementById('lhLogin');
      if (ov) ov.remove();
      if (cfg.onReady) cfg.onReady(user);
    } catch (e) {
      msg.textContent = '오류: ' + e.message;
      btn.disabled = false; btn.textContent = '입장하기';
    }
  }

  // ── 공개 API ────────────────────────────────────────────────
  async function init(options) {
    cfg = options || {};
    var s = await sb.auth.getSession();
    if (s.data.session) {
      setUserFromSession(s.data.session);
      if (cfg.onReady) cfg.onReady(user);
    } else {
      showLogin();
    }
  }

  async function saveScore(score, meta) {
    if (!user) throw new Error('로그인이 필요해요');
    var r = await sb.from('game_scores').insert({
      student_id: user.studentId, unit: cfg.unit, game: cfg.game,
      score: score, meta: meta || {}
    });
    if (r.error) throw new Error(r.error.message);
    return true;
  }

  async function leaderboard(limit) {
    var r = await sb.from('leaderboard').select('*')
      .eq('unit', cfg.unit).eq('game', cfg.game)
      .order('best_score', { ascending: false }).limit(limit || 50);
    if (r.error) throw new Error(r.error.message);
    return r.data;
  }

  function onNewScore(fn) {
    sb.channel('scores-' + cfg.unit + '-' + cfg.game)
      .on('postgres_changes',
        { event: 'INSERT', schema: 'public', table: 'game_scores', filter: 'unit=eq.' + cfg.unit },
        function (payload) { if (payload.new && payload.new.game === cfg.game) fn(payload.new); })
      .subscribe();
  }

  async function logout() {
    await sb.auth.signOut();
    location.reload();
  }

  return {
    init: init, saveScore: saveScore, leaderboard: leaderboard,
    onNewScore: onNewScore, logout: logout,
    get user() { return user; }, get client() { return sb; }
  };
})();

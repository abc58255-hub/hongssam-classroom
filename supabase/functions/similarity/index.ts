// 닮음 보물찾기 — 학생용 엣지함수 (정답을 클라이언트에 노출하지 않고 서버에서만 채점)
//
// ops:
//   scan   {op:'scan', token}                          → 숨긴 삼각형 도형(정답 제외) + 내 진행상태
//   submit {op:'submit', token, answer:{idx,cond,ratioA,ratioB}} → 서버 채점 + 선착순·점수 부여
//
// 학생 인증: Authorization: Bearer <supabase 세션 JWT> (supabase-js functions.invoke가 자동 첨부)
//   → user_metadata.student_id 로 신원 확인 (점수 조작 방지).

import { createClient } from "npm:@supabase/supabase-js@2";

const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};
function json(body: unknown, status = 200) {
  return new Response(JSON.stringify(body), { status, headers: { ...CORS, "Content-Type": "application/json" } });
}
function gcd(a: number, b: number): number { a = Math.abs(a); b = Math.abs(b); while (b) { [a, b] = [b, a % b]; } return a || 1; }
function reduce(a: number, b: number): [number, number] { const g = gcd(a, b); return g ? [a / g, b / g] : [a, b]; }

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response("ok", { headers: CORS });
  try {
    const admin = createClient(Deno.env.get("SUPABASE_URL")!, Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!);

    // ── 학생 신원 확인 ──
    const token = (req.headers.get("Authorization") || "").replace(/^Bearer\s+/i, "");
    const { data: u } = await admin.auth.getUser(token);
    const sid = String(u?.user?.user_metadata?.student_id || "").trim();
    const sname = String(u?.user?.user_metadata?.name || "").trim();
    if (!sid) return json({ success: false, message: "로그인이 필요해요." }, 401);

    const body = await req.json();
    const op = body.op;
    const qr = String(body.token || "").trim();

    // ── 스캔: 도형 데이터만 반환 (정답 제외) ──
    if (op === "scan") {
      const st = await admin.from("sim_stations")
        .select("id,set_id,label,side_a,side_b,side_c,show,rotate,flip").eq("token", qr).maybeSingle();
      if (st.error || !st.data) return json({ success: false, message: "QR을 찾을 수 없어요." });
      const set = await admin.from("sim_sets").select("id,active,max_tries,tri_count,title").eq("id", st.data.set_id).maybeSingle();
      if (!set.data || !set.data.active) return json({ success: false, message: "활동이 종료됐어요." });
      const at = await admin.from("sim_attempts").select("tries,cleared,points,clear_order")
        .eq("station_id", st.data.id).eq("student_id", sid).maybeSingle();
      const tries = at.data?.tries || 0;
      return json({
        success: true,
        station: { id: st.data.id, label: st.data.label, side_a: st.data.side_a, side_b: st.data.side_b, side_c: st.data.side_c, show: st.data.show, rotate: st.data.rotate, flip: st.data.flip },
        triCount: set.data.tri_count, title: set.data.title,
        maxTries: set.data.max_tries, tries, cleared: !!at.data?.cleared,
        triesLeft: Math.max(0, set.data.max_tries - tries), points: at.data?.points || 0, order: at.data?.clear_order || null,
      });
    }

    // ── 제출: 서버 채점 + 선착순·점수 ──
    if (op === "submit") {
      const st = await admin.from("sim_stations").select("*").eq("token", qr).maybeSingle();
      if (!st.data) return json({ success: false, message: "QR을 찾을 수 없어요." });
      const S = st.data;
      const set = await admin.from("sim_sets").select("*").eq("id", S.set_id).maybeSingle();
      if (!set.data || !set.data.active) return json({ success: false, message: "활동이 종료됐어요." });
      const SET = set.data;

      let row = (await admin.from("sim_attempts").select("*").eq("station_id", S.id).eq("student_id", sid).maybeSingle()).data;
      if (!row) {
        const ins = await admin.from("sim_attempts").insert({ set_id: S.set_id, station_id: S.id, student_id: sid, name: sname, tries: 0 }).select("*").single();
        if (ins.error) return json({ success: false, message: ins.error.message }, 500);
        row = ins.data;
      }
      if (row.cleared) return json({ success: true, already: true, cleared: true, points: row.points, order: row.clear_order });
      if (row.tries >= SET.max_tries) return json({ success: true, locked: true, triesLeft: 0 });

      const tries = row.tries + 1;
      const ans = body.answer || {};
      const idxOk = Number(ans.idx) === S.ans_idx;
      const condOk = String(ans.cond || "").toUpperCase() === String(S.ans_cond).toUpperCase();
      const [ra, rb] = reduce(Number(ans.ratioA) || 0, Number(ans.ratioB) || 0);
      const ratioOk = ra === S.ratio_a && rb === S.ratio_b;
      const correct = idxOk && condOk && ratioOk;

      if (correct) {
        const { data: ord } = await admin.rpc("sim_award_clear", { p_station: S.id });
        const order = Number(ord) || 1;
        const tb = (SET.try_bonus || [])[tries - 1] || 0;
        const fb = (SET.first_bonus || [])[order - 1] || 0;
        const pts = (SET.base_pts || 0) + tb + fb;
        await admin.from("sim_attempts").update({ tries, cleared: true, clear_order: order, points: pts, name: sname, updated_at: new Date().toISOString() }).eq("id", row.id);
        return json({ success: true, correct: true, cleared: true, order, points: pts, tries });
      } else {
        const triesLeft = Math.max(0, SET.max_tries - tries);
        await admin.from("sim_attempts").update({ tries, name: sname, updated_at: new Date().toISOString() }).eq("id", row.id);
        // 부분 피드백(어느 항목이 맞았는지) — 3번 제한이라 정답 노출 위험 낮고 형성평가에 도움
        return json({ success: true, correct: false, triesLeft, feedback: { idx: idxOk, cond: condOk, ratio: ratioOk } });
      }
    }

    return json({ success: false, message: "알 수 없는 요청" }, 400);
  } catch (e) {
    return json({ success: false, message: "서버 오류: " + (e as Error).message }, 500);
  }
});

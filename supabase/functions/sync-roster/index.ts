// 명부 동기화 수신 — GAS(학생-수학)가 학번·이름·비번해시를 밀어넣는 엔드포인트
// 인증: X-Sync-Key 헤더 == 환경변수 SYNC_KEY (supabase secrets set으로 주입)
// 주의: auth_user 컬럼은 건드리지 않는다 (로그인 함수가 관리)
// 배포: supabase functions deploy sync-roster --no-verify-jwt  ← GAS는 JWT가 없으므로 필수

import { createClient } from "npm:@supabase/supabase-js@2";

Deno.serve(async (req) => {
  if (req.method !== "POST") return Response.json({ success: false }, { status: 405 });

  const key = req.headers.get("x-sync-key") ?? "";
  const expect = Deno.env.get("SYNC_KEY") ?? "";
  if (!expect || key !== expect) {
    return Response.json({ success: false, message: "unauthorized" }, { status: 401 });
  }

  try {
    const body = await req.json();
    const admin0 = createClient(
      Deno.env.get("SUPABASE_URL")!,
      Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!,
    );

    // 게임 목록 조회 (교사 관리용 — hidden 포함 전체)
    if (body.op === "listGames") {
      const { data, error: lErr } = await admin0.from("games")
        .select("unit, game, title, url, scope, status, created_at")
        .order("created_at", { ascending: false });
      if (lErr) return Response.json({ success: false, message: lErr.message }, { status: 500 });
      return Response.json({ success: true, games: data });
    }

    // 게임 등록/수정 모드: {games:[{unit,game,title,url,scope,status}]}
    if (Array.isArray(body.games)) {
      const okStatus = ["hidden", "open", "free"];
      const gRows = body.games.map((g: Record<string, unknown>) => ({
        unit: String(g.unit ?? "").trim(),
        game: String(g.game ?? "").trim(),
        title: String(g.title ?? "").trim(),
        url: String(g.url ?? "").trim(),
        scope: g.scope === "grade" ? "grade" : "class",
        status: okStatus.includes(String(g.status)) ? String(g.status) : "hidden",
      })).filter((g: { unit: string; game: string }) => g.unit && g.game);
      const { error: gErr } = await admin0.from("games").upsert(gRows, { onConflict: "unit,game" });
      if (gErr) return Response.json({ success: false, message: gErr.message }, { status: 500 });
      return Response.json({ success: true, games: gRows.length });
    }

    const { students, partial } = body; // [{sid, name, pwHash}] / partial=true면 비활성화 생략
    if (!Array.isArray(students) || students.length === 0) {
      return Response.json({ success: false, message: "빈 명부" }, { status: 400 });
    }
    const rows = students
      .filter((s) => s && String(s.sid ?? "").trim() && String(s.name ?? "").trim())
      .map((s) => ({
        student_id: String(s.sid).trim(),
        name: String(s.name).trim(),
        pw_hash: String(s.pwHash ?? "").trim().toLowerCase(), // 미가입 학생은 '' (로그인 시 안내됨)
        active: true,
        synced_at: new Date().toISOString(),
      }));

    const admin = admin0;

    // 명부에서 사라진 학생(전출 등)은 삭제 대신 비활성화 — partial 모드에선 생략
    if (!partial) {
      const ids = rows.map((r) => r.student_id);
      const { error: dErr } = await admin.from("students")
        .update({ active: false })
        .not("student_id", "in", `(${ids.join(",")})`);
      if (dErr) return Response.json({ success: false, message: "비활성화 오류: " + dErr.message }, { status: 500 });
    }

    const { error } = await admin.from("students")
      .upsert(rows, { onConflict: "student_id" });
    if (error) return Response.json({ success: false, message: error.message }, { status: 500 });

    return Response.json({ success: true, upserted: rows.length });
  } catch (e) {
    return Response.json({ success: false, message: (e as Error).message }, { status: 500 });
  }
});

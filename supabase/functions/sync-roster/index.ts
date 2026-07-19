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
    const { students } = await req.json(); // [{sid, name, pwHash}]
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

    const admin = createClient(
      Deno.env.get("SUPABASE_URL")!,
      Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!,
    );

    // 명부에서 사라진 학생(전출 등)은 삭제 대신 비활성화
    const ids = rows.map((r) => r.student_id);
    const { error: dErr } = await admin.from("students")
      .update({ active: false })
      .not("student_id", "in", `(${ids.join(",")})`);
    if (dErr) return Response.json({ success: false, message: "비활성화 오류: " + dErr.message }, { status: 500 });

    const { error } = await admin.from("students")
      .upsert(rows, { onConflict: "student_id" });
    if (error) return Response.json({ success: false, message: error.message }, { status: 500 });

    return Response.json({ success: true, upserted: rows.length });
  } catch (e) {
    return Response.json({ success: false, message: (e as Error).message }, { status: 500 });
  }
});

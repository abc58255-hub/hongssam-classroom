// 학생 로그인 검증 — 기존 StudentAuth 비밀번호(SHA-256 해시) 그대로 사용
//
// 흐름: 페이지가 {studentId, pwHash}를 보냄 (pwHash = 브라우저에서 sha256(비밀번호))
//   1. students 테이블의 pw_hash와 비교
//   2. 맞으면 Supabase Auth 계정을 보장 (email = {학번}@st.local, password = pwHash)
//      - 비번이 바뀌었으면(교사 초기화 등) Auth 비밀번호도 갱신
//   3. {success, email} 반환 → 클라이언트가 signInWithPassword(email, pwHash)로 세션 획득
//
// 규칙: 원문 비밀번호는 이 함수에도 오지 않는다 (해시만 왕복).
//       Auth 비밀번호를 해시(64자)로 쓰는 이유: GoTrue 최소 길이 제한 회피 + 원문 비노출.

import { createClient } from "npm:@supabase/supabase-js@2";

const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};

function json(body: unknown, status = 200) {
  return new Response(JSON.stringify(body), {
    status,
    headers: { ...CORS, "Content-Type": "application/json" },
  });
}

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response("ok", { headers: CORS });
  try {
    const { studentId, pwHash } = await req.json();
    const sid = String(studentId ?? "").trim();
    const hash = String(pwHash ?? "").trim().toLowerCase();
    if (!sid || !/^[0-9a-f]{64}$/.test(hash)) {
      return json({ success: false, message: "잘못된 요청입니다." }, 400);
    }

    const admin = createClient(
      Deno.env.get("SUPABASE_URL")!,
      Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!,
    );

    const { data: st, error } = await admin
      .from("students")
      .select("student_id, name, pw_hash, auth_user, active")
      .eq("student_id", sid)
      .maybeSingle();
    if (error) return json({ success: false, message: "조회 오류: " + error.message }, 500);
    if (!st || !st.active) return json({ success: false, message: "등록되지 않은 학번이에요." });
    if (!st.pw_hash) return json({ success: false, message: "비밀번호가 아직 설정되지 않았어요. 수학교실 앱에서 먼저 로그인해 주세요." });
    if (st.pw_hash.toLowerCase() !== hash) {
      return json({ success: false, message: "비밀번호가 올바르지 않아요." });
    }

    const email = `${sid}@st.local`;

    if (!st.auth_user) {
      // 최초 로그인: Auth 계정 생성 + students에 연결
      const { data: created, error: cErr } = await admin.auth.admin.createUser({
        email,
        password: hash,
        email_confirm: true,
        user_metadata: { student_id: sid, name: st.name },
      });
      if (cErr) return json({ success: false, message: "계정 생성 오류: " + cErr.message }, 500);
      const { error: uErr } = await admin
        .from("students")
        .update({ auth_user: created.user.id })
        .eq("student_id", sid);
      if (uErr) return json({ success: false, message: "연결 오류: " + uErr.message }, 500);
    } else {
      // 재로그인: 비번이 바뀌었을 수 있으니 Auth 비밀번호를 현재 해시로 동기화
      const { error: pErr } = await admin.auth.admin.updateUserById(st.auth_user, { password: hash });
      if (pErr) return json({ success: false, message: "동기화 오류: " + pErr.message }, 500);
    }

    return json({ success: true, email, name: st.name });
  } catch (e) {
    return json({ success: false, message: "서버 오류: " + (e as Error).message }, 500);
  }
});

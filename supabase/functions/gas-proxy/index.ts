// GAS 웹앱 프록시 — iOS 사파리 CORS 우회용
//
// 문제: PWA(github.io)가 GAS(script.google.com)로 직접 fetch하면
//   GAS가 googleusercontent.com으로 교차도메인 리다이렉트 → iOS 웹킷이 "Load failed"로 실패.
// 해결: iOS는 Supabase(CORS 정상)에만 요청하고, 리다이렉트 추적은 서버(Deno)가 대신 함.
//   브라우저 → 이 함수 → (서버끼리) GAS → 결과 JSON을 CORS 헤더 얹어 반환.
//
// 본문은 그대로 전달({fn, args}), 응답도 그대로 반환({ok, result} / {ok, error}).

const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};

const GAS_URL =
  "https://script.google.com/macros/s/AKfycbyR1whn6f90-kJEAaJg_O34uP8v-KvyEqsRky58idjoxVDS5cWj80p2ScJp6V2dnz_0hA/exec";

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response("ok", { headers: CORS });
  try {
    const body = await req.text(); // {"fn":..., "args":[...]}
    const r = await fetch(GAS_URL, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body,
      redirect: "follow", // GAS의 302 → googleusercontent 추적(서버끼리라 CORS 무관)
    });
    const text = await r.text();
    return new Response(text, {
      status: 200,
      headers: { ...CORS, "Content-Type": "application/json" },
    });
  } catch (e) {
    return new Response(
      JSON.stringify({ ok: false, error: "프록시 오류: " + (e as Error).message }),
      { status: 200, headers: { ...CORS, "Content-Type": "application/json" } },
    );
  }
});

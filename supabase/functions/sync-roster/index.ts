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
        .select("*")
        .order("created_at", { ascending: false });
      if (lErr) return Response.json({ success: false, message: lErr.message }, { status: 500 });
      return Response.json({ success: true, games: data });
    }

    // 특정 게임 기록 목록 (교사 기록 관리용 — 이름·반 포함, 점수순)
    if (body.op === "listScores") {
      const { data, error: sErr } = await admin0.from("game_scores")
        .select("id, student_id, score, meta, created_at, students(name, class_name)")
        .eq("unit", String(body.unit ?? "")).eq("game", String(body.game ?? ""))
        .order("score", { ascending: false }).limit(300);
      if (sErr) return Response.json({ success: false, message: sErr.message }, { status: 500 });
      return Response.json({ success: true, scores: data });
    }

    // 기록 1건 삭제 (장난·조작 기록 제거)
    if (body.op === "deleteScore") {
      const { error: dErr } = await admin0.from("game_scores").delete().eq("id", Number(body.id));
      if (dErr) return Response.json({ success: false, message: dErr.message }, { status: 500 });
      return Response.json({ success: true });
    }

    // 게임 전체 기록 리셋 (새 학기 등)
    if (body.op === "resetGame") {
      const { error: rErr, count } = await admin0.from("game_scores")
        .delete({ count: "exact" })
        .eq("unit", String(body.unit ?? "")).eq("game", String(body.game ?? ""));
      if (rErr) return Response.json({ success: false, message: rErr.message }, { status: 500 });
      return Response.json({ success: true, deleted: count ?? 0 });
    }

    // ── 협력 직선그래프 탐구 ──
    // 방 생성: {op:'graphCreate', room:{class_scope,a,b,c,display,hide_eq,max_points,show_wrong,int_only}}
    //   같은 반의 기존 활성방은 종료(active=false) 후 새로 생성
    if (body.op === "graphCreate") {
      const R = body.room ?? {};
      const cls = String(R.class_scope ?? "").trim();
      if (!cls) return Response.json({ success: false, message: "반이 없습니다." }, { status: 400 });
      await admin0.from("graph_rooms").update({ active: false }).eq("class_scope", cls).eq("active", true);
      const row = {
        class_scope: cls,
        a: Number(R.a) || 0, b: Number(R.b) || 0, c: Number(R.c) || 0,
        display: String(R.display ?? "").trim(),
        hide_eq: !!R.hide_eq, max_points: Number(R.max_points) || 0,
        show_wrong: R.show_wrong !== false, int_only: !!R.int_only,
        show_names: !!R.show_names, kind: String(R.kind ?? "line"),
        no_dup: R.no_dup !== false,
        revealed: false, active: true,
      };
      let ins = await admin0.from("graph_rooms").insert(row).select("id").single();
      if (ins.error) {
        // 활성 방 유니크 충돌(동시 생성 레이스) → 기존 활성 방 정리 후 1회 재시도
        await admin0.from("graph_rooms").update({ active: false }).eq("class_scope", cls).eq("active", true);
        ins = await admin0.from("graph_rooms").insert(row).select("id").single();
      }
      if (ins.error) return Response.json({ success: false, message: ins.error.message }, { status: 500 });
      return Response.json({ success: true, id: ins.data.id });
    }
    // 방 설정/개형 갱신: {op:'graphUpdate', id, patch:{revealed?,hide_eq?,max_points?,show_wrong?,int_only?,display?}}
    if (body.op === "graphUpdate") {
      const p = body.patch ?? {};
      const patch: Record<string, unknown> = {};
      ["revealed","hide_eq","show_wrong","int_only","show_names","no_dup"].forEach((k)=>{ if (p[k] !== undefined) patch[k] = !!p[k]; });
      if (p.max_points !== undefined) patch.max_points = Number(p.max_points) || 0;
      if (p.display !== undefined) patch.display = String(p.display).trim();
      const { error } = await admin0.from("graph_rooms").update(patch).eq("id", Number(body.id));
      if (error) return Response.json({ success: false, message: error.message }, { status: 500 });
      return Response.json({ success: true });
    }
    // 점 전체 지우기: {op:'graphClear', id}
    if (body.op === "graphClear") {
      const { error, count } = await admin0.from("graph_points").delete({ count: "exact" }).eq("room_id", Number(body.id));
      if (error) return Response.json({ success: false, message: error.message }, { status: 500 });
      return Response.json({ success: true, deleted: count ?? 0 });
    }
    // 방 종료: {op:'graphClose', id}
    if (body.op === "graphClose") {
      const { error } = await admin0.from("graph_rooms").update({ active: false }).eq("id", Number(body.id));
      if (error) return Response.json({ success: false, message: error.message }, { status: 500 });
      return Response.json({ success: true });
    }
    // 프리셋 저장(신규): {op:'graphPresetSave', preset:{title,a,b,c,display,kind,max_points,show_wrong,int_only,show_names,hide_eq}}
    if (body.op === "graphPresetSave") {
      const P = body.preset ?? {};
      const row = {
        title: String(P.title ?? "").trim(),
        a: Number(P.a) || 0, b: Number(P.b) || 0, c: Number(P.c) || 0,
        display: String(P.display ?? "").trim(), kind: String(P.kind ?? "line"),
        max_points: Number(P.max_points) || 0,
        show_wrong: P.show_wrong !== false, int_only: !!P.int_only,
        show_names: !!P.show_names, hide_eq: !!P.hide_eq, no_dup: P.no_dup !== false,
      };
      const { data, error } = await admin0.from("graph_presets").insert(row).select("id").single();
      if (error) return Response.json({ success: false, message: error.message }, { status: 500 });
      return Response.json({ success: true, id: data.id });
    }
    // 프리셋 삭제: {op:'graphPresetDelete', id}
    if (body.op === "graphPresetDelete") {
      const { error } = await admin0.from("graph_presets").delete().eq("id", Number(body.id));
      if (error) return Response.json({ success: false, message: error.message }, { status: 500 });
      return Response.json({ success: true });
    }

    // ── 수업 링크 보드 ──
    // 링크 목록 (교사 관리용 — 비공개 포함 전체)
    if (body.op === "listLinks") {
      const { data, error: llErr } = await admin0.from("links")
        .select("*").order("sort", { ascending: true }).order("created_at", { ascending: false });
      if (llErr) return Response.json({ success: false, message: llErr.message }, { status: 500 });
      return Response.json({ success: true, links: data });
    }
    // 링크 저장(신규/수정) — id 있으면 update, 없으면 insert
    if (body.op === "saveLink") {
      const L = body.link ?? {};
      const row: Record<string, unknown> = {
        title: String(L.title ?? "").trim(),
        url: String(L.url ?? "").trim(),
        class_scope: String(L.class_scope ?? "all").trim() || "all",
        active: L.active === false ? false : true,
        sort: Number(L.sort) || 0,
      };
      if (!row.title || !row.url) return Response.json({ success: false, message: "제목과 URL은 필수예요." }, { status: 400 });
      if (L.id) {
        const { error: uErr } = await admin0.from("links").update(row).eq("id", Number(L.id));
        if (uErr) return Response.json({ success: false, message: uErr.message }, { status: 500 });
      } else {
        const { error: iErr } = await admin0.from("links").insert(row);
        if (iErr) return Response.json({ success: false, message: iErr.message }, { status: 500 });
      }
      return Response.json({ success: true });
    }
    // 링크 삭제
    if (body.op === "deleteLink") {
      const { error: dlErr } = await admin0.from("links").delete().eq("id", Number(body.id));
      if (dlErr) return Response.json({ success: false, message: dlErr.message }, { status: 500 });
      return Response.json({ success: true });
    }

    // ── 하브루타(짝 질문) 교사 제어 ──
    // 세션 열기: {op:'havrutaOpen', session:{class_scope, unit}} — 기존 활성 세션 종료 후 새로
    if (body.op === "havrutaOpen") {
      const S = body.session ?? {};
      const cls = String(S.class_scope ?? "").trim();
      if (!cls) return Response.json({ success: false, message: "반이 없습니다." }, { status: 400 });
      await admin0.from("havruta_sessions").update({ active: false }).eq("class_scope", cls).eq("active", true);
      const row = { class_scope: cls, unit: String(S.unit ?? "").trim(), sub_units: String(S.sub_units ?? "").trim(),
        n_words: Math.max(1, Number(S.n_words) || 2), n_questions: Math.max(1, Number(S.n_questions) || 2),
        n_votes: Math.max(1, Number(S.n_votes) || 2),
        show_names: S.show_names !== false, stage: "submit", active: true };
      let ins = await admin0.from("havruta_sessions").insert(row).select("id").single();
      if (ins.error) {
        await admin0.from("havruta_sessions").update({ active: false }).eq("class_scope", cls).eq("active", true);
        ins = await admin0.from("havruta_sessions").insert(row).select("id").single();
      }
      if (ins.error) return Response.json({ success: false, message: ins.error.message }, { status: 500 });
      return Response.json({ success: true, id: ins.data.id });
    }
    // 단계 전환: {op:'havrutaStage', id, stage:'submit'|'vote'|'discuss'}
    if (body.op === "havrutaStage") {
      const st = ["submit", "vote", "discuss"].includes(String(body.stage)) ? String(body.stage) : "submit";
      const { error } = await admin0.from("havruta_sessions").update({ stage: st }).eq("id", Number(body.id));
      if (error) return Response.json({ success: false, message: error.message }, { status: 500 });
      return Response.json({ success: true });
    }
    // 오늘의 질문 선정(소단원별): {op:'havrutaFocus', id(세션), qid, sub_unit, on}
    if (body.op === "havrutaFocus") {
      const sid = Number(body.id), qid = Number(body.qid), su = Number(body.sub_unit) || 0, on = body.on !== false;
      if (on) {
        // 같은 소단원의 기존 선정 해제 후 이 질문만 선정 (소단원당 1개)
        await admin0.from("havruta_questions").update({ focus: false }).eq("session_id", sid).eq("sub_unit", su);
        const { error } = await admin0.from("havruta_questions").update({ focus: true }).eq("id", qid);
        if (error) return Response.json({ success: false, message: error.message }, { status: 500 });
      } else {
        const { error } = await admin0.from("havruta_questions").update({ focus: false }).eq("id", qid);
        if (error) return Response.json({ success: false, message: error.message }, { status: 500 });
      }
      return Response.json({ success: true });
    }
    // 이름/익명 전체 전환: {op:'havrutaNames', id, on}
    if (body.op === "havrutaNames") {
      const { error } = await admin0.from("havruta_sessions").update({ show_names: !!body.on }).eq("id", Number(body.id));
      if (error) return Response.json({ success: false, message: error.message }, { status: 500 });
      return Response.json({ success: true });
    }
    // 후보 토글(투표에 올릴 질문): {op:'havrutaCandidate', qid, on}
    if (body.op === "havrutaCandidate") {
      const { error } = await admin0.from("havruta_questions").update({ candidate: !!body.on }).eq("id", Number(body.qid));
      if (error) return Response.json({ success: false, message: error.message }, { status: 500 });
      return Response.json({ success: true });
    }
    // 질문 삭제(부적절): {op:'havrutaDeleteQ', qid}
    if (body.op === "havrutaDeleteQ") {
      const { error } = await admin0.from("havruta_questions").delete().eq("id", Number(body.qid));
      if (error) return Response.json({ success: false, message: error.message }, { status: 500 });
      return Response.json({ success: true });
    }
    // 세션 닫기: {op:'havrutaClose', id}
    if (body.op === "havrutaClose") {
      const { error } = await admin0.from("havruta_sessions").update({ active: false }).eq("id", Number(body.id));
      if (error) return Response.json({ success: false, message: error.message }, { status: 500 });
      return Response.json({ success: true });
    }
    // 세션 영구 삭제(질문·투표 cascade): {op:'havrutaDeleteSession', id}
    if (body.op === "havrutaDeleteSession") {
      const { error } = await admin0.from("havruta_sessions").delete().eq("id", Number(body.id));
      if (error) return Response.json({ success: false, message: error.message }, { status: 500 });
      return Response.json({ success: true });
    }

    // ── 닮음 보물찾기: 세트 열기(기존 활성 닫고 새로) ──
    // {op:'simOpenSet', set:{class_scope,title,base_pts,try_bonus,first_bonus,max_tries},
    //  triangles:[{idx,side_a,side_b,side_c,show}],
    //  stations:[{label,side_a,side_b,side_c,show,rotate,flip,ans_idx,ans_cond,ratio_a,ratio_b}]}
    if (body.op === "simOpenSet") {
      const S = body.set ?? {};
      const cls = String(S.class_scope ?? "").trim();
      if (!cls) return Response.json({ success: false, message: "반이 없습니다." }, { status: 400 });
      // '모든 반(*)'을 열면 기존 활성 세트를 전부 닫아 전 학급에 확실히 적용. 특정 반이면 그 반 것만.
      if (cls === "*") await admin0.from("sim_sets").update({ active: false }).eq("active", true);
      else await admin0.from("sim_sets").update({ active: false }).eq("class_scope", cls).eq("active", true);
      const setRow = {
        class_scope: cls, title: String(S.title ?? "닮음 보물찾기").trim(),
        base_pts: Number(S.base_pts) || 100,
        try_bonus: Array.isArray(S.try_bonus) ? S.try_bonus.map(Number) : [50, 20, 0],
        first_bonus: Array.isArray(S.first_bonus) ? S.first_bonus.map(Number) : [30, 20, 10],
        max_tries: Math.max(1, Number(S.max_tries) || 3),
        tri_count: (body.triangles ?? []).length, station_count: (body.stations ?? []).length,
        active: true,
      };
      let ins = await admin0.from("sim_sets").insert(setRow).select("id").single();
      if (ins.error) { // 활성 유니크 경합 대비 1회 재시도
        await admin0.from("sim_sets").update({ active: false }).eq("class_scope", cls).eq("active", true);
        ins = await admin0.from("sim_sets").insert(setRow).select("id").single();
      }
      if (ins.error) return Response.json({ success: false, message: ins.error.message }, { status: 500 });
      const setId = ins.data.id;

      const tris = (body.triangles ?? []).map((t: Record<string, unknown>) => ({
        set_id: setId, idx: Number(t.idx),
        side_a: Number(t.side_a), side_b: Number(t.side_b), side_c: Number(t.side_c), show: t.show ?? {},
      }));
      if (tris.length) {
        const te = await admin0.from("sim_triangles").insert(tris);
        if (te.error) return Response.json({ success: false, message: "삼각형 저장 오류: " + te.error.message }, { status: 500 });
      }
      const stns = (body.stations ?? []).map((st: Record<string, unknown>) => ({
        set_id: setId, token: crypto.randomUUID().replace(/-/g, "").slice(0, 8),
        label: String(st.label ?? "").trim(),
        side_a: Number(st.side_a), side_b: Number(st.side_b), side_c: Number(st.side_c),
        show: st.show ?? {}, rotate: Number(st.rotate) || 0, flip: !!st.flip,
        ans_idx: Number(st.ans_idx), ans_cond: String(st.ans_cond ?? "SSS").toUpperCase(),
        ratio_a: Number(st.ratio_a) || 1, ratio_b: Number(st.ratio_b) || 1,
      }));
      let outStations: unknown[] = [];
      if (stns.length) {
        const se = await admin0.from("sim_stations").insert(stns).select("id,token,label");
        if (se.error) return Response.json({ success: false, message: "스테이션 저장 오류: " + se.error.message }, { status: 500 });
        outStations = se.data ?? [];
      }
      return Response.json({ success: true, id: setId, stations: outStations });
    }
    // 세트 닫기: {op:'simCloseSet', id}
    if (body.op === "simCloseSet") {
      const { error } = await admin0.from("sim_sets").update({ active: false }).eq("id", Number(body.id));
      if (error) return Response.json({ success: false, message: error.message }, { status: 500 });
      return Response.json({ success: true });
    }
    // 세트 삭제(cascade): {op:'simDeleteSet', id}
    if (body.op === "simDeleteSet") {
      const { error } = await admin0.from("sim_sets").delete().eq("id", Number(body.id));
      if (error) return Response.json({ success: false, message: error.message }, { status: 500 });
      return Response.json({ success: true });
    }
    // 세트 불러오기(수정용, 정답 포함 — 교사 전용): {op:'simGetSet', id}
    if (body.op === "simGetSet") {
      const id = Number(body.id);
      const set = await admin0.from("sim_sets").select("*").eq("id", id).maybeSingle();
      if (!set.data) return Response.json({ success: false, message: "세트를 찾을 수 없어요." });
      const tris = await admin0.from("sim_triangles").select("*").eq("set_id", id).order("idx");
      const stns = await admin0.from("sim_stations").select("*").eq("set_id", id).order("id");
      return Response.json({ success: true, set: set.data, triangles: tris.data || [], stations: stns.data || [] });
    }
    // 세트 제자리 수정(QR 토큰 유지): {op:'simUpdateSet', id, set, triangles, stations}
    if (body.op === "simUpdateSet") {
      const id = Number(body.id); const S = body.set ?? {};
      const su = await admin0.from("sim_sets").update({
        title: String(S.title ?? "닮음 보물찾기").trim(),
        base_pts: Number(S.base_pts) || 100, max_tries: Math.max(1, Number(S.max_tries) || 3),
        tri_count: (body.triangles ?? []).length, station_count: (body.stations ?? []).length,
      }).eq("id", id);
      if (su.error) return Response.json({ success: false, message: su.error.message }, { status: 500 });
      // 삼각형: 전체 교체(참조 FK 없음)
      await admin0.from("sim_triangles").delete().eq("set_id", id);
      const tris = (body.triangles ?? []).map((t: Record<string, unknown>) => ({ set_id: id, idx: Number(t.idx), side_a: Number(t.side_a), side_b: Number(t.side_b), side_c: Number(t.side_c), show: t.show ?? {} }));
      if (tris.length) { const te = await admin0.from("sim_triangles").insert(tris); if (te.error) return Response.json({ success: false, message: "삼각형 저장 오류: " + te.error.message }, { status: 500 }); }
      // 스테이션: 기존 토큰 유지하며 제자리 업데이트 (순서 매칭)
      const existing = (await admin0.from("sim_stations").select("id,token").eq("set_id", id).order("id")).data ?? [];
      const ns = body.stations ?? [];
      const out: unknown[] = [];
      for (let i = 0; i < ns.length; i++) {
        const st = ns[i] as Record<string, unknown>;
        const fields = { label: String(st.label ?? "").trim(), side_a: Number(st.side_a), side_b: Number(st.side_b), side_c: Number(st.side_c), show: st.show ?? {}, rotate: Number(st.rotate) || 0, flip: !!st.flip, ans_idx: Number(st.ans_idx), ans_cond: String(st.ans_cond ?? "SSS").toUpperCase(), ratio_a: Number(st.ratio_a) || 1, ratio_b: Number(st.ratio_b) || 1 };
        if (i < existing.length) {
          await admin0.from("sim_stations").update(fields).eq("id", existing[i].id);
          out.push({ id: existing[i].id, token: existing[i].token, label: fields.label });
        } else {
          const ins = await admin0.from("sim_stations").insert({ set_id: id, token: crypto.randomUUID().replace(/-/g, "").slice(0, 8), ...fields }).select("id,token,label").single();
          if (ins.error) return Response.json({ success: false, message: "스테이션 저장 오류: " + ins.error.message }, { status: 500 });
          out.push(ins.data);
        }
      }
      if (existing.length > ns.length) { const extra = existing.slice(ns.length).map((e: Record<string, unknown>) => e.id); await admin0.from("sim_stations").delete().in("id", extra); }
      return Response.json({ success: true, id, stations: out });
    }

    // 게임 등록/수정 모드: {games:[{unit,game,title,url,scope,status}]}
    if (Array.isArray(body.games)) {
      const okStatus = ["hidden", "open", "free"];
      const gRows = body.games.map((g: Record<string, unknown>) => {
        const row: Record<string, unknown> = {
          unit: String(g.unit ?? "").trim(),
          game: String(g.game ?? "").trim(),
          title: String(g.title ?? "").trim(),
          url: String(g.url ?? "").trim(),
          scope: g.scope === "grade" ? "grade" : "class",
          status: okStatus.includes(String(g.status)) ? String(g.status) : "hidden",
        };
        // 점수 상한 (0 = 제한 없음) — 마이그레이션 전 호환 위해 값이 있을 때만 포함
        if (g.max_score !== undefined) row.max_score = Number(g.max_score) || 0;
        return row;
      }).filter((g: { unit: string; game: string }) => g.unit && g.game);
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

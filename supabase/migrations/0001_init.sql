-- 수업활동 플랫폼 v1 스키마 (README-설계.md 참고)
-- 적용: supabase db push (CLI) 또는 대시보드 SQL Editor에 붙여넣기

-- ── 학생 명부 미러 (원본 = GAS 학생명부, 주기 동기화) ─────────────
create table if not exists public.students (
  student_id text primary key,          -- 학번 '2611'
  name       text not null,
  pw_hash    text not null,             -- SHA-256 hex (StudentAuth와 동일)
  auth_user  uuid unique references auth.users(id),
  active     boolean not null default true,
  synced_at  timestamptz not null default now()
);

-- ── 게임/활동 기록 ─────────────────────────────────────────────
create table if not exists public.game_scores (
  id         bigint generated always as identity primary key,
  student_id text not null references public.students(student_id),
  unit       text not null,             -- 단원 키 (예: '일차함수')
  game       text not null,             -- 게임 키 (예: 'graph-race')
  score      numeric not null,
  meta       jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now()
);
create index if not exists idx_scores_board on public.game_scores (unit, game, score desc);
create index if not exists idx_scores_student on public.game_scores (student_id, created_at desc);

-- ── AI 중계 사용량 (Edge Function 전용 — 클라이언트 접근 불가) ────
create table if not exists public.ai_usage (
  student_id text not null references public.students(student_id),
  used_on    date not null default (now() at time zone 'Asia/Seoul')::date,
  count      int  not null default 0,
  primary key (student_id, used_on)
);

-- ── 헬퍼: 현재 로그인한 학생의 학번 ──────────────────────────────
create or replace function public.current_student_id()
returns text language sql stable security definer set search_path = public as $$
  select student_id from public.students where auth_user = auth.uid()
$$;

-- ── RLS ───────────────────────────────────────────────────────
alter table public.students   enable row level security;
alter table public.game_scores enable row level security;
alter table public.ai_usage   enable row level security;  -- 정책 없음 = 서비스 롤만 접근

-- students: 로그인한 학생은 명부의 이름·학번만 조회 가능 (등수판 표시용)
--           pw_hash는 컬럼 권한으로 차단 — 어떤 클라이언트에도 노출 금지
revoke select on public.students from anon, authenticated;
grant  select (student_id, name, active) on public.students to authenticated;
create policy students_read on public.students
  for select to authenticated using (true);

-- game_scores: 읽기 = 로그인 학생 전부 (등수판) / 쓰기 = 본인 기록만
create policy scores_read on public.game_scores
  for select to authenticated using (true);
create policy scores_insert_own on public.game_scores
  for insert to authenticated
  with check (student_id = public.current_student_id());

-- ── 등수판 뷰: unit+game별 학생 최고점 ──────────────────────────
create or replace view public.leaderboard
with (security_invoker = true) as
  select s.unit, s.game, s.student_id, st.name,
         max(s.score) as best_score, count(*) as attempts,
         max(s.created_at) as last_played
  from public.game_scores s
  join public.students st on st.student_id = s.student_id
  group by s.unit, s.game, s.student_id, st.name;

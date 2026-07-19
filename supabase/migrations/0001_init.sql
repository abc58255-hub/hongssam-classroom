-- 수업활동 플랫폼 v1 스키마 (README-설계.md 참고)
-- 적용: supabase db push (CLI) 또는 대시보드 SQL Editor에 붙여넣기
-- 2026-07-19 사용자 결정 반영: 등수판 공개범위 = 게임별 선택제(class/grade), 실명 표시, AI 중계는 보류

-- ── 학생 명부 미러 (원본 = GAS 학생명부, 주기 동기화) ─────────────
create table if not exists public.students (
  student_id text primary key,          -- 학번 '2611'
  name       text not null,
  class_name text generated always as
    (substring(student_id, 1, 1) || '학년 ' || substring(student_id, 2, 1) || '반') stored,
  pw_hash    text not null,             -- SHA-256 hex (StudentAuth와 동일)
  auth_user  uuid unique references auth.users(id),
  active     boolean not null default true,
  synced_at  timestamptz not null default now()
);

-- ── 게임(활동) 등록부 — 공개 범위 선택제의 단일 출처 ───────────────
--    선생님이 게임을 열 때 한 줄 등록: scope = 'class'(반 안에서만) | 'grade'(학년 전체)
create table if not exists public.games (
  unit   text not null,                 -- 단원 키 (예: '일차함수')
  game   text not null,                 -- 게임 키 (예: 'graph-race')
  title  text not null default '',
  scope  text not null default 'class' check (scope in ('class', 'grade')),
  active boolean not null default true,
  created_at timestamptz not null default now(),
  primary key (unit, game)
);

-- ── 게임/활동 기록 ─────────────────────────────────────────────
create table if not exists public.game_scores (
  id         bigint generated always as identity primary key,
  student_id text not null references public.students(student_id),
  unit       text not null,
  game       text not null,
  score      numeric not null,
  meta       jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now()
);
create index if not exists idx_scores_board on public.game_scores (unit, game, score desc);
create index if not exists idx_scores_student on public.game_scores (student_id, created_at desc);

-- ── 헬퍼: 현재 로그인한 학생의 학번/반 ───────────────────────────
create or replace function public.current_student_id()
returns text language sql stable security definer set search_path = public as $$
  select student_id from public.students where auth_user = auth.uid()
$$;
create or replace function public.current_class_name()
returns text language sql stable security definer set search_path = public as $$
  select class_name from public.students where auth_user = auth.uid()
$$;

-- ── RLS ───────────────────────────────────────────────────────
alter table public.students    enable row level security;
alter table public.games       enable row level security;
alter table public.game_scores enable row level security;

-- students: 로그인한 학생은 이름·학번·반만 조회 가능 (등수판 실명 표시용)
--           pw_hash는 컬럼 권한으로 차단 — 어떤 클라이언트에도 노출 금지
revoke select on public.students from anon, authenticated;
grant  select (student_id, name, class_name, active) on public.students to authenticated;
create policy students_read on public.students
  for select to authenticated using (true);

-- games: 읽기 = 로그인 학생 전부 / 쓰기 = 서비스 롤(교사 도구)만
create policy games_read on public.games
  for select to authenticated using (true);

-- game_scores 읽기 — 공개범위 선택제를 DB 차원에서 강제:
--   본인 기록은 항상 / scope='grade' 게임은 전체 / scope='class' 게임은 같은 반 기록만
create policy scores_read on public.game_scores
  for select to authenticated using (
    student_id = public.current_student_id()
    or exists (select 1 from public.games g
               where g.unit = game_scores.unit and g.game = game_scores.game
                 and g.scope = 'grade')
    or exists (select 1 from public.students owner
               where owner.student_id = game_scores.student_id
                 and owner.class_name = public.current_class_name())
  );

-- game_scores 쓰기 — 본인 기록만, 등록·활성화된 게임에만
create policy scores_insert_own on public.game_scores
  for insert to authenticated
  with check (
    student_id = public.current_student_id()
    and exists (select 1 from public.games g
                where g.unit = game_scores.unit and g.game = game_scores.game
                  and g.active)
  );

-- ── 등수판 뷰: unit+game별 학생 최고점 (RLS가 뷰 통과해도 그대로 적용됨) ──
create or replace view public.leaderboard
with (security_invoker = true) as
  select s.unit, s.game, s.student_id, st.name, st.class_name,
         max(s.score) as best_score, count(*) as attempts,
         max(s.created_at) as last_played
  from public.game_scores s
  join public.students st on st.student_id = s.student_id
  group by s.unit, s.game, s.student_id, st.name, st.class_name;

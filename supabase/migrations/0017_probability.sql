-- 0017_probability.sql — 확률 체험 (큰 수의 법칙): 동전·주사위 시뮬레이션
-- 교사가 반별 세션 열기(도구+사건 선택) → 학생이 각자 체험 → 교사 TV에서 반 전체 합계
-- 교사 쓰기(세션)=서비스롤(엣지함수). 학생 쓰기(기여)=본인 세션.

-- ── 세션(반당 활성 1개) ──
create table if not exists public.prob_sessions (
  id           bigint generated always as identity primary key,
  class_scope  text    not null,                          -- 'N학년 M반'
  tool         text    not null default 'coin',           -- 'coin' | 'dice'
  event_value  int     not null default 0,                -- 동전=1(앞면), 주사위=1~6(특정 숫자)
  theory_p     numeric not null default 0.5,              -- 이론 확률 (표시용, 동전 1/2, 주사위 1/6)
  active       boolean not null default true,
  created_at   timestamptz not null default now()
);
create unique index if not exists prob_sessions_one_active
  on public.prob_sessions (class_scope) where (active);

-- ── 학생 누적 기여 (세션당 학생 1행, upsert) ──
create table if not exists public.prob_contributions (
  session_id  bigint not null references public.prob_sessions(id) on delete cascade,
  student_id  text   not null,
  name        text   not null default '',
  trials      int    not null default 0,
  successes   int    not null default 0,
  updated_at  timestamptz not null default now(),
  primary key (session_id, student_id)
);
create index if not exists prob_contrib_session on public.prob_contributions(session_id);

-- ── RLS ──
alter table public.prob_sessions      enable row level security;
alter table public.prob_contributions enable row level security;

-- 세션 읽기: 누구나(학생앱 카드 anon 조회 + 로그인 학생 + TV)
drop policy if exists prob_sessions_read on public.prob_sessions;
create policy prob_sessions_read on public.prob_sessions for select to anon, authenticated using (true);

-- 기여 읽기: 누구나 (교사 TV가 anon으로 합계 집계)
drop policy if exists prob_contrib_read on public.prob_contributions;
create policy prob_contrib_read on public.prob_contributions for select to anon, authenticated using (true);

-- 기여 쓰기: 로그인 학생 본인 것만 (학번은 current_student_id()로 확인)
drop policy if exists prob_contrib_upsert on public.prob_contributions;
create policy prob_contrib_upsert on public.prob_contributions
  for insert to authenticated with check (student_id = public.current_student_id());
drop policy if exists prob_contrib_update on public.prob_contributions;
create policy prob_contrib_update on public.prob_contributions
  for update to authenticated using (student_id = public.current_student_id())
                                with check (student_id = public.current_student_id());

grant select on public.prob_sessions to anon, authenticated;
grant select, insert, update on public.prob_contributions to anon, authenticated;

-- 실시간 구독 (교사 TV가 학생 기여를 실시간으로 받음)
do $$ begin alter publication supabase_realtime add table public.prob_sessions;      exception when duplicate_object then null; end $$;
do $$ begin alter publication supabase_realtime add table public.prob_contributions; exception when duplicate_object then null; end $$;

notify pgrst, 'reload schema';

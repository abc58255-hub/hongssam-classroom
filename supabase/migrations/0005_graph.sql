-- 0005_graph.sql — 직선의 방정식 협력 그래프 탐구
-- 교사가 ax+by=c 직선을 내고, 학생들이 그 위의 점(x,y)을 제출해 실시간 좌표평면에 모임.
-- 방=반당 1개(active). 정답 판정은 클라이언트(ax+by=c). 점은 분수(분자/분모)로 정확 저장.

-- ── 방 ──
create table if not exists public.graph_rooms (
  id           bigint generated always as identity primary key,
  class_scope  text not null,                 -- 'N학년 M반' (반 자동입장 키)
  a            numeric not null default 0,     -- ax + by = c
  b            numeric not null default 0,
  c            numeric not null default 0,
  display      text not null default '',       -- 학생에게 보이는 식 (예: 'x = 2')
  hide_eq      boolean not null default false, -- 함수식 숨기기(표시만)
  max_points   int not null default 3,         -- 학생당 점 개수(0=무제한)
  show_wrong   boolean not null default true,  -- 틀린 점도 표시
  int_only     boolean not null default false, -- 정수만 입력
  revealed     boolean not null default false, -- 개형(직선) 공개
  active       boolean not null default true,  -- 진행 중
  created_at   timestamptz not null default now()
);
-- 반당 활성 방 1개 보장 (부분 유니크 인덱스)
create unique index if not exists graph_rooms_one_active
  on public.graph_rooms (class_scope) where (active);

-- ── 제출 점 ──
create table if not exists public.graph_points (
  id          bigint generated always as identity primary key,
  room_id     bigint not null references public.graph_rooms(id) on delete cascade,
  student_id  text not null,
  name        text not null default '',
  x_num       int not null,  x_den int not null default 1,  -- x = x_num/x_den
  y_num       int not null,  y_den int not null default 1,  -- y = y_num/y_den
  correct     boolean not null default false,
  created_at  timestamptz not null default now()
);
create index if not exists graph_points_room on public.graph_points(room_id);

alter table public.graph_rooms  enable row level security;
alter table public.graph_points enable row level security;

-- 방 읽기: 누구나(학생앱 anon 조회 + 로그인 학생). 쓰기=서비스롤(교사도구, 엣지함수)만
drop policy if exists graph_rooms_read on public.graph_rooms;
create policy graph_rooms_read on public.graph_rooms
  for select to anon, authenticated using (true);
grant select on public.graph_rooms to anon, authenticated;

-- 점 읽기: 누구나(실시간 좌표평면 표시)
drop policy if exists graph_points_read on public.graph_points;
create policy graph_points_read on public.graph_points
  for select to anon, authenticated using (true);
-- 점 쓰기: 로그인 학생 본인 것만 (student_id = 세션 학번)
drop policy if exists graph_points_insert on public.graph_points;
create policy graph_points_insert on public.graph_points
  for insert to authenticated
  with check (student_id = public.current_student_id());
grant select, insert on public.graph_points to anon, authenticated;

-- 실시간 구독 대상에 추가 (이미 있으면 무시 — 재실행 안전)
do $$ begin
  alter publication supabase_realtime add table public.graph_points;
exception when duplicate_object then null; end $$;
do $$ begin
  alter publication supabase_realtime add table public.graph_rooms;
exception when duplicate_object then null; end $$;

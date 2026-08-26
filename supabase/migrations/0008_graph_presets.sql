-- 그래프 탐구: 교사가 미리 만든 직선(계수+표시식+설정) 프리셋 저장 → 클릭 한 번에 방 시작
create table if not exists public.graph_presets (
  id          bigint generated always as identity primary key,
  title       text    not null default '',
  a           numeric not null default 0,
  b           numeric not null default 0,
  c           numeric not null default 0,
  display     text    not null default '',
  kind        text    not null default 'line',
  max_points  int     not null default 3,
  show_wrong  boolean not null default true,
  int_only    boolean not null default false,
  show_names  boolean not null default false,
  hide_eq     boolean not null default false,
  sort        int     not null default 0,
  created_at  timestamptz not null default now()
);

alter table public.graph_presets enable row level security;
-- 읽기: 누구나(교사 페이지 anon 조회). 쓰기=서비스롤(엣지함수)만
drop policy if exists graph_presets_read on public.graph_presets;
create policy graph_presets_read on public.graph_presets
  for select to anon, authenticated using (true);
grant select on public.graph_presets to anon, authenticated;

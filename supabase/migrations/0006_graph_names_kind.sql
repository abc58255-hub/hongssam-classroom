-- 그래프 탐구 확장: 이름 표시 옵션 + 그래프 종류(향후 확장 대비)
alter table public.graph_rooms
  add column if not exists show_names boolean not null default false,
  add column if not exists kind text not null default 'line';

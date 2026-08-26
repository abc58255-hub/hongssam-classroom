-- 그래프 탐구: 본인 중복 제출 방지를 '방별 설정(no_dup)'으로 전환
-- 기본 true(기존 동작 유지). OFF일 때 같은 점 재제출 허용되도록 전역 유니크 인덱스는 제거하고
-- 방지는 클라이언트(mySet)에서 no_dup 설정에 따라 판단.
alter table public.graph_rooms   add column if not exists no_dup boolean not null default true;
alter table public.graph_presets add column if not exists no_dup boolean not null default true;

drop index if exists public.graph_points_no_selfdup;

-- 0015_similarity.sql — 닮음 보물찾기 (야외 QR 닮음 찾기 활동)
-- 학생: QR 스캔 → 숨긴 삼각형 화면 표시 → 학습지 1~9 중 [번호+닮음조건+닮음비] 제출.
-- 스테이션당 3번 기회, 선착순 보너스. 정답은 학생이 읽을 수 없게 분리(엣지함수만 조회).

-- ── 세트 (반당 활성 1개) ──
create table if not exists public.sim_sets (
  id           bigint generated always as identity primary key,
  class_scope  text not null,
  title        text not null default '닮음 보물찾기',
  base_pts     int  not null default 100,                 -- 클리어 기본점
  try_bonus    int[] not null default '{50,20,0}',        -- 1·2·3번째 정답 보너스
  first_bonus  int[] not null default '{30,20,10}',       -- 스테이션별 선착 1·2·3등 보너스(이후 0)
  max_tries    int  not null default 3,
  active       boolean not null default true,
  created_at   timestamptz not null default now()
);
create unique index if not exists sim_sets_one_active on public.sim_sets (class_scope) where (active);

-- ── 주어진 삼각형 1~9 (학습지에 인쇄) ──
-- side_*: 실제 모양(세 변). show: 학생에게 보여줄 라벨 {sa,sb,sc,Aa,Ab,Ac:boolean} — 닮음조건을 결정.
create table if not exists public.sim_triangles (
  id        bigint generated always as identity primary key,
  set_id    bigint not null references public.sim_sets(id) on delete cascade,
  idx       int not null,                                 -- 표시 번호 1~9
  side_a    numeric not null, side_b numeric not null, side_c numeric not null,
  show      jsonb not null default '{}',
  created_at timestamptz not null default now()
);
create index if not exists sim_triangles_set on public.sim_triangles(set_id);

-- ── 숨긴 삼각형(스테이션) — 정답 포함, 학생 직접조회 금지 ──
create table if not exists public.sim_stations (
  id        bigint generated always as identity primary key,
  set_id    bigint not null references public.sim_sets(id) on delete cascade,
  token     text not null unique,                         -- QR에 담는 불투명 토큰
  label     text not null default '',                     -- 교사용 지점 표식
  side_a    numeric not null, side_b numeric not null, side_c numeric not null,
  show      jsonb not null default '{}',
  rotate    int not null default 0,                       -- 표시 회전(도)
  flip      boolean not null default false,
  ans_idx   int not null,                                 -- 정답: 대응 주어진 번호 1~9
  ans_cond  text not null,                                -- 정답: 'SSS'|'SAS'|'AA'
  ratio_a   int not null, ratio_b int not null,           -- 정답: 닮음비(숨긴:고른것) 약분형
  created_at timestamptz not null default now()
);
create index if not exists sim_stations_set on public.sim_stations(set_id);

-- ── 학생 시도/클리어 ──
create table if not exists public.sim_attempts (
  id          bigint generated always as identity primary key,
  set_id      bigint not null references public.sim_sets(id) on delete cascade,
  station_id  bigint not null references public.sim_stations(id) on delete cascade,
  student_id  text not null,
  name        text not null default '',
  tries       int not null default 0,
  cleared     boolean not null default false,
  clear_order int,                                        -- 스테이션별 선착 순위(클리어 시 부여)
  points      int not null default 0,
  updated_at  timestamptz not null default now(),
  unique (station_id, student_id)
);
create index if not exists sim_attempts_set on public.sim_attempts(set_id);

-- ── RLS ──
-- 핵심: sim_triangles·sim_stations 는 정답을 담으므로 학생(anon/authenticated) 조회 전면 금지.
--       학생은 엣지함수(similarity, 서비스롤)로만 도형(정답 제외)·채점 접근.
alter table public.sim_sets      enable row level security;
alter table public.sim_triangles enable row level security;
alter table public.sim_stations  enable row level security;
alter table public.sim_attempts  enable row level security;

-- 세트: 활성 여부 등 노출 무해 → 읽기 허용
drop policy if exists sim_sets_read on public.sim_sets;
create policy sim_sets_read on public.sim_sets for select to anon, authenticated using (true);

-- 시도/순위: 리더보드 표시용 읽기 허용 (정답 없음). 쓰기는 엣지(서비스롤)만 → 클라 정책 없음
drop policy if exists sim_attempts_read on public.sim_attempts;
create policy sim_attempts_read on public.sim_attempts for select to anon, authenticated using (true);

-- sim_triangles·sim_stations: select 정책 없음 = 기본 거부 (서비스롤만 조회)

grant select on public.sim_sets to anon, authenticated;
grant select on public.sim_attempts to anon, authenticated;

-- 실시간(리더보드 + 세트 활성 전환)
do $$ begin alter publication supabase_realtime add table public.sim_attempts; exception when duplicate_object then null; end $$;
do $$ begin alter publication supabase_realtime add table public.sim_sets;     exception when duplicate_object then null; end $$;

-- ── 클리어 원자 처리 (선착순 경합 방지) ──
-- 스테이션 행을 잠그고 현재 클리어 수 기준으로 순위 부여 → 동시 제출에도 선착순 정확.
create or replace function public.sim_award_clear(p_station bigint)
  returns int language plpgsql security definer set search_path = public as $$
declare n int;
begin
  perform 1 from public.sim_stations where id = p_station for update;   -- 직렬화
  select count(*) into n from public.sim_attempts where station_id = p_station and cleared;
  return n + 1;
end $$;

-- 0004_links.sql — 교사가 학생에게 보여줄 "수업 링크 보드"
-- 학생 수학앱 대시보드 상단 카드로 노출. 반별 지정 + 공개/비공개 토글.
-- 읽기 = 누구나(anon 포함) 활성 링크만 / 쓰기 = 서비스 롤(교사 도구, 엣지함수)만.

create table if not exists public.links (
  id          bigint generated always as identity primary key,
  title       text not null,
  url         text not null,
  class_scope text not null default 'all',   -- 'all' 또는 '1학년 3반' 같은 반 라벨
  active      boolean not null default true,  -- 공개/비공개 토글
  sort        int not null default 0,         -- 표시 순서(작을수록 위)
  created_at  timestamptz not null default now()
);

alter table public.links enable row level security;

-- 읽기: anon/로그인 모두, 단 활성 링크만 (수학앱은 anon 키로 읽음)
drop policy if exists links_read on public.links;
create policy links_read on public.links
  for select to anon, authenticated using (active = true);

grant select on public.links to anon, authenticated;

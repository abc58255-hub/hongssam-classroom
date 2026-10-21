-- 0012_havruta_focus.sql — 소단원별 '오늘의 질문' 선정 플래그
-- 투표는 후보 지정 없이 소단원 전체 질문에서, 토론용 오늘의 질문은 소단원당 1개(focus=true)
alter table public.havruta_questions add column if not exists focus boolean not null default false;

-- 0013_havruta_nvotes.sql — 소단원당 투표 수를 세션에서 설정
alter table public.havruta_sessions add column if not exists n_votes int not null default 2;

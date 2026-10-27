-- 0016_sim_attempt_class.sql — 순위를 반별로 분리하기 위해 시도에 반 정보 저장
-- 세트는 공용('*' = 모든 반)으로 쓰되, 각 시도에 학생의 반을 기록 → 순위판에서 반별 필터.
alter table public.sim_attempts add column if not exists class_scope text not null default '';
create index if not exists sim_attempts_setclass on public.sim_attempts(set_id, class_scope);

-- 0011_havruta_options.sql — 기존 havruta 테이블에 옵션 컬럼 추가 (0010 이후 보강)
-- 소단원당 단어/질문 입력칸 수 + 질문 투표 후보 플래그
alter table public.havruta_sessions  add column if not exists n_words     int not null default 2;
alter table public.havruta_sessions  add column if not exists n_questions int not null default 2;
alter table public.havruta_sessions  add column if not exists show_names  boolean not null default true;  -- 교사가 이름/익명 전체 전환
alter table public.havruta_questions add column if not exists candidate   boolean not null default false;

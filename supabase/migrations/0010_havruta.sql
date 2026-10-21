-- 0010_havruta.sql — 하브루타(짝 질문) 수업 활동
-- 교사가 반별 세션을 열고(제출→투표→토론 단계), 학생이 소단원별 질문을 내고 투표(1인 3표).
-- 교사 쓰기(세션/단계/삭제/오늘질문)=서비스롤(엣지함수 sync-roster). 학생 쓰기(질문·투표)=본인 세션.

-- ── 세션(방): 반당 활성 1개 ──
create table if not exists public.havruta_sessions (
  id           bigint generated always as identity primary key,
  class_scope  text not null,                       -- 'N학년 M반'
  unit         text not null default '',            -- 단원명(표시용)
  sub_units    text not null default '',            -- 소단원 라벨(줄바꿈 구분, 비면 ①~⑤ 번호)
  n_words      int not null default 2,              -- 소단원당 단어 입력칸 수
  n_questions  int not null default 2,              -- 소단원당 질문 입력칸 수
  show_names   boolean not null default true,       -- 교사가 이름/익명 전체 전환
  stage        text not null default 'submit',      -- 'submit' | 'vote' | 'discuss'
  focus_q      bigint,                              -- 오늘의 질문(havruta_questions.id)
  active       boolean not null default true,
  created_at   timestamptz not null default now()
);
create unique index if not exists havruta_sessions_one_active
  on public.havruta_sessions (class_scope) where (active);

-- ── 질문 ──
create table if not exists public.havruta_questions (
  id          bigint generated always as identity primary key,
  session_id  bigint not null references public.havruta_sessions(id) on delete cascade,
  sub_unit    int not null default 0,               -- 소단원 1~5 (0=자유)
  student_id  text not null,
  name        text not null default '',
  anon        boolean not null default false,
  word        text not null default '',             -- 중요한 단어
  question    text not null default '',             -- 질문
  votes       int not null default 0,               -- 득표(트리거로 집계)
  candidate   boolean not null default false,        -- 교사가 투표 후보로 지정
  created_at  timestamptz not null default now()
);
create index if not exists havruta_questions_session on public.havruta_questions(session_id);

-- ── 투표(1인 1질문 1표, 세션당 3표는 클라이언트 제한) ──
create table if not exists public.havruta_votes (
  id          bigint generated always as identity primary key,
  session_id  bigint not null references public.havruta_sessions(id) on delete cascade,
  question_id bigint not null references public.havruta_questions(id) on delete cascade,
  student_id  text not null,
  created_at  timestamptz not null default now()
);
create unique index if not exists havruta_votes_once on public.havruta_votes (question_id, student_id);
create index if not exists havruta_votes_session on public.havruta_votes(session_id);

-- 득표 집계 트리거 (votes 테이블 ↔ questions.votes 동기화)
-- SECURITY DEFINER: 학생 권한으로 트리거가 돌 때 havruta_questions UPDATE가 RLS에 막히므로 소유자 권한으로 우회
create or replace function public.havruta_vote_sync() returns trigger language plpgsql
  security definer set search_path = public as $$
begin
  if (tg_op = 'INSERT') then
    update public.havruta_questions set votes = votes + 1 where id = new.question_id;
  elsif (tg_op = 'DELETE') then
    update public.havruta_questions set votes = greatest(0, votes - 1) where id = old.question_id;
  end if;
  return null;
end $$;
drop trigger if exists havruta_votes_trg on public.havruta_votes;
create trigger havruta_votes_trg after insert or delete on public.havruta_votes
  for each row execute function public.havruta_vote_sync();

-- ── RLS ──
alter table public.havruta_sessions  enable row level security;
alter table public.havruta_questions enable row level security;
alter table public.havruta_votes     enable row level security;

-- 읽기: 누구나 (학생앱 anon 카드 조회 + 로그인 학생 + 칠판)
drop policy if exists havruta_sessions_read on public.havruta_sessions;
create policy havruta_sessions_read on public.havruta_sessions for select to anon, authenticated using (true);
drop policy if exists havruta_questions_read on public.havruta_questions;
create policy havruta_questions_read on public.havruta_questions for select to anon, authenticated using (true);
drop policy if exists havruta_votes_read on public.havruta_votes;
create policy havruta_votes_read on public.havruta_votes for select to anon, authenticated using (true);

-- 질문 쓰기: 로그인 학생 본인 것만
drop policy if exists havruta_questions_insert on public.havruta_questions;
create policy havruta_questions_insert on public.havruta_questions
  for insert to authenticated with check (student_id = public.current_student_id());
-- 질문 삭제: 본인 것 (교사 삭제는 서비스롤)
drop policy if exists havruta_questions_delete on public.havruta_questions;
create policy havruta_questions_delete on public.havruta_questions
  for delete to authenticated using (student_id = public.current_student_id());

-- 투표 쓰기/취소: 로그인 학생 본인 것만
drop policy if exists havruta_votes_insert on public.havruta_votes;
create policy havruta_votes_insert on public.havruta_votes
  for insert to authenticated with check (student_id = public.current_student_id());
drop policy if exists havruta_votes_delete on public.havruta_votes;
create policy havruta_votes_delete on public.havruta_votes
  for delete to authenticated using (student_id = public.current_student_id());

grant select on public.havruta_sessions to anon, authenticated;
grant select, insert, delete on public.havruta_questions to anon, authenticated;
grant select, insert, delete on public.havruta_votes to anon, authenticated;

-- 실시간 구독
do $$ begin alter publication supabase_realtime add table public.havruta_sessions;  exception when duplicate_object then null; end $$;
do $$ begin alter publication supabase_realtime add table public.havruta_questions; exception when duplicate_object then null; end $$;
do $$ begin alter publication supabase_realtime add table public.havruta_votes;     exception when duplicate_object then null; end $$;

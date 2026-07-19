-- 게임 노출 상태 3단계(hidden/open/free) + 페이지 URL
-- hidden=숨김(참여 불가), open=진행 중(수업 활동), free=자유 도전(종료 후 기록경쟁)
-- 주의: active 컬럼을 참조하는 정책을 먼저 제거해야 컬럼 삭제 가능 (의존성 순서)

alter table public.games add column if not exists status text not null default 'hidden'
  check (status in ('hidden', 'open', 'free'));
alter table public.games add column if not exists url text not null default '';

-- 기존 active → open 으로 이관
update public.games set status = 'open' where active;

-- 의존 정책 먼저 제거 → 컬럼 제거
drop policy if exists scores_insert_own on public.game_scores;
drop policy if exists games_read on public.games;
alter table public.games drop column if exists active;

-- 노출 정책: 학생(로그인·비로그인 모두) hidden은 목록에서 안 보임
--   비로그인(anon)도 읽는 이유: 수학교실 앱의 활동 카드 섹션이 로그인 전에 목록을 그림
create policy games_read on public.games
  for select to anon, authenticated using (status in ('open', 'free'));

-- 점수 쓰기: hidden 게임에는 기록 불가 (open/free만)
create policy scores_insert_own on public.game_scores
  for insert to authenticated
  with check (
    student_id = public.current_student_id()
    and exists (select 1 from public.games g
                where g.unit = game_scores.unit and g.game = game_scores.game
                  and g.status in ('open', 'free'))
  );

-- 점수 조작 대비: 게임별 최대 점수 상한 (0 = 제한 없음)
-- 60초 게임에 99999점 같은 개발자도구 조작 기록을 DB 차원에서 차단

alter table public.games add column if not exists max_score numeric not null default 0;

drop policy if exists scores_insert_own on public.game_scores;
create policy scores_insert_own on public.game_scores
  for insert to authenticated
  with check (
    student_id = public.current_student_id()
    and game_scores.score >= 0
    and exists (select 1 from public.games g
                where g.unit = game_scores.unit and g.game = game_scores.game
                  and g.status in ('open', 'free')
                  and (g.max_score <= 0 or game_scores.score <= g.max_score))
  );

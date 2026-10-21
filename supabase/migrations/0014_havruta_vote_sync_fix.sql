-- 0014_havruta_vote_sync_fix.sql — 투표 집계 안 되던 버그 수정
-- 원인: 트리거가 학생(authenticated) 권한으로 havruta_questions UPDATE → RLS에 막혀 0행(조용히 실패).
-- 해결: 함수를 SECURITY DEFINER(소유자 권한)로 재정의 + 기존 투표분 재집계.

create or replace function public.havruta_vote_sync() returns trigger
  language plpgsql security definer set search_path = public as $$
begin
  if (tg_op = 'INSERT') then
    update public.havruta_questions set votes = votes + 1 where id = new.question_id;
  elsif (tg_op = 'DELETE') then
    update public.havruta_questions set votes = greatest(0, votes - 1) where id = old.question_id;
  end if;
  return null;
end $$;

-- 이미 들어간 투표 기준으로 득표 다시 계산
update public.havruta_questions q
  set votes = (select count(*) from public.havruta_votes v where v.question_id = q.id);

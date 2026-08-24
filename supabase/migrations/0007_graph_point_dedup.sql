-- 본인 중복 제출 방지: 같은 방에서 한 학생이 동일 좌표(약분된 값)를 두 번 못 찍게
-- 다른 학생과 겹치는 건 허용 (student_id 포함)
create unique index if not exists graph_points_no_selfdup
  on public.graph_points (room_id, student_id, x_num, x_den, y_num, y_den);

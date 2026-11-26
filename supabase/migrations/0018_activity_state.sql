-- 0018_activity_state.sql — 수업활동 세션 방식 → 상태머신(준비·진행·종료) 전환
--
-- 문제: 지금은 active boolean + UNIQUE(class_scope) WHERE active 구조라
--   ① 만드는 즉시 학생에게 공개되고 ② 반당 1개만 가능해 미리 준비해둘 수 없었음.
-- 해결: state = 'draft'(준비) | 'live'(진행) | 'done'(종료).
--   draft는 몇 개든 미리 만들어둘 수 있고(학생에게 안 보임), live만 반당 1개.
--
-- 호환: 기존 코드가 아직 active를 읽고 쓰므로 트리거로 active ↔ state 양방향 동기화.
--   (active=true ↔ state='live'). 전환이 끝나면 나중에 active 컬럼 제거 가능.

-- ── 공통 헬퍼: 테이블 하나를 상태머신으로 전환 ──
do $$
declare
  t text;
  tabs text[] := array['prob_sessions','havruta_sessions','graph_rooms','sim_sets'];
begin
  foreach t in array tabs loop
    -- 1) state 컬럼 추가
    execute format('alter table public.%I add column if not exists state text not null default ''live''', t);
    -- 2) 기존 데이터 변환: active=true → live, false → done
    execute format('update public.%I set state = case when active then ''live'' else ''done'' end', t);
    -- 3) 값 제약
    execute format('alter table public.%I drop constraint if exists %I', t, t || '_state_chk');
    execute format('alter table public.%I add constraint %I check (state in (''draft'',''live'',''done''))', t, t || '_state_chk');
    -- 4) 조회 인덱스
    execute format('create index if not exists %I on public.%I (class_scope, state)', t || '_cls_state', t);
  end loop;
end $$;

-- ── 유니크 인덱스 교체: active 기준 → state='live' 기준 ──
drop index if exists public.prob_sessions_one_active;
drop index if exists public.havruta_sessions_one_active;
drop index if exists public.graph_rooms_one_active;
drop index if exists public.sim_sets_one_active;

create unique index if not exists prob_sessions_one_live    on public.prob_sessions    (class_scope) where (state = 'live');
create unique index if not exists havruta_sessions_one_live on public.havruta_sessions (class_scope) where (state = 'live');
create unique index if not exists graph_rooms_one_live      on public.graph_rooms      (class_scope) where (state = 'live');
-- 닮음은 '*'(모든 반) 세트를 쓰므로 반별 유니크를 그대로 유지
create unique index if not exists sim_sets_one_live         on public.sim_sets         (class_scope) where (state = 'live');

-- ── active ↔ state 동기화 트리거 (전환기 호환) ──
create or replace function public.sync_active_state() returns trigger language plpgsql as $$
begin
  if (tg_op = 'INSERT') then
    -- state를 명시했으면 그걸 따르고, active만 왔으면 active로 유도
    if new.state is null then new.state := case when new.active then 'live' else 'done' end; end if;
    new.active := (new.state = 'live');
  elsif (tg_op = 'UPDATE') then
    if new.state is distinct from old.state then
      new.active := (new.state = 'live');            -- state가 바뀌면 active를 맞춤
    elsif new.active is distinct from old.active then
      new.state := case when new.active then 'live'  -- 옛 코드가 active만 바꾼 경우
                        else 'done' end;
    end if;
  end if;
  return new;
end $$;

do $$
declare t text; tabs text[] := array['prob_sessions','havruta_sessions','graph_rooms','sim_sets'];
begin
  foreach t in array tabs loop
    execute format('drop trigger if exists %I on public.%I', t || '_sync_state', t);
    execute format('create trigger %I before insert or update on public.%I for each row execute function public.sync_active_state()', t || '_sync_state', t);
  end loop;
end $$;

notify pgrst, 'reload schema';

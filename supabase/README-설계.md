# 수업활동 플랫폼 — Supabase 설계 (0단계)

> 2026-07-19 초안. GAS 코어(과제·채점·성적)는 그대로 두고, 새 수업활동(단원 게임·등수·프로젝트·AI 중계)만 Supabase에 신축한다.

## 전체 구조

```
GitHub Pages (fcm-sw/lessons/…)  ← 화면: 단원 페이지·게임 (순수 HTML)
        │  supabase-js로 직접 질의
        ▼
Supabase (서울 리전)
  ├─ Postgres + RLS  : 학생·게임기록·참여기록
  ├─ Auth            : 학번 로그인 (기존 비밀번호 그대로)
  ├─ Realtime        : 실시간 등수판·멘티미터
  └─ Edge Functions  : 로그인 검증, AI API 중계(키 은닉 + 사용량 제한)

GAS (기존)             ← 명부·비번해시의 원본. 주기 동기화로 Supabase에 미러
```

## 학생 로그인 흐름 (재가입 없음)

1. GAS 학생명부(학번·이름·SHA256 비번해시)를 `students` 테이블로 동기화
2. 학생이 단원 페이지에서 학번+비밀번호 입력
3. Edge Function `login`: 해시 검증 → Supabase Auth 계정 자동 생성/연결
   (이메일 형식 필요해서 내부용 `{학번}@st.local` 사용 — 학생에겐 안 보임)
4. 이후 같은 주소(github.io)의 모든 단원 페이지에서 자동 로그인 유지

## 테이블 설계 (v1)

### students — 명부 미러
| 컬럼 | 타입 | 설명 |
|---|---|---|
| student_id | text PK | 학번 '2611' |
| name | text | 이름 |
| pw_hash | text | SHA-256 hex (기존 StudentAuth와 동일 방식) |
| auth_user | uuid | Supabase Auth 사용자 연결 (최초 로그인 시 채움) |
| active | boolean | 전출 등 비활성 처리 |

### game_scores — 게임/활동 기록
| 컬럼 | 타입 | 설명 |
|---|---|---|
| id | bigint PK | |
| student_id | text FK | |
| unit | text | 단원 키 (예: '일차함수') |
| game | text | 게임 키 (예: 'graph-race') |
| score | numeric | 점수 |
| meta | jsonb | 게임별 부가정보 (클리어 시간, 단계 등) |
| created_at | timestamptz | |

등수판 = 이 테이블 집계 (최고점 기준, unit+game별). 학생별 최고 기록만 남기지 않고
전부 쌓는다 — "도전 횟수" 같은 통계도 뽑을 수 있게.

### ai_usage — AI 중계 사용량 (요금 통제)
| 컬럼 | 타입 | 설명 |
|---|---|---|
| student_id | text | |
| used_on | date | |
| count | int | 하루 사용 횟수 (Edge Function이 증가·제한) |

### 프로젝트 과제용 테이블은 v2에서 (요구사항 확정 후)

## 보안 규칙 (RLS) 원칙

- `students`: 본인 행만 조회 가능. pw_hash는 어떤 클라이언트에도 노출 금지 (Edge Function 전용)
- `game_scores`: 쓰기 = 본인 것만 · 로그인 필수 / 읽기 = 로그인한 학생 전부 (등수판용)
- 등수판 노출 범위는 **미정** → 사용자 결정 필요 (아래 질문)

## 결정 사항 (2026-07-19 사용자 확정)

1. **등수판 공개 범위 = 선택제**: 게임마다 `games.scope`로 지정 — 'class'(반 안에서만) / 'grade'(학년 전체). RLS가 DB 차원에서 강제
2. **실명 표시** (닉네임 없음)
3. **AI 중계는 보류** — v1에서 제외, 나중에 AI 기능 만들 때 재검토 (ai_usage 테이블도 그때 추가)

## 구축 순서

1. ✅ 설계 (이 문서)
2. ⬜ 사용자: Supabase 프로젝트 생성 (서울 리전)
3. ⬜ 스키마 적용 (migrations/0001_init.sql — CLI 또는 대시보드 SQL Editor)
4. ⬜ 명부 동기화 (GAS에서 1회 내보내기 + 이후 주기 동기화)
5. ⬜ Edge Function: login, ai-proxy
6. ⬜ 공용 JS 조각 (lessons-common.js): 로그인 UI + 점수 저장 + 등수판 위젯
7. ⬜ 템플릿 단원 페이지 + 시범 게임 1개
8. ⬜ 수학교실 앱에 "이번 주 활동" 안내 연결
9. ⬜ GitHub Actions 주간 핑 (무료 티어 일시정지 방지)

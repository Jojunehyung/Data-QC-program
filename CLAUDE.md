# CLAUDE.md — 프로젝트 원칙 및 컨텍스트

## 필수 참조 문서

코드 수정 전 반드시 확인:
- `AGENTS.md` — AI 행동 원칙
- `ARCHITECTURE.md` — 전체 구조
- `docs/design-docs/core-beliefs.md` — 절대 불변 원칙
- 해당 기능의 `docs/product-specs/*.md`

## 핵심 원칙 요약

1. 식별번호 우선순위: P > F > 기타 (같은 병록번호 내)
2. P없고 F만 있을 때: 최신 F가 메인, 나머지 F → 가족데이터
3. 생성ID: 내부 로직에 의해 자동 생성 (형식 비공개)
4. 출생연월 2자리 → 4자리: `YEAR_THRESHOLD`(24) 이하면 2000년대, 초과면 1900년대
5. R-ID·생성ID: 미매칭·중복 시트에서 제거, 매칭성공 시트에만 R-ID 포함
6. 병록번호: 항상 8자리 (7자리면 앞에 0)
7. 오류 알림: 데모 환경에서 실제 발송 없이 미리보기만 표시
8. 수집일지 자동 수정: 수집관리시스템=내부코드 기준(개인번호 재조합), 데이터추출시스템=병록번호 기준(컬럼 덮어쓰기)
9. 접수일자 유효성: 누락 또는 미래 날짜 → 오기데이터 분류
10. 보정 매칭 결과: 매칭결과.xlsx 인플레이스 업데이트 (별도 파일 미생성)

## 코드 위치 참조

| 기능 | 함수명 |
|---|---|
| Master DB 생성 | `run_build_master_db()` |
| 병록번호·접수일자 추출 | `run_extract_for_rid_request()` |
| R-ID 매칭 | `run_rid_matching()` |
| 오기입 보정 매칭 | `run_correction_matching()` |
| 오류 알림 발송 | `run_send_error_notification()` |
| 수집일지 자동 수정 | `run_update_collection_log()` |
| 식별번호 우선순위 | `id_priority()` (인라인 함수) |
| 컬럼 자동 감지 | `_detect_columns()` |
| 내부 식별키 생성 | `build_internal_key()` |
| 개인번호 재조합 | `_build_pers_no()` |

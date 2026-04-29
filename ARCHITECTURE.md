# ARCHITECTURE.md — 시스템 구조

## 전체 구조

```
clinical_data_qc.py
│
├── [CONFIG & CONSTANTS]        컬럼명 상수, 시트명 상수, 키워드 설정
├── [SHARED UTILS]              공통 유틸리티 함수
├── [GUI COMPONENTS]            ProgressWindow, 파일 선택 다이얼로그
│
├── [STAGE 1] Master DB 생성     run_build_master_db()
│   ├── 다중 파일 읽기 & 병합
│   ├── 이름 정제 & 가족 분류 (is_family_name)
│   ├── 개인번호 파싱 (parse_personal_no)
│   ├── 생성ID 자동 생성 (build_internal_key)
│   └── 식별번호 우선순위 정렬 (O > X > 기타)
│
├── [STAGE 2] 병록번호·접수일자 추출  run_extract_for_rid_request()
│   └── Master DB → R-ID 요청용 파일 생성
│
├── [STAGE 3] R-ID 매칭           run_rid_matching()
│   ├── 컬럼 자동 감지 (_detect_columns)
│   ├── 내부식별키 매핑 구축 (_build_rid_map)
│   └── 매칭성공 / 미매칭 / 중복 분류
│
├── [STAGE 4] 오기입 보정 매칭     run_correction_matching()
│   ├── 외부 데이터소스(Hubis) 기준 성별·생년월 보정
│   └── 생성ID 재생성 & 재매칭
│
├── [STAGE 5] 오류 알림 발송       run_send_error_notification()
│   ├── 미매칭 bCODE / 병록번호 추출
│   ├── 이메일 미리보기 (_EmailPreviewDialog)
│   └── 데모: 실제 발송 없이 미리보기만 표시
│
├── [STAGE 6] 수집일지 자동 수정   run_update_collection_log()
│   ├── 휴비스쌤 수정파일 읽기 (_read_hubis_fix)
│   │   └── bCODE 기준 매칭 → 개인번호 재조합 (_build_pers_no)
│   ├── 슈프림 수정파일 읽기 (_read_supreme_fix)
│   │   └── 병록번호 기준 매칭 → 변경 컬럼 덮어쓰기
│   └── 원본 포맷 그대로 출력
│
└── [MAIN GUI]                  main()
    ├── 버튼 1: Master DB 생성
    ├── 버튼 2: 병록번호·접수일자 추출
    ├── 버튼 3: R-ID 매칭
    ├── 버튼 4: 오기입 보정 매칭
    ├── 버튼 5: 오류 알림 발송
    └── 버튼 6: 수집일지 자동 수정
```

---

## 데이터 흐름

```
수집일지 엑셀 파일(들)
        │
        ▼
   [1단계: Master DB 생성]
   병합 → 정제 → 가족 분류 → 생성ID → 우선순위 정렬
        │
        ▼
   Master_DB.xlsx
   ├── 통합데이터  (유효 환자 데이터)
   ├── 오기데이터  (이상 데이터)
   ├── 가족데이터  (가족/보호자)
   └── 변경이력    (bCODE 변경 기록)
        │
        ▼
   [2단계: 병록번호·접수일자 추출]
   → 외부 시스템에 R-ID 요청
        │
        ▼
   R-ID 파일 수령
        │
        ▼
   [3단계: R-ID 매칭]
   생성ID 매핑 → 매칭성공 / 미매칭 / 중복
        │
   ┌────┴──────────┐
   │               │
  매칭성공       미매칭
   │               │
   ▼               ▼
  완료      [5단계: 오류 알림 발송]
            미매칭 목록 → 이메일 미리보기
                   │
                   ▼
            [4단계: 오기입 보정 매칭]
            Hubis 기준 성별·생년월 보정 → 재매칭
                   │
                   ▼
            [6단계: 수집일지 자동 수정]
            수정파일(휴비스쌤/슈프림) → 원본 업데이트
```

---

## 핵심 데이터 모델

| 컬럼 상수 | 컬럼명 | 설명 |
|---|---|---|
| `COL_BCODE` | bCODE | 기관 내 환자 코드 |
| `COL_ID_NO` | 식별번호 | O/X 접두사 식별자 |
| `COL_HOSP_NUM` | 병록번호 | 8자리 입원 번호 |
| `COL_NAME` | 이름 | 환자 이름 |
| `COL_PERS_NO` | 개인번호 | 성별+생년월 (M93.05 형식) |
| `COL_DATE` | 접수일자 | 데이터 수집일 |
| `COL_GEN_ID` | 생성ID | 내부 식별키 |
| `COL_RID` | R-ID | IRB 승인 기반 연구 ID |

---

## 외부 의존성

| 라이브러리 | 용도 |
|---|---|
| pandas | 데이터 처리 |
| openpyxl | 엑셀 쓰기 |
| python-calamine | 엑셀 빠른 읽기 |
| tkinter | GUI (표준 라이브러리) |

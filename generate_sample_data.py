# -*- coding: utf-8 -*-
"""
가상 샘플 데이터 생성 스크립트
실행하면 sample_data/ 폴더에 테스트용 Excel 파일 4종을 생성합니다.
모든 데이터는 완전한 가상 데이터이며 실제 환자 정보를 포함하지 않습니다.
"""
from pathlib import Path

import pandas as pd
from openpyxl.styles import numbers as _n

OUT_DIR = Path(__file__).parent / 'sample_data'
OUT_DIR.mkdir(exist_ok=True)

# ---------------------------------------------------------------------------
# 가상 환자 원본 (병록번호 기준 10명)
# ---------------------------------------------------------------------------
PATIENTS = [
    # bCODE  식별번호   병록번호    이름      개인번호    접수일자        비고
    (10001, 'O-0001', '12340001', '김가나', 'M93.05', '2024-03-15', ''),
    (10002, 'O-0002', '12340002', '이다라', 'F75.11', '2024-03-15', ''),
    (10003, 'X-0001', '12340003', '박마바', 'M82.07', '2024-03-15', ''),
    (10004, 'O-0003', '12340004', '최사아', 'F90.04', '2024-03-16', ''),
    (10005, 'X-0002', '1234005',  '정자차', 'M01.09', '2024-03-16', ''),   # 7자리 병록번호
    (10006, 'O-0004', '12340006', '강카타', 'F55.12', '2024-03-17', ''),
    (10007, 'O-0005', '12340007', '조파하', 'M79.06', '2024-03-17', ''),
    (10008, 'O-0006', '12340008', '윤거너', 'F93.02', '2024-03-18', ''),
    # 가족 데이터 (보호자 키워드 포함)
    (10001, 'X-0003', '12340001', '보호자 김가나', 'M65.03', '2024-03-16', ''),
    # 같은 병록번호에 O/X 공존 → O가 대표
    (10003, 'O-0007', '12340003', '박마바', 'M82.07', '2024-03-18', ''),
]

COLS = ['bCODE', '식별번호', '병록번호', '이름', '개인번호', '접수일자', '비고']


def _write_text_col(ws, col_idx: int):
    for r in range(2, ws.max_row + 1):
        ws.cell(r, col_idx).number_format = _n.FORMAT_TEXT


def make_collection_log():
    """수집일지_샘플.xlsx — 시트 2개 (수집일지가 여러 시트로 구성된 경우 재현)"""
    rows = [list(p) for p in PATIENTS]
    df = pd.DataFrame(rows, columns=COLS)
    df1 = df.iloc[:6].reset_index(drop=True)
    df2 = df.iloc[6:].reset_index(drop=True)

    out = OUT_DIR / '수집일지_샘플.xlsx'
    with pd.ExcelWriter(out, engine='openpyxl') as w:
        df1.to_excel(w, sheet_name='3월1주', index=False)
        df2.to_excel(w, sheet_name='3월2주', index=False)
        for sname, df_s in [('3월1주', df1), ('3월2주', df2)]:
            ws = w.sheets[sname]
            _write_text_col(ws, list(df_s.columns).index('병록번호') + 1)
    print(f'생성: {out.name}')


def make_rid_file():
    """rid_파일_샘플.xlsx — 외부 시스템에서 수령한 R-ID 파일"""
    data = [
        ('RID-2024-001', '김가나', '남', '1993-05', '2024-03-15'),
        ('RID-2024-002', '이다라', '여', '1975-11', '2024-03-15'),
        ('RID-2024-003', '박마바', '남', '1982-07', '2024-03-15'),
        ('RID-2024-004', '최사아', '여', '1990-04', '2024-03-16'),
        ('RID-2024-005', '강카타', '여', '1955-12', '2024-03-17'),
        ('RID-2024-006', '조파하', '남', '1979-06', '2024-03-17'),
        # 정자차·윤거너는 오기입으로 인해 미매칭 발생
    ]
    df = pd.DataFrame(data, columns=['환자ID', '환자명', '성별', '생년월일', '방문일'])
    out = OUT_DIR / 'rid_파일_샘플.xlsx'
    df.to_excel(out, index=False)
    print(f'생성: {out.name}')


def make_hubis_file():
    """
    hubis_샘플.xlsx — 외부 데이터소스 형식
    구조: 헤더 행 + 내부필드명 행(skip 대상) + 데이터
    매칭 키: bCODE, 병록번호 없음
    """
    col_names = ['donor_code', 'sex', 'birth_year', 'birth_month']
    field_row  = ['donor_code', 'sex', 'birth_year', 'birth_month']
    data = [
        [10005, '남', 2001, 9],
        [10008, '여', 1993, 2],
    ]
    df = pd.DataFrame([field_row] + data, columns=col_names)
    out = OUT_DIR / 'hubis_샘플.xlsx'
    df.to_excel(out, index=False)
    print(f'생성: {out.name}')


def make_supreme_fix_file():
    """
    슈프림_수정파일_샘플.xlsx — 슈프림 추출 데이터 형식
    시트명: '슈프림 추출 데이터(DEMO-000-000)'
    병록번호 기준 (R-ID 아님), 성별·생년월일 정정값 포함
    """
    data = [
        ('12340005', '정자차', '남', '2001-09-01'),
        ('12340008', '윤거너', '여', '1993-02-14'),
    ]
    df = pd.DataFrame(data, columns=['병록번호', '환자명', '성별', '생년월일'])
    out = OUT_DIR / '슈프림_수정파일_샘플.xlsx'
    with pd.ExcelWriter(out, engine='openpyxl') as w:
        df.to_excel(w, sheet_name='슈프림 추출 데이터(DEMO-000-000)', index=False)
        ws = w.sheets['슈프림 추출 데이터(DEMO-000-000)']
        _write_text_col(ws, 1)
    print(f'생성: {out.name}')


def make_master_db():
    """
    Master_DB_샘플.xlsx — 1단계 결과물 (2·3단계 입력)
    수집일지_샘플에서 식별번호 우선순위·가족 분류 적용 후 상태
    내부 식별키(생성ID)는 실제 실행 시 자동 생성되므로 샘플에서는 제외
    """
    통합_cols = ['bCODE', '식별번호', '병록번호', '이름', '개인번호', '접수일자']

    # 통합데이터: 유효 환자 8명 (가족 제외, O우선)
    통합 = [
        (10001, 'O-0001', '12340001', '김가나', 'M93.05', '2024-03-15'),
        (10002, 'O-0002', '12340002', '이다라', 'F75.11', '2024-03-15'),
        (10003, 'O-0007', '12340003', '박마바', 'M82.07', '2024-03-18'),
        (10004, 'O-0003', '12340004', '최사아', 'F90.04', '2024-03-16'),
        (10005, 'X-0002', '01234005', '정자차', 'M01.09', '2024-03-16'),
        (10006, 'O-0004', '12340006', '강카타', 'F55.12', '2024-03-17'),
        (10007, 'O-0005', '12340007', '조파하', 'M79.06', '2024-03-17'),
        (10008, 'O-0006', '12340008', '윤거너', 'F93.02', '2024-03-18'),
    ]
    df_통합 = pd.DataFrame(통합, columns=통합_cols)

    # 가족데이터: 보호자 1명
    가족 = [(10001, 'X-0003', '12340001', '김가나', 'M65.03', '2024-03-16', '보호자')]
    df_가족 = pd.DataFrame(가족, columns=통합_cols + ['비고'])

    # 오기데이터: 없음 (빈 시트)
    df_오기 = pd.DataFrame(columns=통합_cols + ['비고'])

    # 변경이력: 없음 (빈 시트)
    df_이력 = pd.DataFrame(columns=['병록번호', '이름', '구분', 'bCODE', '식별번호', '접수일자'])

    out = OUT_DIR / 'Master_DB_샘플.xlsx'
    with pd.ExcelWriter(out, engine='openpyxl') as w:
        df_통합.to_excel(w, sheet_name='통합데이터', index=False)
        df_오기.to_excel(w, sheet_name='오기데이터', index=False)
        df_가족.to_excel(w, sheet_name='가족데이터', index=False)
        df_이력.to_excel(w, sheet_name='변경이력', index=False)
        for sheet in ['통합데이터', '오기데이터', '가족데이터']:
            ws = w.sheets[sheet]
            col_idx = 통합_cols.index('병록번호') + 1
            _write_text_col(ws, col_idx)
    print(f'생성: {out.name}')


if __name__ == '__main__':
    make_collection_log()
    make_rid_file()
    make_hubis_file()
    make_supreme_fix_file()
    make_master_db()
    print('\n모든 샘플 데이터가 sample_data/ 폴더에 생성되었습니다.')

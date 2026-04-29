# -*- coding: utf-8 -*-
"""
핵심 로직 단위 테스트
GUI·파일 I/O 없이 순수 함수만 검증합니다.
"""
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).parent.parent))

from clinical_data_qc import (
    COL_BCODE,
    COL_HOSP_NUM,
    SUPREME_PREFIX,
    _build_email_body,
    _detect_columns,
    _read_supreme_fix,
    build_internal_key,
    clean_str,
    is_family_name,
    pad_hosp_num,
    parse_personal_no,
    resolve_year,
)


# =============================================================================
# clean_str
# =============================================================================

class TestCleanStr:
    def test_nan_returns_empty(self):
        import math
        assert clean_str(float('nan')) == ''

    def test_strips_whitespace(self):
        assert clean_str('  김철수  ') == '김철수'

    def test_removes_trailing_dot_zero(self):
        assert clean_str('12345.0') == '12345'

    def test_plain_string(self):
        assert clean_str('abc') == 'abc'

    def test_integer(self):
        assert clean_str(123) == '123'


# =============================================================================
# pad_hosp_num
# =============================================================================

class TestPadHospNum:
    def test_seven_digits_padded(self):
        assert pad_hosp_num('1234567') == '01234567'

    def test_eight_digits_unchanged(self):
        assert pad_hosp_num('12345678') == '12345678'

    def test_non_digit_unchanged(self):
        assert pad_hosp_num('ABC1234') == 'ABC1234'

    def test_float_input_cleaned(self):
        # clean_str이 .0 제거 → 7자리 숫자 → 0 패딩
        assert pad_hosp_num('1234567.0') == '01234567'


# =============================================================================
# resolve_year
# =============================================================================

class TestResolveYear:
    def test_below_threshold_is_2000s(self):
        assert resolve_year(24) == 2024

    def test_above_threshold_is_1900s(self):
        assert resolve_year(25) == 1925

    def test_zero_is_2000(self):
        assert resolve_year(0) == 2000

    def test_over_100_returned_as_is(self):
        assert resolve_year(1993) == 1993

    def test_99_is_1999(self):
        assert resolve_year(99) == 1999


# =============================================================================
# parse_personal_no
# =============================================================================

class TestParsePersonalNo:
    def test_male_93_05(self):
        g, y, m = parse_personal_no('M93.05')
        assert g == 'M' and y == 1993 and m == 5

    def test_female_75_11(self):
        g, y, m = parse_personal_no('F75.11')
        assert g == 'F' and y == 1975 and m == 11

    def test_2000s_birth(self):
        g, y, m = parse_personal_no('M01.09')
        assert y == 2001

    def test_lowercase_normalized(self):
        g, y, m = parse_personal_no('m82.07')
        assert g == 'M' and y == 1982

    def test_dash_separator(self):
        g, y, m = parse_personal_no('F90-04')
        assert g == 'F' and y == 1990 and m == 4

    def test_empty_returns_none(self):
        g, y, m = parse_personal_no('')
        assert g is None and y is None and m is None

    def test_no_birth_returns_none(self):
        g, y, m = parse_personal_no('M')
        assert y is None and m is None


# =============================================================================
# is_family_name
# =============================================================================

class TestIsFamilyName:
    def test_patient_name_false(self):
        is_fam, _ = is_family_name('김철수')
        assert not is_fam

    def test_guardian_keyword_true(self):
        is_fam, name = is_family_name('보호자 김철수')
        assert is_fam and name == '김철수'

    def test_husband_keyword_true(self):
        is_fam, _ = is_family_name('남편')
        assert is_fam

    def test_donor_keyword_true(self):
        is_fam, _ = is_family_name('공여자이영희')
        assert is_fam

    def test_exclude_keyword_false(self):
        is_fam, _ = is_family_name('개명김철수')
        assert not is_fam


# =============================================================================
# build_internal_key
# =============================================================================

class TestBuildInternalKey:
    def test_format(self):
        key = build_internal_key('M', 1993, 5, '김가나', '2024-03-15')
        assert key == 'M199305김20240315'

    def test_female(self):
        key = build_internal_key('F', 1975, 11, '이다라', '2024-03-15')
        assert key == 'F197511이20240315'

    def test_month_zero_padded(self):
        key = build_internal_key('M', 2001, 9, '정자차', '2024-03-16')
        assert key == 'M200109정20240316'

    def test_invalid_date_returns_empty(self):
        key = build_internal_key('M', 1993, 5, '김가나', 'not-a-date')
        assert key == ''

    def test_empty_name_uses_empty_char(self):
        key = build_internal_key('F', 1990, 4, '', '2024-03-16')
        assert key == 'F199004' + '20240316'


# =============================================================================
# _detect_columns
# =============================================================================

class TestDetectColumns:
    def _make_df(self, cols):
        import pandas as pd
        return pd.DataFrame(columns=cols)

    def test_detects_rid(self):
        df = self._make_df(['환자ID', '환자명', '성별', '생년월일', '방문일'])
        col_map = _detect_columns(df)
        assert col_map.get('rid') == '환자ID'

    def test_detects_gender(self):
        df = self._make_df(['R-ID', '이름', '성별', '생년월일', '방문일'])
        col_map = _detect_columns(df)
        assert col_map.get('gender') == '성별'

    def test_detects_date_variants(self):
        df = self._make_df(['RID', '성명', '성별', '생년월일', '검사일'])
        col_map = _detect_columns(df)
        assert col_map.get('date') == '검사일'

    def test_missing_columns(self):
        df = self._make_df(['A', 'B'])
        col_map = _detect_columns(df)
        assert 'rid' not in col_map


# =============================================================================
# _build_email_body
# =============================================================================

class TestBuildEmailBody:
    def test_contains_system_name(self):
        body = _build_email_body('휴비스쌤', 'bCODE', ['10001', '10002'], 'test.xlsx')
        assert '휴비스쌤' in body

    def test_contains_count(self):
        ids = [str(i) for i in range(60)]
        body = _build_email_body('슈프림', '병록번호', ids, 'test.xlsx')
        assert '60건' in body
        assert '외 10건' in body

    def test_contains_filename(self):
        body = _build_email_body('슈프림', '병록번호', ['12340001'], '매칭결과.xlsx')
        assert '매칭결과.xlsx' in body


# =============================================================================
# SUPREME_PREFIX 상수
# =============================================================================

def test_supreme_prefix_constant():
    assert SUPREME_PREFIX == '슈프림 추출 데이터'

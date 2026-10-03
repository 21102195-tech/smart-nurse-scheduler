"""스마트 널스 스케줄러 v2 — 관리자용 초안 생성 및 독립 검수.
실행: python -m streamlit run smart_nurse_scheduler_v2.py --server.address 127.0.0.1
핵심 함수는 Streamlit 없이 import하여 테스트할 수 있습니다.
"""
from __future__ import annotations
import calendar
import copy
import hashlib
import io
import math
import random
import re
import time
from dataclasses import dataclass, asdict
from datetime import date
from typing import Callable
import numpy as np
import pandas as pd

SHIFTS = ('D', 'E', 'N', 'DE', 'OFF', '교육')
DUTIES = ('D', 'E', 'N', 'DE')
ALIASES = {
    'D':'D', '데이':'D', 'DAY':'D', 'E':'E', '이브':'E', '이브닝':'E', 'EVENING':'E',
    'N':'N', 'NC1':'N', '나이트':'N', 'NIGHT':'N', 'DE':'DE',
    'OFF':'OFF', '오프':'OFF', '휴무':'OFF', '휴':'OFF',
    **{x:'OFF' for x in ['C','OF','OF1','V','H','H1','NV','BV','B','S','S10','S2','S3']},
    **{x:'교육' for x in ['CPE','P','PE','PE1','PL','교육','EDU']},
}
FORBIDDEN = (
    ('D','D','N','N','N'), ('D','D','D','N','N'), ('D','D','D','D','N'),
    ('D','E','N','N','N'), ('E','E','N','N','N'), ('D','D','E','N','N'),
    ('DE','DE','N','N','N'), ('DE','DE','DE','N','N'),
    ('DE','DE','DE','DE','N'), ('DE','E','N','N','N'), ('DE','DE','E','N','N'),
)
ISSUE_COLUMNS = ['구분','직원','날짜','항목','내용']

@dataclass(frozen=True)
class Settings:
    max_work: int = 5
    max_night: int = 6
    max_consecutive_night: int = 3
    night_keeper_target: int = 15
    two_off_after_five: bool = True
    no_single_night: bool = True
    group_balance: bool = True
    two_off_after_night: bool = True
    no_single_work: bool = True  # 희망조건: 필수조건과 구분
    education_counts_as_day: bool = False
    max_iterations: int = 60000
    time_limit: float = 30.0
    seed: int = 42

@dataclass
class Problem:
    source: pd.DataFrame
    year: int
    month: int
    names: list[str]
    groups: list[str]
    keepers: list[bool]
    allowed: list[set[str]]
    wanted: list[set[int]]
    fixed: np.ndarray
    fixed_codes: np.ndarray
    histories: list[list[str]]
    warnings: list[dict]
    requirements: dict[str, list[int]]
    row_indices: list[int]

    @property
    def days(self): return calendar.monthrange(self.year, self.month)[1]

    @property
    def dates(self): return [date(self.year, self.month, d) for d in range(1, self.days+1)]


def blank(value):
    return pd.isna(value) or str(value).strip().lower() in ('','nan','none','null','-')


def clean_label(value):
    if blank(value): return ''
    val = str(value).strip()
    return val[:-2] if val.endswith('.0') and val[:-2].isdigit() else val


def normalize_shift(value):
    if blank(value): return None
    token = str(value).strip().upper()
    if token not in ALIASES:
        raise ValueError(f'인식할 수 없는 근무코드 {value!r}. 코드 대응표에 추가하거나 템플릿을 수정하세요.')
    return ALIASES[token]


def align_headers(df):
    out = df.copy()

import streamlit as st
import pandas as pd
import numpy as np
import random
import copy
import io
import re
import os
import json
import requests
from datetime import datetime

# 1. 웹페이지 기본 설정
st.set_page_config(page_title="스마트 널스 스케쥴러", layout="wide")

# ⭐ [프로그램 제목 및 부제목 업데이트 - HTML 오타 완벽 해결!]
st.title("📊 스마트 널스 스케쥴러")
st.markdown("<h5 style='color: gray; font-weight: normal;'>수간호사 관리자 메뉴에서 1.작성규칙, 2. 초기 근무표 템플릿 업로드 후 AI 최종 근무표를 실행할 수 있습니다</h5>", unsafe_allow_html=True)
st.markdown("---")

# 2. 세션 메모리 초기화 (새로고침 시 데이터 및 신청 시간 로그 보존용)
if "schedule_df_state" not in st.session_state:
    st.session_state["schedule_df_state"] = None  # 취합 중인 데이터프레임
if "rules_df_state" not in st.session_state:
    st.session_state["rules_df_state"] = None  # 업로드된 작성 규칙
if "prev_df_state" not in st.session_state:
    st.session_state["prev_df_state"] = None  # 이전 달 근무표 기록
if "optimized_result" not in st.session_state:
    st.session_state["optimized_result"] = None  # 최종 생성된 근무표

# 구글 스프레드시트 API 주소 연동 (선생님의 진짜 주소 고정 적용)
GAS_URL = """https://script.google.com/macros/s/AKfycbzsbX7PygUpBz2kCssQ7x3vuL4rraz_3uM7lQyhSSUsdGIDtxJ08Dwyf3irDy7zn8ZI/exec"""

# 에러 실시간 판독 추적형 데이터 로딩 함수 (이중 그림자 백업 구동)
def fetch_google_sheet_data():
    try:
        bust_url = f"{GAS_URL}&t={random.random()}" if "?" in GAS_URL else f"{GAS_URL}?t={random.random()}"
        response = requests.get(bust_url, timeout=15)
        
        if response.status_code == 200:
            res_json = response.json()
            if isinstance(res_json, dict) and "error" in res_json:
                st.sidebar.error(f"🚨 구글 스크립트 연동 오류: {res_json['error']}")
                return []
            
            with open("google_backup.json", "w", encoding="utf-8") as f:
                json.dump(res_json, f, ensure_ascii=False)
            return res_json
        else:
            st.sidebar.warning("⚠️ 구글 지연으로 서버에 임시 백업된 데이터를 사용합니다.")
    except Exception as e:
        pass
        
    if os.path.exists("google_backup.json"):
        try:
            with open("google_backup.json", "r", encoding="utf-8") as f:
                return json.load(f)
        except: pass
    return []

# 서버 공유 설정값 로드 함수
def load_shared_config():
    default_config = {
        "rule_5_consec_off": True, "rule_no_single_night": True, "rule_group_balance": True, "rule_night_after_2_off": True,
        "limit_max_consec_work": 5, "limit_max_monthly_night": 6, "limit_max_consec_night": 3, "limit_daily_off_request": 2, "target_month": 8
    }
    if os.path.exists("config_shared.json"):
        try:
            with open("config_shared.json", "r", encoding="utf-8") as f:
                return json.load(f)
        except: pass
    return default_config

# 서버 공유 설정값 저장 함수
def save_shared_config(config):
    try:
        with open("config_shared.json", "w", encoding="utf-8") as f:
            json.dump(config, f, ensure_ascii=False, indent=4)
    except: pass

shared_config = load_shared_config()

# ----------------- ⭐ [서버 하드디스크 영구 저장소 자동 로드 시스템] -----------------
if st.session_state["schedule_df_state"] is None and os.path.exists("template_shared.xlsx"):
    try:
        loaded_df = pd.read_excel("template_shared.xlsx")
        st.session_state["schedule_df_state"] = loaded_df
    except: pass

if st.session_state["rules_df_state"] is None and os.path.exists("rules_shared.xlsx"):
    try:
        st.session_state["rules_df_state"] = pd.read_excel("rules_shared.xlsx")
    except: pass

if st.session_state["prev_df_state"] is None and os.path.exists("prev_month_shared.xlsx"):
    try:
        st.session_state["prev_df_state"] = pd.read_excel("prev_month_shared.xlsx")
    except: pass

# 구글 시트 실시간 신청내역 자동 병합 상태 추적
if "synced_once" not in st.session_state:
    st.session_state["synced_once"] = False

# [지능형 가변 이름 및 야간전담 판정 파서]
def parse_nurse_row(name, group):
    name_str = str(name).strip() if pd.notna(name) else ""
    group_str = str(group).strip() if pd.notna(group) else ""
    
    if not name_str or name_str.lower() in ['nan', '이름', '그룹', 'none', 'null', '', '토', '일', '월', '화', '수', '목', '금', '요일']:
        return None, False
        
    if "듀티별" in group_str or "인원수" in group_str:
        return None, False
        
    is_keeper = "야간전담" in name_str or "야간전담" in group_str
    
    clean_name = name_str
    if clean_name.endswith('.0'):
        clean_name = clean_name[:-2]
        
    return clean_name, is_keeper

# [보정 함수] 헤더 밀림 방지 보정
def load_and_align_headers(df):
    df.columns = [str(col).strip() for col in df.columns]
    if '그룹' in df.columns:
        df.columns = [str(col).strip().replace('.0', '') for col in df.columns]
        return df
    
    for idx in range(min(5, len(df))):
        row_vals = [str(x).strip() for x in df.iloc[idx].values]
        if '그룹' in row_vals or '그룹 ' in row_vals:
            new_cols = []
            for col_val in df.iloc[idx].values:
                val = str(col_val).strip() if pd.notna(col_val) else ""
                if val.endswith('.0'):



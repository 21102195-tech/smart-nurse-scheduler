import streamlit as st
import pandas as pd
import numpy as np
import random
import copy
import io
import re
import os
import requests
from datetime import datetime

# 1. 웹페이지 기본 설정
st.set_page_config(page_title="간호사 스마트 3교대 스케줄러", layout="wide")

st.title("🏥 3교대 간호사 스마트 근무표 작성 & 신청 시스템")
st.subheader("모바일 원티드 신청 포털과 수간호사 제어판이 구글 데이터베이스를 통해 실시간 연동됩니다.")

# ⭐ [구글 API 주소 연동]: 구글에서 복사해 둔 웹 앱 URL 주소를 여기에 넣어주세요!
GAS_URL = "https://script.google.com/macros/s/AKfycbzsbX7PygUpBz2kCssQ7x3vuL4rraz_3uM7lQyhSSUsdGIDtxJO8Dwyf3irDy7zn8ZI/exec"

# 구글 데이터 읽어오기 함수
def fetch_google_sheet_data():
    try:
        response = requests.get(GAS_URL)
        if response.status_code == 200:
            return response.json()
    except Exception as e:
        st.sidebar.error(f"구글 연동 실패: {e}")
    return []

# 2. ⭐ [서버 디스크 공유 데이터베이스 엔진]: 
# 다른 사용자나 스마트폰이 접속해도 서버에 저장된 엑셀 파일을 읽어와 실시간 공유합니다!
if "schedule_df_state" not in st.session_state:
    if os.path.exists("template_shared.xlsx"):
        try:
            st.session_state["schedule_df_state"] = pd.read_excel("template_shared.xlsx")
        except:
            st.session_state["schedule_df_state"] = None
    else:
        st.session_state["schedule_df_state"] = None

# 자동 싱크 로직 (최초 접속 시 구글 시트의 최신 신청 내역 자동 결합)
if "synced_once" not in st.session_state:
    st.session_state["synced_once"] = False

if st.session_state["schedule_df_state"] is not None and not st.session_state["synced_once"]:
    sheet_data = fetch_google_sheet_data()
    if len(sheet_data) > 0:
        df_temp = st.session_state["schedule_df_state"].copy()
        for item in sheet_data:
            nurse_name = str(item['name']).strip()
            wanted_val = str(item['wanted_off']).strip()
            for idx, row in df_temp.iterrows():
                nurse_id, _ = parse_nurse_row(row['이름'], row['그룹'])
                if str(nurse_id) == nurse_name:
                    df_temp.loc[idx, '원티드 오프'] = wanted_val
                    break
        st.session_state["schedule_df_state"] = df_temp
    st.session_state["synced_once"] = True

# 3. 사이드바 - 관리자 메뉴 및 업로드 버튼
st.sidebar.header("⚙️ 수간호사 관리자 메뉴")
uploaded_rules = st.sidebar.file_uploader("1. 작성 규칙 파일 업로드 (xlsx/csv)", type=["xlsx", "csv"])
uploaded_schedule = st.sidebar.file_uploader("2. 초기 근무표 템플릿 업로드 (xlsx/csv)", type=["xlsx", "csv"])
uploaded_prev_month = st.sidebar.file_uploader("3. 이전 달 근무표 업로드 (선택사항)", type=["xlsx", "csv"])

# 🛠 [부서별 맞춤 근무 조건 설정]
st.sidebar.markdown("---")
st.sidebar.header("🛠️ 부서별 맞춤 근무 조건 설정")

rule_5_consec_off = st.sidebar.toggle("5일 연속 근무 시 후속 2 OFF 강제 보장", value=True)
rule_no_single_night = st.sidebar.toggle("단독 나이트(하루짜리 N) 금지", value=True)
rule_group_balance = st.sidebar.toggle("듀티별 그룹(A/B/C) 균등 배치 적용", value=True)
rule_night_after_2_off = st.sidebar.toggle("야간 근무(N) 후 2일 OFF 필수 부여", value=True)

limit_max_consec_work = st.sidebar.slider("최대 연속 근무 일수 제한", 0, 5, 5)
limit_max_monthly_night = st.sidebar.slider("월간 인당 최대 나이트(N) 개수", 0, 7, 6)
limit_max_consec_night = st.sidebar.slider("최대 연속 나이트(N) 제한", 2, 3, 3)
limit_daily_off_request = st.sidebar.slider("📢 하루 최대 오프 신청 허용 인원", 1, 4, 2)

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
                    val = val[:-2]
                new_cols.append(val)
            df.columns = new_cols
            df = df.iloc[idx+1:].reset_index(drop=True)
            break
    return df

# [이전 달 정보 추출기]
def extract_nurse_history(prev_df, nurse_name):
    try:
        df_aligned = load_and_align_headers(prev_df)
        day_cols = []
        for col in df_aligned.columns:
            col_str = str(col).strip()
            if col_str.isdigit():
                day_cols.append(int(col_str))
                
        day_cols.sort()
        last_7_days = day_cols[-7:]
        
        for idx, row in df_aligned.iterrows():
            name = row['이름']
            group = row['그룹'] if '그룹' in df_aligned.columns else ""
            nurse_id, _ = parse_nurse_row(name, group)
            if nurse_id is not None and str(nurse_id) == str(nurse_name):
                shifts = []
                for d in last_7_days:
                    col_name = str(d) if str(d) in df_aligned.columns else (int(d) if int(d) in df_aligned.columns else d)
                    val = str(row[col_name]).strip().upper() if pd.notna(row[col_name]) else "OFF"
                    if val in ['D', '데이', 'DAY']: val = 'D'
                    elif val in ['E', '이브', '이브닝', 'EVENING']: val = 'E'
                    elif val in ['N', '나이트', 'NIGHT']: val = 'N'
                    elif val in ['DE']: val = 'DE'
                    elif '교육' in val: val = '교육'
                    else: val = 'OFF'
                    shifts.append(val)
                return shifts
    except Exception as e:
        pass
    return ['OFF'] * 7

# 벌점 계산 수식 정의
def get_nurse_penalty(row_current, i, nurse_wanted_off_set, num_days, forbidden_5_patterns, is_night_keeper, history, target_N_min, target_N_max, target_OFF_min, target_OFF_max, is_fixed_row):
    row_norm = []
    for x in (list(history) + list(row_current)):
        if x == '교육':
            row_norm.append('D')
        else:
            row_norm.append(x)
            
    num_total = len(row_norm)
    history_len = len(history)
    
    penalty = 0
    # 규칙 1: 원티드 오프 준수
    for d in range(history_len, num_total):
        current_day = d - history_len + 1
        if current_day in nurse_wanted_off_set and row_norm[d] != 'OFF':
            if not is_fixed_row[current_day - 1]:
                penalty += 1000000
            
    if is_night_keeper:
        for d in range(history_len, num_total):
            if row_norm[d] in ['D', 'E', 'DE']:
                penalty += 1000000
        total_N = sum(1 for x in row_norm[history_len:] if x == 'N')
        if total_N != 15:
            penalty += abs(total_N - 15) * 500000
    else:
        # 규칙 2: 한 달 밤근무(N) 개수 균등화 (수학적으로 실시간 계산된 타겟 반영)
        total_N = sum(1 for x in row_norm[history_len:] if x == 'N')
        if limit_max_monthly_night == 0:
            if total_N > 0:
                penalty += total_N * 1000000
        else:
            if total_N < target_N_min or total_N > target_N_max:
                mid = (target_N_min + target_N_max) / 2.0
                half_w = (target_N_max - target_N_min) / 2.0
                penalty += (abs(total_N - mid) - half_w) * 500000
            
        # 규칙 3: 한 달 총 휴무(OFF) 개수 자동 조정
        total_OFF = sum(1 for x in row_norm[history_len:] if x == 'OFF')
        if total_OFF < target_OFF_min or total_OFF > target_OFF_max:
            mid = (target_OFF_min + target_OFF_max) / 2.0
            half_w = (target_OFF_max - target_OFF_min) / 2.0
            penalty += (abs(total_OFF - mid) - half_w) * 400000
        
    consec_work = 0
    consec_N = 0
    for d in range(num_total):
        shift = row_norm[d]
        if shift != 'OFF':
            consec_work += 1
            if limit_max_consec_work > 0:
                if consec_work > limit_max_consec_work and d >= history_len:
                    penalty += (consec_work - limit_max_consec_work) * 500000
        else:
            # [조건 On/Off] 5일 연속 근무 후 2 OFF 연속 보장
            if rule_5_consec_off:
                if consec_work == 5:
                    if d + 1 < num_total:
                        if row_norm[d+1] != 'OFF' and (d+1) >= history_len:
                            penalty += 300000
            consec_work = 0
            
        if shift == 'N':
            consec_N += 1
            if consec_N > limit_max_consec_night and d >= history_len:  
                penalty += (consec_N - limit_max_consec_night) * 500000
        else:
            consec_N = 0
            
        if d < num_total - 1:
            next_shift = row_norm[d+1]
            if (d+1) >= history_len:
                if shift == 'E' and next_shift in ['D', 'DE']:
                    penalty += 500000
                if shift == 'N' and next_shift in ['D', 'E', 'DE']:
                    penalty += 500000
                
        # [조건 On/Off] 야간 근무(N) 후 2일 OFF 필수 부여
        if shift == 'N' and rule_night_after_2_off:
            if d < num_total - 1:
                if row_norm[d+1] != 'N':
                    if row_norm[d+1] != 'OFF' and (d+1) >= history_len:
                        penalty += 500000
                    if d < num_total - 2:
                        if row_norm[d+2] != 'OFF' and (d+2) >= history_len:
                            penalty += 500000
                            
        if d <= num_total - 5:
            pat = list(row_norm[d:d+5])
            if (d+4) >= history_len:
                if pat in forbidden_5_patterns:
                    penalty += 500000
                
    # [조건 On/Off] 단독 나이트 방지
    if rule_no_single_night:
        for d in range(history_len, num_total):
            if row_norm[d] == 'N':
                prev_is_N = (d > 0 and row_norm[d-1] == 'N')
                next_is_N = (d < num_total - 1 and row_norm[d+1] == 'N')
                if not prev_is_N and not next_is_N:
                    penalty += 300000
                
    return penalty

# 그룹별 일별 듀티 인원수 균형도 평가 함수
def get_day_penalty(col, num_nurses, nurse_groups):
    if not rule_group_balance:
        return 0
        
    penalty = 0
    for duty in ['D', 'E', 'N', 'OFF']:
        counts = {'A': 0, 'B': 0, 'C': 0}
        for i in range(num_nurses):
            if col[i] == duty:
                g = nurse_groups[i] if nurse_groups[i] in ['A', 'B', 'C'] else 'A'
                counts[g] += 1
        tot = sum(counts.values())
        
        if tot == 3:
            penalty += (abs(counts['A'] - 1) + abs(counts['B'] - 1) + abs(counts['C'] - 1)) * 50
        elif tot == 4:
            penalty += (max(0, counts['A'] - 2) + max(0, 1 - counts['A'])) * 50
            penalty += abs(counts['B'] - 1) * 50
            penalty += (max(0, counts['C'] - 2) + max(0, 1 - counts['C'])) * 50
        elif tot == 6:
            penalty += (abs(counts['A'] - 2) + abs(counts['B'] - 2) + abs(counts['C'] - 2)) * 50
        elif tot == 7:
            penalty += (max(0, counts['A'] - 3) + max(0, 2 - counts['A'])) * 50
            penalty += abs(counts['B'] - 2) * 50
            penalty += (max(0, counts['C'] - 3) + max(0, 2 - counts['C'])) * 50
        elif tot == 8:
            penalty += (abs(counts['A'] - 3) + abs(counts['B'] - 2) + abs(counts['C'] - 3)) * 50
            
    return penalty

# 하이브리드 고정형 초기 스케줄 생성 함수
def initialize_schedule_hybrid(num_nurses, num_days, requirements, is_fixed, fixed_shifts):
    sched = np.empty((num_nurses, num_days), dtype=object)
    for d in range(num_days):
        nD = requirements['D'][d]
        nE = requirements['E'][d]
        nN = requirements['N'][d]
        nDE = requirements['DE'][d] if 'DE' in requirements else 0
        
        pD = sum(1 for i in range(num_nurses) if is_fixed[i, d] and fixed_shifts[i, d] == 'D')
        pE = sum(1 for i in range(num_nurses) if is_fixed[i, d] and fixed_shifts[i, d] == 'E')
        pN = sum(1 for i in range(num_nurses) if is_fixed[i, d] and fixed_shifts[i, d] == 'N')
        pDE = sum(1 for i in range(num_nurses) if is_fixed[i, d] and fixed_shifts[i, d] == 'DE')
        pEDU = sum(1 for i in range(num_nurses) if is_fixed[i, d] and fixed_shifts[i, d] == '교육')
        
        rem_D = max(0, nD - pD - pEDU)
        rem_E = max(0, nE - pE)
        rem_N = max(0, nN - pN)
        rem_DE = max(0, nDE - pDE)
        
        num_unfixed = sum(1 for i in range(num_nurses) if not is_fixed[i, d])
        rem_OFF = max(0, num_unfixed - rem_D - rem_E - rem_N - rem_DE)
        
        pool = ['D'] * rem_D + ['E'] * rem_E + ['N'] * rem_N + ['DE'] * rem_DE + ['OFF'] * rem_OFF
        random.shuffle(pool)
        
        pool_idx = 0
        for i in range(num_nurses):
            if is_fixed[i, d]:
                sched[i, d] = fixed_shifts[i, d]
            else:
                sched[i, d] = pool[pool_idx]
                pool_idx += 1
    return sched

# 구글 스프레드시트 데이터베이스로부터 실시간 신청 오프 내역 동기화 함수
def fetch_google_sheet_data():
    try:
        response = requests.get(GAS_URL)
        if response.status_code == 200:
            return response.json()
    except Exception as e:
        pass
    return []

# ⚙️ 수간호사 신규 업로드 시 세션 상태 등록 및 서버 디스크 영구 백업 기능
if uploaded_schedule:
    try:
        if uploaded_schedule.name.endswith('xlsx'):
            raw_df = pd.read_excel(uploaded_schedule)
        else:
            raw_df = pd.read_csv(uploaded_schedule, encoding='utf-8-sig')
        schedule_df = load_and_align_headers(raw_df)
        
        # 1. 서버 하드디스크에 영구 고정 저장 (절대 유실되지 않음!)
        schedule_df.to_excel("template_shared.xlsx", index=False)
        st.session_state["schedule_df_state"] = schedule_df
        
        # 2. 구글 스프레드시트에 최초 업로드된 간호사 명단을 가입 처리
        sheet_data = fetch_google_sheet_data()
        existing_names = [str(item['name']).strip() for item in sheet_data] if sheet_data else []
        
        for idx, row in schedule_df.iterrows():
            nurse_id, _ = parse_nurse_row(row['이름'], row['그룹'])
            if nurse_id is not None and str(nurse_id) not in existing_names:
                requests.post(GAS_URL, json={"name": str(nurse_id), "wanted_off": "-"})
        st.sidebar.success("✅ 새로운 템플릿이 서버와 데이터베이스에 성공적으로 등록되었습니다!")
    except Exception as e:
        st.sidebar.error(f"템플릿 파일 읽기 에러: {e}")

# 5. 메인 인터페이스부
# (이제 schedule_df_state가 None이 아니므로 모든 동료들이 자유롭게 접속해 화면을 정상적으로 볼 수 있습니다!)
if st.session_state["schedule_df_state"] is not None:
    menu = st.radio("👉 사용 유형을 선택하세요", ["🙋‍♀️ [간호사용] 원티드 오프 신청", "⚙️ [수간호사용] 관리자 제어판"], horizontal=True)
    
    # ----------------- [뷰 1] 간호사용 원티드 오프 신청 포털 -----------------
    if menu == "🙋‍♀️ [간호사용] 원티드 오프 신청":
        st.markdown("---")
        st.write("### 📅 원하는 휴무일 직접 신청")
        st.info("💡 원하는 휴가 날짜를 지정하고 신청을 완료하면, 구글 스프레드시트에 신청 시간과 함께 실시간 저장됩니다.")
        
        # 서버 디스크에서 템플릿 읽기
        df_temp = st.session_state["schedule_df_state"].copy()
        sheet_data = fetch_google_sheet_data()
        
        # 간호사 목록을 서버 엑셀에서 원천 추출하므로 절대 비어있지 않습니다!
        nurse_names = []
        for idx, row in df_temp.iterrows():
            nurse_id, _ = parse_nurse_row(row['이름'], row['그룹'])
            if nurse_id is not None:
                nurse_names.append(nurse_id)
        
        # 구글 시트에 실시간 쌓인 오프 갯수 계산
        day_request_counts = {}
        for item in sheet_data:
            wanted = item['wanted_off']
            if pd.notna(wanted) and wanted.strip() != '' and wanted.strip() != 'nan':
                for x in str(wanted).split(','):
                    match = re.search(r'^\s*(\d+)', x.strip())
                    if match:
                        d = int(match.group(1))
                        day_request_counts[d] = day_request_counts.get(d, 0) + 1
                        
        st.write("#### 📢 8월 일자별 오프 신청 현황 및 마감 잔여석")
        cols = st.columns(10)
        for day in range(1, 32):
            col_idx = (day - 1) % 10
            count = day_request_counts.get(day, 0)
            rem = max(0, limit_daily_off_request - count)
            with cols[col_idx]:
                if rem == 0:
                    st.markdown(f"**{day}일**\n\n🔴 **마감**")
                else:
                    st.markdown(f"**{day}일**\n\n🟢 {rem}석")
                    
        st.markdown("---")
        
        col1, col2 = st.columns(2)
        with col1:
            selected_nurse = st.selectbox("1. 본인의 이름을 선택하세요", nurse_names, format_func=lambda x: f"{x}번 간호사" if str(x).replace('.0', '').isdigit() else f"{x} 간호사")
        with col2:
            selected_offs = st.multiselect("2. 희망 휴무일을 복수 선택하세요 (마감된 날짜 제외)", list(range(1, 32)))
            
        if st.button("📝 원티드 오프 신청하기", type="primary"):
            if len(selected_offs) > 0:
                closed_days = []
                for d in selected_offs:
                    # 타인 신청 인원수 구글 시트에서 즉시 계산
                    other_count = sum(1 for item in sheet_data if d in [int(re.search(r'^\s*(\d+)', x.strip()).group(1)) for x in str(item['wanted_off']).split(',') if re.search(r'^\s*(\d+)', x.strip())] and str(item['name']).strip() != str(selected_nurse))
                    if other_count >= limit_daily_off_request:
                        closed_days.append(d)
                        
                if len(closed_days) > 0:
                    st.error(f"❌ 신청 실패: {closed_days}일은 이미 다른 동료들에 의해 선착순 마감되었습니다!")
                else:
                    current_time_str = datetime.now().strftime("%m/%d %H:%M")
                    detail_offs_list = [f"{d}({current_time_str})" for d in sorted(selected_offs)]
                    offs_formatted_str = ", ".join(detail_offs_list)
                    
                    try:
                        # 구글 스프레드시트에 실시간 전송 저장!
                        payload = {"name": str(selected_nurse), "wanted_off": offs_formatted_str}
                        res = requests.post(GAS_URL, json=payload)
                        if res.status_code == 200:
                            st.success(f"🎉 신청 성공! {selected_nurse} 간호사님 오프 수집 완료: [{offs_formatted_str}]")
                            time.sleep(1.5)
                            st.rerun()
                    except Exception as e:
                        st.error(f"구글 서버 전송 중 오류 발생: {e}")
            else:
                st.warning("날짜를 선택해 주세요.")

    # ----------------- [뷰 2] 수간호사용 비밀번호 관리자 제어판 -----------------
    elif menu == "⚙️ [수간호사용] 관리자 제어판":
        st.markdown("---")
        st.sidebar.markdown("### 🔐 관리자 인증")
        admin_password = st.sidebar.text_input("수간호사 비밀번호를 입력하세요", type="password")
        
        if admin_password == "1234":
            st.sidebar.success("🔑 관리자 인증 성공!")
            st.write("### 🛠️ 수간호사 근무표 마스터 통제 제어판")
            
            # 수간호사용 실시간 오프 신청 모니터링 및 작성 제어
            col_tab1, col_tab2 = st.tabs(["📋 실시간 오프 신청 현황판", "📅 AI 최종 근무표 생성"])
            
            with col_tab1:
                st.write("### 📋 현재 구글 시트에 취합된 실시간 오프 현황")
                
                if st.button("🔄 새로운 신청 내역 실시간 동기화"):
                    st.session_state["optimized_result"] = None
                    
                sheet_data = fetch_google_sheet_data()
                df_temp = st.session_state["schedule_df_state"].copy()
                
                # 구글 시트 데이터를 로컬 뷰에 실시간 덮어쓰기
                for item in sheet_data:
                    nurse_name = item['name']
                    wanted_val = item['wanted_off']
                    for idx, row in df_temp.iterrows():
                        nurse_id, _ = parse_nurse_row(row['이름'], row['그룹'])
                        if str(nurse_id) == str(nurse_name):
                            df_temp.loc[idx, '원티드 오프'] = wanted_val
                            break
                st.session_state["schedule_df_state"] = df_temp
                st.dataframe(df_temp)
                
            with col_tab2:
                st.write("### 🚀 AI 3교대 스마트 근무표 작성 시작")
                max_iter_val = st.slider("최대 탐색 횟수", 10000, 150000, 60000, step=10000)
                
                if st.button("🔮 최종 AI 근무표 생성 시작", type="primary"):
                    if uploaded_rules is None:
                        st.error("에러: 규칙 파일을 업로드해 주세요.")
                    else:
                        with st.spinner("야간전담 분류 및 이전 달 근태 연동 연산 중..."):
                            df_clean = st.session_state["schedule_df_state"].copy()
                            df_clean = df_clean.replace(r'^\s*$', np.nan, regex=True)
                            df_clean['그룹'] = df_clean['그룹'].ffill()
                            
                            # 이전 달 로드
                            prev_df = None
                            if uploaded_prev_month is not None:
                                try:
                                    prev_df = pd.read_excel(uploaded_prev_month) if uploaded_prev_month.name.endswith('xlsx') else pd.read_csv(uploaded_prev_month, encoding='utf-8-sig')
                                except: pass
                            
                            # 1. 간호사 추출 및 야간전담 분류 + 이전달 근무 기록 매핑
                            nurses = []
                            is_night_keepers = []
                            nurse_histories = []
                            
                            for idx, row in df_clean.iterrows():
                                name = row['이름']
                                group = row['그룹']
                                nurse_id, is_keeper = parse_nurse_row(name, group)
                                
                                if nurse_id is not None:
                                    wanted = row['원티드 오프']
                                    wanted_days = []
                                    # 정규식 일자 파싱
                                    if pd.notna(wanted) and str(wanted).strip() != '-':
                                        for x in str(wanted).split(','):
                                            match = re.search(r'^\s*(\d+)', x.strip())
                                            if match:
                                                wanted_days.append(int(match.group(1)))
                                    
                                    history = ['OFF'] * 7
                                    if prev_df is not None:
                                        history = extract_nurse_history(prev_df, nurse_id)
                                        
                                    nurses.append({
                                        'id': nurse_id,
                                        'group': row['그룹'],
                                        'wanted_off': wanted_days,
                                        'row_idx': idx,
                                        'is_keeper': is_keeper
                                    })
                                    is_night_keepers.append(is_keeper)
                                    nurse_histories.append(history)
                                    
                            # 요구수 파싱
                            requirements = {}
                            default_values = {
                                'D': [3, 3, 4, 4, 4, 4, 4, 3, 3, 4, 4, 4, 4, 4, 3, 3, 3, 4, 4, 4, 4, 3, 3, 4, 4, 4, 4, 4, 3, 3, 4],
                                'E': [3, 3, 4, 4, 4, 4, 4, 3, 3, 4, 4, 4, 4, 4, 3, 3, 3, 4, 4, 4, 4, 3, 3, 4, 4, 4, 4, 4, 3, 3, 4],
                                'N': [3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3],
                                'DE': [0] * 31
                            }
                            start_idx = None
                            for idx, row in enumerate(df_clean.values):
                                row_str = " ".join([str(x) for x in row])
                                if "듀티별" in row_str or "인원수" in row_str:
                                    start_idx = idx
                                    break
                            if start_idx is not None:
                                for i in range(4):
                                    if start_idx + i < len(df_clean):
                                        row = df_clean.iloc[start_idx + i]
                                        duty = None
                                        for col in df_clean.columns:
                                            if str(col).strip() not in [str(d) for d in range(1, 32)]:
                                                val_str = str(row[col]).strip().upper()
                                                if val_str in ['D', 'E', 'N', 'DE']:
                                                    duty = val_str
                                                    break
                                        if duty:
                                            day_values = []
                                            for d in range(1, 32):
                                                col_name = str(d) if str(d) in df_clean.columns else (int(d) if int(d) in df_clean.columns else d)
                                                val = int(float(row[col_name])) if pd.notna(row[col_name]) and str(row[col_name]).strip() != '' else default_values[duty][d-1]
                                                day_values.append(val)
                                            requirements[duty] = day_values

                            for duty in ['D', 'E', 'N', 'DE']:
                                if duty not in requirements: requirements[duty] = default_values[duty]
                                
                            num_keepers = sum(is_night_keepers)
                            num_normal = num_nurses - num_keepers
                            total_shifts_required = sum(requirements['D']) + sum(requirements['E']) + sum(requirements['N']) + sum(requirements['DE'])
                            total_N_required = sum(requirements['N'])
                            total_keeper_N = num_keepers * 15
                            total_normal_N = max(0, total_N_required - total_keeper_N)
                            total_normal_shifts = max(0, total_shifts_required - (num_keepers * 15))
                            total_normal_OFF = max(0, (num_normal * 31) - total_normal_shifts)
                            
                            if num_normal > 0:
                                avg_normal_N = total_normal_N / num_normal
                                avg_normal_OFF = total_normal_OFF / num_normal
                                target_N_min = int(avg_normal_N)
                                target_N_max = min(int(avg_normal_N) + 1, limit_max_monthly_night)
                                target_N_min = min(target_N_min, target_N_max)
                                target_OFF_min, target_OFF_max = int(avg_normal_OFF), int(avg_normal_OFF) + 1
                            else:
                                target_N_min, target_N_max, target_OFF_min, target_OFF_max = 0, 0, 0, 0
                                
                            is_fixed = np.zeros((num_nurses, num_days), dtype=bool)
                            fixed_shifts = np.empty((num_nurses, num_days), dtype=object)
                            for i, nurse in enumerate(nurses):
                                row_idx = nurse['row_idx']
                                row = df_clean.iloc[row_idx]
                                for d in range(num_days):
                                    col_name = str(d+1) if str(d+1) in df_clean.columns else (int(d+1) if int(d+1) in df_clean.columns else d+1)
                                    raw_val = str(row[col_name]).strip().upper() if pd.notna(row[col_name]) else ""
                                    val = ""
                                    if raw_val in ['D', '데이']: val = 'D'
                                    elif raw_val in ['E', '이브', '이브닝']: val = 'E'
                                    elif raw_val in ['N', '나이트']: val = 'N'
                                    elif raw_val in ['DE']: val = 'DE'
                                    elif raw_val in ['OFF', '오프', '휴무']: val = 'OFF'
                                    elif '교육' in raw_val: val = '교육'
                                    if val in ['D', 'E', 'N', 'DE', 'OFF', '교육']:
                                        is_fixed[i, d] = True
                                        fixed_shifts[i, d] = val
                                        
                            sched = initialize_schedule_hybrid(num_nurses, num_days, requirements, is_fixed, fixed_shifts)
                            nurse_wanted_off = [set(n['wanted_off']) for n in nurses]
                            
                            row_penalties = [get_nurse_penalty(sched[i], i, nurse_wanted_off[i], num_days, forbidden_5_patterns, is_night_keepers[i], nurse_histories[i], target_N_min, target_N_max, target_OFF_min, target_OFF_max, is_fixed[i]) for i in range(num_nurses)]
                            col_penalties = [get_day_penalty(sched[:, d], num_nurses, nurse_groups) for d in range(num_days)]
                            total_penalty = sum(row_penalties) + sum(col_penalties)
                            
                            best_sched = copy.deepcopy(sched)
                            best_penalty = total_penalty
                            best_hard = sum(row_penalties)
                            temp = 25.0
                            cooling_rate = 0.9999
                            
                            for step in range(max_iter_val):
                                d = random.randint(0, num_days - 1)
                                i1 = random.randint(0, num_nurses - 1)
                                i2 = random.randint(0, num_nurses - 1)
                                while i1 == i2: i2 = random.randint(0, num_nurses - 1)
                                if is_fixed[i1, d] or is_fixed[i2, d]: continue
                                if sched[i1, d] == sched[i2, d]: continue
                                
                                old_shift_i1, old_shift_i2 = sched[i1, d], sched[i2, d]
                                old_row_pen_i1, old_row_pen_i2 = row_penalties[i1], row_penalties[i2]
                                old_col_pen = col_penalties[d]
                                
                                sched[i1, d], sched[i2, d] = old_shift_i2, old_shift_i1
                                
                                new_row_pen_i1 = get_nurse_penalty(sched[i1], i1, nurse_wanted_off[i1], num_days, forbidden_5_patterns, is_night_keepers[i1], nurse_histories[i1], target_N_min, target_N_max, target_OFF_min, target_OFF_max, is_fixed[i1])
                                new_row_pen_i2 = get_nurse_penalty(sched[i2], i2, nurse_wanted_off[i2], num_days, forbidden_5_patterns, is_night_keepers[i2], nurse_histories[i2], target_N_min, target_N_max, target_OFF_min, target_OFF_max, is_fixed[i2])
                                new_col_pen = get_day_penalty(sched[:, d], num_nurses, nurse_groups)
                                
                                new_total_penalty = (total_penalty - old_row_pen_i1 - old_row_pen_i2 - old_col_pen + new_row_pen_i1 + new_row_pen_i2 + new_col_pen)
                                delta = new_total_penalty - total_penalty
                                
                                accept = False
                                if delta < 0: accept = True
                                elif temp > 0.05: accept = (random.random() < np.exp(-delta / temp))
                                
                                if accept:
                                    total_penalty = new_total_penalty
                                    row_penalties[i1] = new_row_pen_i1
                                    row_penalties[i2] = new_row_pen_i2
                                    col_penalties[d] = new_col_pen
                                    if total_penalty < best_penalty:
                                        best_sched = copy.deepcopy(sched)
                                        best_penalty = total_penalty
                                        best_hard = sum(row_penalties)
                                else:
                                    sched[i1, d], sched[i2, d] = old_shift_i1, old_shift_i2
                                    
                                temp *= cooling_rate
                                if best_hard == 0 and step > 45000: break
                                    
                            for i, nurse in enumerate(nurses):
                                row_idx = nurse['row_idx']
                                for d in range(num_days):
                                    col_name = str(d+1) if str(d+1) in df_clean.columns else (int(d+1) if int(d+1) in df_clean.columns else d+1)
                                    df_clean.loc[row_idx, col_name] = best_sched[i, d]
                                    
                            st.session_state["optimized_result"] = df_clean
                            st.balloons()
                            
                    if st.session_state["optimized_result"] is not None:
                        st.write("### 🎉 생성 완료된 최종 근무표")
                        st.dataframe(st.session_state["optimized_result"])
                        
                        towrite = io.BytesIO()
                        st.session_state["optimized_result"].to_excel(towrite, index=False, header=True)
                        towrite.seek(0)
                        
                        st.download_button(
                            label="📥 최종 근무표 Excel 다운로드",
                            data=towrite,
                            file_name="최종_근무표_야간전담_이전달연동.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
        else:
            st.warning("⚠️ 관리자 비밀번호가 틀렸거나 입력되지 않았습니다.")
            st.info("💡 사이드바의 자물쇠 입력창에 올바른 비밀번호를 입력해 주시면 관리 제어판이 열립니다.")
else:
    # ⚠️ 최초 등록 전에는 관리자 제어판의 업로드 창만 가이드로 띄워둡니다.
    st.warning("📊 아직 등록된 템플릿(간호사 목록)이 없습니다.")
    st.info("💡 먼저 **[⚙️ 관리자 제어판]** 탭을 누르시고, 비밀번호 **`1234`**를 입력하신 후 사이드바 메뉴에 '2. 초기 근무표 템플릿 업로드' 파일을 최소 한 번 이상 올려주세요! 등록되는 즉시 모든 사용자에게 오프 신청 포털이 자동으로 개설됩니다.")

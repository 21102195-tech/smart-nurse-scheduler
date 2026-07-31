import streamlit as st
import pandas as pd
import numpy as np
import random
import copy
import io
import re
from datetime import datetime

# 1. 웹페이지 기본 설정
st.set_page_config(page_title="간호사 스마트 3교대 스케줄러", layout="wide")

st.title("🏥 3교대 간호사 스마트 근무표 작성 & 신청 시스템")
st.subheader("실시간 일별 선착순 오프 마감 및 신청 일시 기록 기능이 완벽하게 반영되었습니다.")

# 2. 세션 메모리 초기화 (새로고침 시 데이터 및 신청 시간 로그 보존용)
if "schedule_df_state" not in st.session_state:
    st.session_state["schedule_df_state"] = None  # 취합 중인 데이터프레임
if "optimized_result" not in st.session_state:
    st.session_state["optimized_result"] = None  # 최종 생성된 근무표
if "wanted_off_log" not in st.session_state:
    st.session_state["wanted_off_log"] = {}  # 간호사별 상세 신청 로그: { "간호사명": { day: "MM/DD HH:MM" } }

# 3. 사이드바 - 관리자 메뉴 및 업로드 버튼
st.sidebar.header("⚙️ 수간호사 관리자 메뉴")
uploaded_rules = st.sidebar.file_uploader("1. 작성 규칙 파일 업로드 (xlsx/csv)", type=["xlsx", "csv"])
uploaded_schedule = st.sidebar.file_uploader("2. 초기 근무표 템플릿 업로드 (xlsx/csv)", type=["xlsx", "csv"])
uploaded_prev_month = st.sidebar.file_uploader("3. 이전 달 근무표 업로드 (선택사항)", type=["xlsx", "csv"])

# 🛠 [부서별 맞춤 근무 조건 설정]
st.sidebar.markdown("---")
st.sidebar.header("🛠️ 부서별 맞춤 근무 조건 설정")

# 1. On/Off 토글 규칙들
rule_5_consec_off = st.sidebar.toggle("5일 연속 근무 시 후속 2 OFF 강제 보장", value=True)
rule_no_single_night = st.sidebar.toggle("단독 나이트(하루짜리 N) 금지", value=True)
rule_group_balance = st.sidebar.toggle("듀티별 그룹(A/B/C) 균등 배치 적용", value=True)
rule_night_after_2_off = st.sidebar.toggle("야간 근무(N) 후 2일 OFF 필수 부여", value=True)

# 2. 원하는 일수 슬라이더 조절 기능
limit_max_consec_work = st.sidebar.slider("최대 연속 근무 일수 제한", 0, 5, 5)
limit_max_monthly_night = st.sidebar.slider("월간 인당 최대 나이트(N) 개수", 0, 7, 6)
limit_max_consec_night = st.sidebar.slider("최대 연속 나이트(N) 제한", 2, 3, 3)

# ⭐ [실시간 마감] 하루 최대 오프 신청 한도 설정 슬라이더
limit_daily_off_request = st.sidebar.slider(
    "📢 하루 최대 오프 신청 허용 인원", 
    min_value=1, max_value=4, value=2, 
    help="하루에 최대로 오프를 신청할 수 있는 간호사 인원수입니다. 초과 시 마감 처리됩니다. (권장: 2명)"
)

# [새 파일 업로드 감지] 파일 교체 시 메모리 리셋
if uploaded_schedule:
    file_key = f"{uploaded_schedule.name}_{uploaded_schedule.size}"
    if "last_file_key" not in st.session_state or st.session_state["last_file_key"] != file_key:
        st.session_state["last_file_key"] = file_key
        st.session_state["schedule_df_state"] = None  
        st.session_state["optimized_result"] = None   
        st.session_state["wanted_off_log"] = {}

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
        # 규칙 2: 한 달 밤근무(N) 개수 균등화
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

# 4. 파일 데이터 로드 및 갱신 시스템
if uploaded_schedule and st.session_state["schedule_df_state"] is None:
    try:
        if uploaded_schedule.name.endswith('xlsx'):
            raw_df = pd.read_excel(uploaded_schedule)
        else:
            raw_df = pd.read_csv(uploaded_schedule, encoding='utf-8-sig')
        st.session_state["schedule_df_state"] = load_and_align_headers(raw_df)
    except Exception as e:
        st.error(f"템플릿 파일을 읽는 중 오류가 발생했습니다: {e}")

# 5. 메인 인터페이스부
if st.session_state["schedule_df_state"] is not None:
    tab_apply, tab_check, tab_result = st.tabs([
        "🙋‍♀️ [간호사용] 원티드 오프 신청", 
        "📋 [관리자용] 신청 및 고정 근무 확인", 
        "📅 [관리자용] AI 최종 근무표 생성"
    ])
    
    # ---------------- 탭 1: 간호사 자가 신청 포털 (선착순 마감 기능 탑재!) ----------------
    with tab_apply:
        st.write("### 📅 원하는 휴무일 직접 신청")
        st.info("💡 선착순으로 오프를 신청할 수 있으며, 일별 최대 허용 한도를 초과하면 마감됩니다.")
        df_temp = st.session_state["schedule_df_state"]
        
        nurse_names = []
        for idx, row in df_temp.iterrows():
            nurse_id, _ = parse_nurse_row(row['이름'], row['그룹'])
            if nurse_id is not None:
                nurse_names.append(nurse_id)
                
        # 1. 일별 실시간 신청 현황 요약 테이블 작성
        day_request_counts = {}
        for nurse, days_dict in st.session_state["wanted_off_log"].items():
            for d in days_dict.keys():
                day_request_counts[d] = day_request_counts.get(d, 0) + 1
                
        # 간호사들이 직관적으로 볼 수 있게 달력 형태의 신청 현황판 렌더링
        st.write("#### 📢 8월 일자별 오프 신청 현황 및 마감 잔여석")
        cols = st.columns(10) # 10개 열로 나누어 일수 나열
        for day in range(1, 32):
            col_idx = (day - 1) % 10
            count = day_request_counts.get(day, 0)
            rem = max(0, limit_daily_off_request - count)
            
            with cols[col_idx]:
                if rem == 0:
                    st.markdown(f"**{day}일**\n\n🔴 **마감**")
                else:
                    st.markdown(f"**{day}일**\n\n🟢 {rem}석 남음")
        
        st.markdown("---")
        
        col1, col2 = st.columns(2)
        with col1:
            selected_nurse = st.selectbox(
                "1. 본인의 이름을 선택하세요", 
                nurse_names, 
                format_func=lambda x: f"{x}번 간호사" if str(x).replace('.0', '').isdigit() else f"{x} 간호사"
            )
        with col2:
            selected_offs = st.multiselect("2. 희망 휴무일을 선택하세요 (마감되지 않은 날짜만 선택 가능)", list(range(1, 32)))
            
        if st.button("📝 원티드 오프 신청하기", type="primary"):
            if len(selected_offs) > 0:
                # 선택한 오프 일정 중 마감된 날짜가 있는지 최종 체크
                closed_days = []
                for d in selected_offs:
                    # 나를 제외한 다른 사람들의 신청 수 합산
                    other_count = sum(1 for n_id, days in st.session_state["wanted_off_log"].items() if d in days and str(n_id) != str(selected_nurse))
                    if other_count >= limit_daily_off_request:
                        closed_days.append(d)
                        
                if len(closed_days) > 0:
                    st.error(f"❌ 신청 실패: 선택하신 날짜 중 {closed_days}일은 이미 선착순 마감되었습니다. 다른 날짜를 선택해 주세요.")
                else:
                    # 신청 시간 기록 (Timestamp) 생성 (예: "08/15 14:35")
                    current_time_str = datetime.now().strftime("%m/%d %H:%M")
                    
                    # 세션 오프 로그 갱신
                    if selected_nurse not in st.session_state["wanted_off_log"]:
                        st.session_state["wanted_off_log"][selected_nurse] = {}
                        
                    # 이번에 선택한 날짜와 시간을 신규 등록
                    st.session_state["wanted_off_log"][selected_nurse] = { d: current_time_str for d in selected_offs }
                    
                    # 템플릿 표의 '원티드 오프' 칸에 저장할 텍스트 제작: "3일(08/20 14:22), 5일(08/20 14:23)"
                    detail_offs_list = []
                    for d in sorted(selected_offs):
                        t_str = st.session_state["wanted_off_log"][selected_nurse][d]
                        detail_offs_list.append(f"{d}({t_str})")
                    offs_formatted_str = ", ".join(detail_offs_list)
                    
                    # 데이터프레임 업데이트
                    for idx, row in df_temp.iterrows():
                        nurse_id, _ = parse_nurse_row(row['이름'], row['그룹'])
                        if str(nurse_id) == str(selected_nurse):
                            df_temp.loc[idx, '원티드 오프'] = offs_formatted_str
                            break
                            
                    st.session_state["schedule_df_state"] = df_temp # 세션 갱신
                    st.success(f"🎉 신청 완료: {selected_nurse} 간호사님 오프가 성공적으로 수집되었습니다! [기록: {offs_formatted_str}]")
                    st.rerun() # 화면 즉시 갱신
            else:
                st.warning("날짜를 최소 1개 이상 선택해 주세요.")
                
    # ---------------- 탭 2: 수간호사 확인 대시보드 ----------------
    with tab_check:
        st.write("### 📋 현재 업로드된 템플릿 및 오프 취합 현황")
        st.info("💡 간호사들의 원티드 신청 및 시간(Timestamp)이 실시간 취합된 결과입니다.")
        st.dataframe(st.session_state["schedule_df_state"])
        
        towrite_temp = io.BytesIO()
        st.session_state["schedule_df_state"].to_excel(towrite_temp, index=False, header=True)
        towrite_temp.seek(0)
        st.download_button(
            label="📥 현재까지 실시간 취합된 템플릿 파일 다운로드",
            data=towrite_temp,
            file_name="실시간_취합_템플릿.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    # ---------------- 탭 3: AI 최종 생성판 ----------------
    with tab_result:
        st.write("### 🚀 고정 근무 및 야간전담이 연동된 AI 근무표 작성")
        st.info("⚙️ 팁: 수집된 시간 기록에서 숫자 일자만 똑똑하게 파싱하여 AI 오프 연동에 가동합니다.")
        max_iter = st.slider("최대 탐색 횟수 (탐색 횟수가 높을수록 정밀해집니다)", 10000, 150000, 60000, step=10000)
        
        if st.button("🔮 최종 AI 근무표 생성 시작", type="primary"):
            if uploaded_rules is None:
                st.error("에러: 규칙 파일을 먼저 사이드바에 업로드해 주세요.")
            else:
                with st.spinner("야간전담 분류 및 이전 달 근태 연동 연산 중..."):
                    df_clean = st.session_state["schedule_df_state"].copy()
                    
                    df_clean = df_clean.replace(r'^\s*$', np.nan, regex=True)
                    df_clean['그룹'] = df_clean['그룹'].ffill()
                    
                    # 1. 이전 달 근무 데이터 로드 처리
                    prev_df = None
                    if uploaded_prev_month is not None:
                        try:
                            if uploaded_prev_month.name.endswith('xlsx'):
                                prev_df = pd.read_excel(uploaded_prev_month)
                            else:
                                prev_df = pd.read_csv(uploaded_prev_month, encoding='utf-8-sig')
                        except Exception as e:
                            st.warning(f"경고: 이전 달 근무표를 파싱하는 과정에서 오류가 발생했습니다. (오류내용: {e})")
                    
                    # 2. 간호사 추출 및 야간전담 분류 + 이전달 근무 기록 매핑
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
                            # ⭐ [정규식 가변 숫자 날짜 추출]: "3(08/20 14:22)" 과 같은 기록에서 "3"이라는 순수 숫자 일자만 안전하게 정규식 파싱합니다.
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
                            
                    # 3. 요구 인원수 파싱
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
                                        col_name = None
                                        for col in df_clean.columns:
                                            if str(col).strip().replace('.0', '') == str(d):
                                                col_name = col
                                                break
                                        
                                        val = None
                                        if col_name is not None:
                                            try:
                                                val = int(float(row[col_name]))
                                            except (ValueError, TypeError):
                                                pass
                                        
                                        if val is None or np.isnan(val):
                                            val = default_values[duty][d-1]
                                            
                                        day_values.append(val)
                                    requirements[duty] = day_values

                    for duty in ['D', 'E', 'N', 'DE']:
                        if duty not in requirements or len(requirements[duty]) != 31:
                            requirements[duty] = default_values[duty]
                            
                    num_nurses = len(nurses)
                    num_days = 31
                    nurse_groups = [n['group'] for n in nurses]
                    nurse_wanted_off = [set(n['wanted_off']) for n in nurses]
                    forbidden_5_patterns = [
                        ['D', 'D', 'N', 'N', 'N'], ['D', 'D', 'D', 'N', 'N'], ['D', 'D', 'D', 'D', 'N'],
                        ['D', 'E', 'N', 'N', 'N'], ['E', 'E', 'N', 'N', 'N'], ['D', 'D', 'E', 'N', 'N']
                    ]
                    
                    # [수학적 벌점 충돌 차단]
                    num_keepers = sum(is_night_keepers)
                    num_normal = num_nurses - num_keepers
                    
                    total_shifts_required = sum(requirements['D']) + sum(requirements['E']) + sum(requirements['N']) + sum(requirements['DE'])
                    total_N_required = sum(requirements['N'])
                    
                    total_keeper_N = num_keepers * 15
                    total_keeper_shifts = num_keepers * 15
                    
                    total_normal_N = max(0, total_N_required - total_keeper_N)
                    total_normal_shifts = max(0, total_shifts_required - total_keeper_shifts)
                    
                    total_normal_nurse_days = num_normal * num_days
                    total_normal_OFF = max(0, total_normal_nurse_days - total_normal_shifts)
                    
                    if num_normal > 0:
                        avg_normal_N = total_normal_N / num_normal
                        avg_normal_OFF = total_normal_OFF / num_normal
                        
                        target_N_min = int(avg_normal_N)
                        target_N_max = int(avg_normal_N) + 1 if avg_normal_N % 1 != 0 else int(avg_normal_N)
                        
                        target_OFF_min = int(avg_normal_OFF)
                        target_OFF_max = int(avg_normal_OFF) + 1 if avg_normal_OFF % 1 != 0 else int(avg_normal_OFF)
                    else:
                        target_N_min, target_N_max = 0, 0
                        target_OFF_min, target_OFF_max = 0, 0
                    
                    target_N_max = min(target_N_max, limit_max_monthly_night)
                    target_N_min = min(target_N_min, target_N_max)
                    
                    # 4. 사용자가 수동으로 채워둔 고정 근무의 위치(Lock Mask) 식별 및 한글 예외처리
                    is_fixed = np.zeros((num_nurses, num_days), dtype=bool)
                    fixed_shifts = np.empty((num_nurses, num_days), dtype=object)
                    
                    for i, nurse in enumerate(nurses):
                        row_idx = nurse['row_idx']
                        row = df_clean.iloc[row_idx]
                        for d in range(num_days):
                            col_name = str(d+1) if str(d+1) in df_clean.columns else (int(d+1) if int(d+1) in df_clean.columns else d+1)
                            raw_val = str(row[col_name]).strip().upper() if pd.notna(row[col_name]) else ""
                            
                            val = ""
                            if raw_val in ['D', '데이', 'DAY']: val = 'D'
                            elif raw_val in ['E', '이브', '이브닝', 'EVENING']: val = 'E'
                            elif raw_val in ['N', '나이트', 'NIGHT']: val = 'N'
                            elif raw_val in ['DE']: val = 'DE'
                            elif raw_val in ['OFF', '오프', '휴무', '휴']: val = 'OFF'
                            elif '교육' in raw_val: val = '교육'
                            
                            if val in ['D', 'E', 'N', 'DE', 'OFF', '교육']:
                                is_fixed[i, d] = True
                                fixed_shifts[i, d] = val
                    
                    # 5. 하이브리드 고정 스케줄 초기화
                    sched = initialize_schedule_hybrid(num_nurses, num_days, requirements, is_fixed, fixed_shifts)
                    
                    row_penalties = [get_nurse_penalty(sched[i], i, nurse_wanted_off[i], num_days, forbidden_5_patterns, is_night_keepers[i], nurse_histories[i], target_N_min, target_N_max, target_OFF_min, target_OFF_max, is_fixed[i]) for i in range(num_nurses)]
                    col_penalties = [get_day_penalty(sched[:, d], num_nurses, nurse_groups) for d in range(num_days)]
                    total_penalty = sum(row_penalties) + sum(col_penalties)
                    
                    best_sched = copy.deepcopy(sched)
                    best_penalty = total_penalty
                    best_hard = sum(row_penalties)
                    
                    temp = 25.0
                    cooling_rate = 0.9999
                    
                    # 6. 최적화 루프
                    for step in range(max_iter):
                        d = random.randint(0, num_days - 1)
                        i1 = random.randint(0, num_nurses - 1)
                        i2 = random.randint(0, num_nurses - 1)
                        while i1 == i2:
                            i2 = random.randint(0, num_nurses - 1)
                            
                        if is_fixed[i1, d] or is_fixed[i2, d]:
                            continue
                            
                        if sched[i1, d] == sched[i2, d]:
                            continue
                            
                        old_shift_i1, old_shift_i2 = sched[i1, d], sched[i2, d]
                        old_row_pen_i1, old_row_pen_i2 = row_penalties[i1], row_penalties[i2]
                        old_col_pen = col_penalties[d]
                        
                        sched[i1, d], sched[i2, d] = old_shift_i2, old_shift_i1
                        
                        new_row_pen_i1 = get_nurse_penalty(sched[i1], i1, nurse_wanted_off[i1], num_days, forbidden_5_patterns, is_night_keepers[i1], nurse_histories[i1], target_N_min, target_N_max, target_OFF_min, target_OFF_max, is_fixed[i1])
                        new_row_pen_i2 = get_nurse_penalty(sched[i2], i2, nurse_wanted_off[i2], num_days, forbidden_5_patterns, is_night_keepers[i2], nurse_histories[i2], target_N_min, target_N_max, target_OFF_min, target_OFF_max, is_fixed[i2])
                        new_col_pen = get_day_penalty(sched[:, d], num_nurses, nurse_groups)
                        
                        new_total_penalty = (total_penalty 
                                             - old_row_pen_i1 - old_row_pen_i2 - old_col_pen 
                                             + new_row_pen_i1 + new_row_pen_i2 + new_col_pen)
                        
                        delta = new_total_penalty - total_penalty
                        
                        accept = False
                        if delta < 0:
                            accept = True
                        elif temp > 0.05:
                            accept = (random.random() < np.exp(-delta / temp))
                            
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
                        if best_hard == 0 and step > 45000:
                            break
                            
                    # 결과를 데이터프레임 매핑
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
                
                # Excel 다운로드 기능 제공
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
    st.info("👈 시작하려면 왼쪽 사이드바에서 '2. 초기 근무표 템플릿 업로드' 파일을 가장 먼저 업로드해 주세요.")

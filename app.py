import streamlit as st
import pandas as pd
import numpy as np
import random
import copy
import io
import re

# 1. 웹페이지 기본 설정
st.set_page_config(page_title="간호사 스마트 3교대 스케줄러", layout="wide")

st.title("🏥 3교대 간호사 스마트 근무표 작성 & 신청 시스템")

# 2. 세션 메모리 초기화
if "schedule_df_state" not in st.session_state:
    st.session_state["schedule_df_state"] = None
if "optimized_result" not in st.session_state:
    st.session_state["optimized_result"] = None

# 사이드바 - 관리자 메뉴
st.sidebar.header("⚙️ 수간호사 관리자 메뉴")
uploaded_rules = st.sidebar.file_uploader("1. 작성 규칙 파일 업로드 (xlsx/csv)", type=["xlsx", "csv"])
uploaded_schedule = st.sidebar.file_uploader("2. 초기 근무표 템플릿 업로드 (xlsx/csv)", type=["xlsx", "csv"])

# [새 파일 업로드 감지 시스템] 파일이 교체되면 세션을 즉시 초기화하여 갱신
if uploaded_schedule:
    file_key = f"{uploaded_schedule.name}_{uploaded_schedule.size}"
    if "last_file_key" not in st.session_state or st.session_state["last_file_key"] != file_key:
        st.session_state["last_file_key"] = file_key
        st.session_state["schedule_df_state"] = None  
        st.session_state["optimized_result"] = None   

# [초강력 지능형 이름 및 야간전담 판정 파서]
def parse_nurse_row(name, group):
    name_str = str(name).strip() if pd.notna(name) else ""
    group_str = str(group).strip() if pd.notna(group) else ""
    
    if not name_str or name_str.lower() == 'nan':
        return None, False
        
    is_keeper = "야간전담" in name_str or "야간전담" in group_str
    
    clean_name = name_str
    if clean_name.endswith('.0'):
        clean_name = clean_name[:-2]
        
    match = re.search(r'\d+', clean_name)
    if match:
        nurse_id = int(match.group())
        return nurse_id, is_keeper
    return None, is_keeper

# [보정 함수] 업로드 엑셀의 헤더 위치 밀림 현상 자동 보정
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

if uploaded_schedule and st.session_state["schedule_df_state"] is None:
    try:
        if uploaded_schedule.name.endswith('xlsx'):
            raw_df = pd.read_excel(uploaded_schedule)
        else:
            raw_df = pd.read_csv(uploaded_schedule, encoding='utf-8-sig')
        st.session_state["schedule_df_state"] = load_and_align_headers(raw_df)
    except Exception as e:
        st.error(f"템플릿 파일을 읽는 중 오류가 발생했습니다: {e}")

# 벌점 계산 수식 정의
def get_nurse_penalty(row, i, nurse_wanted_off, num_days, forbidden_5_patterns, is_night_keeper):
    penalty = 0
    # 규칙 1: 원티드 오프 준수
    for d in range(num_days):
        if (d+1) in nurse_wanted_off[i] and row[d] != 'OFF':
            penalty += 1000000
            
    # [야간전담 전용 규칙 분기 처리]
    if is_night_keeper:
        # 야간전담 규칙 A: D, E 근무 절대 배정 금지
        for d in range(num_days):
            if row[d] in ['D', 'E']:
                penalty += 1000000
        
        # 야간전담 규칙 B: 한 달 총 밤근무(N)는 정확히 15개여야 함
        total_N = sum(1 for x in row if x == 'N')
        if total_N != 15:
            penalty += abs(total_N - 15) * 500000
    else:
        # [일반 교대 간호사 규칙]
        # 규칙 2: 한 달 밤근무(N) 5~6개 균등화
        total_N = sum(1 for x in row if x == 'N')
        if total_N < 5 or total_N > 6:
            penalty += (abs(total_N - 5.5) - 0.5) * 500000
            
        # 규칙 3: 한 달 총 휴무(OFF) 개수 균등화 (12~13일)
        total_OFF = sum(1 for x in row if x == 'OFF')
        if total_OFF < 12 or total_OFF > 13:
            penalty += (abs(total_OFF - 12.5) - 0.5) * 400000
        
    # 공통 피로도 제어 규칙
    consec_work = 0
    consec_N = 0
    for d in range(num_days):
        shift = row[d]
        if shift != 'OFF':
            consec_work += 1
            if consec_work >= 6:
                penalty += (consec_work - 5) * 500000
        else:
            # 규칙 4: 5일 연속 근무 후 2 OFF 연속 보장
            if consec_work == 5:
                if d + 1 < num_days:
                    if row[d+1] != 'OFF':
                        penalty += 300000
            consec_work = 0
            
        if shift == 'N':
            consec_N += 1
            if consec_N > 3: # 연속 밤근무 최대 3일 차단
                penalty += (consec_N - 3) * 500000
        else:
            consec_N = 0
            
        if d < num_days - 1:
            next_shift = row[d+1]
            if shift == 'E' and next_shift == 'D':
                penalty += 500000
            if shift == 'N' and next_shift in ['D', 'E']:
                penalty += 500000
                
        if shift == 'N':
            if d < num_days - 1:
                if row[d+1] != 'N':
                    if row[d+1] != 'OFF':
                        penalty += 500000
                    if d < num_days - 2:
                        if row[d+2] != 'OFF':
                            penalty += 500000
                            
        if d <= num_days - 5:
            pat = list(row[d:d+5])
            if pat in forbidden_5_patterns:
                penalty += 500000
                
    for d in range(num_days):
        if row[d] == 'N':
            prev_is_N = (d > 0 and row[d-1] == 'N')
            next_is_N = (d < num_days - 1 and row[d+1] == 'N')
            if not prev_is_N and not next_is_N:
                penalty += 300000
                
    return penalty

def get_day_penalty(col, num_nurses, nurse_groups):
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
        
        pD = sum(1 for i in range(num_nurses) if is_fixed[i, d] and fixed_shifts[i, d] == 'D')
        pE = sum(1 for i in range(num_nurses) if is_fixed[i, d] and fixed_shifts[i, d] == 'E')
        pN = sum(1 for i in range(num_nurses) if is_fixed[i, d] and fixed_shifts[i, d] == 'N')
        
        rem_D = max(0, nD - pD)
        rem_E = max(0, nE - pE)
        rem_N = max(0, nN - pN)
        
        num_unfixed = sum(1 for i in range(num_nurses) if not is_fixed[i, d])
        rem_OFF = max(0, num_unfixed - rem_D - rem_E - rem_N)
        
        pool = ['D'] * rem_D + ['E'] * rem_E + ['N'] * rem_N + ['OFF'] * rem_OFF
        random.shuffle(pool)
        
        pool_idx = 0
        for i in range(num_nurses):
            if is_fixed[i, d]:
                sched[i, d] = fixed_shifts[i, d]
            else:
                sched[i, d] = pool[pool_idx]
                pool_idx += 1
    return sched

# 3. 메인 인터페이스부
if st.session_state["schedule_df_state"] is not None:
    tab_apply, tab_check, tab_result = st.tabs([
        "🙋‍♀️ [간호사용] 원티드 오프 신청", 
        "📋 [관리자용] 신청 및 고정 근무 확인", 
        "📅 [관리자용] AI 최종 근무표 생성"
    ])
    
    # ---------------- 탭 1: 자가 신청 포털 ----------------
    with tab_apply:
        st.write("### 📅 원하는 휴무일 직접 신청")
        df_temp = st.session_state["schedule_df_state"]
        
        nurse_names = []
        for idx, row in df_temp.iterrows():
            nurse_id, _ = parse_nurse_row(row['이름'], row['그룹'])
            if nurse_id is not None:
                nurse_names.append(nurse_id)
                
        col1, col2 = st.columns(2)
        with col1:
            selected_nurse = st.selectbox("1. 본인의 이름을 선택하세요", nurse_names, format_func=lambda x: f"{x}번 간호사")
        with col2:
            selected_offs = st.multiselect("2. 희망 휴무일을 복수 선택하세요", list(range(1, 32)))
            
        if st.button("📝 원티드 오프 신청하기", type="primary"):
            if len(selected_offs) > 0:
                offs_str = ", ".join(map(str, sorted(selected_offs)))
                for idx, row in df_temp.iterrows():
                    nurse_id, _ = parse_nurse_row(row['이름'], row['그룹'])
                    if nurse_id == selected_nurse:
                        df_temp.loc[idx, '원티드 오프'] = offs_str
                        break
                st.session_state["schedule_df_state"] = df_temp
                st.success(f"✔️ {selected_nurse}번 간호사: {offs_str}일 OFF 신청 완료!")
            else:
                st.warning("날짜를 선택해 주세요.")
                
    # ---------------- 탭 2: 수간호사 확인 대시보드 ----------------
    with tab_check:
        st.write("### 📋 현재 업로드된 템플릿 현황")
        st.info("💡 팁 1: 템플릿 엑셀에 미리 기입해 둔 'N'(나이트)이나 'OFF' 등은 AI가 건드리지 않고 그대로 유지(Lock)됩니다.")
        st.info("💡 팁 2: 간호사 이름 옆이나 그룹 칸에 '야간전담'이라고 적으면, 자동으로 D/E가 제외되며 월 15일 고정 N이 배정됩니다.")
        st.dataframe(st.session_state["schedule_df_state"])
        
    # ---------------- 탭 3: AI 최적화 연산 실행판 ----------------
    with tab_result:
        st.write("### 🚀 고정 근무 및 야간전담이 연동된 AI 근무표 작성")
        max_iter = st.slider("최대 탐색 횟수 (탐색 횟수가 높을수록 정밀해집니다)", 10000, 150000, 60000, step=10000)
        
        if st.button("🔮 최종 AI 근무표 생성 시작", type="primary"):
            if uploaded_rules is None:
                st.error("에러: 규칙 파일을 먼저 사이드바에 업로드해 주세요.")
            else:
                with st.spinner("야간전담 대상자 분류 및 잠금 마스크 활성화 중..."):
                    df_clean = st.session_state["schedule_df_state"].copy()
                    df_clean['그룹'] = df_clean['그룹'].ffill()
                    
                    # 1. 간호사 추출 및 야간전담 분류
                    nurses = []
                    is_night_keepers = []
                    
                    for idx, row in df_clean.iterrows():
                        name = row['이름']
                        group = row['그룹']
                        nurse_id, is_keeper = parse_nurse_row(name, group)
                        
                        if nurse_id is not None:
                            wanted = row['원티드 오프']
                            wanted_days = []
                            if pd.notna(wanted) and str(wanted).strip() != '-':
                                wanted_days = [int(float(x.strip())) for x in str(wanted).split(',') if x.strip().replace('.0', '').isdigit()]
                            
                            nurses.append({
                                'id': nurse_id,
                                'group': row['그룹'],
                                'wanted_off': wanted_days,
                                'row_idx': idx,
                                'is_keeper': is_keeper
                            })
                            is_night_keepers.append(is_keeper)
                            
                    # 2. 요구 인원수 파싱
                    requirements = {}
                    start_idx = None
                    for idx, row in enumerate(df_clean.values):
                        row_str = " ".join([str(x) for x in row])
                        if "듀티별" in row_str or "인원수" in row_str:
                            start_idx = idx
                            break
                            
                    if start_idx is not None:
                        for i in range(3):
                            if start_idx + i < len(df_clean):
                                row = df_clean.iloc[start_idx + i]
                                duty = None
                                for col in df_clean.columns:
                                    if str(col).strip() not in [str(d) for d in range(1, 32)]:
                                        val_str = str(row[col]).strip().upper()
                                        if val_str in ['D', 'E', 'N']:
                                            duty = val_str
                                            break
                                if duty:
                                    day_values = []
                                    for d in range(1, 32):
                                        col_name = str(d) if str(d) in df_clean.columns else (int(d) if int(d) in df_clean.columns else d)
                                        if col_name in df_clean.columns:
                                            try:
                                                day_values.append(int(float(row[col_name])))
                                            except:
                                                pass
                                    if len(day_values) == 31:
                                        requirements[duty] = day_values

                    # [스마트 폴백 장치]
                    default_D = [3, 3, 4, 4, 4, 4, 4, 3, 3, 4, 4, 4, 4, 4, 3, 3, 3, 4, 4, 4, 4, 3, 3, 4, 4, 4, 4, 4, 3, 3, 4]
                    default_E = [3, 3, 4, 4, 4, 4, 4, 3, 3, 4, 4, 4, 4, 4, 3, 3, 3, 4, 4, 4, 4, 3, 3, 4, 4, 4, 4, 4, 3, 3, 4]
                    default_N = [3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3]
                    
                    if 'D' not in requirements or len(requirements['D']) != 31:
                        requirements['D'] = default_D
                    if 'E' not in requirements or len(requirements['E']) != 31:
                        requirements['E'] = default_E
                    if 'N' not in requirements or len(requirements['N']) != 31:
                        requirements['N'] = default_N
                            
                    num_nurses = len(nurses)
                    num_days = 31
                    nurse_groups = [n['group'] for n in nurses]
                    nurse_wanted_off = [set(n['wanted_off']) for n in nurses]
                    forbidden_5_patterns = [
                        ['D', 'D', 'N', 'N', 'N'], ['D', 'D', 'D', 'N', 'N'], ['D', 'D', 'D', 'D', 'N'],
                        ['D', 'E', 'N', 'N', 'N'], ['E', 'E', 'N', 'N', 'N'], ['D', 'D', 'E', 'N', 'N']
                    ]
                    
                    # 3. 사용자가 수동으로 채워둔 고정 근무의 위치(Lock Mask) 식별 및 한글 예외처리
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
                            elif raw_val in ['OFF', '오프', '휴무', '휴']: val = 'OFF'
                            
                            if val in ['D', 'E', 'N', 'OFF']:
                                is_fixed[i, d] = True
                                fixed_shifts[i, d] = val
                    
                    # 4. 하이브리드 고정 스케줄 초기화
                    sched = initialize_schedule_hybrid(num_nurses, num_days, requirements, is_fixed, fixed_shifts)
                    
                    row_penalties = [get_nurse_penalty(sched[i], i, nurse_wanted_off, num_days, forbidden_5_patterns, is_night_keepers[i]) for i in range(num_nurses)]
                    col_penalties = [get_day_penalty(sched[:, d], num_nurses, nurse_groups) for d in range(num_days)]
                    total_penalty = sum(row_penalties) + sum(col_penalties)
                    
                    best_sched = copy.deepcopy(sched)
                    best_penalty = total_penalty
                    best_hard = sum(row_penalties)
                    
                    temp = 25.0
                    cooling_rate = 0.9999
                    
                    # 5. 최적화 루프 (고정 스케줄 자리 잠금 체크)
                    for step in range(max_iter):
                        d = random.randint(0, num_days - 1)
                        i1 = random.randint(0, num_nurses - 1)
                        i2 = random.randint(0, num_nurses - 1)
                        while i1 == i2:
                            i2 = random.randint(0, num_nurses - 1)
                            
                        # [고정 근무 보호] 고정된 자리가 하나라도 있다면 패스
                        if is_fixed[i1, d] or is_fixed[i2, d]:
                            continue
                            
                        if sched[i1, d] == sched[i2, d]:
                            continue
                            
                        old_shift_i1, old_shift_i2 = sched[i1, d], sched[i2, d]
                        old_row_pen_i1, old_row_pen_i2 = row_penalties[i1], row_penalties[i2]
                        old_col_pen = col_penalties[d]
                        
                        sched[i1, d], sched[i2, d] = old_shift_i2, old_shift_i1
                        
                        new_row_pen_i1 = get_nurse_penalty(sched[i1], i1, nurse_wanted_off, num_days, forbidden_5_patterns, is_night_keepers[i1])
                        new_row_pen_i2 = get_nurse_penalty(sched[i2], i2, nurse_wanted_off, num_days, forbidden_5_patterns, is_night_keepers[i2])
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
                    file_name="최종_근무표_야간전담반영.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
else:
    st.info("👈 시작하려면 왼쪽 사이드바에서 '2. 근무표 템플릿 파일'을 가장 먼저 업로드해 주세요.")
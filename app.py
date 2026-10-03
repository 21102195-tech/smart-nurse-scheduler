import streamlit as st
import pandas as pd
import numpy as np
import random
import copy
import io
import re

# 1. 웹페이지 기본 설정
st.set_page_config(page_title="스마트 널스 스케쥴러", layout="wide")

# ⭐ [프로그램 제목 및 설명 부제목 적용 완료]
st.title("📊 스마트 널스 스케쥴러")
st.markdown("<h5 style='color: gray; font-weight: normal;'>수간호사 관리자 메뉴에서 초기 근무표 템플릿 업로드 후 AI 최종 근무표를 실행할 수 있습니다</h5>", unsafe_allow_html=True)
st.markdown("---")

# 2. 세션 메모리 초기화
if "schedule_df_state" not in st.session_state:
    st.session_state["schedule_df_state"] = None
if "optimized_result" not in st.session_state:
    st.session_state["optimized_result"] = None

# 3. 사이드바 - 관리자 메뉴 및 업로드 버튼
st.sidebar.header("⚙ 수간호사 관리자 메뉴")
uploaded_schedule = st.sidebar.file_uploader("1. 초기 근무표 템플릿 업로드 (xlsx/csv)", type=["xlsx", "csv"])
uploaded_prev_month = st.sidebar.file_uploader("2. 이전 달 근무표 업로드 (선택사항)", type=["xlsx", "csv"])

# 🛠 [부서별 맞춤 근무 조건 설정]
st.sidebar.markdown("---")
st.sidebar.header("🛠 부서별 맞춤 근무 조건 설정")

# 1. On/Off 토글 규칙들
rule_5_consec_off = st.sidebar.toggle("5일 연속 근무 시 후속 2 OFF 강제 보장", value=True, help="체크 시 5일 일하면 무조건 이틀 쉽니다.")
rule_no_single_night = st.sidebar.toggle("단독 나이트(하루짜리 N) 금지", value=True, help="체크 시 밤근무는 무조건 연속 2~3일로 묶어서 배정됩니다.")
rule_group_balance = st.sidebar.toggle("듀티별 그룹(A/B/C) 균등 배치 적용", value=True, help="체크 시 특정 경력의 간호사가 한 듀티에 쏠리지 않도록 분산합니다.")
rule_night_after_2_off = st.sidebar.toggle("야간 근무(N) 후 2일 OFF 필수 부여", value=True, help="체크 시 야간 근무 종료 후 최소 2일 연속 OFF를 필수로 보장합니다.")
rule_no_single_work = st.sidebar.toggle("단독 근무(하루짜리 근무) 금지", value=True, help="체크 시 근무는 최소 연속 2일 이상 배정되도록 유도하여 퐁당퐁당 근무를 방지합니다.")

# 2. 원하는 일수 슬라이더 조절 기능
limit_max_consec_work = st.sidebar.slider(
    "최대 연속 근무 일수 제한", 
    min_value=0, max_value=5, value=5, 
    help="연속 일할 수 있는 한도를 지정합니다. 0일 지정 시 연속 근무 일수 제한 규칙이 꺼집니다. (0일 ~ 최대 5일)"
)
limit_max_monthly_night = st.sidebar.slider(
    "월간 인당 최대 나이트(N) 개수", 
    min_value=0, max_value=7, value=6, 
    help="교대 간호사 기준 한 달 최대 밤근무 한도를 설정합니다. (0개 ~ 최대 7개)"
)
limit_max_consec_night = st.sidebar.slider(
    "최대 연속 나이트(N) 제한", 
    min_value=2, max_value=3, value=3, 
    help="연속으로 밤근무를 서는 최대 일수를 제어합니다. (최소 2일 ~ 최대 3일)"
)

# [새 파일 업로드 감지] 파일 교체 시 메모리 리셋
if uploaded_schedule:
    file_key = f"{uploaded_schedule.name}_{uploaded_schedule.size}"
    if "last_file_key" not in st.session_state or st.session_state["last_file_key"] != file_key:
        st.session_state["last_file_key"] = file_key
        st.session_state["schedule_df_state"] = None  
        st.session_state["optimized_result"] = None   

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

# [지능형 가변 근무코드 추출 파서]
def parse_allowed_shifts(val):
    if pd.isna(val) or str(val).strip() in ["", "-", "전체", "nan"]:
        return {"D", "E", "N", "DE", "OFF", "교육"}
        
    val_str = str(val).strip().upper()
    tokens = re.split(r'[,/\s]+', val_str)
    allowed = set()
    for t in tokens:
        if t in ['D', '데이', 'DAY']: allowed.add('D')
        elif t in ['E', '이브', '이브닝', 'EVENING']: allowed.add('E')
        elif t in ['N', '나이트', 'NIGHT']: allowed.add('N')
        elif t in ['DE']: allowed.add('DE')
        elif t in ['OFF', '오프', '휴무', '휴']: allowed.add('OFF')
        elif t in ['교육', 'EDU']: allowed.add('교육')
        
    allowed.add('OFF')
    return allowed

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
                    
                    # 1. 야간 근무 판정 (N, NC1)
                    if val in ['N', 'NC1', '나이트', 'NIGHT']: 
                        val = 'N'
                    # 2. 오프 및 휴무 판정 (C, OF, OF1, V, H, H1, NV, BV, B, S, S10, S2, S3 등)
                    elif val in ['C', 'OF', 'OF1', 'V', 'H', 'H1', 'NV', 'BV', 'B', 'S', 'S10', 'S2', 'S3', 'OFF', '오프', '휴무', '휴']: 
                        val = 'OFF'
                    # 3. 교육 근무 판정 (CPE, P, PE, PE1, PL 등)
                    elif val in ['CPE', 'P', 'PE', 'PE1', 'PL', '교육', 'EDU']: 
                        val = '교육'
                    # 4. 기본 데이 근무 판정
                    elif val in ['D', '데이', 'DAY']: 
                        val = 'D'
                    # 5. 기본 이브닝 근무 판정
                    elif val in ['E', '이브', '이브닝', 'EVENING']: 
                        val = 'E'
                    # 6. 기본 DE 근무 판정
                    elif val in ['DE']: 
                        val = 'DE'
                    # 7. 예외 코드의 경우 OFF 기본값 처리
                    else: 
                        val = 'OFF'
                    shifts.append(val)
                return shifts
    except Exception as e:
        pass
    return ['OFF'] * 7

# 벌점 계산 수식 정의
def get_nurse_penalty(row_current, i, nurse_wanted_off_set, num_days, forbidden_5_patterns, is_night_keeper, history, target_N_min, target_N_max, target_OFF_min, target_OFF_max, is_fixed_row, allowed_shifts_set):
    # ⭐ [교육 근무의 D 치환]
    row_norm = []
    for x in (list(history) + list(row_current)):
        if x == '교육':
            row_norm.append('D')
        else:
            row_norm.append(x)
            
    num_total = len(row_norm)
    history_len = len(history)
    
    penalty = 0
    HARD_PENALTY = 10000000
    
    # 규칙 1: 원티드 오프 준수
    for d in range(history_len, num_total):
        current_day = d - history_len + 1
        if current_day in nurse_wanted_off_set and row_norm[d] != 'OFF':
            if not is_fixed_row[current_day - 1]:
                penalty += 1000000
                
    # 규칙 2: 간호사별 허용 근무코드 필터링
    for d in range(history_len, num_total):
        shift = row_norm[d]
        if shift not in allowed_shifts_set:
            penalty += HARD_PENALTY
            
    if is_night_keeper:
        for d in range(history_len, num_total):
            if row_norm[d] in ['D', 'E', 'DE']:
                penalty += HARD_PENALTY
        total_N = sum(1 for x in row_norm[history_len:] if x == 'N')
        if total_N != 15:
            penalty += abs(total_N - 15) * HARD_PENALTY
    else:
        # 규칙 2: 한 달 밤근무(N) 개수 균등화
        total_N = sum(1 for x in row_norm[history_len:] if x == 'N')
        if limit_max_monthly_night == 0:
            if total_N > 0:
                penalty += total_N * HARD_PENALTY
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
            
        # [근무 다양성 보장 규칙]
        total_D = sum(1 for x in row_norm[history_len:] if x in ['D', '교육'])
        total_E = sum(1 for x in row_norm[history_len:] if x in ['E', 'DE'])
        if 'D' in allowed_shifts_set and total_D < 3:
            penalty += (3 - total_D) * 100000
        if 'E' in allowed_shifts_set and total_E < 3:
            penalty += (3 - total_E) * 100000
        
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
                            penalty += HARD_PENALTY
            consec_work = 0
            
        if shift == 'N':
            consec_N += 1
            if consec_N > limit_max_consec_night and d >= history_len:  
                penalty += (consec_N - limit_max_consec_night) * HARD_PENALTY
        else:
            consec_N = 0
            
        # 교대 제한 (E->D, E->DE, N->D, N->E, N->DE, N->교육, DE->D 등 자동 제어)
        if d < num_total - 1:
            next_shift = row_norm[d+1]
            if (d+1) >= history_len:
                if shift == 'E' and next_shift in ['D', 'DE']:
                    penalty += HARD_PENALTY
                if shift == 'N' and next_shift in ['D', 'E', 'DE']:
                    penalty += HARD_PENALTY
                # [신규 규칙]: DE 근무 다음날 D 근무 금지 (교육 포함)
                if shift == 'DE' and next_shift == 'D':
                    penalty += HARD_PENALTY
                    
        # [신규 규칙]: E -> OFF -> D 근무 금지 (교육 포함)
        if d < num_total - 2:
            next_shift = row_norm[d+1]
            day_after_next = row_norm[d+2]
            if (d+2) >= history_len:
                if shift == 'E' and next_shift == 'OFF' and day_after_next == 'D':
                    penalty += HARD_PENALTY
                # [신규 규칙] N ➡ OFF ➡ D 근무 금지
                if shift == 'N' and next_shift == 'OFF' and day_after_next == 'D':
                    penalty += HARD_PENALTY
                # [신규 규칙] N ➡ OFF ➡ N 근무 금지
                if not is_night_keeper and shift == 'N' and next_shift == 'OFF' and day_after_next == 'N':
                    penalty += HARD_PENALTY
                
        # [조건 On/Off] 야간 근무(N) 후 2일 OFF 필수 부여
        if shift == 'N' and rule_night_after_2_off:
            if d < num_total - 1:
                if row_norm[d+1] != 'N':
                    if row_norm[d+1] != 'OFF' and (d+1) >= history_len:
                        penalty += HARD_PENALTY
                    if d < num_total - 2:
                        if row_norm[d+2] != 'OFF' and (d+2) >= history_len:
                            penalty += HARD_PENALTY
                            
        if d <= num_total - 5:
            pat = list(row_norm[d:d+5])
            if (d+4) >= history_len:
                if pat in forbidden_5_patterns:
                    penalty += HARD_PENALTY
                
    # [조건 On/Off] 단독 나이트 방지
    if rule_no_single_night:
        for d in range(history_len, num_total):
            if row_norm[d] == 'N':
                prev_is_N = (d > 0 and row_norm[d-1] == 'N')
                next_is_N = (d < num_total - 1 and row_norm[d+1] == 'N')
                if not prev_is_N and not next_is_N:
                    penalty += HARD_PENALTY
                    
    # ⭐ [조건 On/Off] 단독 근무 금지 검사로 퐁당퐁당 방지! (연속 2일 미만 근무 금지)
    if rule_no_single_work:
        for d in range(history_len, num_total):
            if row_norm[d] != 'OFF':
                prev_is_off = (d == 0 or row_norm[d-1] == 'OFF')
                next_is_off = (d == num_total - 1 or row_norm[d+1] == 'OFF')
                if prev_is_off and next_is_off:
                    penalty += 1000000  # 강한 soft 벌점 부과
                
    return penalty

# [가변 그룹 균등 분배 연산 패널]
def get_day_penalty(col, num_nurses, nurse_groups, unique_groups):
    if not rule_group_balance:
        return 0
        
    penalty = 0
    num_groups = len(unique_groups)
    if num_groups == 0:
        return 0
        
    for duty in ['D', 'E', 'N', 'OFF']:
        counts = {g: 0 for g in unique_groups}
        for i in range(num_nurses):
            if col[i] == duty:
                g = nurse_groups[i]
                if g in counts:
                    counts[g] += 1
        tot = sum(counts.values())
        if tot == 0:
            continue
            
        ideal_min = tot // num_groups
        ideal_max = ideal_min if tot % num_groups == 0 else ideal_min + 1
        
        duty_penalty = 0
        for g in unique_groups:
            c = counts[g]
            if c < ideal_min:
                duty_penalty += (ideal_min - c)
            elif c > ideal_max:
                duty_penalty += (c - ideal_max)
                
        penalty += duty_penalty * 50
        
    return penalty

# [수학적 인원 규칙 100% 만족형] 하이브리드 고정형 초기 스케줄 생성 함수
def initialize_schedule_hybrid(num_nurses, num_days, requirements, is_fixed, fixed_shifts, is_night_keepers):
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
        
        rem_D = max(0, nD - pD) 
        rem_E = max(0, nE - pE)
        rem_N = max(0, nN - pN)
        rem_DE = max(0, nDE - pDE)
        
        unfixed_keepers = [i for i in range(num_nurses) if not is_fixed[i, d] and is_night_keepers[i]]
        unfixed_normals = [i for i in range(num_nurses) if not is_fixed[i, d] and not is_night_keepers[i]]
        
        keeper_N_assign = min(len(unfixed_keepers), rem_N)
        rem_N -= keeper_N_assign
        
        keeper_pool = ['N'] * keeper_N_assign + ['OFF'] * (len(unfixed_keepers) - keeper_N_assign)
        random.shuffle(keeper_pool)
        
        pool = ['D'] * rem_D + ['E'] * rem_E + ['N'] * rem_N + ['DE'] * rem_DE
        rem_OFF = max(0, len(unfixed_normals) - len(pool))
        pool += ['OFF'] * rem_OFF
        pool = pool[:len(unfixed_normals)]
        random.shuffle(pool)
        
        keeper_idx = 0
        normal_idx = 0
        for i in range(num_nurses):
            if is_fixed[i, d]:
                sched[i, d] = fixed_shifts[i, d]
            elif is_night_keepers[i]:
                sched[i, d] = keeper_pool[keeper_idx]
                keeper_idx += 1
            else:
                sched[i, d] = pool[normal_idx]
                normal_idx += 1
    return sched


# 검수 전용: 생성 로직에 영향을 주지 않습니다.
def audit_fixed_inputs(df, nurses, num_days):
    fixed = np.zeros((len(nurses), num_days), dtype=bool)
    values = np.full(fixed.shape, None, dtype=object)
    aliases = {'데이':'D', 'DAY':'D', '이브':'E', '이브닝':'E', 'EVENING':'E', '나이트':'N', 'NIGHT':'N', 'NC1':'N', '오프':'OFF', '휴무':'OFF', '휴':'OFF', 'OF':'OFF', 'OF1':'OFF', 'C':'OFF', 'V':'OFF', 'H':'OFF', 'H1':'OFF', 'NV':'OFF', 'BV':'OFF', 'B':'OFF', 'S':'OFF', 'S10':'OFF', 'S2':'OFF', 'S3':'OFF', 'EDU':'교육', 'CPE':'교육', 'P':'교육', 'PE':'교육', 'PE1':'교육', 'PL':'교육'}
    for i, nurse in enumerate(nurses):
        row = df.loc[nurse['row_idx']]
        for d in range(num_days):
            col = str(d+1) if str(d+1) in df.columns else d+1
            raw = str(row.get(col, '')).strip().upper()
            value = aliases.get(raw, '교육' if '교육' in raw else raw)
            if value in ['D','E','N','DE','OFF','교육']:
                fixed[i,d] = True
                values[i,d] = value
    return fixed, values

def audit_schedule_violations(sched, nurses, requirements, fixed, fixed_values, histories, allowed, patterns, history_available):
    issues = []
    n, days = sched.shape
    def add(name, dates, rule, detail, level='위반'):
        issues.append({'구분': level, '직원': str(name), '날짜': dates, '조건': rule, '상세': detail})
    for d in range(days):
        for duty, counts in requirements.items():
            actual = sum(sched[:,d] == duty)
            if actual != counts[d]:
                add('전체', f'{d+1}일', '고정 N 인원 확인' if duty == 'N' else '필요인원', f'{duty}: 필요 {counts[d]}명 / 배정 {actual}명', '확인 필요' if duty == 'N' else '위반')
    for i, nurse in enumerate(nurses):
        name = nurse['id']
        raw = list(sched[i])
        row = ['D' if x == '교육' else x for x in list(histories[i]) + raw]
        h = len(histories[i])
        for d, shift in enumerate(raw):
            if fixed[i,d] and shift != fixed_values[i,d]:
                add(name, f'{d+1}일', '고정근무·승인 OFF', f'{fixed_values[i,d]} → {shift}')
            approved_fixed = fixed[i,d] and shift in ['N','교육'] and shift == fixed_values[i,d]
            if shift not in allowed[i] and not approved_fixed:
                add(name, f'{d+1}일', '허용 근무', f'{shift}은 허용되지 않음')
            if d+1 in nurse['wanted_off'] and shift != 'OFF':
                add(name, f'{d+1}일', '원티드 OFF', f'{shift} 배정', '선호 미충족')
        work = nights = 0
        for j, shift in enumerate(row):
            work = work+1 if shift != 'OFF' else 0
            nights = nights+1 if shift == 'N' else 0
            if j < h:
                continue
            day = j-h+1
            if limit_max_consec_work > 0 and work > limit_max_consec_work:
                add(name, f'{day}일', '연속근무 금지', f'{work}일 연속근무 (최대 {limit_max_consec_work}일)')
            if rule_5_consec_off and j >= 6 and row[j-1] == 'OFF' and all(x != 'OFF' for x in row[j-6:j-1]) and (j == 6 or row[j-7] == 'OFF') and shift != 'OFF':
                add(name, f'{day}일', '5일 근무 후 2 OFF', f'{shift} 배정')
            if nights > limit_max_consec_night and not fixed[i,day-1]:
                add(name, f'{day}일', '연속 야간', f'{nights}일 연속 N')
            if j:
                prev = row[j-1]
                if (prev == 'E' and shift in ['D','DE']) or (prev == 'N' and shift in ['D','E','DE']) or (prev == 'DE' and shift == 'D'):
                    add(name, f'{day}일 (전날 포함)', '금지 근무조합', f'{prev} → {shift}')
            if j >= 2:
                triple = row[j-2:j+1]
                if triple in [['E','OFF','D'], ['N','OFF','D']] or (not nurse['is_keeper'] and triple == ['N','OFF','N']):
                    add(name, f'{max(1,day-2)}~{day}일 (이전 달 포함 가능)', '금지 근무조합', ' → '.join(triple))
            if j >= 4 and row[j-4:j+1] in patterns:
                add(name, f'{max(1,day-4)}~{day}일', '금지 5일 패턴', ' → '.join(row[j-4:j+1]))
            if rule_night_after_2_off:
                if (j >= 1 and row[j-1] == 'N' and shift not in ['N','OFF']) or (j >= 2 and row[j-2] == 'N' and row[j-1] != 'N' and shift != 'OFF'):
                    add(name, f'{day}일', '야간 후 2 OFF', f'{shift} 배정')
            if rule_no_single_night and shift == 'N' and not fixed[i,day-1] and (j == 0 or row[j-1] != 'N'):
                if j+1 < len(row) and row[j+1] != 'N':
                    add(name, f'{day}일', '단독 나이트', '하루짜리 N')
                elif j+1 == len(row):
                    add(name, f'{day}일', '월말 야간 연결', '다음 달 N 또는 휴무 확인 필요', '확인 필요')
            if rule_no_single_work and shift != 'OFF' and (j == 0 or row[j-1] == 'OFF'):
                if j+1 < len(row) and row[j+1] == 'OFF':
                    add(name, f'{day}일', '단독 근무', '하루짜리 근무', '선호 미충족')
        total_n = raw.count('N')
        # 모든 N은 고정 입력입니다. 월 N 개수 자체는 위반으로 보지 않습니다.
        if rule_night_after_2_off and 'N' in raw[-2:]:
            add(name, f'{max(1,days-1)}~{days}일', '월말 야간 후 휴무', '다음 달 근무표에서 야간 종료 후 2 OFF 확인 필요', '확인 필요')
        if not history_available or not nurse.get('history_known', True):
            add(name, '1일 (이전 달 연결)', '월초 연결', '이전 달 근무자료 미제공: 월초 연속근무·근무조합 추가 확인 필요', '확인 필요')
    result = pd.DataFrame(issues, columns=['구분','직원','날짜','조건','상세'])
    return result[result['구분'] == '위반'].drop(columns=['구분']).reset_index(drop=True)

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
    df_temp = st.session_state["schedule_df_state"]
    
    # 가변 날짜 동적 감지
    day_cols_detected = [int(col) for col in df_temp.columns if str(col).strip().isdigit()]
    num_days_dynamic = max(day_cols_detected) if day_cols_detected else 31

    tab_apply, tab_check, tab_result = st.tabs([
        "🙋‍♀ [간호사용] 원티드 오프 신청", 
        "📋 [관리자용] 신청 및 고정 근무 확인", 
        "📅 [관리자용] AI 최종 근무표 생성"
    ])
    
    # ---------------- 탭 1: 자가 신청 포털 ----------------
    with tab_apply:
        st.write("### 📅 원하는 휴무일 직접 신청")
        
        nurse_names = []
        for idx, row in df_temp.iterrows():
            nurse_id, _ = parse_nurse_row(row['이름'], row['그룹'])
            if nurse_id is not None:
                nurse_names.append(nurse_id)
                
        col1, col2 = st.columns(2)
        with col1:
            selected_nurse = st.selectbox(
                "1. 본인의 이름을 선택하세요", 
                nurse_names, 
                format_func=lambda x: f"{x}번 간호사" if str(x).replace('.0', '').isdigit() else f"{x} 간호사"
            )
        with col2:
            selected_offs = st.multiselect("2. 희망 휴무일을 복수 선택하세요", list(range(1, num_days_dynamic + 1)))
            
        if st.button("📝 원티드 오프 신청하기", type="primary"):
            if len(selected_offs) > 0:
                offs_str = ", ".join(map(str, sorted(selected_offs)))
                for idx, row in df_temp.iterrows():
                    nurse_id, _ = parse_nurse_row(row['이름'], row['그룹'])
                    if str(nurse_id) == str(selected_nurse):
                        df_temp.loc[idx, '원티드 오프'] = offs_str
                        break
                st.session_state["schedule_df_state"] = df_temp
                st.success(f"✔ {selected_nurse} 간호사: {offs_str}일 OFF 신청 완료!")
            else:
                st.warning("날짜를 선택해 주세요.")
                
    # ---------------- 탭 2: 수간호사 확인 대시보드 ----------------
    with tab_check:
        st.write("### 📋 현재 업로드된 템플릿 현황")
        st.info("💡 팁 1: 템플릿 엑셀에 미리 기입해 둔 'D', 'E', 'N', 'DE', '교육', 'OFF' 등은 AI가 건드리지 않고 그대로 유지(Lock)됩니다.")
        st.info("💡 팁 2: 간호사 이름 옆이나 그룹 칸에 '야간전담'이라고 적으면, 자동으로 D/E가 제외되며 월 15일 고정 N이 배정됩니다.")
        st.dataframe(st.session_state["schedule_df_state"])
        
    # ---------------- 탭 3: AI 최적화 연산 실행판 ----------------
    with tab_result:
        st.write("### 🚀 고정 근무 및 야간전담이 연동된 AI 근무표 작성")
        st.info("⚙ 팁: 왼쪽 사이드바 메뉴에서 부서 맞춤 조건을 변경하시면 즉시 알고리즘 연산에 반영됩니다!")
        max_iter = st.slider("최대 탐색 횟수 (탐색 횟수가 높을수록 정밀해집니다)", 10000, 150000, 60000, step=10000)
        
        if st.button("🔮 최종 AI 근무표 생성 시작", type="primary"):
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
                        st.warning(f"경고: 이전 달 근무표를 파싱하는 과정에서 오류가 발생했습니다. 이전 달 근무 연동 없이 연산을 시작합니다. (오류내용: {e})")
                
                # 2. 간호사 추출 및 야간전담 분류 + 이전달 근무 기록 매핑
                nurses = []
                is_night_keepers = []
                nurse_histories = []
                allowed_shifts_list = []
                
                allowed_col = '가능 근무' if '가능 근무' in df_clean.columns else ('가능근무' if '가능근무' in df_clean.columns else None)
                
                for idx, row in df_clean.iterrows():
                    name = row['이름']
                    group = row['그룹']
                    nurse_id, is_keeper = parse_nurse_row(name, group)
                    
                    if nurse_id is not None:
                        wanted = row['원티드 오프'] if '원티드 오프' in df_clean.columns else None
                        wanted_days = []
                        if pd.notna(wanted) and str(wanted).strip() != '-':
                            wanted_days = [int(float(x.strip())) for x in str(wanted).split(',') if x.strip().replace('.0', '').isdigit()]
                        
                        # 이전 달 근무 내역 추출
                        history = extract_nurse_history(prev_df, nurse_id) if prev_df is not None else ['OFF'] * 7
                        
                        allowed = parse_allowed_shifts(row[allowed_col]) if allowed_col is not None else {"D", "E", "N", "DE", "OFF", "교육"}
                        if is_keeper:
                            allowed = {"N", "OFF"}
                            
                        nurses.append({
                            'id': nurse_id,
                            'group': row['그룹'],
                            'wanted_off': wanted_days,
                            'row_idx': idx,
                            'is_keeper': is_keeper
                        })
                        is_night_keepers.append(is_keeper)
                        nurse_histories.append(history)
                        allowed_shifts_list.append(allowed)
                        
                # 요구량 파싱
                requirements = {}
                default_values = {
                    'D': [3, 3, 4, 4, 4, 4, 4, 3, 3, 4, 4, 4, 4, 4, 3, 3, 3, 4, 4, 4, 4, 3, 3, 4, 4, 4, 4, 4, 3, 3, 4],
                    'E': [3, 3, 4, 4, 4, 4, 4, 3, 3, 4, 4, 4, 4, 4, 3, 3, 3, 4, 4, 4, 4, 3, 3, 4, 4, 4, 4, 4, 3, 3, 4],
                    'N': [3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3],
                    'DE': [0] * 31
                }
                
                de_allowed_groups = None
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
                                if str(col).strip() not in [str(d) for d in range(1, num_days_dynamic + 1)]:
                                    val_str = str(row[col]).strip().upper()
                                    if val_str in ['D', 'E', 'N', 'DE']:
                                        duty = val_str
                                        break
                            if duty:
                                # ⭐ DE 근무조의 가능 그룹 동적 파싱 및 추출 (한글/영문 전부 완벽 매핑!)
                                if duty == 'DE':
                                    for col in df_clean.columns:
                                        if str(col).strip() not in [str(d) for d in range(1, num_days_dynamic + 1)]:
                                            val = str(row[col]).strip()
                                            if pd.notna(row[col]) and val != "" and val.upper() not in ['DE', '듀티별 인원수', '듀티별인원수', 'NAN']:
                                                tokens = re.split(r'[^A-Za-z0-9가-힣]+', val.upper())
                                                ignore = {'DE', '듀티별', '인원수', '듀티별인원수', 'NAN', ''}
                                                groups = {t for t in tokens if t not in ignore}
                                                if groups:
                                                    de_allowed_groups = groups
                                                    break
                                                    
                                day_values = []
                                for d in range(1, num_days_dynamic + 1):
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
                    if duty not in requirements or len(requirements[duty]) != num_days_dynamic:
                        requirements[duty] = default_values[duty][:num_days_dynamic]
                        
                # ⭐ [동적 그룹 연동]: 템플릿에서 가져온 de_allowed_groups로 DE 근무코드 목록에서 실시간 차단!
                if de_allowed_groups is not None:
                    for i, nurse in enumerate(nurses):
                        nurse_group = str(nurse['group']).strip().upper() if pd.notna(nurse['group']) else ""
                        if nurse_group not in de_allowed_groups:
                            if "DE" in allowed_shifts_list[i]:
                                allowed_shifts_list[i].remove("DE")
                        
                num_nurses = len(nurses)
                num_days = num_days_dynamic
                nurse_groups = [n['group'] for n in nurses]
                nurse_wanted_off = [set(n['wanted_off']) for n in nurses]
                
                # 40시간 초과 금지 패턴
                forbidden_5_patterns = [
                    ['D', 'D', 'N', 'N', 'N'], ['D', 'D', 'D', 'N', 'N'], ['D', 'D', 'D', 'D', 'N'],
                    ['D', 'E', 'N', 'N', 'N'], ['E', 'E', 'N', 'N', 'N'], ['D', 'D', 'E', 'N', 'N'],
                    ['DE', 'DE', 'N', 'N', 'N'], 
                    ['DE', 'DE', 'DE', 'N', 'N'], 
                    ['DE', 'DE', 'DE', 'DE', 'N'],
                    ['DE', 'E', 'N', 'N', 'N'], 
                    ['DE', 'DE', 'E', 'N', 'N']
                ]
                
                # [수학적 벌점 충돌 차단 - DE 수량 포함 전면 리팩토링]
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
                    target_N_min, target_N_max, target_OFF_min, target_OFF_max = 0, 0, 0, 0
                
                target_N_max = min(target_N_max, limit_max_monthly_night)
                target_N_min = min(target_N_min, target_N_max)
                
                # 원본 고정입력 검수용 사본 (기존 배정 로직은 변경하지 않음)
                audit_fixed, audit_fixed_values = audit_fixed_inputs(df_clean, nurses, num_days)
                # 고정 근무 보호 처리
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
                        
                        if nurse['is_keeper']:
                            if val == 'N':
                                is_fixed[i, d] = True
                                fixed_shifts[i, d] = 'N'
                            else:
                                is_fixed[i, d] = False 
                                fixed_shifts[i, d] = None
                        else:
                            if val in ['D', 'E', 'N', 'DE', 'OFF', '교육']:
                                is_fixed[i, d] = True
                                fixed_shifts[i, d] = val
                
                # 수동 고정 OFF 해제 방어책
                for d in range(num_days):
                    nD_req = requirements['D'][d]
                    nE_req = requirements['E'][d]
                    nN_req = requirements['N'][d]
                    nDE_req = requirements['DE'][d] if 'DE' in requirements else 0
                    
                    pD = sum(1 for i in range(num_nurses) if is_fixed[i, d] and fixed_shifts[i, d] == 'D')
                    pE = sum(1 for i in range(num_nurses) if is_fixed[i, d] and fixed_shifts[i, d] == 'E')
                    pN = sum(1 for i in range(num_nurses) if is_fixed[i, d] and fixed_shifts[i, d] == 'N')
                    pDE = sum(1 for i in range(num_nurses) if is_fixed[i, d] and fixed_shifts[i, d] == 'DE')
                    
                    rem_D = max(0, nD_req - pD)
                    rem_E = max(0, nE_req - pE)
                    rem_N = max(0, nN_req - pN)
                    rem_DE = max(0, nDE_req - pDE)
                    
                    unfixed_normals_count = sum(1 for i in range(num_nurses) if not is_fixed[i, d] and not is_night_keepers[i])
                    required_normal_slots = rem_D + rem_E + rem_DE
                    
                    if unfixed_normals_count < required_normal_slots:
                        deficit = required_normal_slots - unfixed_normals_count
                        unlocked_count = 0
                        for i in range(num_nurses):
                            if not is_night_keepers[i] and is_fixed[i, d] and fixed_shifts[i, d] == 'OFF':
                                if (d+1) not in nurse_wanted_off[i]:
                                    is_fixed[i, d] = False
                                    fixed_shifts[i, d] = None
                                    unlocked_count += 1
                                    if unlocked_count >= deficit:
                                        break
                                        
                        unfixed_normals_count = sum(1 for i in range(num_nurses) if not is_fixed[i, d] and not is_night_keepers[i])
                        if unfixed_normals_count < required_normal_slots:
                            deficit = required_normal_slots - unfixed_normals_count
                            unlocked_count_wanted = 0
                            for i in range(num_nurses):
                                if not is_night_keepers[i] and is_fixed[i, d] and fixed_shifts[i, d] == 'OFF':
                                    is_fixed[i, d] = False
                                    fixed_shifts[i, d] = None
                                    unlocked_count_wanted += 1
                                    if unlocked_count_wanted >= deficit:
                                        break
                                        
                    unfixed_keepers_count = sum(1 for i in range(num_nurses) if not is_fixed[i, d] and is_night_keepers[i])
                    unfixed_normals_count = sum(1 for i in range(num_nurses) if not is_fixed[i, d] and not is_night_keepers[i])
                    total_unfixed = unfixed_normals_count + unfixed_keepers_count
                    total_required_slots = rem_D + rem_E + rem_N + rem_DE
                    
                    if total_unfixed < total_required_slots:
                        deficit = total_required_slots - total_unfixed
                        unlocked_count = 0
                        for i in range(num_nurses):
                            if not is_night_keepers[i] and is_fixed[i, d] and fixed_shifts[i, d] == 'OFF':
                                is_fixed[i, d] = False
                                fixed_shifts[i, d] = None
                                unlocked_count += 1
                                if unlocked_count >= deficit:
                                    break
                
                # 하이브리드 고정 스케줄 초기화
                sched = initialize_schedule_hybrid(num_nurses, num_days, requirements, is_fixed, fixed_shifts, is_night_keepers)
                
                # 벌점 계산기 함수 호출
                row_penalties = [get_nurse_penalty(sched[i], i, nurse_wanted_off[i], num_days, forbidden_5_patterns, is_night_keepers[i], nurse_histories[i], target_N_min, target_N_max, target_OFF_min, target_OFF_max, is_fixed[i], allowed_shifts_list[i]) for i in range(num_nurses)]
                
                # [가변 그룹 자동 균등 배치 연동]
                unique_groups = sorted(list(set(nurse_groups)))
                col_penalties = [get_day_penalty(sched[:, d], num_nurses, nurse_groups, unique_groups) for d in range(num_days)]
                
                total_penalty = sum(row_penalties) + sum(col_penalties)
                
                best_sched = copy.deepcopy(sched)
                best_penalty = total_penalty
                best_hard = sum(row_penalties)
                
                temp = 25.0
                cooling_rate = 0.99995  
                
                # 최적화 루프
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
                    
                    new_row_pen_i1 = get_nurse_penalty(sched[i1], i1, nurse_wanted_off[i1], num_days, forbidden_5_patterns, is_night_keepers[i1], nurse_histories[i1], target_N_min, target_N_max, target_OFF_min, target_OFF_max, is_fixed[i1], allowed_shifts_list[i1])
                    new_row_pen_i2 = get_nurse_penalty(sched[i2], i2, nurse_wanted_off[i2], num_days, forbidden_5_patterns, is_night_keepers[i2], nurse_histories[i2], target_N_min, target_N_max, target_OFF_min, target_OFF_max, is_fixed[i2], allowed_shifts_list[i2])
                    new_col_pen = get_day_penalty(sched[:, d], num_nurses, nurse_groups, unique_groups)
                    
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
                        
                st.session_state['violations_only'] = audit_schedule_violations(best_sched, nurses, requirements, audit_fixed, audit_fixed_values, nurse_histories, allowed_shifts_list, forbidden_5_patterns, prev_df is not None)
                st.session_state["optimized_result"] = df_clean
                st.balloons()
                
        if st.session_state["optimized_result"] is not None:
            st.write("### 🎉 생성 완료된 최종 근무표")
            st.dataframe(st.session_state["optimized_result"])

            # 추가 기능: 실제 위반 내역만 표시 (생성·다운로드는 유지)
            st.write('### 🔎 검수 결과 — 위반 사항')
            violations = st.session_state.get('violations_only')
            if violations is None:
                st.info('위반 검수를 보려면 근무표를 다시 생성해 주세요.')
            elif violations.empty:
                st.success('검수한 조건에서 위반 사항이 없습니다. 고정 N 자체와 선호사항·월 경계 확인 항목은 이 표에 포함하지 않습니다.')
            else:
                st.warning(f'위반 사항 {len(violations)}건이 있습니다. 직원·날짜별 내용을 확인해 주세요.')
                st.dataframe(violations, hide_index=True)
            
            # Excel 다운로드 기능 제공
            towrite = io.BytesIO()
            st.session_state["optimized_result"].to_excel(towrite, index=False, header=True)
            towrite.seek(0)
            
            st.download_button(
                label="📥 최종 근무표 Excel 다운로드",
                data=towrite,
                file_name="최종_근무표_맞춤형조건반영.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
else:
    st.info("👈 시작하려면 왼쪽 사이드바에서 '2. 초기 근무표 템플릿 업로드' 파일을 가장 먼저 업로드해 주세요.")
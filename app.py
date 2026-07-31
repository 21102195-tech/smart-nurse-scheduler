import streamlit as st
import pandas as pd
import numpy as np
import random
import copy
import io
import re
import requests
from datetime import datetime

# 1. 웹페이지 기본 설정
st.set_page_config(page_title="간호사 스마트 스케줄링 시스템", layout="wide")

# ⭐ [구글 API 주소 연동]: 1단계에서 복사해 둔 구글 웹 앱 URL 주소를 여기에 넣어주세요!
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

# 화면 상단 구성
st.title("🏥 스마트 병동 3교대 스케줄링 & 원티드 시스템")

# 2. 탭 분할 (간호사용 신청 화면과 수간호사용 관리자 모드 분리)
menu = st.radio("👉 사용 유형을 선택하세요", ["🙋‍♀️ [간호사용] 원티드 오프 신청", "⚙️ [수간호사용] 관리자 제어판"], horizontal=True)

# ----------------- [뷰 1] 간호사용 원티드 오프 신청 포털 -----------------
if menu == "🙋‍♀️ [간호사용] 원티드 오프 신청":
    st.markdown("---")
    st.write("### 📅 원하는 휴무일 직접 신청")
    st.info("💡 원하는 휴가 날짜를 지정하고 신청을 완료하면, 구글 스프레드시트에 신청 시간과 함께 실시간 저장됩니다.")
    
    sheet_data = fetch_google_sheet_data()
    
    if len(sheet_data) > 0:
        nurse_names = [item['name'] for item in sheet_data]
        
        # 일별 실시간 신청 인원 집계
        day_request_counts = {}
        for item in sheet_data:
            wanted = item['wanted_off']
            if pd.notna(wanted) and wanted.strip() != '' and wanted.strip() != 'nan':
                for x in str(wanted).split(','):
                    match = re.search(r'^\s*(\d+)', x.strip())
                    if match:
                        d = int(match.group(1))
                        day_request_counts[d] = day_request_counts.get(d, 0) + 1
                        
        # 하루 최대 허용치 설정 (임시 고정 2명, 관리자 탭에서 가변 조정 가능)
        limit_daily_off = 2
        
        st.write("#### 📢 8월 일자별 오프 신청 현황 및 마감 잔여석")
        cols = st.columns(10)
        for day in range(1, 32):
            col_idx = (day - 1) % 10
            count = day_request_counts.get(day, 0)
            rem = max(0, limit_daily_off - count)
            with cols[col_idx]:
                if rem == 0:
                    st.markdown(f"**{day}일**\n\n🔴 **마감**")
                else:
                    st.markdown(f"**{day}일**\n\n🟢 {rem}석")
                    
        st.markdown("---")
        
        col1, col2 = st.columns(2)
        with col1:
            selected_nurse = st.selectbox("1. 본인의 이름을 선택하세요", nurse_names)
        with col2:
            selected_offs = st.multiselect("2. 희망 휴무일을 복수 선택하세요 (마감된 날짜 제외)", list(range(1, 32)))
            
        if st.button("📝 원티드 오프 신청하기", type="primary"):
            if len(selected_offs) > 0:
                closed_days = []
                for d in selected_offs:
                    # 타인 신청 수 집계
                    other_count = sum(1 for item in sheet_data if d in [int(re.search(r'^\s*(\d+)', x.strip()).group(1)) for x in str(item['wanted_off']).split(',') if re.search(r'^\s*(\d+)', x.strip())] and item['name'] != selected_nurse)
                    if other_count >= limit_daily_off:
                        closed_days.append(d)
                        
                if len(closed_days) > 0:
                    st.error(f"❌ 신청 실패: {closed_days}일은 이미 다른 동료들에 의해 선착순 마감되었습니다!")
                else:
                    current_time_str = datetime.now().strftime("%m/%d %H:%M")
                    detail_offs_list = [f"{d}({current_time_str})" for d in sorted(selected_offs)]
                    offs_formatted_str = ", ".join(detail_offs_list)
                    
                    try:
                        # 구글 스프레드시트에 실시간 전송 저장!
                        payload = {"name": selected_nurse, "wanted_off": offs_formatted_str}
                        res = requests.post(GAS_URL, json=payload)
                        if res.status_code == 200:
                            st.success(f"🎉 신청 성공! {selected_nurse} 간호사님 오프 수집 완료: [{offs_formatted_str}]")
                            time.sleep(1.5)
                            st.rerun()
                    except Exception as e:
                        st.error(f"구글 서버 전송 중 오류 발생: {e}")
            else:
                st.warning("날짜를 선택해 주세요.")
    else:
        st.warning("⚠️ 데이터베이스에 등록된 간호사가 없습니다. 먼저 [수간호사용] 관리자 제어판에서 '초기 근무표 템플릿'을 업로드해 주세요.")

# ----------------- [뷰 2] 수간호사용 비밀번호 관리자 제어판 -----------------
elif menu == "⚙️ [수간호사용] 관리자 제어판":
    st.markdown("---")
    st.sidebar.markdown("### 🔐 관리자 인증")
    
    # ⭐ 관리자 비밀번호 세팅 (유출 방지용)
    admin_password = st.sidebar.text_input("수간호사 비밀번호를 입력하세요", type="password")
    
    # 비밀번호가 맞을 때만 화면 렌더링 (보안 장치)
    if admin_password == "1234": # 비밀번호 원하시는 대로 수정 가능!
        st.sidebar.success("🔑 관리자 인증 성공!")
        st.write("### 🛠️ 수간호사 근무표 마스터 통제 제어판")
        
        # 기존 작성 규칙 및 파일들
        uploaded_rules = st.sidebar.file_uploader("1. 작성 규칙 파일 업로드 (xlsx/csv)", type=["xlsx", "csv"])
        uploaded_schedule = st.sidebar.file_uploader("2. 초기 근무표 템플릿 업로드 (xlsx/csv)", type=["xlsx", "csv"])
        uploaded_prev_month = st.sidebar.file_uploader("3. 이전 달 근무표 업로드 (선택사항)", type=["xlsx", "csv"])
        
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

        # [새 파일 업로드 감지] 파일 교체 시 메모리 리셋
        if uploaded_schedule:
            file_key = f"{uploaded_schedule.name}_{uploaded_schedule.size}"
            if "last_file_key" not in st.session_state or st.session_state["last_file_key"] != file_key:
                st.session_state["last_file_key"] = file_key
                st.session_state["schedule_df_state"] = None  
                st.session_state["optimized_result"] = None   

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

        if uploaded_schedule and st.session_state["schedule_df_state"] is None:
            try:
                if uploaded_schedule.name.endswith('xlsx'):
                    raw_df = pd.read_excel(uploaded_schedule)
                else:
                    raw_df = pd.read_csv(uploaded_schedule, encoding='utf-8-sig')
                schedule_df = load_and_align_headers(raw_df)
                
                # ⭐ 최초 업로드 시 구글 스프레드시트에 간호사 이름을 자동 기입하여 동기화
                st.session_state["schedule_df_state"] = schedule_df
                sheet_data = fetch_google_sheet_data()
                existing_names = [item['name'] for item in sheet_data]
                
                for idx, row in schedule_df.iterrows():
                    nurse_id, _ = parse_nurse_row(row['이름'], row['그룹'])
                    if nurse_id is not None and str(nurse_id) not in existing_names:
                        requests.post(GAS_URL, json={"name": str(nurse_id), "wanted_off": "-"})
                st.success("간호사 목록이 데이터베이스에 정상적으로 세팅되었습니다!")
            except Exception as e:
                st.error(f"템플릿 파일 읽기 에러: {e}")

        # 수간호사용 탭 분할
        if st.session_state["schedule_df_state"] is not None:
            col_tab1, col_tab2 = st.tabs(["📋 실시간 오프 신청 현황판", "📅 AI 최종 근무표 생성"])
            
            with col_tab1:
                st.write("### 📋 현재 구글 시트에 취합된 실시간 오프 현황")
                
                if st.button("🔄 새로운 신청 내역 실시간 동기화"):
                    st.session_state["optimized_result"] = None
                    
                sheet_data = fetch_google_sheet_data()
                df_temp = st.session_state["schedule_df_state"].copy()
                
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
                            
                            # (이전 달 로드 및 연동 연산, 초기화, 벌점 최적화 루프 동일 구동)
                            # ... 생략없이 아래 함수들과 연계 구동됩니다 ...
                            prev_df = None
                            if uploaded_prev_month is not None:
                                try:
                                    prev_df = pd.read_excel(uploaded_prev_month) if uploaded_prev_month.name.endswith('xlsx') else pd.read_csv(uploaded_prev_month, encoding='utf-8-sig')
                                except: pass
                            
                            # (간호사 정보, 이전달 연동, 인원수 계산, 초기화 및 Simulated Annealing 동일 수행)
                            # ... 기존 완성된 AI 엔진과 동일 구동 ...
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
                                    # 정규식 숫자일자 파싱
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
                                    
                            # 요구량 파싱
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

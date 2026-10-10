import calendar
from datetime import datetime
import json
import os
import pandas as pd
import openpyxl
import streamlit as st
from ortools.sat.python import cp_model

# 페이지 기본 설정
st.set_page_config(page_title="물류센터 근무계획표 시스템", layout="wide")

DATA_FILE = "employees.json"

# ==============================================================================
# 0. 데이터 영구 저장/로드 함수 (JSON 파일 방식)
# ==============================================================================
def load_employees():
    """파일에서 근무자 명단을 불러옵니다. 파일이 없으면 초기 데이터 생성."""
    if os.path.exists(DATA_FILE):
        try:
            with open(DATA_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
            
    # 초기 기본 명단 (admin 계정 포함)
    initial_data = [
        {'사번': 'admin', '비밀번호': 'admin', '성명': '관리자', '직급': '팀장', '구분': '주5일', '기본휴무일': [], '선호짝꿍': '', '비선호짝꿍': '', '입사일': '2023-01-01', '연차': 15, '휴가': 3, '관리자': True},
        {'사번': '1001', '비밀번호': '1001', '성명': '최광호', '직급': '팀장', '구분': '주5일', '기본휴무일': [], '선호짝꿍': '', '비선호짝꿍': '', '입사일': '2023-01-01', '연차': 15, '휴가': 3, '관리자': True},
        {'사번': '1002', '비밀번호': '1002', '성명': '권민중', '직급': '부팀장', '구분': '야간', '기본휴무일': [], '선호짝꿍': '이영희', '비선호짝꿍': '김철수', '입사일': '2024-05-01', '연차': 15, '휴가': 3, '관리자': False},
        {'사번': '1003', '비밀번호': '1003', '성명': '김철수', '직급': '대리', '구분': '야간', '기본휴무일': [], '선호짝꿍': '', '비선호짝꿍': '권민중', '입사일': '2025-02-01', '연차': 15, '휴가': 3, '관리자': False},
        {'사번': '1004', '비밀번호': '1004', '성명': '이영희', '직급': '주임', '구분': '야간', '기본휴무일': [], '선호짝꿍': '권민중', '비선호짝꿍': '', '입사일': '2025-11-01', '연차': 15, '휴가': 3, '관리자': False},
        {'사번': '1005', '비밀번호': '1005', '성명': '박민수', '직급': '사원', '구분': '주간스케줄', '기본휴무일': ['수', '목'], '선호짝꿍': '', '비선호짝꿍': '', '입사일': '2026-02-01', '연차': 10, '휴가': 3, '관리자': False},
    ]
    save_employees(initial_data)
    return initial_data

def save_employees(data):
    """근무자 명단을 JSON 파일로 저장합니다."""
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4)

def get_sorted_employees(emp_list):
    """주5일 -> 주간스케줄 -> 야간 순, 2순위 사번 오름차순 정렬"""
    df = pd.DataFrame(emp_list)
    if df.empty:
        return df
    
    # 근무 구분 우선순위 맵핑
    type_order = {'주5일': 1, '주간스케줄': 2, '야간': 3}
    df['type_rank'] = df['구분'].map(lambda x: type_order.get(x, 99))
    
    # 1순위: type_rank 오름차순, 2순위: 사번 오름차순 정렬
    df_sorted = df.sort_values(by=['type_rank', '사번'], ascending=[True, True])
    df_sorted = df_sorted.drop(columns=['type_rank'])
    return df_sorted

# ==============================================================================
# 1. 세션 상태 초기화
# ==============================================================================
if 'logged_in_user' not in st.session_state:
    st.session_state.logged_in_user = None

if 'is_admin' not in st.session_state:
    st.session_state.is_admin = False

if 'employees' not in st.session_state:
    st.session_state.employees = load_employees()

if 'requests' not in st.session_state:
    st.session_state.requests = []

# ==============================================================================
# 2. 첫 페이지: 로그인 화면 (미인증 시 게이트)
# ==============================================================================
if not st.session_state.logged_in_user:
    st.title("🔒 물류센터 ERP 시스템 - 로그인")
    st.markdown("---")
    
    col_a, col_b, col_c = st.columns([1, 2, 1])
    with col_b:
        st.subheader("계정 정보 입력")
        login_id = st.text_input("아이디 (사원번호)", placeholder="예: admin 또는 1001")
        login_pw = st.text_input("비밀번호", type="password", placeholder="비밀번호 입력")
        
        if st.button("로그인", use_container_width=True):
            user_info = next((e for e in st.session_state.employees if e['사번'] == login_id.strip()), None)
            
            if user_info and user_info.get('비밀번호') == login_pw.strip():
                st.session_state.logged_in_user = user_info
                st.session_state.is_admin = user_info.get('관리자', False)
                st.success(f"{user_info['성명']} 님 환영합니다!")
                st.rerun()
            else:
                st.error("아이디(사번) 또는 비밀번호가 올바르지 않습니다.")
                
    st.stop()  # 로그인되지 않았으면 메인화면 미출력

# ==============================================================================
# 3. 메인 시스템 화면 (로그인 후 접근 가능)
# ==============================================================================
user = st.session_state.logged_in_user
role_badge = "👑 관리자" if st.session_state.is_admin else "👤 일반 사용자"

st.sidebar.markdown(f"### 👤 {user['성명']} 님")
st.sidebar.write(f"- **사번:** {user['사번']}")
st.sidebar.write(f"- **직급/구분:** {user['직급']} / {user['구분']}")
st.sidebar.write(f"- **권한:** {role_badge}")

if st.sidebar.button("🚪 로그아웃", use_container_width=True):
    st.session_state.logged_in_user = None
    st.session_state.is_admin = False
    st.rerun()

st.sidebar.markdown("---")

st.title("📅 물류센터 월간 근무계획표 자동 생성 시스템")

# 관리자 여부에 따라 탭 목록 생성
tab_titles = ["👥 1. 근무자 명단", "📝 2. 근무조정 요청 입력", "🚀 3. 근무표 자동 생성"]
if st.session_state.is_admin:
    tab_titles.append("🛠️ 4. [관리자] 근무자 정보 관리")

tabs = st.tabs(tab_titles)

# ------------------------------------------------------------------------------
# TAB 1: 근무자 명단 (정렬 적용)
# ------------------------------------------------------------------------------
with tabs[0]:
    st.subheader("📋 전체 근무자 명단")
    
    if st.session_state.employees:
        df_sorted = get_sorted_employees(st.session_state.employees)
        
        df_sorted['성명(표식)'] = df_sorted.apply(
            lambda r: f"👑 {r['성명']} (관리자)" if r.get('관리자', False) else r['성명'], axis=1
        )
        df_sorted['기본휴무일_표시'] = df_sorted['기본휴무일'].apply(
            lambda x: ", ".join(x) if isinstance(x, list) and x else "-"
        )
        
        target_cols = ['사번', '성명(표식)', '직급', '구분', '기본휴무일_표시', '선호짝꿍', '비선호짝꿍', '입사일', '관리자']
        valid_cols = [c for c in target_cols if c in df_sorted.columns]
        
        rename_map = {'기본휴무일_표시': '기본휴무일(주간스케줄)'}
        st.dataframe(df_sorted[valid_cols].rename(columns=rename_map), use_container_width=True)

# ------------------------------------------------------------------------------
# TAB 2: 근무조정 요청 입력
# ------------------------------------------------------------------------------
with tabs[1]:
    st.subheader("📝 팀원별 근무조정 요청 입력")
    col1, col2, col3, col4 = st.columns(4)
    emp_names = [e['성명'] for e in st.session_state.employees]
    
    with col1:
        req_emp = st.selectbox("근무자 선택", emp_names, key="req_emp_sb")
    with col2:
        req_day = st.number_input("요청 일자(일)", min_value=1, max_value=31, value=13)
    with col3:
        req_type = st.selectbox("요청 근무 구분", ["대휴", "연차", "휴가", "비번", "공가", "휴무"])
    with col4:
        st.write(" ")
        if st.button("요청 추가"):
            st.session_state.requests.append({'성명': req_emp, '일자': req_day, '요청근무': req_type})
            st.success(f"{req_emp} 님 {req_day}일 {req_type} 요청 등록 완료!")

    if st.session_state.requests:
        st.write("현재 등록된 요청 목록:")
        st.dataframe(pd.DataFrame(st.session_state.requests), use_container_width=True)
        if st.button("요청 목록 전체 초기화"):
            st.session_state.requests = []
            st.rerun()

# ------------------------------------------------------------------------------
# TAB 3: 근무표 자동 생성 Engine (OR-Tools)
# ------------------------------------------------------------------------------
with tabs[2]:
    st.subheader("🚀 월간 스케줄 자동 생성")
    
    col_y, col_m = st.columns(2)
    with col_y:
        target_year = st.number_input("연도 선택", min_value=2025, max_value=2030, value=2026)
    with col_m:
        target_month = st.number_input("월 선택", min_value=1, max_value=12, value=11)
        
    if st.button("🔥 스케줄 최적화 실행 및 근무표 생성"):
        with st.spinner("구글 OR-Tools 엔진이 최적 조합을 계산 중입니다..."):
            num_days = calendar.monthrange(target_year, target_month)[1]
            days = list(range(1, num_days + 1))
            
            model = cp_model.CpModel()
            shift_map = {'주': 0, '야3': 1, '대': 2, '연차': 3, '휴가': 4, '공가': 5, '비': 6, '휴무': 7}
            code_to_name = {v: k for k, v in shift_map.items()}
            
            employees = [e for e in st.session_state.employees if e['성명']]
            num_emp = len(employees)
            
            x = {}
            for e in range(num_emp):
                for d in days:
                    for s in range(8):
                        x[(e, d, s)] = model.NewBoolVar(f'x_{e}_{d}_{s}')
                        
            # 1. 하루 1인 1근무
            for e in range(num_emp):
                for d in days:
                    model.AddExactlyOne(x[(e, d, s)] for s in range(8))
                    
            # 2. 신청 요청사항 반영
            for req in st.session_state.requests:
                e_idx = next((i for i, e in enumerate(employees) if e['성명'] == req['성명']), None)
                if e_idx is not None and req['일자'] <= num_days:
                    s_code = shift_map.get(req['요청근무'], 2)
                    model.Add(x[(e_idx, req['일자'], s_code)] == 1)
                    
            # 3. 야간 근무자 조건
            night_indices = [i for i, e in enumerate(employees) if e['구분'] == '야간']
            for d in days:
                if len(night_indices) >= 2:
                    model.Add(sum(x[(e, d, 1)] for e in night_indices) == 2)
                    for e in night_indices:
                        dis_name = employees[e].get('비선호짝꿍', '')
                        if dis_name:
                            dis_idx = next((i for i, emp in enumerate(employees) if emp['성명'] == dis_name), None)
                            if dis_idx is not None:
                                model.AddBoolOr([x[(e, d, 1)].Not(), x[(dis_idx, d, 1)].Not()])

            # 4. 주5일 및 주간 스케줄 기본 조건
            for e, emp in enumerate(employees):
                for d in days:
                    dt = datetime(target_year, target_month, d)
                    weekday_str = ["월", "화", "수", "목", "금", "토", "일"][dt.weekday()]
                    
                    if emp['구분'] == '주5일':
                        if weekday_str in ["토", "일"]:
                            model.Add(x[(e, d, 2)] == 1)
                    elif emp['구분'] == '주간스케줄':
                        pref_off_list = emp.get('기본휴무일', [])
                        if weekday_str in pref_off_list:
                            model.Add(x[(e, d, 2)] == 1)

            solver = cp_model.CpSolver()
            solver.parameters.max_time_in_seconds = 10.0
            status = solver.Solve(model)
            
            if status in (cp_model.OPTIMAL, cp_model.FEASIBLE):
                st.success("🎉 최적의 근무표가 성공적으로 생성되었습니다!")
                
                schedule_data = []
                for e, emp in enumerate(employees):
                    row = {'사번': emp['사번'], '성명': emp['성명'], '직급': emp['직급'], '구분': emp['구분']}
                    for d in days:
                        assigned = [s for s in range(8) if solver.Value(x[(e, d, s)]) == 1][0]
                        row[f"{d}일"] = code_to_name[assigned]
                    schedule_data.append(row)
                    
                df_res = pd.DataFrame(schedule_data)
                st.dataframe(df_res, use_container_width=True)
                
                output_filename = f"{target_year}년_{target_month}월_근무계획표.xlsx"
                df_res.to_excel(output_filename, index=False)
                
                with open(output_filename, "rb") as file:
                    st.download_button(
                        label="📥 엑셀 파일(.xlsx) 다운로드",
                        data=file,
                        file_name=output_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            else:
                st.error("❌ 제약 조건이 너무 까다로워 스케줄을 생성하지 못했습니다.")

# ------------------------------------------------------------------------------
# TAB 4: [관리자 전용] 근무자 정보 관리 (입사일 수정 포함)
# ------------------------------------------------------------------------------
if st.session_state.is_admin:
    with tabs[3]:
        st.subheader("🛠️ [관리자 전용] 근무자 정보 추가 / 수정 / 삭제")
        
        edit_mode = st.radio("작업 선택", ["신규 근무자 등록", "기존 근무자 정보 수정/삭제"], horizontal=True)
        days_list = ["월", "화", "수", "목", "금", "토", "일"]
        
        # 1. 신규 근무자 등록
        if edit_mode == "신규 근무자 등록":
            c1, c2, c3 = st.columns(3)
            with c1:
                new_id = st.text_input("아이디 (사번)")
                new_pw = st.text_input("비밀번호 (미입력 시 사번과 동일)", type="password")
                new_name = st.text_input("성명")
                new_rank = st.selectbox("직급", ["팀장", "부팀장", "대리", "주임", "사원"])
            with c2:
                new_type = st.selectbox("근무 구분", ["주5일", "주간스케줄", "야간"])
                
                is_schedule = (new_type == "주간스케줄")
                st.write("**기본 휴무일 선택 (주간스케줄 전용)**")
                new_off_days = []
                cb_cols = st.columns(7)
                for idx, day_str in enumerate(days_list):
                    with cb_cols[idx]:
                        checked = st.checkbox(day_str, disabled=not is_schedule, key=f"new_cb_{day_str}")
                        if checked and is_schedule:
                            new_off_days.append(day_str)
                            
            with c3:
                new_pref = st.text_input("선호 짝꿍 (야간 전용)")
                new_dislike = st.text_input("비선호 짝꿍 (야간 전용)")
                new_join = st.date_input("입사일").strftime("%Y-%m-%d")
                is_admin_check = st.checkbox("관리자 권한 부여")
                
            if st.button("신규 근무자 등록 저장", use_container_width=True):
                if not new_id.strip() or not new_name.strip():
                    st.error("사번과 성명은 필수 입력 사항입니다.")
                elif any(e['사번'] == new_id.strip() for e in st.session_state.employees):
                    st.error("이미 등록된 사번입니다.")
                else:
                    pw_final = new_pw.strip() if new_pw.strip() else new_id.strip()
                    st.session_state.employees.append({
                        '사번': new_id.strip(), '비밀번호': pw_final, '성명': new_name.strip(),
                        '직급': new_rank, '구분': new_type, '기본휴무일': new_off_days,
                        '선호짝꿍': new_pref.strip(), '비선호짝꿍': new_dislike.strip(),
                        '입사일': new_join, '연차': 15, '휴가': 3, '관리자': is_admin_check
                    })
                    save_employees(st.session_state.employees)
                    st.success(f"근무자 '{new_name}'(사번: {new_id}) 님이 새로 등록되었습니다.")
                    st.rerun()

        # 2. 기존 근무자 정보 수정 및 삭제 (입사일 수정 가능)
        else:
            emp_options = [f"{e['사번']} - {e['성명']}" for e in st.session_state.employees]
            selected_emp_str = st.selectbox("수정할 근무자 선택", emp_options)
            
            if selected_emp_str:
                target_id = selected_emp_str.split(" - ")[0]
                target_emp = next((e for e in st.session_state.employees if e['사번'] == target_id), None)
                
                if target_emp:
                    c1, c2, c3 = st.columns(3)
                    with c1:
                        mod_id = st.text_input("아이디 (사번)", value=target_emp['사번'])
                        mod_pw = st.text_input("비밀번호", value=target_emp.get('비밀번호', ''), type="password")
                        mod_name = st.text_input("성명", value=target_emp['성명'])
                        mod_rank = st.selectbox("직급", ["팀장", "부팀장", "대리", "주임", "사원"], index=["팀장", "부팀장", "대리", "주임", "사원"].index(target_emp['직급']))
                    with c2:
                        mod_type = st.selectbox("근무 구분", ["주5일", "주간스케줄", "야간"], index=["주5일", "주간스케줄", "야간"].index(target_emp['구분']))
                        
                        is_mod_schedule = (mod_type == "주간스케줄")
                        st.write("**기본 휴무일 선택 (주간스케줄 전용)**")
                        mod_off_days = []
                        cb_cols2 = st.columns(7)
                        curr_off = target_emp.get('기본휴무일', [])
                        
                        for idx, day_str in enumerate(days_list):
                            with cb_cols2[idx]:
                                is_default = (day_str in curr_off) if is_mod_schedule else False
                                checked = st.checkbox(day_str, value=is_default, disabled=not is_mod_schedule, key=f"mod_cb_{day_str}")
                                if checked and is_mod_schedule:
                                    mod_off_days.append(day_str)
                                    
                    with c3:
                        mod_pref = st.text_input("선호 짝꿍", value=target_emp.get('선호짝꿍', ''))
                        mod_dislike = st.text_input("비선호 짝꿍", value=target_emp.get('비선호짝꿍', ''))
                        
                        # 기존 입사일 파싱하여 date_input 기본값으로 지정
                        curr_join_str = target_emp.get('입사일', '2026-01-01')
                        try:
                            curr_join_dt = datetime.strptime(curr_join_str, "%Y-%m-%d").date()
                        except ValueError:
                            curr_join_dt = datetime.now().date()
                        
                        mod_join = st.date_input("입사일", value=curr_join_dt).strftime("%Y-%m-%d")
                        mod_admin = st.checkbox("관리자 권한 부여", value=target_emp.get('관리자', False))
                        
                    btn_c1, btn_c2 = st.columns(2)
                    with btn_c1:
                        if st.button("수정 내용 저장", use_container_width=True):
                            target_emp['사번'] = mod_id.strip()
                            target_emp['비밀번호'] = mod_pw.strip()
                            target_emp['성명'] = mod_name.strip()
                            target_emp['직급'] = mod_rank
                            target_emp['구분'] = mod_type
                            target_emp['기본휴무일'] = mod_off_days
                            target_emp['선호짝꿍'] = mod_pref.strip()
                            target_emp['비선호짝꿍'] = mod_dislike.strip()
                            target_emp['입사일'] = mod_join
                            target_emp['관리자'] = mod_admin
                            save_employees(st.session_state.employees)
                            st.success(f"'{mod_name}' 님의 정보가 수정되었습니다.")
                            st.rerun()
                    with btn_c2:
                        if st.button("❌ 근무자 삭제", use_container_width=True):
                            st.session_state.employees = [e for e in st.session_state.employees if e['사번'] != target_id]
                            save_employees(st.session_state.employees)
                            st.success(f"사번 '{target_id}' 근무자가 삭제되었습니다.")
                            st.rerun()

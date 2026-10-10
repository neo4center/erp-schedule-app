import calendar
from datetime import datetime
import pandas as pd
import openpyxl
import streamlit as st
from ortools.sat.python import cp_model

# 페이지 기본 설정
st.set_page_config(page_title="물류센터 근무계획표 시스템", layout="wide")

# ==============================================================================
# 1. 세션 상태 및 초기 데이터 설정 (사번 체계 반영)
# ==============================================================================
if 'logged_in_user' not in st.session_state:
    st.session_state.logged_in_user = None  # 로그인한 사용자 정보 (dict)

if 'is_admin' not in st.session_state:
    st.session_state.is_admin = False  # 관리자 권한 여부

# 초기 근무자 명단 (사번/비밀번호/직급/구분 등)
if 'employees' not in st.session_state:
    st.session_state.employees = [
        {'사번': '1001', '비밀번호': '1001', '성명': '최광호', '직급': '팀장', '구분': '주5일', '기본휴무일_선호짝꿍': '', '비선호짝꿍': '', '입사일': '2023-01-01', '연차': 15, '휴가': 3, '관리자': True},
        {'사번': '1002', '비밀번호': '1002', '성명': '권민중', '직급': '부팀장', '구분': '야간', '기본휴무일_선호짝꿍': '이영희', '비선호짝꿍': '김철수', '입사일': '2024-05-01', '연차': 15, '휴가': 3, '관리자': False},
        {'사번': '1003', '비밀번호': '1003', '성명': '김철수', '직급': '대리', '구분': '야간', '기본휴무일_선호짝꿍': '', '비선호짝꿍': '권민중', '입사일': '2025-02-01', '연차': 15, '휴가': 3, '관리자': False},
        {'사번': '1004', '비밀번호': '1004', '성명': '이영희', '직급': '주임', '구분': '야간', '기본휴무일_선호짝꿍': '권민중', '비선호짝꿍': '', '입사일': '2025-11-01', '연차': 15, '휴가': 3, '관리자': False},
        {'사번': '1005', '비밀번호': '1005', '성명': '박민수', '직급': '사원', '구분': '주간스케줄', '기본휴무일_선호짝꿍': '수,목', '비선호짝꿍': '', '입사일': '2026-02-01', '연차': 10, '휴가': 3, '관리자': False},
    ]

if 'requests' not in st.session_state:
    st.session_state.requests = []

# ==============================================================================
# 2. 사이드바 사번 기반 로그인 시스템
# ==============================================================================
st.sidebar.title("🔐 계정 로그인")

if not st.session_state.logged_in_user:
    input_emp_id = st.sidebar.text_input("사원번호(사번)")
    input_pw = st.sidebar.text_input("비밀번호", type="password")
    
    if st.sidebar.button("로그인"):
        user_info = next((e for e in st.session_state.employees if e['사번'] == input_emp_id.strip()), None)
        
        if user_info:
            if user_info.get('비밀번호') == input_pw.strip():
                st.session_state.logged_in_user = user_info
                st.session_state.is_admin = user_info.get('관리자', False)
                role_msg = "👑 관리자" if st.session_state.is_admin else "👤 일반 사용자"
                st.sidebar.success(f"{user_info['성명']} 님 로그인 완료 ({role_msg})")
                st.rerun()
            else:
                st.sidebar.error("비밀번호가 일치하지 않습니다.")
        else:
            st.sidebar.error("존재하지 않는 사원번호입니다.")
else:
    user = st.session_state.logged_in_user
    role_label = "👑 관리자" if st.session_state.is_admin else "👤 일반 사용자"
    st.sidebar.write(f"**접속자:** {user['성명']} ({user['사번']})")
    st.sidebar.write(f"**권한:** {role_label}")
    
    if st.sidebar.button("로그아웃"):
        st.session_state.logged_in_user = None
        st.session_state.is_admin = False
        st.rerun()

# ==============================================================================
# 3. 메인 화면 및 탭 구성
# ==============================================================================
st.title("📅 물류센터 월간 근무계획표 자동 생성 시스템")

tab1, tab2, tab3 = st.tabs(["👥 1. 근무자 명단 및 권한 관리", "📝 2. 근무조정 요청 입력", "🚀 3. 근무표 자동 생성"])

# ------------------------------------------------------------------------------
# TAB 1: 근무자 명단 관리
# ------------------------------------------------------------------------------
with tab1:
    st.subheader("📋 전체 근무자 명단")
    
    if st.session_state.employees:
        df_display = pd.DataFrame(st.session_state.employees)
        
        # 관리자 표식 컬럼 생성
        df_display['성명(표식)'] = df_display.apply(
            lambda r: f"👑 {r['성명']} (관리자)" if r.get('관리자', False) else r['성명'], axis=1
        )
        
        # 안전한 열 선택 (KeyError 방지)
        target_cols = ['사번', '성명(표식)', '직급', '구분', '기본휴무일_선호짝꿍', '비선호짝꿍', '입사일', '연차', '휴가', '관리자']
        valid_cols = [col for col in target_cols if col in df_display.columns]
        
        st.dataframe(df_display[valid_cols], use_container_width=True)
    else:
        st.info("등록된 근무자가 없습니다.")
        
    st.markdown("---")
    
    # --- 관리자 전용 편집 기능 ---
    if st.session_state.is_admin:
        st.subheader("🛠️ [관리자 전용] 근무자 정보 추가 / 수정 / 삭제")
        
        edit_mode = st.radio("작업 선택", ["신규 근무자 등록", "근무자 정보 수정/삭제"], horizontal=True)
        
        # 1. 신규 등록
        if edit_mode == "신규 근무자 등록":
            with st.form("add_emp_form"):
                c1, c2, c3 = st.columns(3)
                with c1:
                    new_id = st.text_input("사원번호(사번)")
                    new_pw = st.text_input("비밀번호 (미입력 시 사번과 동일)", type="password")
                    new_name = st.text_input("성명")
                with c2:
                    new_rank = st.selectbox("직급", ["팀장", "부팀장", "대리", "주임", "사원"])
                    new_type = st.selectbox("근무 구분", ["주5일", "주간스케줄", "야간"])
                    new_pref = st.text_input("기본휴무일(수,목) 또는 선호짝꿍")
                with c3:
                    new_dislike = st.text_input("비선호짝꿍")
                    new_join = st.date_input("입사일").strftime("%Y-%m-%d")
                    is_admin_check = st.checkbox("관리자 권한 부여")
                    
                submit_add = st.form_submit_button("신규 근무자 추가")
                
                if submit_add:
                    if new_id.strip() == "" or new_name.strip() == "":
                        st.error("사번과 성명은 필수 입력 항목입니다.")
                    elif any(e['사번'] == new_id.strip() for e in st.session_state.employees):
                        st.error("이미 존재하는 사원번호입니다.")
                    else:
                        pw_final = new_pw.strip() if new_pw.strip() else new_id.strip()
                        st.session_state.employees.append({
                            '사번': new_id.strip(), '비밀번호': pw_final, '성명': new_name.strip(),
                            '직급': new_rank, '구분': new_type,
                            '기본휴무일_선호짝꿍': new_pref.strip(), '비선호짝꿍': new_dislike.strip(),
                            '입사일': new_join, '연차': 15, '휴가': 3, '관리자': is_admin_check
                        })
                        st.success(f"'{new_name}'(사번: {new_id}) 님이 등록되었습니다!")
                        st.rerun()

        # 2. 기존 정보 수정 및 삭제
        else:
            emp_list = [f"{e['사번']} - {e['성명']}" for e in st.session_state.employees]
            selected_emp_str = st.selectbox("수정/삭제할 근무자 선택", emp_list)
            
            if selected_emp_str:
                selected_id = selected_emp_str.split(" - ")[0]
                target_emp = next((e for e in st.session_state.employees if e['사번'] == selected_id), None)
                
                if target_emp:
                    with st.form("edit_emp_form"):
                        c1, c2, c3 = st.columns(3)
                        with c1:
                            mod_name = st.text_input("성명", value=target_emp['성명'])
                            mod_pw = st.text_input("비밀번호 변경", value=target_emp.get('비밀번호', ''), type="password")
                            mod_rank = st.selectbox("직급", ["팀장", "부팀장", "대리", "주임", "사원"], index=["팀장", "부팀장", "대리", "주임", "사원"].index(target_emp['직급']))
                        with c2:
                            mod_type = st.selectbox("근무 구분", ["주5일", "주간스케줄", "야간"], index=["주5일", "주간스케줄", "야간"].index(target_emp['구분']))
                            mod_pref = st.text_input("기본휴무일 또는 선호짝꿍", value=target_emp['기본휴무일_선호짝꿍'])
                        with c3:
                            mod_dislike = st.text_input("비선호짝꿍", value=target_emp['비선호짝꿍'])
                            mod_admin = st.checkbox("관리자 권한 부여", value=target_emp.get('관리자', False))
                            
                        c_btn1, c_btn2 = st.columns(2)
                        with c_btn1:
                            submit_mod = st.form_submit_button("정보 수정 저장")
                        with c_btn2:
                            submit_del = st.form_submit_button("❌ 근무자 삭제")
                            
                        if submit_mod:
                            target_emp['성명'] = mod_name.strip()
                            target_emp['비밀번호'] = mod_pw.strip()
                            target_emp['직급'] = mod_rank
                            target_emp['구분'] = mod_type
                            target_emp['기본휴무일_선호짝꿍'] = mod_pref.strip()
                            target_emp['비선호짝꿍'] = mod_dislike.strip()
                            target_emp['관리자'] = mod_admin
                            st.success(f"'{mod_name}' 님의 정보가 수정되었습니다.")
                            st.rerun()
                            
                        if submit_del:
                            st.session_state.employees = [e for e in st.session_state.employees if e['사번'] != selected_id]
                            st.success(f"사번 '{selected_id}' 근무자가 삭제되었습니다.")
                            st.rerun()
    else:
        st.info("🔒 근무자 등록/수정/삭제 및 권한 관리는 **관리자 계정으로 로그인** 시 가능합니다. (초기 관리자: 사번 1001 / 비번 1001)")

# ------------------------------------------------------------------------------
# TAB 2: 근무조정 요청 입력
# ------------------------------------------------------------------------------
with tab2:
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
with tab3:
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
                        pref_off = str(emp.get('기본휴무일_선호짝꿍', ''))
                        if weekday_str in pref_off:
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

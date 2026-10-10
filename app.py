import calendar
from datetime import datetime
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import streamlit as st
from ortools.sat.python import cp_model

# 페이지 기본 설정
st.set_page_config(page_title="물류센터 월간 근무계획표 자동 생성기", layout="wide")

st.title("📅 월간 근무계획표 자동 생성 & 스케줄 관리 시스템")
st.markdown("---")

# 세션 상태 초기화 (데이터 유지용)
if 'employees' not in st.session_state:
    st.session_state.employees = [
        {'성명': '최광호', '직급': '팀장', '구분': '주5일', '기본휴무일_선호짝꿍': '', '비선호짝꿍': '', '입사일': '2023-01-01', '연차': 15, '휴가': 3},
        {'성명': '권민중', '직급': '선임', '구분': '야간', '기본휴무일_선호짝꿍': '이영희', '비선호짝꿍': '김철수', '입사일': '2024-05-01', '연차': 15, '휴가': 3},
        {'성명': '김철수', '직급': '후임', '구분': '야간', '기본휴무일_선호짝꿍': '', '비선호짝꿍': '권민중', '입사일': '2025-02-01', '연차': 15, '휴가': 3},
        {'성명': '이영희', '직급': '후임', '구분': '야간', '기본휴무일_선호짝꿍': '권민중', '비선호짝꿍': '', '입사일': '2025-11-01', '연차': 15, '휴가': 3},
        {'성명': '박민수', '직급': '사원', '구분': '주간스케줄', '기본휴무일_선호짝꿍': '수,목', '비선호짝꿍': '', '입사일': '2026-02-01', '연차': 10, '휴가': 3},
    ]

if 'requests' not in st.session_state:
    st.session_state.requests = []

# 탭 구성 (1. 근무자 관리 / 2. 근무조정 요청 / 3. 근무표 생성)
tab1, tab2, tab3 = st.tabs(["👥 1. 근무자 명단 및 조건 관리", "📝 2. 팀원별 근무조정 요청 입력", "🚀 3. 근무표 자동 생성 및 다운로드"])

# ==============================================================================
# TAB 1: 근무자 명단 관리
# ==============================================================================
with tab1:
    st.subheader("근무자 기본 정보 및 야간 짝꿍 / 휴무 조건 설정")
    df_emp = pd.DataFrame(st.session_state.employees)
    edited_df = st.data_editor(df_emp, num_rows="dynamic", use_container_width=True)
    st.session_state.employees = edited_df.to_dict('records')
    st.info("💡 '구분' 항목: 주5일, 주간스케줄, 야간 중 선택 / 주간스케줄은 '수,목' 형태 입력 / 야간은 선호/비선호 짝꿍 이름을 입력하세요.")

# ==============================================================================
# TAB 2: 근무조정 요청사항 입력
# ==============================================================================
with tab2:
    st.subheader("팀원별 개별 근태 및 휴무 요청 등록")
    
    col1, col2, col3, col4 = st.columns(4)
    emp_names = [e['성명'] for e in st.session_state.employees if e['성명']]
    
    with col1:
        req_emp = st.selectbox("근무자 선택", emp_names)
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
            st.experimental_rerun()

# ==============================================================================
# TAB 3: 스케줄 생성 엔진 (OR-Tools)
# ==============================================================================
with tab3:
    st.subheader("월간 스케줄 자동 생성")
    
    col_y, col_m = st.columns(2)
    with col_y:
        target_year = st.number_input("연도 선택", min_value=2025, max_value=2030, value=2026)
    with col_m:
        target_month = st.number_input("월 선택", min_value=1, max_value=12, value=11)
        
    if st.button("🔥 스케줄 최적화 실행 및 근무표 생성"):
        with st.spinner("구글 OR-Tools 엔진이 제약 조건을 분석하여 근무표를 생성 중입니다..."):
            
            # 해당 월의 일수 및 요일 계산
            num_days = calendar.monthrange(target_year, target_month)[1]
            days = list(range(1, num_days + 1))
            
            # OR-Tools 모델 초기화
            model = cp_model.CpModel()
            
            # 근무 코드: 0:주간, 1:야간3, 2:대휴, 3:연차, 4:휴가, 5:공가, 6:비번, 7:휴무
            shift_map = {'주': 0, '야3': 1, '대': 2, '연차': 3, '휴가': 4, '공가': 5, '비': 6, '휴무': 7}
            code_to_name = {v: k for k, v in shift_map.items()}
            
            employees = [e for e in st.session_state.employees if e['성명']]
            num_emp = len(employees)
            
            # 변수 선언: x[(e, d, s)] = 0 또는 1
            x = {}
            for e in range(num_emp):
                for d in days:
                    for s in range(8):
                        x[(e, d, s)] = model.NewBoolVar(f'x_{e}_{d}_{s}')
                        
            # 1. 하루 1인 1근무
            for e in range(num_emp):
                for d in days:
                    model.AddExactlyOne(x[(e, d, s)] for s in range(8))
                    
            # 2. 팀원별 신청 요청사항 반영
            for req in st.session_state.requests:
                e_idx = next((i for i, e in enumerate(employees) if e['성명'] == req['성명']), None)
                if e_idx is not None and req['일자'] <= num_days:
                    s_code = shift_map.get(req['요청근무'], 2)
                    model.Add(x[(e_idx, req['일자'], s_code)] == 1)
                    
            # 3. 야간 근무자 조건 (2인 1조 매무 및 선호/비선호)
            night_indices = [i for i, e in enumerate(employees) if e['구분'] == '야간']
            for d in days:
                if len(night_indices) >= 2:
                    # 하루 야간3(1) 근무자는 정확히 2명
                    model.Add(sum(x[(e, d, 1)] for e in night_indices) == 2)
                    
                    # 비선호 짝꿍 제약 (동시 야간3 근무 금지)
                    for e in night_indices:
                        dis_name = employees[e].get('비선호짝꿍', '')
                        if dis_name:
                            dis_idx = next((i for i, emp in enumerate(employees) if emp['성명'] == dis_name), None)
                            if dis_idx is not None:
                                model.AddBoolOr([x[(e, d, 1)].Not(), x[(dis_idx, d, 1)].Not()])

            # 4. 주5일 및 주간 스케줄러 기본 조건
            for e, emp in enumerate(employees):
                for d in days:
                    dt = datetime(target_year, target_month, d)
                    weekday_str = ["월", "화", "수", "목", "금", "토", "일"][dt.weekday()]
                    
                    if emp['구분'] == '주5일':
                        if weekday_str in ["토", "일"]:
                            # 주말은 대휴(2) 우선
                            model.Add(x[(e, d, 2)] == 1)
                    elif emp['구분'] == '주간스케줄':
                        pref_off = str(emp.get('기본휴무일_선호짝꿍', ''))
                        if weekday_str in pref_off:
                            model.Add(x[(e, d, 2)] == 1) # 설정 휴무일은 대휴(2)

            # Solver 실행
            solver = cp_model.CpSolver()
            solver.parameters.max_time_in_seconds = 10.0
            status = solver.Solve(model)
            
            if status in (cp_model.OPTIMAL, cp_model.FEASIBLE):
                st.success("🎉 최적의 근무표가 성공적으로 생성되었습니다!")
                
                # 결과를 DataFrame으로 변환
                schedule_data = []
                for e, emp in enumerate(employees):
                    row = {'성명': emp['성명'], '직급': emp['직급'], '구분': emp['구분']}
                    for d in days:
                        assigned = [s for s in range(8) if solver.Value(x[(e, d, s)]) == 1][0]
                        row[f"{d}일"] = code_to_name[assigned]
                    schedule_data.append(row)
                    
                df_res = pd.DataFrame(schedule_data)
                st.dataframe(df_res, use_container_width=True)
                
                # 엑셀 다운로드 파일 생성
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
                st.error("❌ 설정된 제약 조건이 너무 까다로워 근무표를 생성할 수 없습니다. 요청사항이나 짝꿍 제약을 완화해 보세요.")
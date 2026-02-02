import streamlit as st
import pandas as pd
import openpyxl
from datetime import datetime, time
import io

# ==========================================
# 1. 파일 이름 설정
# ==========================================
FILE_NAME = "2026_2월 급여대장.xlsx"
SHEET_NAME = "근무표(입력)"

st.set_page_config(page_title="OUR 급여관리", layout="wide")
st.title("🍞 OUR 베이커리 급여 입력 (2월)")

# ==========================================
# 2. 엑셀 파일 읽기
# ==========================================
if "wb" not in st.session_state:
    try:
        st.session_state.wb = openpyxl.load_workbook(FILE_NAME, data_only=False)
    except:
        st.error(f"폴더에 '{FILE_NAME}' 파일이 없습니다. 파일 이름을 확인해주세요.")
        st.stop()

wb = st.session_state.wb
try:
    sheet = wb[SHEET_NAME]
except:
    st.error(f"엑셀 안에 '{SHEET_NAME}' 시트가 없습니다.")
    st.stop()

# ==========================================
# 3. 직원 목록 가져오기 (수정됨: 집계표에서 읽기)
# ==========================================
def get_employees(sheet):
    names = set()
    # Y열(25번째 열)에 있는 집계표 이름 목록을 읽어옵니다.
    # 3행부터 50행 정도까지만 확인
    for row in sheet.iter_rows(min_row=3, max_row=50, min_col=25, max_col=25, values_only=True):
        if row[0]: # 이름이 있으면 추가
            names.add(row[0])
    return sorted(list(names))

employee_list = get_employees(sheet)
if not employee_list: 
    # 만약 집계표도 비어있다면 기본값
    employee_list = []

# ==========================================
# 4. 날짜 목록 가져오기
# ==========================================
date_row_map = {}
date_options = []
for i, row in enumerate(sheet.iter_rows(min_row=3, min_col=1, max_col=2), start=3):
    d, day = row[0].value, row[1].value
    if d:
        d_str = d.strftime("%Y-%m-%d") if isinstance(d, datetime) else str(d)
        label = f"{d_str} ({day})"
        date_row_map[label] = i
        date_options.append(label)

# ==========================================
# 5. 입력 화면 (타임 1~5)
# ==========================================
selected_label = st.selectbox("📅 날짜 선택", date_options)
target_row = date_row_map[selected_label]

st.markdown("---")

slot_configs = [
    {"name": "타임 1 (오전)", "col": 3},
    {"name": "타임 2 (미들1)", "col": 7},
    {"name": "타임 3 (미들2)", "col": 11},
    {"name": "타임 4 (오후)", "col": 15},
    {"name": "타임 5 (마감)", "col": 19},
]

with st.form("input_form"):
    st.write(f"**📝 {selected_label} 근무자 입력**")
    st.info("💡 목록에 없는 직원은 '(직접 입력)'을 선택하세요.")
    
    cols = st.columns(2) 
    updates = {}
    
    for idx, slot in enumerate(slot_configs):
        col_ui = cols[idx % 2]
        base = slot["col"]
        
        # 엑셀 값 읽기
        curr_name = sheet.cell(row=target_row, column=base).value
        curr_s = sheet.cell(row=target_row, column=base+1).value
        curr_e = sheet.cell(row=target_row, column=base+2).value
        
        def to_time(v):
            if isinstance(v, datetime): return v.time()
            if isinstance(v, time): return v
            return None

        with col_ui:
            is_expanded = (curr_name is not None)
            with st.expander(f"{slot['name']}", expanded=is_expanded):
                
                # --- 이름 선택 로직 (직접 입력 추가) ---
                # 기존 값이 리스트에 없으면 (신규 직원이면) 직접 입력으로 간주
                list_options = ["(선택없음)", "(직접 입력)"] + employee_list
                
                default_idx = 0
                if curr_name:
                    if curr_name in employee_list:
                        default_idx = list_options.index(curr_name)
                    else:
                        default_idx = 1 # (직접 입력)

                selected_option = st.selectbox(
                    "이름", 
                    list_options, 
                    index=default_idx, 
                    key=f"sel_{base}"
                )
                
                final_name = None
                if selected_option == "(직접 입력)":
                    # 직접 입력 창 보여주기 (기존 값이 있으면 채워줌)
                    input_val = curr_name if (curr_name and curr_name not in employee_list) else ""
                    final_name = st.text_input("이름 직접 입력", value=input_val, key=f"txt_{base}")
                elif selected_option != "(선택없음)":
                    final_name = selected_option
                
                # --- 시간 선택 ---
                c1, c2 = st.columns(2)
                new_s = c1.time_input("출근", value=to_time(curr_s), key=f"s_{base}")
                new_e = c2.time_input("퇴근", value=to_time(curr_e), key=f"e_{base}")
                
                updates[base] = {"n": final_name, "s": new_s, "e": new_e}
    
    st.markdown("###")
    applied = st.form_submit_button("✅ 입력 내용 반영하기", use_container_width=True)

if applied:
    for base, data in updates.items():
        if data["n"]: # 이름이 있으면 저장
            sheet.cell(row=target_row, column=base).value = data["n"]
            sheet.cell(row=target_row, column=base+1).value = data["s"]
            sheet.cell(row=target_row, column=base+2).value = data["e"]
        else: # 없으면 지우기
            sheet.cell(row=target_row, column=base).value = None
            sheet.cell(row=target_row, column=base+1).value = None
            sheet.cell(row=target_row, column=base+2).value = None
    
    st.success(f"{selected_label} 저장 완료! 엑셀 파일을 다운로드하세요.")

# ==========================================
# 6. 다운로드 버튼
# ==========================================
st.markdown("---")
output = io.BytesIO()
wb.save(output)
processed_data = output.getvalue()

st.download_button(
    label="📥 엑셀 파일 다운로드 (카톡 전송용)",
    data=processed_data,
    file_name=f"2월급여대장_OUR_{datetime.now().strftime('%m%d')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True
)

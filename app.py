import streamlit as st
import pandas as pd
import openpyxl
from datetime import datetime, time
import io

# ==========================================
# 1. 파일 이름 설정 (수정됨)
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
# 3. 직원 및 날짜 목록
# ==========================================
def get_employees(sheet):
    names = set()
    for row in sheet.iter_rows(min_row=3, values_only=True):
        # 타임 1~5 (C, G, K, O, S열)
        for col_idx in [2, 6, 10, 14, 18]:
            if col_idx < len(row) and row[col_idx]:
                names.add(row[col_idx])
    return sorted(list(names))

employee_list = get_employees(sheet)
if not employee_list: employee_list = ["직원 명단 없음"]

# 날짜 매핑
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
# 4. 입력 화면 (타임 1~5)
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
    cols = st.columns(2) 
    updates = {}
    
    for idx, slot in enumerate(slot_configs):
        col_ui = cols[idx % 2]
        base = slot["col"]
        
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
                n_idx = employee_list.index(curr_name) + 1 if curr_name in employee_list else 0
                new_n = st.selectbox("이름", ["(선택없음)"] + employee_list, index=n_idx, key=f"n_{base}")
                
                c1, c2 = st.columns(2)
                new_s = c1.time_input("출근", value=to_time(curr_s), key=f"s_{base}")
                new_e = c2.time_input("퇴근", value=to_time(curr_e), key=f"e_{base}")
                updates[base] = {"n": new_n, "s": new_s, "e": new_e}
    
    st.markdown("###")
    applied = st.form_submit_button("✅ 입력 내용 반영하기", use_container_width=True)

if applied:
    for base, data in updates.items():
        if data["n"] != "(선택없음)":
            sheet.cell(row=target_row, column=base).value = data["n"]
            sheet.cell(row=target_row, column=base+1).value = data["s"]
            sheet.cell(row=target_row, column=base+2).value = data["e"]
        else:
            sheet.cell(row=target_row, column=base).value = None
            sheet.cell(row=target_row, column=base+1).value = None
            sheet.cell(row=target_row, column=base+2).value = None
    st.success(f"{selected_label} 입력 완료! 아래 버튼으로 파일을 받으세요.")

st.markdown("---")

# 파일 다운로드
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
import streamlit as st
import pandas as pd
import openpyxl
from datetime import datetime, time
import io

# ==========================================
# 1. 파일 설정
# ==========================================
FILE_NAME = "2026_2월 급여대장.xlsx"
SHEET_NAME = "근무표(입력)"

st.set_page_config(page_title="OUR 급여관리", layout="wide")
st.title("🍞 OUR 베이커리 급여 입력 (2월)")

# ==========================================
# 2. 엑셀 파일 로드
# ==========================================
if "wb" not in st.session_state:
    try:
        st.session_state.wb = openpyxl.load_workbook(FILE_NAME, data_only=False)
    except:
        st.error(f"폴더에 '{FILE_NAME}' 파일이 없습니다.")
        st.stop()

wb = st.session_state.wb
try:
    sheet = wb[SHEET_NAME]
except:
    st.error(f"엑셀에 '{SHEET_NAME}' 시트가 없습니다.")
    st.stop()

# 읽기 전용 데이터 (날짜 수식 해결용)
@st.cache_data
def load_display_data():
    try:
        return pd.read_excel(FILE_NAME, sheet_name=SHEET_NAME, header=None)
    except:
        return None

df = load_display_data()
if df is None:
    st.error("엑셀 파일을 읽을 수 없습니다.")
    st.stop()

# ==========================================
# 3. 직원 목록 & 날짜 목록
# ==========================================

# 직원 목록 (Y열 집계표에서 읽기)
def get_employees(sheet_obj):
    names = set()
    # Y열은 엑셀에서 25번째 열입니다.
    # 3행부터 50행까지 탐색
    for r in range(3, 51):
        val = sheet_obj.cell(row=r, column=25).value # Y열
        if val and str(val).strip() != "":
            names.add(str(val))
    return sorted(list(names))

# 현재 등록된 직원 리스트 불러오기
employee_list = get_employees(sheet)

# 날짜 목록 (Pandas로 값 읽기)
date_row_map = {}
date_options = []

if len(df) > 2:
    for idx, row in df.iloc[2:].iterrows():
        val_date = row[0] # A열
        val_day = row[1]  # B열
        
        if pd.notna(val_date):
            d_str = ""
            if isinstance(val_date, datetime):
                d_str = val_date.strftime("%Y-%m-%d")
            else:
                d_str = str(val_date).split(" ")[0]

            label = f"{d_str} ({val_day})"
            excel_row = idx + 1
            date_row_map[label] = excel_row
            date_options.append(label)

if not date_options:
    st.error("날짜를 찾을 수 없습니다.")
    st.stop()

# ==========================================
# 4. 입력 화면
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
    st.caption("※ '(직접 입력)'으로 이름을 넣으면 집계표에 자동 등록됩니다.")
    
    cols = st.columns(2) 
    updates = {}
    
    # 신규 등록할 직원 이름을 담을 집합
    new_names_to_add = set()

    for idx, slot in enumerate(slot_configs):
        col_ui = cols[idx % 2]
        base = slot["col"]
        
        # 기존 값 (openpyxl)
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
                
                # 이름 선택 UI
                list_opts = ["(선택없음)", "(직접 입력)"] + employee_list
                def_idx = 0
                
                # 기존에 입력된 이름이 리스트에 있으면 선택, 없으면 직접입력 상태로
                if curr_name:
                    str_name = str(curr_name)
                    if str_name in employee_list:
                        def_idx = list_opts.index(str_name)
                    else:
                        def_idx = 1 # (직접 입력)

                sel = st.selectbox("이름", list_opts, index=def_idx, key=f"sel_{base}")
                
                final_n = None
                if sel == "(직접 입력)":
                    # 기존 값이 리스트에 없는 값이었다면 그 값을 보여줌
                    val_show = curr_name if (curr_name and str(curr_name) not in employee_list) else ""
                    final_n = st.text_input("이름 입력", value=val_show, key=f"txt_{base}")
                    
                    # 입력된 이름이 유효하고, 기존 리스트에 없다면 추가 대기열에 등록
                    if final_n and final_n.strip() != "" and final_n not in employee_list:
                        new_names_to_add.add(final_n)
                elif sel != "(선택없음)":
                    final_n = sel
                
                # 시간 입력
                c1, c2 = st.columns(2)
                new_s = c1.time_input("출근", value=to_time(curr_s), key=f"s_{base}")
                new_e = c2.time_input("퇴근", value=to_time(curr_e), key=f"e_{base}")
                
                updates[base] = {"n": final_n, "s": new_s, "e": new_e}

    st.markdown("###")
    applied = st.form_submit_button("✅ 저장 및 반영하기", use_container_width=True)

# ==========================================
# 5. 저장 로직 (핵심 수정 부분)
# ==========================================
if applied:
    # 1. 신규 직원 집계표(Y열)에 자동 등록
    if new_names_to_add:
        for new_name in new_names_to_add:
            # Y열(25번 열)에서 3행부터 빈칸 찾기
            for r in range(3, 100):
                cell = sheet.cell(row=r, column=25)
                if cell.value is None or str(cell.value).strip() == "":
                    cell.value = new_name
                    # 번호(X열)도 넣어주면 좋음 (선택사항)
                    # sheet.cell(row=r, column=24).value = r - 2 
                    st.toast(f"🎉 신규 직원 '{new_name}'님이 명단에 등록되었습니다!")
                    break
        
        # 리스트 갱신을 위해 메모리 상의 리스트 업데이트
        employee_list.extend(list(new_names_to_add))
        employee_list.sort()

    # 2. 근무표 데이터 저장
    for base, data in updates.items():
        if data["n"]:
            sheet.cell(row=target_row, column=base).value = data["n"]
            sheet.cell(row=target_row, column=base+1).value = data["s"]
            sheet.cell(row=target_row, column=base+2).value = data["e"]
        else:
            sheet.cell(row=target_row, column=base).value = None
            sheet.cell(row=target_row, column=base+1).value = None
            sheet.cell(row=target_row, column=base+2).value = None

    st.success(f"{selected_label} 저장 완료! (신규 직원 자동 등록 포함)")

# ==========================================
# 6. 다운로드
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

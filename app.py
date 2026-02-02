import streamlit as st
import pandas as pd
import openpyxl
from datetime import datetime, time
import io

# ==========================================
# 1. 기본 설정
# ==========================================
FILE_NAME = "2026_2월 급여대장.xlsx"
SHEET_NAME = "근무표(입력)"

st.set_page_config(page_title="OUR 급여관리", layout="wide")
st.title("🍞 OUR 베이커리 급여 입력 (2월)")

# ==========================================
# 2. 엑셀 파일 로드 (수식 보호 + 날짜 읽기)
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

# 화면 표시용 데이터 읽기
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
# 3. 직원 목록 & 날짜 목록 준비
# ==========================================

# 직원 목록 (집계표 Y열에서 읽기)
def get_employees(sheet_obj):
    names = set()
    # Y열(25열) 3행~100행 탐색
    for r in range(3, 101):
        val = sheet_obj.cell(row=r, column=25).value 
        if val and str(val).strip() != "":
            names.add(str(val))
    return sorted(list(names))

employee_list = get_employees(sheet)

# 날짜 목록 (Pandas로 값 읽기)
date_row_map = {}
date_options = []

if len(df) > 2:
    for idx, row in df.iloc[2:].iterrows():
        val_date = row[0]
        val_day = row[1]
        
        if pd.notna(val_date):
            d_str = val_date.strftime("%Y-%m-%d") if isinstance(val_date, datetime) else str(val_date).split(" ")[0]
            label = f"{d_str} ({val_day})"
            date_row_map[label] = idx + 1 # 엑셀 행 번호
            date_options.append(label)

if not date_options:
    st.error("날짜를 찾을 수 없습니다.")
    st.stop()

# ==========================================
# 4. 입력 화면 (사라짐 방지 기능 적용)
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
    st.caption("※ 빈칸일 땐 바로 이름을 입력하시면 됩니다.")
    
    cols = st.columns(2) 
    updates = {}
    new_names_to_add = set()

    for idx, slot in enumerate(slot_configs):
        col_ui = cols[idx % 2]
        base = slot["col"]
        
        # 엑셀 현재 값 읽기
        curr_name = sheet.cell(row=target_row, column=base).value
        curr_s = sheet.cell(row=target_row, column=base+1).value
        curr_e = sheet.cell(row=target_row, column=base+2).value
        
        # 시간 변환
        def to_time(v):
            if isinstance(v, datetime): return v.time()
            if isinstance(v, time): return v
            return None

        # --- UI 그리기 ---
        with col_ui:
            # 값이 있거나, 사용자가 뭔가 입력 중이라면 펼쳐두기
            is_expanded = (curr_name is not None)
            with st.expander(f"{slot['name']}", expanded=is_expanded):
                
                # 1. 이름 선택 리스트 구성 (선택없음 제거!)
                # "직접 입력"을 맨 앞에 둬서 기본값으로 만듦
                list_opts = ["(직접 입력)"] + employee_list
                
                # 기본 선택값 결정
                def_idx = 0 # 기본은 (직접 입력)
                if curr_name and str(curr_name) in employee_list:
                    def_idx = list_opts.index(str(curr_name))
                
                # ★중요★ key에 날짜(selected_label)를 붙여서 날짜별로 입력창을 따로 관리함 (증발 방지)
                unique_key_sel = f"sel_{base}_{selected_label}"
                unique_key_txt = f"txt_{base}_{selected_label}"
                unique_key_s = f"s_{base}_{selected_label}"
                unique_key_e = f"e_{base}_{selected_label}"

                sel_val = st.selectbox("이름 선택", list_opts, index=def_idx, key=unique_key_sel)
                
                final_n = None
                
                # 2. 직접 입력 로직
                if sel_val == "(직접 입력)":
                    # 기존 엑셀에 값이 있는데 리스트에 없는 이름(신규)이라면, 그 값을 입력창에 채워줌
                    val_to_show = ""
                    if curr_name and str(curr_name) not in employee_list:
                        val_to_show = str(curr_name)
                    
                    # 텍스트 입력창
                    input_name = st.text_input("이름 입력", value=val_to_show, key=unique_key_txt)
                    
                    if input_name.strip():
                        final_n = input_name.strip()
                        # 신규 이름이면 등록 대기
                        if final_n not in employee_list:
                            new_names_to_add.add(final_n)
                    else:
                        final_n = None # 비워두면 삭제
                else:
                    final_n = sel_val

                # 3. 시간 입력
                c1, c2 = st.columns(2)
                new_s = c1.time_input("출근", value=to_time(curr_s), key=unique_key_s)
                new_e = c2.time_input("퇴근", value=to_time(curr_e), key=unique_key_e)
                
                updates[base] = {"n": final_n, "s": new_s, "e": new_e}

    st.markdown("###")
    applied = st.form_submit_button("✅ 저장 및 반영하기", use_container_width=True)

# ==========================================
# 5. 저장 로직 (신규직원 자동등록 포함)
# ==========================================
if applied:
    # 1. 신규 직원 집계표 자동 등록
    if new_names_to_add:
        for new_name in new_names_to_add:
            # Y열(25번 열) 빈칸 찾기
            for r in range(3, 101):
                cell = sheet.cell(row=r, column=25)
                if cell.value is None or str(cell.value).strip() == "":
                    cell.value = new_name
                    # 번호(X열) 매기기
                    sheet.cell(row=r, column=24).value = r - 2
                    break
        st.toast(f"🎉 신규 직원 {len(new_names_to_add)}명이 명단에 추가되었습니다!")

    # 2. 근무표 저장
    for base, data in updates.items():
        if data["n"]:
            sheet.cell(row=target_row, column=base).value = data["n"]
            sheet.cell(row=target_row, column=base+1).value = data["s"]
            sheet.cell(row=target_row, column=base+2).value = data["e"]
        else:
            # 이름 없으면 지우기
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

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
# 2. 엑셀 파일 로드 (두 가지 방식 혼용)
# ==========================================

# (1) 저장용: 수식을 깨뜨리지 않기 위해 openpyxl로 로드 (세션에 저장)
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

# (2) 읽기용: 날짜가 수식(IF...)으로 되어 있어도 '값'을 읽기 위해 pandas 사용
# @st.cache_data는 매번 파일을 읽지 않고 속도를 높여줍니다.
@st.cache_data
def load_display_data():
    try:
        # header=None으로 해서 있는 그대로 다 읽어옴
        return pd.read_excel(FILE_NAME, sheet_name=SHEET_NAME, header=None)
    except:
        return None

df = load_display_data()

if df is None:
    st.error("엑셀 파일을 읽을 수 없습니다.")
    st.stop()

# ==========================================
# 3. 직원 목록 & 날짜 목록 만들기
# ==========================================

# 직원 목록 (집계표 쪽에서 읽기 - 값만 읽어오므로 수식 문제 없음)
def get_employees(df_data):
    names = set()
    # pandas 데이터프레임에서 직접 읽기 (24번째 컬럼 = Y열)
    # 데이터는 2행(엑셀3행)부터 시작
    # Y열 데이터가 있는지 확인하고 가져옴
    if len(df_data.columns) > 24:
        for val in df_data.iloc[2:, 24]: 
            if pd.notna(val) and str(val).strip() != "":
                names.add(str(val))
    return sorted(list(names))

employee_list = get_employees(df)
if not employee_list: employee_list = []

# 날짜 목록 (여기가 핵심! Pandas로 값을 읽어서 수식 문제 해결)
date_row_map = {}
date_options = []

# df.iloc[2:] -> 엑셀 3행부터 끝까지 반복
for idx, row in df.iloc[2:].iterrows():
    # row[0]: A열(날짜), row[1]: B열(요일)
    val_date = row[0]
    val_day = row[1]
    
    if pd.notna(val_date): # 날짜가 비어있지 않으면
        # 날짜 포맷 예쁘게 다듬기
        d_str = ""
        if isinstance(val_date, datetime):
            d_str = val_date.strftime("%Y-%m-%d")
        else:
            # 혹시 날짜가 글자로 되어 있을 경우 대비
            d_str = str(val_date).split(" ")[0] 

        # 라벨 생성: "2026-02-02 (월)"
        label = f"{d_str} ({val_day})"
        
        # 엑셀 행 번호 계산
        # Pandas index는 0부터 시작, 엑셀은 1부터 시작.
        # df는 header 없이 읽었으므로 idx 0 = 엑셀 1행.
        # 따라서 엑셀 행 번호 = idx + 1
        excel_row = idx + 1
        
        date_row_map[label] = excel_row
        date_options.append(label)

# ==========================================
# 4. 입력 화면 구성
# ==========================================
if not date_options:
    st.error("날짜를 찾을 수 없습니다. 엑셀 A열에 날짜가 있는지 확인해주세요.")
    st.stop()

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
    st.caption("※ 목록에 없는 직원은 '(직접 입력)'을 선택하세요.")
    
    cols = st.columns(2) 
    updates = {}
    
    for idx, slot in enumerate(slot_configs):
        col_ui = cols[idx % 2]
        base = slot["col"]
        
        # 기존 값 읽기 (여기는 openpyxl 사용 - 입력된 값은 보통 수식이 아니므로 괜찮음)
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
                
                # 이름 선택 로직
                list_options = ["(선택없음)", "(직접 입력)"] + employee_list
                default_idx = 0
                
                if curr_name:
                    if str(curr_name) in employee_list:
                        default_idx = list_options.index(str(curr_name))
                    else:
                        default_idx = 1 # 목록에 없으면 직접 입력으로

                sel_name = st.selectbox("이름", list_options, index=default_idx, key=f"sel_{base}")
                
                final_name = None
                if sel_name == "(직접 입력)":
                    # 기존 값이 목록에 없는 값이라면 입력창에 보여주기
                    val_to_show = curr_name if (curr_name and str(curr_name) not in employee_list) else ""
                    final_name = st.text_input("직원 이름 입력", value=val_to_show, key=f"txt_{base}")
                elif sel_name != "(선택없음)":
                    final_name = sel_name
                
                # 시간 선택
                c1, c2 = st.columns(2)
                new_s = c1.time_input("출근", value=to_time(curr_s), key=f"s_{base}")
                new_e = c2.time_input("퇴근", value=to_time(curr_e), key=f"e_{base}")
                
                updates[base] = {"n": final_name, "s": new_s, "e": new_e}
    
    st.markdown("###")
    applied = st.form_submit_button("✅ 입력 내용 반영하기", use_container_width=True)

# ==========================================
# 5. 반영 및 다운로드
# ==========================================
if applied:
    for base, data in updates.items():
        if data["n"]:
            sheet.cell(row=target_row, column=base).value = data["n"]
            sheet.cell(row=target_row, column=base+1).value = data["s"]
            sheet.cell(row=target_row, column=base+2).value = data["e"]
        else:
            sheet.cell(row=target_row, column=base).value = None
            sheet.cell(row=target_row, column=base+1).value = None
            sheet.cell(row=target_row, column=base+2).value = None
            
    st.success(f"저장 완료! ({selected_label})")

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

import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import time

# ==========================================
# 1. 기본 설정
# ==========================================
st.set_page_config(page_title="OUR 급여관리", layout="wide")
st.title("🍞 OUR 베이커리 급여 입력 (5타임)")

# 👇 [필수] 아까 복사한 '새로운 ID'를 따옴표 안에 붙여넣으세요!
SHEET_ID = "10JJhbH6QuW4esqq1wEvo-0s4-fhNS_7MQQlyk-sVdLg"

# 시간 선택지 (06:00 ~ 02:00, 30분 단위)
TIME_OPTIONS = [""] 
for h in range(6, 26):
    hour = h if h < 24 else h - 24
    TIME_OPTIONS.append(f"{hour:02d}:00")
    TIME_OPTIONS.append(f"{hour:02d}:30")

# 구글 시트 연결
def connect_to_gsheet():
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    try:
        key_dict = dict(st.secrets["gcp_service_account"])
        creds = Credentials.from_service_account_info(key_dict, scopes=scopes)
        client = gspread.authorize(creds)
        return client.open_by_key(SHEET_ID)
    except Exception as e:
        st.error(f"⚠️ 연결 오류: {e}")
        st.stop()

sh = connect_to_gsheet()
worksheet = sh.worksheet("근무표(입력)")

# ==========================================
# 2. 데이터 읽어오기
# ==========================================
all_values = worksheet.get_all_values()

# 날짜 목록 만들기
date_row_map = {}
date_options = []
if len(all_values) > 2:
    for idx, row in enumerate(all_values[2:], start=2):
        val_date = row[0]
        val_day = row[1]
        if val_date and str(val_date).strip():
            d_str = str(val_date).split(" ")[0]
            label = f"{d_str} ({val_day})"
            date_row_map[label] = idx + 1
            date_options.append(label)

if not date_options:
    st.error("날짜를 찾을 수 없습니다. 시트 A열에 날짜가 있는지 확인하세요.")
    st.stop()

selected_label = st.selectbox("📅 날짜 선택", date_options)
target_row = date_row_map[selected_label]

# 데이터 읽기 (칸이 늘어났으니 넉넉하게 읽음)
current_row_values = worksheet.row_values(target_row)
while len(current_row_values) < 30: 
    current_row_values.append("")

# ==========================================
# 3. 근무 현황판 (5타임)
# ==========================================
st.markdown("---")
st.info(f"👇 **{selected_label}** 근무 현황")

# 5개 타임 설정 (S,T,U열이 타임 5입니다)
slot_configs = [
    {"name": "타임 1 (오전)", "col": 3},   # C열
    {"name": "타임 2 (미들1)", "col": 7},  # G열
    {"name": "타임 3 (미들2)", "col": 11}, # K열
    {"name": "타임 4 (오후)", "col": 15},  # O열
    {"name": "타임 5 (마감)", "col": 19},  # S열
]

status_data = []
for slot in slot_configs:
    idx = slot["col"] - 1
    if current_row_values[idx].strip():
        status_data.append({
            "타임": slot["name"],
            "이름": current_row_values[idx],
            "출근": current_row_values[idx+1],
            "퇴근": current_row_values[idx+2]
        })

if status_data:
    st.table(pd.DataFrame(status_data))
else:
    st.write("등록된 근무자가 없습니다.")

# ==========================================
# 4. 입력 폼
# ==========================================
# [중요] 직원 명단을 가져오는 곳입니다.
# 5타임(S,T,U) 뒤에 명단이 있어야 하므로, 넉넉하게 'Y열(25)'로 잡았습니다.
# 엑셀에서 명단이 Y열이 아니라면, 아래 숫자 25를 해당 열 번호로 바꾸세요.
# (V=22, W=23, X=24, Y=25, Z=26)
name_list = [n for n in worksheet.col_values(25)[2:] if n.strip()]
employee_options = ["(직접 입력)"] + sorted(list(set(name_list)))

with st.form("input_form"):
    st.write(f"**✏️ {selected_label} 근무 입력/수정**")
    st.caption("※ 시간을 선택하고 저장 버튼을 누르세요.")
    
    cols = st.columns(2)
    updates = {}
    new_names = set()

    for i, slot in enumerate(slot_configs):
        with cols[i % 2]:
            idx = slot["col"] - 1
            curr_n = current_row_values[idx]
            curr_s = current_row_values[idx+1]
            curr_e = current_row_values[idx+2]
            
            with st.expander(f"{slot['name']}", expanded=(curr_n != "")):
                # 이름
                def_idx = employee_options.index(curr_n) if curr_n in employee_options else 0
                sel_n = st.selectbox("이름", employee_options, index=def_idx, key=f"n_{slot['col']}")
                
                final_n = ""
                if sel_n == "(직접 입력)":
                    val = curr_n if curr_n not in employee_options else ""
                    txt_n = st.text_input("이름 직접 입력", value=val, key=f"txt_{slot['col']}")
                    final_n = txt_n.strip()
                    if final_n and final_n not in employee_options:
                        new_names.add(final_n)
                else:
                    final_n = sel_n

                # 시간
                c1, c2 = st.columns(2)
                s_idx = TIME_OPTIONS.index(curr_s) if curr_s in TIME_OPTIONS else 0
                e_idx = TIME_OPTIONS.index(curr_e) if curr_e in TIME_OPTIONS else 0
                
                sel_s = c1.selectbox("출근", TIME_OPTIONS, index=s_idx, key=f"s_{slot['col']}")
                sel_e = c2.selectbox("퇴근", TIME_OPTIONS, index=e_idx, key=f"e_{slot['col']}")
                
                updates[slot["col"]] = {"n": final_n, "s": sel_s, "e": sel_e}

    st.markdown("###")
    applied = st.form_submit_button("✅ 저장하기", use_container_width=True)

if applied:
    msg = st.empty()
    msg.info("저장 중...")
    
    # 신규 직원 명단 추가 (Y열 아래에)
    if new_names:
        col_vals = worksheet.col_values(25)
        next_r = 3
        for v in col_vals[2:]:
            if v == "": break
            next_r += 1
        
        for nm in new_names:
            worksheet.update_cell(next_r, 25, nm)
            next_r += 1
            
    # 근무 내용 저장
    cells = []
    for col, data in updates.items():
        cells.append(gspread.Cell(target_row, col, data["n"]))
        cells.append(gspread.Cell(target_row, col+1, data["s"]))
        cells.append(gspread.Cell(target_row, col+2, data["e"]))
        
    worksheet.update_cells(cells)
    msg.success("완료! 내용이 반영되었습니다.")
    time.sleep(1)
    st.rerun()

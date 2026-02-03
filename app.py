import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials # [변경] 최신 부품 사용
import time

# ==========================================
# 1. 구글 시트 연결 설정 (최신 V4 방식)
# ==========================================
st.set_page_config(page_title="OUR 급여관리(실시간)", layout="wide")
st.title("🍞 OUR 베이커리 급여 입력 (구글연동)")

def connect_to_gsheet():
    # [변경] 구글 시트 주소를 최신 버전(V4)으로 변경
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    try:
        # Secrets에서 키 정보 가져오기
        key_dict = dict(st.secrets["gcp_service_account"])
        
        # [변경] 최신 방식(google-auth)으로 로그인
        creds = Credentials.from_service_account_info(key_dict, scopes=scopes)
        client = gspread.authorize(creds)
        
        # 파일 열기
        sh = client.open('OUR_급여장부_2월')
        return sh
    except Exception as e:
        st.error(f"⚠️ 연결 오류! 설정을 확인해주세요.\n에러내용: {e}")
        st.stop()

sh = connect_to_gsheet()
try:
    worksheet = sh.worksheet("근무표(입력)")
except:
    st.error("시트 이름을 찾을 수 없습니다. '근무표(입력)' 탭이 있는지 확인하세요.")
    st.stop()

# ==========================================
# 2. 데이터 처리 및 화면 구성 (이하 동일)
# ==========================================
all_values = worksheet.get_all_values()
df = pd.DataFrame(all_values)

date_row_map = {}
date_options = []

if len(df) > 2:
    for idx, row in df.iloc[2:].iterrows():
        val_date = row[0]
        val_day = row[1]
        if val_date and str(val_date).strip() != "":
            d_str = str(val_date).split(" ")[0]
            label = f"{d_str} ({val_day})"
            date_row_map[label] = idx + 1
            date_options.append(label)

if not date_options:
    st.error("날짜 데이터가 없습니다.")
    st.stop()

selected_label = st.selectbox("📅 날짜 선택", date_options)
target_row = date_row_map[selected_label]

st.markdown("---")
st.info(f"👇 **{selected_label}** 현재 구글 장부 내용")

slot_configs = [
    {"name": "타임 1 (오전)", "col": 3},
    {"name": "타임 2 (미들1)", "col": 7},
    {"name": "타임 3 (미들2)", "col": 11},
    {"name": "타임 4 (오후)", "col": 15},
    {"name": "타임 5 (마감)", "col": 19},
]

current_row_values = worksheet.row_values(target_row)
while len(current_row_values) < 25:
    current_row_values.append("")

current_status = []
for slot in slot_configs:
    base_idx = slot["col"] - 1
    if current_row_values[base_idx] and current_row_values[base_idx].strip():
        current_status.append({
            "타임": slot["name"],
            "이름": current_row_values[base_idx],
            "출근": current_row_values[base_idx+1],
            "퇴근": current_row_values[base_idx+2]
        })

if current_status:
    st.table(pd.DataFrame(current_status))
else:
    st.write("🚫 근무자가 없습니다.")

# ==========================================
# 3. 입력 폼
# ==========================================
name_col_values = worksheet.col_values(25)[2:] 
employee_list = sorted(list(set([n for n in name_col_values if n.strip()])))

with st.form("input_form"):
    st.write(f"**✏️ {selected_label} 근무 수정**")
    cols = st.columns(2)
    updates = {}
    new_names_to_add = set()
    
    for idx, slot in enumerate(slot_configs):
        col_ui = cols[idx % 2]
        base_idx = slot["col"] - 1
        curr_n = current_row_values[base_idx]
        curr_s = current_row_values[base_idx+1]
        curr_e = current_row_values[base_idx+2]
        
        with col_ui:
            is_expanded = (curr_n != "")
            with st.expander(f"{slot['name']}", expanded=is_expanded):
                list_opts = ["(직접 입력)"] + employee_list
                def_idx = list_opts.index(curr_n) if curr_n in employee_list else 0
                
                k_base = f"{base_idx}_{selected_label}"
                sel_val = st.selectbox("이름", list_opts, index=def_idx, key=f"sel_{k_base}")
                
                final_n = ""
                if sel_val == "(직접 입력)":
                    val_show = curr_n if curr_n not in employee_list else ""
                    input_n = st.text_input("이름 입력", value=val_show, key=f"txt_{k_base}")
                    final_n = input_n.strip()
                    if final_n and final_n not in employee_list:
                        new_names_to_add.add(final_n)
                else:
                    final_n = sel_val
                
                c1, c2 = st.columns(2)
                new_s = c1.text_input("출근", value=curr_s, key=f"s_{k_base}")
                new_e = c2.text_input("퇴근", value=curr_e, key=f"e_{k_base}")
                
                updates[slot["col"]] = {"n": final_n, "s": new_s, "e": new_e}

    st.markdown("###")
    applied = st.form_submit_button("✅ 구글 장부에 저장하기", use_container_width=True)

if applied:
    status_msg = st.empty()
    status_msg.info("⏳ 저장 중...")
    
    if new_names_to_add:
        y_values = worksheet.col_values(25)
        next_r = 3
        for val in y_values[2:]:
            if val.strip() == "": break
            next_r += 1
        for name in new_names_to_add:
            worksheet.update_cell(next_r, 25, name)
            worksheet.update_cell(next_r, 24, next_r - 2)
            next_r += 1
            
    cells_to_update = []
    for col_num, data in updates.items():
        cells_to_update.append(gspread.Cell(target_row, col_num, data["n"]))
        cells_to_update.append(gspread.Cell(target_row, col_num+1, data["s"]))
        cells_to_update.append(gspread.Cell(target_row, col_num+2, data["e"]))
    
    worksheet.update_cells(cells_to_update)
    status_msg.success("✅ 저장 완료!")
    time.sleep(1)
    st.rerun()
import streamlit as st
import pandas as pd
import urllib.parse
import streamlit.components.v1 as components
import re
from collections import defaultdict

# 페이지 설정
st.set_page_config(layout="wide", page_title="통합 데이터 뷰어")

@st.cache_data
def load_data():
    url = 'https://docs.google.com/spreadsheets/d/1efbahF_we5GN8uXeEu4C9CPCj6tniaLn3bNagqwA-ks/export?format=xlsx'
    try:
        sheets = pd.read_excel(url, sheet_name=None, engine='openpyxl', header=6)
        combined_data = pd.DataFrame()
        for sheet_name, df in sheets.items():
            if not df.empty:
                df['출처_시트'] = sheet_name
                combined_data = pd.concat([combined_data, df], ignore_index=True)
        return combined_data
    except Exception as e:
        st.error(f"데이터 로드 중 오류 발생: {e}")
        return pd.DataFrame()

# 부서 프로그램의 무게 및 규격 분리 로직
def parse_packing_string(val, is_size=False):
    if pd.isna(val) or str(val).strip() in ['', '-', '0', '""', 'nan', 'None']:
        return {'is_ditto': False, 'total_qty': 0, 'total_val': 0.0, 'formatted': '', 'line_data': []}
        
    val = str(val).strip()
    if val in ['"', '〃', '”', '“']:
        return {'is_ditto': True, 'total_qty': 0, 'total_val': 0.0, 'formatted': '" (합포장)', 'line_data': []}
        
    lines = [line.strip() for line in val.split('\n') if line.strip()]
    total_qty = 0
    total_val_sum = 0.0
    formatted_lines = []
    line_data = [] 
    
    for line in lines:
        qty = 1
        value_part = line.strip()
        
        if is_size:
            parts = [p.strip() for p in re.split(r'[xX*]', value_part)]
            if len(parts) >= 4:
                last_part = parts[-1]
                qty_match = re.match(r'^(\d+)', last_part)
                if qty_match:
                    qty = int(qty_match.group(1))
                    value_part = f"{parts[0]}*{parts[1]}*{parts[2]}"
                else:
                    value_part = f"{parts[0]}*{parts[1]}*{parts[2]}"
            elif len(parts) == 3:
                value_part = f"{parts[0]}*{parts[1]}*{parts[2]}"
            else:
                fallback_match = re.search(r'[xX*]\s*(\d+)\s*(ea|box|bxs)?\b', value_part, re.IGNORECASE)
                if fallback_match:
                    qty = int(fallback_match.group(1))
                    value_part = value_part[:fallback_match.start()].strip()
        else:
            explicit_match = re.search(r'[xX*]\s*(\d+)\s*(ea|box|bxs|ctns?|boxes)?\b', value_part, re.IGNORECASE)
            if explicit_match:
                qty = int(explicit_match.group(1))
                value_part = value_part[:explicit_match.start()].strip()
                
            clean_str = value_part.replace(',', '')
            nums = re.findall(r'\d+(?:\.\d+)?', clean_str)
            if nums:
                numeric_val = max([float(n) for n in nums])
                total_val_sum += numeric_val * qty
                
        total_qty += qty
        formatted_lines.append(f"{value_part} x {qty}EA" if qty > 1 else value_part)
        line_data.append({'value': value_part, 'qty': qty})
        
    return {
        'is_ditto': False, 
        'total_qty': total_qty, 
        'total_val': round(total_val_sum, 2), 
        'formatted': '\n'.join(formatted_lines),
        'line_data': line_data
    }

st.title("📊 통합 데이터 뷰어 및 메일 전송")

try:
    with st.spinner('데이터를 불러오는 중입니다...'):
        raw_df = load_data()

    if raw_df.empty:
        st.stop()

    col_indices = [0, 1, 2, 3, 4, 5, 15, 16, 17, 20, 21]
    max_col = len(raw_df.columns)
    valid_indices = [i for i in col_indices if i < max_col]

    # 세션 키 업데이트 (my_data_v7)
    if 'my_data_v7' not in st.session_state:
        selected_df = raw_df.iloc[:, valid_indices].copy()
        
        if 15 < max_col and 16 < max_col:
            p_col_name = raw_df.columns[15]
            q_col_name = raw_df.columns[16]
            
            def apply_packing_logic(row):
                w_val = row[p_col_name]
                s_val = row[q_col_name]
                
                p_weight = parse_packing_string(w_val, is_size=False)
                p_size = parse_packing_string(s_val, is_size=True)
                
                is_ditto = p_weight['is_ditto'] or p_size['is_ditto']
                box_qty = 0 if is_ditto else max(p_weight['total_qty'], p_size['total_qty'])
                
                return pd.Series({
                    '계산된 박스수': '합포장' if is_ditto else f"{box_qty} BOX",
                    '계산된 총 무게': p_weight['total_val'],
                    '정리된 규격': p_size['formatted']
                })
            
            calc_df = selected_df.apply(apply_packing_logic, axis=1)
            selected_df = pd.concat([selected_df, calc_df], axis=1)

        selected_df.insert(0, '선택', False)
        if '출처_시트' not in selected_df.columns:
            selected_df.insert(1, '출처_시트', raw_df['출처_시트'])
            
        st.session_state.my_data_v7 = selected_df

    if st.button("🔄 구글 시트 최신 데이터로 새로고침"):
        st.cache_data.clear()
        if 'my_data_v7' in st.session_state:
            del st.session_state.my_data_v7
        st.rerun()

    st.divider()

    # --- 검색 및 필터 UI ---
    st.markdown("### 🔎 검색 및 상세 필터")
    
    # 🌟 [수정됨] 체크박스 텍스트 간소화
    u_col_name = raw_df.columns[20] if 20 < max_col else None
    show_received_only = False
    if u_col_name:
        show_received_only = st.checkbox("📦 입고 완료 항목", value=False)
        st.write("") # 간격 띄우기

    col_search, col_filter1, col_filter2 = st.columns([2, 1, 1])
    
    with col_search:
        search_query = st.text_input("⌨️ 텍스트 통합 검색", placeholder="거래처명, INVOICE 번호 등 아무거나 입력하세요")
        
    with col_filter1:
        filter_cols = [col for col in st.session_state.my_data_v7.columns if col != '선택']
        selected_filter_col = st.selectbox("📂 필터 걸 열 선택", ["선택 안 함"] + filter_cols)
        
    with col_filter2:
        selected_items = []
        if selected_filter_col != "선택 안 함":
            unique_values = st.session_state.my_data_v7[selected_filter_col].dropna().astype(str).unique().tolist()
            selected_items = st.multiselect("📌 항목 체크", unique_values, placeholder="원하는 항목 고르기")

    st.write("")
    all_columns = [col for col in st.session_state.my_data_v7.columns if col != '선택']
    selected_view_cols = st.multiselect(
        "👁️ 표에 보여줄 열 선택 (여기서 제외하면 메일 양식에서도 빠집니다)", 
        options=all_columns, 
        default=all_columns
    )

    # --- 필터 로직 적용 ---
    temp_df = st.session_state.my_data_v7.copy()
    
    # 1. 입고 여부 필터 (U열 내용 있음)
    if show_received_only and u_col_name in temp_df.columns:
        mask_received = temp_df[u_col_name].notna() & \
                        (temp_df[u_col_name].astype(str).str.strip() != '') & \
                        (temp_df[u_col_name].astype(str).str.strip().str.lower() != 'nan') & \
                        (temp_df[u_col_name].astype(str).str.strip().str.lower() != 'nat')
        temp_df = temp_df[mask_received]

    # 2. 텍스트 검색어 필터
    if search_query:
        mask = temp_df.drop(columns=['선택']).astype(str).apply(lambda x: x.str.contains(search_query, case=False, na=False)).any(axis=1)
        temp_df = temp_df[mask]

    # 3. 상세 항목 체크 필터
    if selected_filter_col != "선택 안 함" and selected_items:
        temp_df = temp_df[temp_df[selected_filter_col].astype(str).isin(selected_items)]

    display_df = temp_df[['선택'] + selected_view_cols]

    st.write("")
    col_btn1, col_btn2, _ = st.columns([2, 2, 6])
    with col_btn1:
        if st.button("✅ 현재 보이는 목록 전체 선택"):
            st.session_state.my_data_v7.loc[display_df.index, '선택'] = True
            st.rerun()
    with col_btn2:
        if st.button("❌ 전체 선택 해제"):
            st.session_state.my_data_v7.loc[display_df.index, '선택'] = False
            st.rerun()

    st.caption("💡 팁: 표 맨 위 제목을 클릭해 실수로 정렬이 꼬였다면, 제목을 한두 번 더 눌러 화살표(↑,↓)를 없애면 원래대로 돌아옵니다.")

    # --- 데이터 표 ---
    r_col_name = raw_df.columns[17] if 17 < max_col else None
    disabled_cols = [col for col in display_df.columns if col not in ['선택', r_col_name]]

    edited_df = st.data_editor(
        display_df,
        disabled=disabled_cols,
        use_container_width=True,
        height=400,
        hide_index=True,
        key="main_editor"
    )

    if not edited_df.equals(display_df):
        st.session_state.my_data_v7.update(edited_df)
        st.rerun()

    # --- 📧 메일 양식 생성 ---
    st.divider()
    selected_rows = st.session_state.my_data_v7[st.session_state.my_data_v7['선택'] == True]

    if not selected_rows.empty:
        st.markdown("### 📧 메일 양식 복사 및 보내기")
        
        r_col_name = raw_df.columns[17] if 17 < max_col else "INVOICE"
        invoice_list = selected_rows[r_col_name].dropna().astype(str).unique()
        invoice_text = ", ".join(invoice_list)
        
        def extract_box_num(val):
            if pd.isna(val) or val == '합포장' or val == '': return 0
            nums = re.findall(r'\d+', str(val))
            return int(nums[0]) if nums else 0
            
        total_boxes = selected_rows['계산된 박스수'].apply(extract_box_num).sum() if '계산된 박스수' in selected_rows.columns else 0
        total_weight = selected_rows['계산된 총 무게'].fillna(0).sum() if '계산된 총 무게' in selected_rows.columns else 0
        
        size_html_table = ""
        size_text_for_mailto = ""
        q_col_name = raw_df.columns[16] if 16 < max_col else None
        
        if q_col_name and q_col_name in selected_rows.columns:
            size_map = defaultdict(int)
            for val in selected_rows[q_col_name]:
                parsed = parse_packing_string(val, is_size=True)
                if not parsed.get('is_ditto', False) and parsed.get('line_data'):
                    for ld in parsed['line_data']:
                        if ld['value']:
                            size_map[ld['value']] += ld['qty']
            
            if size_map:
                size_table_rows = ""
                size_text_for_mailto = "\n[ 규격별 박스 수 ]\n"
                
                for size_name, qty in size_map.items():
                    size_table_rows += f"<tr><td style='padding: 6px 12px; border: 1px solid #c0c0c0;'>{size_name}</td><td style='padding: 6px 12px; border: 1px solid #c0c0c0; text-align: center;'>{qty}</td></tr>"
                    size_text_for_mailto += f"- {size_name} : {qty}\n"
                
                size_html_table = f"""
                <br>
                <table style="border-collapse: collapse; margin-top: 5px; margin-bottom: 15px; font-size: 13px; min-width: 250px;">
                    <thead>
                        <tr>
                            <th style="padding: 6px 12px; border: 1px solid #c0c0c0; background-color: #f0f4f8;">박스 규격</th>
                            <th style="padding: 6px 12px; border: 1px solid #c0c0c0; background-color: #f0f4f8;">박스 수</th>
                        </tr>
                    </thead>
                    <tbody>
                        {size_table_rows}
                    </tbody>
                </table>
                """

        mail_data = selected_rows[selected_view_cols]
        html_table = mail_data.to_html(index=False, border=1)

        full_mail_content = f"""안녕하세요,<br>하기의 건으로 출하요청 드립니다.<br><br><b>출하요청일 :</b> <br><b>INVOICE :</b> {invoice_text}<br><br><b>총 박스 수 / 무게 :</b> {total_boxes} BOX / {total_weight:,.2f} kg{size_html_table}<br>{html_table}"""

        copy_and_preview_html = f"""
        <div id="email-area" style="background-color: white; padding: 20px; border: 2px dashed #007BFF; border-radius: 8px; color: black; font-family: 'Malgun Gothic', sans-serif;">
            {full_mail_content}
        </div>
        <br>
        <button onclick="copyToClipboard()" style="
            padding: 10px 20px; background-color: #007BFF; color: white; border: none; border-radius: 5px; cursor: pointer; font-weight: bold;
        ">📋 양식 복사하기</button>

        <script>
        function copyToClipboard() {{
            var container = document.getElementById('email-area');
            var range = document.createRange();
            range.selectNode(container);
            window.getSelection().removeAllRanges();
            window.getSelection().addRange(range);
            document.execCommand('copy');
            window.getSelection().removeAllRanges();
            alert('메일 양식이 클립보드에 복사되었습니다! 아웃룩에 붙여넣기(Ctrl+V) 하세요.');
        }}
        </script>
        """
        components.html(copy_and_preview_html, height=550, scrolling=True)

        st.markdown("#### 🚀 2단계: 메일 프로그램 열기")
        st.caption("위에서 [📋 양식 복사하기]를 눌러 복사한 후, 아래 버튼을 눌러 아웃룩을 띄워 붙여넣으세요.")
        
        subject = urllib.parse.quote("출하요청 드립니다.")
        body_text = f"안녕하세요,\n\n출하요청 건입니다.\nINVOICE: {invoice_text}\n총 박스 수 / 무게 : {total_boxes} BOX / {total_weight:,.2f} kg\n{size_text_for_mailto}\n(여기에 복사한 내용을 붙여넣으세요)"
        mailto_link = f"mailto:?subject={subject}&body={urllib.parse.quote(body_text)}"

        st.markdown(f"""
            <a href="{mailto_link}" target="_self" style="
                display: inline-block; padding: 12px 24px; background-color: #28a745; color: white; 
                text-decoration: none; border-radius: 8px; font-weight: bold;
            ">아웃룩(메일) 창 열기 ✉️</a>
        """, unsafe_allow_html=True)

    else:
        st.info("💡 위 표에서 보낼 항목의 **'선택'** 체크박스를 클릭하면 복사 버튼과 메일 양식이 나타납니다.")

except Exception as e:
    st.error(f"오류: {e}")

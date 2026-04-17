import streamlit as st
import pandas as pd
import urllib.parse
import streamlit.components.v1 as components
import re

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

# 부서 프로그램의 무게 및 박스수 계산 로직
def parse_packing_string(val, is_size=False):
    if pd.isna(val) or str(val).strip() in ['', '-', '0', '""']:
        return {'is_ditto': False, 'total_qty': 0, 'total_val': 0.0, 'formatted': ''}
        
    val = str(val).strip()
    if val in ['"', '〃', '”', '“']:
        return {'is_ditto': True, 'total_qty': 0, 'total_val': 0.0, 'formatted': '" (합포장)'}
        
    lines = [line.strip() for line in val.split('\n') if line.strip()]
    total_qty = 0
    total_val_sum = 0.0
    formatted_lines = []
    
    for line in lines:
        qty = 1
        value_part = line
        
        # 'x 2 ea', '* 3 boxes' 같은 패턴 찾기
        explicit_match = re.search(r'[xX*]\s*(\d+)\s*(ea|box|bxs|ctns?|boxes)?\b', line, re.IGNORECASE)
        
        if explicit_match:
            if not (is_size and not explicit_match.group(2)):
                qty = int(explicit_match.group(1))
                value_part = line[:explicit_match.start()].strip()
                
        # 규격(Size)인 경우: 가로x세로x높이x수량 구조 파악
        if is_size and qty == 1:
            parts = [p.strip() for p in re.split(r'[xX*]', line)]
            if len(parts) == 4 and parts[3].isdigit():
                qty = int(parts[3])
                value_part = ' x '.join(parts[:3])
                
        total_qty += qty
        
        # 무게인 경우: 숫자 추출 후 가장 큰 값을 단위 무게로 인식
        if not is_size:
            clean_str = value_part.replace(',', '')
            nums = re.findall(r'\d+(?:\.\d+)?', clean_str)
            if nums:
                numeric_val = max([float(n) for n in nums])
                total_val_sum += numeric_val * qty
                
        formatted_lines.append(f"{value_part} x {qty}EA" if qty > 1 else value_part)
        
    return {
        'is_ditto': False, 
        'total_qty': total_qty, 
        'total_val': round(total_val_sum, 2), 
        'formatted': '\n'.join(formatted_lines)
    }

st.title("📊 통합 데이터 뷰어 및 메일 전송")

try:
    with st.spinner('데이터를 불러오는 중입니다...'):
        raw_df = load_data()

    if raw_df.empty:
        st.stop()

    # 데이터 설정 및 세션 저장
    col_indices = [0, 1, 2, 3, 4, 5, 15, 16, 17, 20, 21]
    max_col = len(raw_df.columns)
    valid_indices = [i for i in col_indices if i < max_col]

    if 'my_data' not in st.session_state:
        selected_df = raw_df.iloc[:, valid_indices].copy()
        
        # P열(15)과 Q열(16) 데이터 자동 계산 적용
        if 15 < max_col and 16 < max_col:
            p_col_name = raw_df.columns[15] # 무게 열
            q_col_name = raw_df.columns[16] # 치수/박스 열
            
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
            
            # 계산된 결과를 새 열로 추가
            calc_df = selected_df.apply(apply_packing_logic, axis=1)
            selected_df = pd.concat([selected_df, calc_df], axis=1)

        selected_df.insert(0, '선택', False)
        st.session_state.my_data = selected_df

    if st.button("🔄 데이터 새로고침"):
        st.cache_data.clear()
        del st.session_state.my_data
        st.rerun()

    # --- 필터 및 검색 ---
    st.markdown("### 🔎 검색 및 필터")
    col_search, col_f1 = st.columns([2, 2])
    with col_search:
        search_query = st.text_input("⌨️ 통합 검색", placeholder="검색어 입력")
    
    all_columns = [col for col in st.session_state.my_data.columns if col != '선택']
    selected_view_cols = st.multiselect("👁️ 표에 보여줄 열 선택", options=all_columns, default=all_columns)

    # 필터 적용
    temp_df = st.session_state.my_data.copy()
    if search_query:
        mask = temp_df.drop(columns=['선택']).astype(str).apply(lambda x: x.str.contains(search_query, case=False, na=False)).any(axis=1)
        temp_df = temp_df[mask]

    display_df = temp_df[['선택'] + selected_view_cols]

    # --- 데이터 표 ---
    st.caption("💡 새롭게 추가된 **'계산된 박스수', '계산된 총 무게', '정리된 규격'** 열을 확인해 보세요.")
    edited_df = st.data_editor(display_df, use_container_width=True, height=400, hide_index=True, key="main_editor")

    if not edited_df.equals(display_df):
        st.session_state.my_data.update(edited_df)
        st.rerun()

    # --- 📧 메일 양식 및 복사 기능 ---
    st.divider()
    selected_rows = st.session_state.my_data[st.session_state.my_data['선택'] == True]

    if not selected_rows.empty:
        st.markdown("### 📧 메일 양식")
        
        # INVOICE 정리
        r_col_name = raw_df.columns[17] if 17 < max_col else "INVOICE"
        invoice_list = selected_rows[r_col_name].dropna().astype(str).unique()
        invoice_text = ", ".join(invoice_list)
        
        # 🌟 [추가됨] 총 박스 수 및 총 무게 계산
        def extract_box_num(val):
            if pd.isna(val) or val == '합포장': return 0
            nums = re.findall(r'\d+', str(val))
            return int(nums[0]) if nums else 0
            
        total_boxes = selected_rows['계산된 박스수'].apply(extract_box_num).sum() if '계산된 박스수' in selected_rows.columns else 0
        total_weight = selected_rows['계산된 총 무게'].sum() if '계산된 총 무게' in selected_rows.columns else 0
        
        # 테이블 HTML 생성
        mail_data = selected_rows[selected_view_cols]
        html_table = mail_data.to_html(index=False, border=1)

        # 🌟 [수정됨] 실제 복사될 텍스트 구성 (HTML 포함) - 총 박스수/무게 추가
        full_mail_content = f"""안녕하세요,<br>하기의 건으로 출하요청 드립니다.<br><br><b>출하요청일 :</b> <br><b>INVOICE :</b> {invoice_text}<br><b>총 박스 수 / 무게 :</b> {total_boxes} BOX / {total_weight:,.2f} kg<br><br>{html_table}"""

        # 복사하기 버튼 + 미리보기 박스 (JavaScript 포함)
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
        components.html(copy_and_preview_html, height=450, scrolling=True)

        # --- 메일 앱 호출 버튼 ---
        st.markdown("#### 🚀 아웃룩이 안 뜬다면?")
        st.caption("아래 버튼을 눌러도 반응이 없다면, 윈도우 설정에서 '기본 메일 앱'을 Outlook으로 설정해야 합니다.")
        
        subject = urllib.parse.quote("출하요청 드립니다.")
        body = urllib.parse.quote(f"안녕하세요,\n\n출하요청 건입니다.\nINVOICE: {invoice_text}\n총 박스 수 / 무게 : {total_boxes} BOX / {total_weight:,.2f} kg\n\n(여기에 복사한 내용을 붙여넣으세요)")
        mailto_link = f"mailto:?subject={subject}&body={body}"

        st.markdown(f"""
            <a href="{mailto_link}" target="_self" style="
                display: inline-block; padding: 12px 24px; background-color: #28a745; color: white; 
                text-decoration: none; border-radius: 8px; font-weight: bold;
            ">아웃룩(메일) 창 열기 ✉️</a>
        """, unsafe_allow_html=True)

    else:
        st.info("💡 항목을 선택하면 복사 버튼과 메일 양식이 나타납니다.")

except Exception as e:
    st.error(f"오류: {e}")

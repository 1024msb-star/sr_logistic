import streamlit as st
import pandas as pd
import urllib.parse
import streamlit.components.v1 as components

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
    selected_view_cols = st.multiselect("👁️ 보여줄 열 선택", options=all_columns, default=all_columns)

    # 필터 적용
    temp_df = st.session_state.my_data.copy()
    if search_query:
        mask = temp_df.drop(columns=['선택']).astype(str).apply(lambda x: x.str.contains(search_query, case=False, na=False)).any(axis=1)
        temp_df = temp_df[mask]

    display_df = temp_df[['선택'] + selected_view_cols]

    # --- 데이터 표 ---
    edited_df = st.data_editor(display_df, use_container_width=True, height=400, hide_index=True, key="main_editor")

    if not edited_df.equals(display_df):
        st.session_state.my_data.update(edited_df)
        st.rerun()

    # --- 📧 메일 양식 및 복사 기능 ---
    st.divider()
    selected_rows = st.session_state.my_data[st.session_state.my_data['선택'] == True]

    if not selected_rows.empty:
        st.markdown("### 📧 메일 양식")
        
        # 데이터 정리
        r_col_name = raw_df.columns[17] if 17 < max_col else "INVOICE"
        invoice_list = selected_rows[r_col_name].dropna().astype(str).unique()
        invoice_text = ", ".join(invoice_list)
        
        mail_data = selected_rows[selected_view_cols]
        html_table = mail_data.to_html(index=False, border=1)

        # 🌟 실제 복사될 텍스트 구성 (HTML 포함)
        full_mail_content = f"""안녕하세요,<br>하기의 건으로 출하요청 드립니다.<br><br><b>출하요청일 :</b> <br><b>INVOICE :</b> {invoice_text}<br><br>{html_table}"""

        # 🌟 복사하기 버튼 + 미리보기 박스 (JavaScript 포함)
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
        body = urllib.parse.quote(f"안녕하세요,\n\n출하요청 건입니다.\nINVOICE: {invoice_text}\n\n(여기에 복사한 내용을 붙여넣으세요)")
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

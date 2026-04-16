import streamlit as st
import pandas as pd
import urllib.parse

st.set_page_config(layout="wide")

@st.cache_data
def load_data():
    url = 'https://docs.google.com/spreadsheets/d/1efbahF_we5GN8uXeEu4C9CPCj6tniaLn3bNagqwA-ks/export?format=xlsx'
    sheets = pd.read_excel(url, sheet_name=None, engine='openpyxl', header=6)
    combined_data = pd.DataFrame()
    for sheet_name, df in sheets.items():
        if not df.empty:
            df['출처_시트'] = sheet_name
            combined_data = pd.concat([combined_data, df], ignore_index=True)
    return combined_data

st.title("📊 통합 데이터 뷰어 및 메일 전송")

try:
    with st.spinner('데이터를 불러오는 중입니다...'):
        raw_df = load_data()
        
    col_indices = [0, 1, 2, 3, 4, 5, 15, 16, 17, 20, 21]
    max_col = len(raw_df.columns)
    valid_indices = [i for i in col_indices if i < max_col]
    
    if valid_indices:
        # 데이터 고정 (세션 상태 저장)
        if 'my_data' not in st.session_state:
            selected_df = raw_df.iloc[:, valid_indices].copy()
            selected_df.insert(0, '선택', False)
            selected_df.insert(1, '출처_시트', raw_df['출처_시트'])
            st.session_state.my_data = selected_df

        if st.button("🔄 구글 시트 최신 데이터로 새로고침"):
            st.cache_data.clear()
            del st.session_state.my_data
            st.rerun()
            
        st.divider()

        # --- 🌟 검색 및 필터 ---
        st.markdown("### 🔎 검색 및 상세 필터")
        col_search, col_filter1, col_filter2 = st.columns([2, 1, 1])
        
        with col_search:
            search_query = st.text_input("⌨️ 텍스트 통합 검색", placeholder="거래처명, INVOICE 번호 등 아무거나 입력하세요")
            
        with col_filter1:
            filter_cols = [col for col in st.session_state.my_data.columns if col != '선택']
            selected_filter_col = st.selectbox("📂 필터 걸 열 선택", ["선택 안 함"] + filter_cols)
            
        with col_filter2:
            selected_items = []
            if selected_filter_col != "선택 안 함":
                unique_values = st.session_state.my_data[selected_filter_col].dropna().astype(str).unique().tolist()
                selected_items = st.multiselect("📌 항목 체크", unique_values, placeholder="원하는 항목 고르기")

        # --- 🌟 [신규] 열 숨기기/보이기 기능 ---
        st.write("") # 여백
        all_columns = [col for col in st.session_state.my_data.columns if col != '선택']
        selected_view_cols = st.multiselect(
            "👁️ 표에 보여줄 열(컬럼) 선택 (여기서 제외하면 메일 양식에서도 빠집니다)", 
            options=all_columns, 
            default=all_columns, 
            placeholder="보여줄 열을 선택하세요"
        )
        
        # 1. 텍스트 검색어 적용
        if search_query:
            mask = st.session_state.my_data.drop(columns=['선택']).astype(str).apply(lambda x: x.str.contains(search_query, case=False, na=False)).any(axis=1)
            display_df = st.session_state.my_data[mask]
        else:
            display_df = st.session_state.my_data

        # 2. 상세 필터 적용 (교집합)
        if selected_filter_col != "선택 안 함" and selected_items:
            display_df = display_df[display_df[selected_filter_col].astype(str).isin(selected_items)]
            
        # 3. 선택한 열만 화면에 표시 ('선택' 열은 무조건 포함)
        final_display_cols = ['선택'] + selected_view_cols
        display_df = display_df[final_display_cols]

        # --- 선택 버튼 ---
        st.write("")
        col_btn1, col_btn2, _ = st.columns([2, 2, 6])
        with col_btn1:
            if st.button("✅ 현재 보이는 목록 전체 선택"):
                st.session_state.my_data.loc[display_df.index, '선택'] = True
                st.rerun()
        with col_btn2:
            if st.button("❌ 전체 선택 해제"):
                st.session_state.my_data.loc[display_df.index, '선택'] = False
                st.rerun()

        st.caption("💡 팁: 표 맨 위 제목을 클릭해 실수로 정렬이 꼬였다면, 제목을 한두 번 더 눌러 화살표(↑,↓)를 없애면 원래대로 돌아옵니다.")

        # --- 데이터 표 ---
        r_col_name = raw_df.columns[17] if 17 < max_col else None
        disabled_cols = [col for col in display_df.columns if col not in ['선택', r_col_name]]
        
        # 🌟 화면 튕김 방지를 위해 key 추가
        edited_df = st.data_editor(
            display_df,
            disabled=disabled_cols,
            use_container_width=True,
            height=400,
            hide_index=True,
            key="main_data_editor" 
        )
        
        # 수정된 내용(체크박스 등)을 원본 데이터에 업데이트
        st.session_state.my_data.update(edited_df)
        selected_rows = st.session_state.my_data[st.session_state.my_data['선택'] == True]
        
        # --- 메일 양식 ---
        st.divider()
        st.markdown("### 📧 메일 양식 복사 및 보내기")
        
        if not selected_rows.empty:
            if r_col_name and r_col_name in selected_rows.columns:
                invoices = selected_rows[r_col_name].dropna().astype(str)
                invoices = invoices[~invoices.isin(['', 'nan', 'None'])]
                unique_invoices = invoices.unique()
                invoice_text = ", ".join(unique_invoices)
            else:
                invoice_text = ""

            # 🌟 [신규] 화면에서 선택한 열(selected_view_cols)만 메일 표에 포함
            mail_data = selected_rows[selected_view_cols]
            html_table = mail_data.to_html(index=False, border=1)
            
            st.success("👇 **아래 점선 박스 안의 '인사말부터 표 끝까지' 쭉 드래그해서 복사(Ctrl+C)하세요!**")
            
            st.markdown(
                f"""
                <div style="background-color: white; padding: 20px; border: 2px dashed #007BFF; border-radius: 8px; color: black; font-family: 'Malgun Gothic', sans-serif; line-height: 1.6;">
                    안녕하세요,<br>
                    하기의 건으로 출하요청 드립니다.<br><br>
                    출하요청일 : <br>
                    INVOICE : {invoice_text}<br><br>
                    {html_table}
                </div>
                """, 
                unsafe_allow_html=True
            )
            
            st.markdown("---")
            st.markdown("#### 🚀 메일 창 띄우기")
            
            plain_text_body = f"안녕하세요,\n하기의 건으로 출하요청 드립니다.\n\n출하요청일 : \nINVOICE : {invoice_text}\n\n(여기에 복사하신 표를 붙여넣기 해주세요 Ctrl+V)\n"
            subject = urllib.parse.quote("출하요청 드립니다.")
            body = urllib.parse.quote(plain_text_body)
            mailto_link = f"mailto:?subject={subject}&body={body}"
            
            st.markdown(
                f"""
                <a href="{mailto_link}" style="
                    display: inline-block;
                    padding: 10px 20px;
                    font-size: 16px;
                    color: white;
                    background-color: #28a745;
                    text-align: center;
                    text-decoration: none;
                    border-radius: 8px;
                    font-weight: bold;
                " target="_blank">메일 창 열기 ✉️</a>
                """,
                unsafe_allow_html=True
            )
        else:
            st.info("💡 위에서 보낼 행의 체크박스를 선택하시면 자동으로 메일 양식이 완성됩니다.")
            
except Exception as e:
    st.error(f"오류가 발생했습니다: {e}")
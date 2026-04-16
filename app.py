import streamlit as st
import pandas as pd
import urllib.parse

# 페이지 설정
st.set_page_config(layout="wide", page_title="통합 데이터 뷰어")

@st.cache_data
def load_data():
    # 실제 구글 시트 URL (XLSX 내보내기 형식)
    url = 'https://docs.google.com/spreadsheets/d/1efbahF_we5GN8uXeEu4C9CPCj6tniaLn3bNagqwA-ks/export?format=xlsx'
    try:
        # header=6은 7행부터 데이터를 읽음을 의미 (0부터 시작하므로)
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
        st.warning("데이터를 불러오지 못했습니다. URL이나 시트 형식을 확인해주세요.")
        st.stop()

    # 표시할 열 인덱스 설정
    col_indices = [0, 1, 2, 3, 4, 5, 15, 16, 17, 20, 21]
    max_col = len(raw_df.columns)
    valid_indices = [i for i in col_indices if i < max_col]

    # 세션 상태에 데이터 저장 (원본 유지 및 체크박스 상태 보존)
    if 'my_data' not in st.session_state:
        selected_df = raw_df.iloc[:, valid_indices].copy()
        selected_df.insert(0, '선택', False)
        # '출처_시트' 열이 인덱스에 없더라도 추가 보장
        if '출처_시트' not in selected_df.columns:
            selected_df['출처_시트'] = raw_df['출처_시트']
        st.session_state.my_data = selected_df

    # 새로고침 버튼
    if st.button("🔄 구글 시트 최신 데이터로 새로고침"):
        st.cache_data.clear()
        if 'my_data' in st.session_state:
            del st.session_state.my_data
        st.rerun()

    st.divider()

    # --- 🔎 검색 및 상세 필터 ---
    st.markdown("### 🔎 검색 및 상세 필터")
    col_search, col_filter1, col_filter2 = st.columns([2, 1, 1])

    with col_search:
        search_query = st.text_input("⌨️ 텍스트 통합 검색", placeholder="거래처명, INVOICE 번호 등 입력")

    with col_filter1:
        filter_cols = [col for col in st.session_state.my_data.columns if col != '선택']
        selected_filter_col = st.selectbox("📂 필터링할 열 선택", ["선택 안 함"] + filter_cols)

    with col_filter2:
        selected_items = []
        if selected_filter_col != "선택 안 함":
            unique_values = st.session_state.my_data[selected_filter_col].dropna().astype(str).unique().tolist()
            selected_items = st.multiselect("📌 항목 체크", unique_values)

    # --- 👁️ 열 숨기기/보이기 기능 ---
    all_columns = [col for col in st.session_state.my_data.columns if col != '선택']
    selected_view_cols = st.multiselect(
        "👁️ 표에 보여줄 열 선택 (선택된 열만 메일에 포함됩니다)",
        options=all_columns,
        default=all_columns
    )

    # 필터링 로직 적용
    temp_df = st.session_state.my_data.copy()
    
    if search_query:
        mask = temp_df.drop(columns=['선택']).astype(str).apply(
            lambda x: x.str.contains(search_query, case=False, na=False)
        ).any(axis=1)
        temp_df = temp_df[mask]

    if selected_filter_col != "선택 안 함" and selected_items:
        temp_df = temp_df[temp_df[selected_filter_col].astype(str).isin(selected_items)]

    # 현재 보이는 데이터프레임 (선택 열 포함)
    display_df = temp_df[['선택'] + selected_view_cols]

    # --- ✅ 선택 버튼 컨트롤 ---
    col_btn1, col_btn2, _ = st.columns([2, 2, 6])
    with col_btn1:
        if st.button("✅ 현재 보이는 목록 전체 선택"):
            st.session_state.my_data.loc[display_df.index, '선택'] = True
            st.rerun()
    with col_btn2:
        if st.button("❌ 전체 선택 해제"):
            st.session_state.my_data.loc[display_df.index, '선택'] = False
            st.rerun()

    # --- 📑 데이터 표 (에디터) ---
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

    # 수정사항 반영
    if not edited_df.equals(display_df):
        st.session_state.my_data.update(edited_df)
        st.rerun()

    # --- 📧 메일 양식 생성 ---
    st.divider()
    st.markdown("### 📧 메일 양식 복사 및 보내기")
    
    selected_rows = st.session_state.my_data[st.session_state.my_data['선택'] == True]

    if not selected_rows.empty:
        # Invoice 번호 추출 (17번 인덱스 열 기준)
        if r_col_name and r_col_name in selected_rows.columns:
            invoices = selected_rows[r_col_name].dropna().astype(str)
            invoices = invoices[~invoices.isin(['', 'nan', 'None'])]
            invoice_text = ", ".join(invoices.unique())
        else:
            invoice_text = "(없음)"

        # 메일용 테이블 생성 (사용자가 선택한 열만 포함)
        mail_data = selected_rows[selected_view_cols]
        html_table = mail_data.to_html(index=False, border=1, justify='center')
        
        # 복사용 가이드 UI
        st.info("💡 **1단계:** 아래 박스 내부를 마우스로 드래그하여 **복사(Ctrl+C)** 하세요.")
        
        mail_content_html = f"""
        <div style="background-color: white; padding: 20px; border: 2px dashed #007BFF; border-radius: 8px; color: black; font-family: 'Malgun Gothic', sans-serif;">
            안녕하세요,<br>
            하기의 건으로 출하요청 드립니다.<br><br>
            <b>출하요청일 :</b> <br>
            <b>INVOICE :</b> {invoice_text}<br><br>
            {html_table}
        </div>
        """
        st.markdown(mail_content_html, unsafe_allow_html=True)
        
        st.write("")
        st.markdown("#### 🚀 2단계: 메일 프로그램 열기")
        
        # Mailto 링크 생성 (본문 길이를 최소화하여 빈 창 방지)
        subject = "출하요청 드립니다."
        short_body = f"안녕하세요,\n\n출하요청 드립니다.\nINVOICE: {invoice_text}\n\n(여기에 아까 복사한 표를 붙여넣어 주세요 - Ctrl+V)"
        
        mailto_link = f"mailto:?subject={urllib.parse.quote(subject)}&body={urllib.parse.quote(short_body)}"
        
        st.markdown(
            f"""
            <a href="{mailto_link}" target="_self" style="
                display: inline-block;
                padding: 12px 24px;
                background-color: #28a745;
                color: white;
                text-decoration: none;
                border-radius: 8px;
                font-weight: bold;
            ">메일 창 열기 ✉️</a>
            """, 
            unsafe_allow_html=True
        )
    else:
        st.info("💡 위 표에서 보낼 항목의 **'선택'** 체크박스를 클릭하면 메일 양식이 나타납니다.")

except Exception as e:
    st.error(f"오류가 발생했습니다: {e}")

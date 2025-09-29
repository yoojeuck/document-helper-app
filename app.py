import streamlit as st
import pandas as pd
from jinja2 import Environment, FileSystemLoader
from weasyprint import HTML
from datetime import datetime
import streamlit.components.v1 as components
import google.generativeai as genai

# --- AI 설정 ---
# Streamlit Secrets에서 API 키를 가져와 설정합니다.
try:
    genai.configure(api_key=st.secrets["GOOGLE_API_KEY"])
except Exception as e:
    st.error("⚠️ AI 기능을 사용하려면 Streamlit Secrets에 GOOGLE_API_KEY를 등록해야 합니다.")

def generate_purpose_with_ai(keywords):
    """AI를 사용하여 품의 목적 문장을 생성하는 함수"""
    model = genai.GenerativeModel('gemini-pro')
    prompt = f"""
    당신은 한국 기업의 유능한 사원입니다. 다음 핵심 키워드를 바탕으로, 상급자에게 정중하게 보고하는 '품의 목적' 문장을 완성해주세요.
    문장은 "ㅇㅇ하고자 아래와 같이 품의하오니 검토 후 재가 바랍니다." 와 같은 형식으로, 격식 있고 간결하게 작성해주세요.

    핵심 키워드: {keywords}

    완성된 문장:
    """
    try:
        response = model.generate_content(prompt)
        return response.text.strip()
    except Exception as e:
        return f"AI 생성 중 오류가 발생했습니다: {e}"

# --- 기본 앱 설정 (이전과 동일) ---
st.set_page_config(page_title="문서 작성 도우미", layout="wide")
env = Environment(loader=FileSystemLoader('.'))

def load_template(template_name):
    return env.get_template(template_name)
def generate_html(template, context):
    return template.render(context)
def generate_pdf(html_content):
    return HTML(string=html_content).write_pdf()

st.sidebar.title("📑 문서 종류 선택")
doc_type = st.sidebar.radio(
    "작성할 문서의 종류를 선택하세요.",
    ('품의서', '공지문', '공문', '비즈니스 이메일'),
    label_visibility="collapsed"
)

st.title("✍️ AI 문서 작성 도우미 v2.0")
st.markdown(f"**'{doc_type}'** 작성을 시작합니다. 아래 양식에 내용을 입력하거나 AI의 도움을 받아 문서를 완성하세요.")
st.divider()

# ==============================================================================
# --- 품의서 ---
# ==============================================================================
if doc_type == '품의서':
    st.header("품의서 작성")
    if 'pumui_data' not in st.session_state:
        st.session_state.pumui_data = {
            "title": "영업팀 신규 노트북 구매에 관한 건",
            "purpose": "영업팀의 업무 효율성 증대를 위해 노후화된 노트북을 교체하고자 아래와 같이 품의하오니 검토 후 재가 바랍니다.",
            "remarks": "1. 결제 방식: 법인카드 결제\n2. 납품 업체: (주)디지털존\n3. 납품 예정일: 2025년 10월 15일",
            "items_df": pd.DataFrame([
                {"No": 1, "거래처": "(주)디지털존", "품목": "ABC 노트북 모델-15", "단가": 1500000, "수량": 5, "합계": 7500000, "비고": "영업팀"},
                {"No": 2, "거래처": "(주)디지털존", "품목": "무선 마우스", "단가": 30000, "수량": 5, "합계": 150000, "비고": ""},
            ])
        }
    p_data = st.session_state.pumui_data

    # --- ✨ NEW AI FEATURE SECTION ---
    with st.container(border=True):
        st.subheader("✨ AI로 목적 자동 생성")
        st.info("핵심 단어만 입력하고 버튼을 누르면, AI가 격식에 맞는 품의 목적을 자동으로 작성해줍니다.")
        keywords = st.text_input("핵심 키워드", placeholder="예: 영업팀 노트북 교체, 마케팅 캠페인 예산 증액")
        if st.button("AI로 문장 생성하기", use_container_width=True):
            if keywords:
                with st.spinner("AI가 문장을 작성 중입니다..."):
                    generated_purpose = generate_purpose_with_ai(keywords)
                    p_data["purpose"] = generated_purpose
            else:
                st.warning("핵심 키워드를 입력해주세요.")
    # ----------------------------
    
    with st.container(border=True):
        st.subheader("1. 기본 정보")
        p_data["title"] = st.text_input("제목", value=p_data["title"], help="문서의 핵심 내용이 한눈에 파악되도록 명확하게 작성하세요.")
        p_data["purpose"] = st.text_area("1. 목적 및 개요", value=p_data["purpose"], height=100, help="결재자가 '이 보고의 목적이 무엇인가?'라는 의문을 갖지 않도록 핵심 내용을 명료하게 작성하십시오.")
    
    with st.container(border=True):
        st.subheader("2. 상세 내역 (표)")
        p_data["items_df"] = st.data_editor(p_data["items_df"], num_rows="dynamic", key="pumui_editor")

    with st.container(border=True):
        st.subheader("3. 비고 및 참고사항")
        p_data["remarks"] = st.text_area("비고", value=p_data["remarks"], height=150, help="결제 조건, 특이사항 등 의사결정에 필요한 추가 정보를 기입합니다.")

    if 'final_html' not in st.session_state: st.session_state.final_html = ""
    if st.button("1. 미리보기 및 수정 단계로 이동", type="secondary", use_container_width=True):
        if '단가' in p_data["items_df"].columns and '수량' in p_data["items_df"].columns: p_data["items_df"]['합계'] = p_data["items_df"]['단가'] * p_data["items_df"]['수량']
        items = p_data["items_df"].to_dict('records')
        total_sum = p_data["items_df"]['합계'].sum() if '합계' in p_data["items_df"].columns else 0
        context = { "title": p_data["title"], "purpose": p_data["purpose"].replace('\n', '<br>'), "items": items, "total_sum": f"{total_sum:,.0f}", "remarks": p_data["remarks"].replace('\n', '<br>'), "generation_date": datetime.now().strftime('%Y-%m-%d') }
        template = load_template('pumui_template_v2.html')
        st.session_state.final_html = generate_html(template, context)

    if st.session_state.final_html:
        st.subheader("📄 문서 미리보기")
        components.html(st.session_state.final_html, height=600, scrolling=True)
        st.subheader("✏️ 최종 수정용 텍스트 상자")
        edited_html = st.text_area("HTML 원문 수정", value=st.session_state.final_html, height=300)
        if st.button("2. 수정된 내용으로 최종 PDF 생성", type="primary", use_container_width=True):
            pdf_output = generate_pdf(edited_html)
            st.download_button(label="📥 PDF 파일 다운로드", data=pdf_output, file_name=f"{p_data['title']}.pdf", mime="application/pdf", use_container_width=True)

# (공지문, 공문, 비즈니스 이메일 코드는 이전과 동일하게 유지됩니다)
# ... (이하 생략) ...
# (이전 답변의 공지문, 공문, 이메일 코드를 여기에 붙여넣으세요)

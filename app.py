import streamlit as st
import pandas as pd
from jinja2 import Environment, FileSystemLoader
from weasyprint import HTML
from datetime import datetime
import streamlit.components.v1 as components
import google.generativeai as genai

# --- AI 설정 ---
try:
    genai.configure(api_key=st.secrets["GOOGLE_API_KEY"])
    # 가장 빠르고 효율적인 최신 모델을 사용합니다.
    model = genai.GenerativeModel('gemini-1.5-flash-latest')
except Exception as e:
    st.error("⚠️ AI 기능을 사용하려면 Google Cloud에서 'Vertex AI API'를 활성화하고, Streamlit Secrets에 GOOGLE_API_KEY를 등록해야 합니다.")

def generate_purpose_with_ai(keywords):
    """AI를 사용하여 품의 목적 문장을 생성하는 함수"""
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
        return f"AI 생성 중 오류가 발생했습니다. Google Cloud 프로젝트에서 'Vertex AI API'가 활성화되었는지 확인해주세요. 오류 상세: {e}"

# --- 기본 앱 설정 ---
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

st.title("✍️ AI 문서 작성 도우미 v2.1")
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

# ... (이하 공지문, 공문, 비즈니스 이메일 코드는 이전 답변과 동일합니다) ...

# ==============================================================================
# --- 공지문, 공문, 이메일도 동일하게 2단계 방식으로 수정됩니다. ---
# ==============================================================================
elif doc_type == '공지문':
    st.header("공지문 작성")
    if 'gongji_data' not in st.session_state:
        st.session_state.gongji_data = {"title": "사내 정보보안 강화 및 PC 클린업 캠페인 안내", "target": "전 임직원", "summary": "최근 증가하는 사이버 위협에 대응하고, 안전한 업무 환경을 조성하기 위해 정보보안 강화 캠페인을 실시합니다.", "details": "1. 캠페인 기간: 2025년 10월 6일(월) ~ 10월 10일(금)\n2. 주요 내용\n   - PC 비밀번호 변경 (영문, 숫자, 특수문자 포함 10자 이상)\n   - 불필요한 프로그램 및 액티브X 제거\n   - 중요 파일 백업 및 개인정보 파일 암호화\n3. 협조 요청: 전 직원은 기간 내 반드시 조치를 완료해주시기 바랍니다.", "contact": "IT지원팀 김철수 대리 (내선 456)"}
    g_data = st.session_state.gongji_data
    
    with st.container(border=True):
        # (입력 양식은 이전과 동일)
        g_data["title"] = st.text_input("제목", value=g_data["title"], help="공지의 내용을 한눈에 파악할 수 있도록 명시적으로 작성합니다.")
        g_data["target"] = st.text_input("대상", value=g_data["target"], help="공지의 적용 범위를 명확히 합니다. (예: 전 직원, 사업본부 임직원 등)")
        g_data["summary"] = st.text_area("핵심 요약", value=g_data["summary"], height=100, help="본문 상단에 한두 문장으로 공지의 핵심을 요약하여 전달력을 높입니다.")
        g_data["details"] = st.text_area("상세 내용", value=g_data["details"], height=200, help="5W1H(누가, 무엇을, 언제, 어디서, 왜, 어떻게) 원칙에 따라 구체적인 정보를 제공합니다.")
        g_data["contact"] = st.text_input("문의처", value=g_data["contact"], help="관련 질문에 답변할 담당자의 이름과 연락처를 명기합니다.")

    if 'final_html_gongji' not in st.session_state: st.session_state.final_html_gongji = ""
    if st.button("1. 미리보기 및 수정 단계로 이동", type="secondary", use_container_width=True):
        context = {k: v.replace('\n', '<br>') for k, v in g_data.items()}
        context["generation_date"] = datetime.now().strftime('%Y. %m. %d.')
        template = load_template('gongji_template.html')
        st.session_state.final_html_gongji = generate_html(template, context)

    if st.session_state.final_html_gongji:
        st.subheader("📄 문서 미리보기")
        components.html(st.session_state.final_html_gongji, height=600, scrolling=True)
        st.subheader("✏️ 최종 수정용 텍스트 상자")
        edited_html = st.text_area("HTML 원문 수정", value=st.session_state.final_html_gongji, height=300)
        if st.button("2. 수정된 내용으로 최종 PDF 생성", type="primary", use_container_width=True):
            pdf_output = generate_pdf(edited_html)
            st.download_button(label="📥 PDF 파일 다운로드", data=pdf_output, file_name=f"{g_data['title']}.pdf", mime="application/pdf", use_container_width=True)

elif doc_type == '공문':
    st.header("공문 작성")
    if 'gongmun_data' not in st.session_state:
        st.session_state.gongmun_data = {"sender_org": "주식회사 몬쉘코리아", "doc_number": "사업-2025-102호", "receiver": "협력사 A 대표이사", "cc": "내부 법무팀", "title": "신제품 개발 관련 업무 협조 요청", "body": "귀사의 무궁한 발전을 기원합니다.\n\n당사는 2026년 상반기 출시를 목표로 신제품 '프로젝트 델타'를 기획하고 있습니다.\n\n본 프로젝트의 성공적인 수행을 위해 귀사의 기술 지원이 필요한 부분이 있어, 아래와 같이 자료 및 기술 미팅을 정중히 요청드립니다.\n\n- 아 래 -\n\n1. 요청 자료: 신규 부품 XYZ의 기술 사양서 및 샘플\n2. 요청 미팅: 2025년 10월 중순, 양사 실무진 기술 미팅 (일정 추후 협의)\n\n바쁘시겠지만, 긍정적인 검토 부탁드립니다.", "sender_name": "주식회사 몬쉘코리아 대표이사 김수근"}
    gm_data = st.session_state.gongmun_data
    
    with st.container(border=True):
        # (입력 양식은 이전과 동일)
        st.subheader("두문 (머리말)")
        col1, col2 = st.columns(2)
        with col1:
            gm_data["sender_org"] = st.text_input("발신 기관명", value=gm_data["sender_org"], help="기관의 공식 명칭을 기입합니다.")
            gm_data["doc_number"] = st.text_input("문서 번호", value=gm_data["doc_number"], help="문서 관리 및 추적을 위한 정보입니다.")
        with col2:
            gm_data["receiver"] = st.text_input("수신", value=gm_data["receiver"], help="문서를 받는 주체를 명확히 기입합니다.")
            gm_data["cc"] = st.text_input("참조", value=gm_data["cc"], help="참고할 대상을 기입합니다.")
        st.subheader("본문")
        gm_data["title"] = st.text_input("제목", value=gm_data["title"], help="공문의 내용을 함축적으로 나타내는 제목입니다.")
        gm_data["body"] = st.text_area("내용", value=gm_data["body"], height=250, help="전달하고자 하는 핵심 내용을 명료하게 서술합니다.")
        st.subheader("결문 (맺음말)")
        gm_data["sender_name"] = st.text_input("발신 명의", value=gm_data["sender_name"], help="발신 주체의 공식 직함과 이름을 기입합니다. (예: OOO 주식회사 대표이사 OOO)")

    if 'final_html_gongmun' not in st.session_state: st.session_state.final_html_gongmun = ""
    if st.button("1. 미리보기 및 수정 단계로 이동", type="secondary", use_container_width=True):
        context = {k: v.replace('\n', '<br>') for k, v in gm_data.items()}
        context["generation_date"] = datetime.now().strftime('%Y. %m. %d.')
        template = load_template('gongmun_template.html')
        st.session_state.final_html_gongmun = generate_html(template, context)

    if st.session_state.final_html_gongmun:
        st.subheader("📄 문서 미리보기")
        components.html(st.session_state.final_html_gongmun, height=800, scrolling=True)
        st.subheader("✏️ 최종 수정용 텍스트 상자")
        edited_html = st.text_area("HTML 원문 수정", value=st.session_state.final_html_gongmun, height=300)
        if st.button("2. 수정된 내용으로 최종 PDF 생성", type="primary", use_container_width=True):
            pdf_output = generate_pdf(edited_html)
            st.download_button(label="📥 PDF 파일 다운로드", data=pdf_output, file_name=f"{gm_data['title']}.pdf", mime="application/pdf", use_container_width=True)

elif doc_type == '비즈니스 이메일':
    st.header("비즈니스 이메일 작성")
    if 'email_data' not in st.session_state:
        st.session_state.email_data = {"to": "manager@partner-company.com", "cc": "team-leader@my-company.com", "bcc": "", "subject": "[몬쉘코리아] 4분기 마케팅 전략 회의 일정 조율 요청", "intro": "안녕하세요, 박서준 부장님.\n몬쉘코리아 마케팅팀 이지은입니다.", "body": "선선한 가을, 평안히 지내고 계신지 궁금합니다.\n\n다름이 아니라, 4분기 공동 마케팅 캠페인 추진을 위한 실무진 회의를 진행하고자 합니다.\n\n아래 후보 시간 중 편하신 시간을 알려주시거나, 다른 시간을 제안해주시면 감사하겠습니다.\n\n1안) 10월 7일(화) 오후 2시\n2안) 10월 8일(수) 오전 10시\n3안) 10월 9일(목) 오후 4시", "closing": "그럼, 답변 기다리겠습니다.\n감사합니다.", "signature_name": "이지은", "signature_title": "대리", "signature_team": "마케팅팀", "signature_company": "주식회사 몬쉘코리아", "signature_phone": "010-9876-5432", "signature_email": "jieun.lee@mon-chouchou.co.kr"}
    e_data = st.session_state.email_data
    
    with st.container(border=True):
        # (입력 양식은 이전과 동일)
        st.subheader("수신 정보")
        e_data["to"] = st.text_input("받는 사람 (To)", value=e_data["to"])
        e_data["cc"] = st.text_input("참조 (CC)", value=e_data["cc"])
        e_data["bcc"] = st.text_input("숨은 참조 (BCC)", value=e_data["bcc"])
        e_data["subject"] = st.text_input("제목", value=e_data["subject"], help="[소속] OOO 관련 OOO 요청과 같은 형식을 사용하면 전달력을 높일 수 있습니다.")
        st.subheader("본문")
        e_data["intro"] = st.text_area("도입", value=e_data["intro"], height=100, help="간단한 인사와 자기소개를 작성합니다.")
        e_data["body"] = st.text_area("본론", value=e_data["body"], height=150, help="핵심 용건을 두괄식으로 먼저 제시하고, 상세 내용은 가독성 있게 작성합니다.")
        e_data["closing"] = st.text_area("결론", value=e_data["closing"], height=100, help="요청 사항이나 다음 행동을 명확히 요약하고 끝인사로 마무리합니다.")
    
    with st.expander("내 서명 정보 수정하기"):
        e_data["signature_name"] = st.text_input("이름", value=e_data["signature_name"])
        e_data["signature_title"] = st.text_input("직책", value=e_data["signature_title"])
        e_data["signature_team"] = st.text_input("부서/팀", value=e_data["signature_team"])
        e_data["signature_company"] = st.text_input("회사명", value=e_data["signature_company"])
        e_data["signature_phone"] = st.text_input("연락처", value=e_data["signature_phone"])
        e_data["signature_email"] = st.text_input("이메일 주소", value=e_data["signature_email"])
    
    if st.button("이메일 본문 생성 및 복사하기", type="primary", use_container_width=True):
        context = {k: v.replace('\n', '<br>') for k, v in e_data.items()}
        template = load_template('email_template.html')
        html_output = generate_html(template, context)

        st.subheader("📧 이메일 미리보기")
        components.html(html_output, height=400, scrolling=True)
        
        st.subheader("📋 복사할 HTML 코드")
        st.info("이메일 클라이언트가 HTML 붙여넣기를 지원하는 경우, 아래 코드를 복사해서 사용하세요.")
        st.code(html_output, language='html')






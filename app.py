import streamlit as st
import pandas as pd
from jinja2 import Environment, FileSystemLoader
from weasyprint import HTML
from datetime import datetime
import streamlit.components.v1 as components
from openai import OpenAI
import json

# --- AI 설정 (OpenAI GPT-4o mini 사용) ---
try:
    client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])
except Exception as e:
    st.error("⚠️ AI 기능을 사용하려면 Streamlit Secrets에 OPENAI_API_KEY를 등록해야 합니다.")

def generate_ai_draft(doc_type, keywords):
    """문서 종류와 키워드에 따라 AI 초안을 생성하는 범용 함수"""
    prompts = {
        "품의서": {
            "system": """
            당신은 한국 기업의 유능한 사원입니다. 보고서 작성의 전문가로서, 사용자의 핵심 키워드를 바탕으로 '품의서' 초안 전체를 생성합니다.
            응답은 반드시 "title", "purpose", "remarks" key를 포함하는 JSON 형식이어야 합니다.
            사용자의 키워드를 분석하여, 내용에 따라 'items' (표) 또는 'body' (줄글) 중 하나를 선택하여 내용을 구성합니다.
            'items'는 구매, 견적 등 목록화가 필요할 때 사용하며, 내용에 맞는 table header를 자율적으로 정하고, 그에 맞춰 각 항목을 객체 리스트로 작성합니다.
            'body'는 정책 제안, 결과 보고 등 서술이 필요할 때 사용하며, 상세 내용을箇条書き 형식의 문자열로 작성합니다.
            "title", "purpose", "remarks"와 함께 "items" 또는 "body" 중 하나만 JSON에 포함시켜야 합니다.
            """,
            "user": f"핵심 키워드: '{keywords}'를 바탕으로 품의서 초안을 JSON 형식으로 생성해주세요."
        },
        "공지문": {
            "system": """
            당신은 한국 기업의 사내 커뮤니케이션 담당자입니다. 사용자의 핵심 키워드를 바탕으로 명확하고 간결한 '사내 공지문' 초안을 생성합니다.
            응답은 반드시 "title", "target", "summary", "details", "contact" key를 포함하는 JSON 형식이어야 합니다.
            "details"는箇条書き 형식으로 명확하게 작성해주세요.
            """,
            "user": f"핵심 키워드: '{keywords}'를 바탕으로 공지문 초안을 JSON 형식으로 생성해주세요."
        },
        "공문": {
            "system": """
            당신은 대외 문서를 담당하는 총무팀 직원입니다. 사용자의 핵심 키워드를 바탕으로 격식과 규정에 맞는 '공문' 초안을 생성합니다.
            응답은 반드시 "sender_org", "receiver", "cc", "title", "body", "sender_name" key를 포함하는 JSON 형식이어야 합니다.
            "body"에는 정중한 인사말과 '- 아 래 -' 형식의 본문, 그리고 맺음말을 포함해야 합니다.
            """,
            "user": f"핵심 키워드: '{keywords}'를 바탕으로 공문 초안을 JSON 형식으로 생성해주세요."
        },
        "비즈니스 이메일": {
            "system": """
            당신은 비즈니스 커뮤니케이션 전문가입니다. 사용자의 핵심 키워드를 바탕으로 전문적이고 정중한 '비즈니스 이메일' 초안을 생성합니다.
            응답은 반드시 "to", "cc", "subject", "intro", "body", "closing" key를 포함하는 JSON 형식이어야 합니다.
            받는 사람의 이메일 주소는 '이름@회사명.com' 형식으로 추정하여 작성해주세요.
            """,
            "user": f"핵심 키워드: '{keywords}'를 바탕으로 이메일 초안을 JSON 형식으로 생성해주세요."
        }
    }

    if doc_type not in prompts:
        return None

    try:
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            response_format={"type": "json_object"},
            messages=[
                {"role": "system", "content": prompts[doc_type]["system"]},
                {"role": "user", "content": prompts[doc_type]["user"]}
            ],
            temperature=0.7,
        )
        return json.loads(response.choices[0].message.content)
    except Exception as e:
        st.error(f"AI 생성 중 오류가 발생했습니다: {e}")
        return None

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

st.title(f"✍️ AI {doc_type} 자동 생성")
st.markdown(f"핵심 키워드를 입력하면, AI가 **'{doc_type}'** 초안 전체를 자동으로 작성해줍니다.")
st.divider()

keyword_examples = {
    "품의서": "표 생성 예시: '영업팀 태블릿 5대 구매' / 줄글 생성 예시: '사내 휴게공간 개선 건의'",
    "공지문": "예: 10월 전사 워크숍, 제주도, 1박 2일, 참석 여부 회신 요청",
    "공문": "예: A사에 신제품 기술 자료 요청, B팀 참조",
    "비즈니스 이메일": "예: 박서준 부장님께, 4분기 마케팅 회의 일정 조율 요청"
}
keywords = st.text_input("핵심 키워드", placeholder=keyword_examples.get(doc_type, ""))

# 세션 상태 초기화 버튼
if st.button("새 문서 작성 시작 (양식 초기화)"):
    st.session_state.ai_draft = {}
    st.session_state.final_html = ""
    st.rerun()

if st.button(f"AI로 {doc_type} 전체 생성하기", type="primary", use_container_width=True):
    if keywords:
        with st.spinner(f"AI가 {doc_type} 전체를 작성 중입니다..."):
            ai_result = generate_ai_draft(doc_type, keywords)
            if ai_result:
                st.session_state.ai_draft = ai_result
                st.session_state.final_html = "" # 이전 미리보기 초기화
                st.success("AI가 문서 초안을 모두 작성했습니다. 아래 내용을 확인하고 수정하세요.")
    else:
        st.warning("핵심 키워드를 입력해주세요.")

st.divider()

if 'ai_draft' not in st.session_state:
    st.session_state.ai_draft = {}
if 'final_html' not in st.session_state:
    st.session_state.final_html = ""

draft = st.session_state.ai_draft

if draft:
    if doc_type == '품의서':
        p_data = draft
        p_data["title"] = st.text_input("제목", value=p_data.get("title", ""))
        p_data["purpose"] = st.text_area("목적 및 개요", value=p_data.get("purpose", ""), height=100)
        
        if "items" in p_data and p_data["items"]:
            p_data["df"] = pd.DataFrame(p_data["items"])
            st.subheader("상세 내역 (표)")
            p_data["df_edited"] = st.data_editor(p_data["df"], num_rows="dynamic")
            p_data["body_edited"] = ""
        else:
            st.subheader("상세 내용 (줄글)")
            p_data["body_edited"] = st.text_area("내용", value=p_data.get("body", ""), height=200)
            p_data["df_edited"] = pd.DataFrame()

        p_data["remarks"] = st.text_area("비고 및 참고사항", value=p_data.get("remarks", ""), height=150)

        if st.button("미리보기 생성", use_container_width=True):
            context = { "title": p_data["title"], "purpose": p_data["purpose"].replace('\n', '<br>'), "remarks": p_data["remarks"].replace('\n', '<br>'), "generation_date": datetime.now().strftime('%Y-%m-%d') }
            if not p_data["df_edited"].empty:
                context["table_headers"] = list(p_data["df_edited"].columns)
                context["items"] = p_data["df_edited"].to_dict('records')
            else:
                context["body"] = p_data["body_edited"].replace('\n', '<br>')
            
            template = load_template('pumui_template_final.html')
            st.session_state.final_html = generate_html(template, context)

    elif doc_type == '공지문':
        g_data = draft
        g_data["title"] = st.text_input("제목", value=g_data.get("title", ""))
        g_data["target"] = st.text_input("대상", value=g_data.get("target", ""))
        g_data["summary"] = st.text_area("핵심 요약", value=g_data.get("summary", ""), height=100)
        g_data["details"] = st.text_area("상세 내용", value=g_data.get("details", ""), height=200)
        g_data["contact"] = st.text_input("문의처", value=g_data.get("contact", ""))
        
        if st.button("미리보기 생성", use_container_width=True):
            context = {k: v.replace('\n', '<br>') for k, v in g_data.items() if isinstance(v, str)}
            context["generation_date"] = datetime.now().strftime('%Y. %m. %d.')
            template = load_template('gongji_template.html')
            st.session_state.final_html = generate_html(template, context)

    elif doc_type == '공문':
        gm_data = draft
        gm_data["sender_org"] = st.text_input("발신 기관명", value=gm_data.get("sender_org", ""))
        gm_data["receiver"] = st.text_input("수신", value=gm_data.get("receiver", ""))
        gm_data["cc"] = st.text_input("참조", value=gm_data.get("cc", ""))
        gm_data["title"] = st.text_input("제목", value=gm_data.get("title", ""))
        gm_data["body"] = st.text_area("내용", value=gm_data.get("body", ""), height=250)
        gm_data["sender_name"] = st.text_input("발신 명의", value=gm_data.get("sender_name", ""))

        if st.button("미리보기 생성", use_container_width=True):
            context = {k: v.replace('\n', '<br>') for k, v in gm_data.items() if isinstance(v, str)}
            context["generation_date"] = datetime.now().strftime('%Y. %m. %d.')
            template = load_template('gongmun_template.html')
            st.session_state.final_html = generate_html(template, context)

    elif doc_type == '비즈니스 이메일':
        e_data = draft
        e_data["to"] = st.text_input("받는 사람", value=e_data.get("to", ""))
        e_data["cc"] = st.text_input("참조", value=e_data.get("cc", ""))
        e_data["subject"] = st.text_input("제목", value=e_data.get("subject", ""))
        e_data["intro"] = st.text_area("도입", value=e_data.get("intro", ""), height=100)
        e_data["body"] = st.text_area("본론", value=e_data.get("body", ""), height=150)
        e_data["closing"] = st.text_area("결론", value=e_data.get("closing", ""), height=100)

        with st.expander("내 서명 정보 입력/수정"):
            e_data["signature_name"] = st.text_input("이름", value="홍길동")
            e_data["signature_title"] = st.text_input("직책", value="대리")
            e_data["signature_team"] = st.text_input("부서/팀", value="마케팅팀")

        if st.button("이메일 본문 생성", use_container_width=True):
            e_data["signature_company"] = "주식회사 몬쉘코리아" # 고정값 예시
            context = {k: v.replace('\n', '<br>') for k, v in e_data.items() if isinstance(v, str)}
            template = load_template('email_template_final.html')
            st.session_state.final_html = generate_html(template, context)

if st.session_state.final_html:
    st.divider()
    st.subheader("📄 최종 미리보기")
    components.html(st.session_state.final_html, height=600, scrolling=True)

    if doc_type == "비즈니스 이메일":
        st.subheader("📋 복사할 HTML 코드")
        st.code(st.session_state.final_html, language='html')
    else:
        st.subheader("📥 PDF 다운로드")
        pdf_output = generate_pdf(st.session_state.final_html)
        title_for_file = st.session_state.ai_draft.get("title", "document")
        st.download_button(label="PDF 파일 다운로드", data=pdf_output, file_name=f"{title_for_file}.pdf", mime="application/pdf", use_container_width=True)

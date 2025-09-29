import streamlit as st
import pandas as pd
from jinja2 import Environment, FileSystemLoader
from weasyprint import HTML
from datetime import datetime
import streamlit.components.v1 as components
from openai import OpenAI
import json
import markdown

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
            당신은 한국의 '주식회사 몬쉘코리아' 소속의 유능한 사원입니다. 지금부터 제공하는 규칙과 예시를 완벽하게 숙지하고, 사용자의 키워드를 바탕으로 품의서 초안 전체를 생성합니다.

            ### 문서 작성 규칙 (문서작성매뉴얼 기반)
            1.  **번호 매기기:** 본문 항목 구분 시 `1.`, `1)`, `(1)` 의 위계질서를 반드시 준수합니다.
            2.  **가독성:** 의미 단위로 명확하게 줄을 바꾸고, 문장은 간결하게 작성합니다.
            3.  **내용:** 결론을 먼저 제시하고, 이유나 상세 설명을 뒤에 붙이는 두괄식 구성을 선호합니다.
            4.  **종결:** 본문이 끝나면 "끝." 표시를 사용합니다.
            5.  **출력 형식:** 본문(`body`) 또는 상세내역(`items`) 내용은 반드시 마크다운(Markdown) 형식으로 작성해야 합니다.

            ### 품의서 작성 예시 (실제 샘플 기반 학습)
            #### 예시 1: 목록이 필요한 경우 (물품 구매 등)
            - **키워드:** "선정릉점 리뉴얼 상품 공급"
            - **출력 JSON:**
              ```json
              {
                "title": "선정릉점 리뉴얼에 따른 상품 공급의 건",
                "purpose": "당 본부에서는 선정릉점 리뉴얼에 따른 상품 공급을 아래와 같이 진행하였기에 보고드리오니 검토 후 재가 부탁드립니다.",
                "items": [
                  {"No": 1, "거래처": "라온", "품목": "35박스 냉동고", "매입금액": 1298000, "가맹공급금액": 1394800, "비고": "배송/설치비 포함"},
                  {"No": 2, "거래처": "카멜", "품목": "DID 모니터", "매입금액": 1642000, "가맹공급금액": 1768200, "비고": "배송/설치비 포함"}
                ],
                "remarks": "1. 대금결제방식\\n  1) 라온 : 세금계산서 수취 후 10월 5일 결제\\n  2) 카멜 : 법인카드 결제"
              }
              ```

            #### 예시 2: 서술이 필요한 경우 (정책 변경 등)
            - **키워드:** "신규 브랜드 로스율 조정"
            - **출력 JSON:**
              ```json
              {
                "title": "신규 브랜드 기본 로스율 조정 품의",
                "purpose": "신규브랜드 런칭에 따라 안정적인 매출을 위해 기본 로스율을 조정하여 중간관리자 부담을 완화 하고자 함.",
                "body": "### 1. 현상황\\n1) 제품 판매가격 대비 매출 저조로 인해 소극적인 운영이 불가피함.\\n2) 중간관리자 로스부담액 과다로 인해 매장 내 제품 구색이 떨어짐.\\n\\n### 2. 조정 방안\\n- 기본 로스율 조정: **3% → 5%**\\n- 단, 강남점은 매출금액과 운영기간을 반영해 **4%**로 조정함.\\n\\n### 3. 추후 대처 방안\\n1) SNS 마케팅을 통한 브랜드 인지도 향상\\n2) 브랜드 안정화 이후 로스율 재조정",
                "remarks": "브랜드의 성공적인 시장 안착을 위한 한시적 조정임."
              }
              ```

            ### 최종 지시
            이제 사용자의 키워드를 분석하여, 위 규칙과 예시 스타일에 맞춰 'items'(표) 또는 'body'(줄글) 중 하나를 선택하여 품의서 초안 전체를 JSON 형식으로 생성해주세요. "title", "purpose", "remarks"는 항상 포함되어야 합니다.
            """,
            "user": f"핵심 키워드: '{keywords}'"
        },
        "공지문": {
            "system": "당신은 한국 기업의 사내 커뮤니케이션 담당자입니다. 사용자의 키워드를 바탕으로, 번호 매기기(`1.`, `1)`)와 줄바꿈을 명확히 사용한 '사내 공지문' 초안을 생성합니다. 응답은 'title', 'target', 'summary', 'details', 'contact' key를 포함하는 JSON 형식이어야 하며, 'details'는 마크다운 형식으로 작성해주세요.",
            "user": f"핵심 키워드: '{keywords}'"
        },
        "공문": {
            "system": "당신은 대외 문서를 담당하는 총무팀 직원입니다. 사용자의 키워드를 바탕으로, '- 아 래 -' 형식과 번호 매기기를 사용하여 격식에 맞는 '공문' 초안을 생성합니다. 응답은 'sender_org', 'receiver', 'cc', 'title', 'body', 'sender_name' key를 포함하는 JSON 형식이어야 하며, 'body'는 마크다운 형식으로 작성해주세요.",
            "user": f"핵심 키워드: '{keywords}'"
        },
        "비즈니스 이메일": {
            "system": "당신은 비즈니스 커뮤니케이션 전문가입니다. 사용자의 키워드를 바탕으로, 줄바꿈과 가독성을 고려한 전문적인 '비즈니스 이메일' 초안을 생성합니다. 응답은 'to', 'cc', 'subject', 'intro', 'body', 'closing' key를 포함하는 JSON 형식이어야 하며, 'body'는 마크다운 형식으로 작성해주세요.",
            "user": f"핵심 키워드: '{keywords}'"
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

def md_to_html(text):
    return markdown.markdown(text, extensions=['fenced_code', 'tables'])

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

if "ai_draft" not in st.session_state: st.session_state.ai_draft = {}
if "final_html" not in st.session_state: st.session_state.final_html = ""
    
if st.button(f"AI로 {doc_type} 전체 생성하기", type="primary", use_container_width=True):
    if keywords:
        with st.spinner(f"AI가 {doc_type} 전체를 작성 중입니다..."):
            ai_result = generate_ai_draft(doc_type, keywords)
            if ai_result:
                st.session_state.ai_draft = ai_result
                st.session_state.final_html = ""
                st.success("AI가 문서 초안을 모두 작성했습니다. 아래 내용을 확인하고 수정하세요.")
    else:
        st.warning("핵심 키워드를 입력해주세요.")

if st.session_state.ai_draft:
    st.divider()
    draft = st.session_state.ai_draft

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
            p_data["body_edited"] = st.text_area("내용 (마크다운 지원)", value=p_data.get("body", ""), height=200)
            p_data["df_edited"] = pd.DataFrame()

        p_data["remarks"] = st.text_area("비고 및 참고사항 (마크다운 지원)", value=p_data.get("remarks", ""), height=150)

        if st.button("미리보기 생성", use_container_width=True):
            context = { "title": p_data["title"], "purpose": md_to_html(p_data["purpose"]), "remarks": md_to_html(p_data["remarks"]), "generation_date": datetime.now().strftime('%Y-%m-%d') }
            if not p_data["df_edited"].empty:
                context["table_headers"] = list(p_data["df_edited"].columns)
                context["items"] = p_data["df_edited"].to_dict('records')
            else:
                context["body"] = md_to_html(p_data["body_edited"])
            
            template = load_template('pumui_template_final.html')
            st.session_state.final_html = generate_html(template, context)

    # ... 이하 공지문, 공문, 이메일 양식 코드 ...
    elif doc_type == '공지문':
        g_data = draft
        g_data["title"] = st.text_input("제목", value=g_data.get("title", ""))
        g_data["target"] = st.text_input("대상", value=g_data.get("target", ""))
        g_data["summary"] = st.text_area("핵심 요약", value=g_data.get("summary", ""), height=100)
        g_data["details"] = st.text_area("상세 내용 (마크다운 지원)", value=g_data.get("details", ""), height=200)
        g_data["contact"] = st.text_input("문의처", value=g_data.get("contact", ""))
        
        if st.button("미리보기 생성", use_container_width=True):
            context = { "title": g_data["title"], "target": g_data["target"], "summary": md_to_html(g_data["summary"]), "details": md_to_html(g_data["details"]), "contact": g_data["contact"] }
            context["generation_date"] = datetime.now().strftime('%Y. %m. %d.')
            template = load_template('gongji_template.html')
            st.session_state.final_html = generate_html(template, context)

    elif doc_type == '공문':
        gm_data = draft
        gm_data["sender_org"] = st.text_input("발신 기관명", value=gm_data.get("sender_org", ""))
        gm_data["receiver"] = st.text_input("수신", value=gm_data.get("receiver", ""))
        gm_data["cc"] = st.text_input("참조", value=gm_data.get("cc", ""))
        gm_data["title"] = st.text_input("제목", value=gm_data.get("title", ""))
        gm_data["body"] = st.text_area("내용 (마크다운 지원)", value=gm_data.get("body", ""), height=250)
        gm_data["sender_name"] = st.text_input("발신 명의", value=gm_data.get("sender_name", ""))

        if st.button("미리보기 생성", use_container_width=True):
            context = { "sender_org": gm_data["sender_org"], "receiver": gm_data["receiver"], "cc": gm_data["cc"], "title": gm_data["title"], "body": md_to_html(gm_data["body"]), "sender_name": gm_data["sender_name"] }
            context["generation_date"] = datetime.now().strftime('%Y. %m. %d.')
            template = load_template('gongmun_template.html')
            st.session_state.final_html = generate_html(template, context)

    elif doc_type == '비즈니스 이메일':
        e_data = draft
        e_data["to"] = st.text_input("받는 사람", value=e_data.get("to", ""))
        e_data["cc"] = st.text_input("참조", value=e_data.get("cc", ""))
        e_data["subject"] = st.text_input("제목", value=e_data.get("subject", ""))
        e_data["intro"] = st.text_area("도입 (마크다운 지원)", value=e_data.get("intro", ""), height=100)
        e_data["body"] = st.text_area("본론 (마크다운 지원)", value=e_data.get("body", ""), height=150)
        e_data["closing"] = st.text_area("결론 (마크다운 지원)", value=e_data.get("closing", ""), height=100)

        with st.expander("내 서명 정보 입력/수정"):
            e_data["signature_name"] = st.text_input("이름", value="홍길동")
            e_data["signature_title"] = st.text_input("직책", value="대리")
            e_data["signature_team"] = st.text_input("부서/팀", value="마케팅팀")

        if st.button("이메일 본문 생성", use_container_width=True):
            e_data["signature_company"] = "주식회사 몬쉘코리아"
            context = { "to": e_data["to"], "cc": e_data["cc"], "subject": e_data["subject"], "intro": md_to_html(e_data["intro"]), "body": md_to_html(e_data["body"]), "closing": md_to_html(e_data["closing"]), "signature_name": e_data["signature_name"], "signature_title": e_data["signature_title"], "signature_team": e_data["signature_team"], "signature_company": e_data["signature_company"] }
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

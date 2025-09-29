import streamlit as st
import pandas as pd
from jinja2 import Environment, FileSystemLoader
from weasyprint import HTML, CSS
from datetime import datetime
import streamlit.components.v1 as components
from openai import OpenAI
import json
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import re

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

            ### 문서 작성 규칙 (반드시 준수)
            1.  **번호 매기기 상세 규칙:** 본문 항목 구분 시 `1. 첫째 수준`, `  1) 둘째 수준`, `    (1) 셋째 수준` 의 위계질서와 들여쓰기를 일반 텍스트 형식으로 완벽하게 준수합니다. `#` 과 같은 마크다운 제목 기호는 절대로 사용하지 마세요.
            2.  **가독성:** 의미 단위로 명확하게 줄을 바꾸고(`\\n` 사용), 문장은 간결하게 작성합니다.
            3.  **출력 형식:** 'items'(표) 또는 'body'(줄글) 중 하나를 선택하여 품의서 초안 전체를 JSON 형식으로 생성해주세요. "title", "purpose", "remarks"는 항상 포함되어야 합니다.
            """,
            "user": f"핵심 키워드: '{keywords}'"
        },
        "공지문": { "system": "당신은 한국 기업의 사내 커뮤니케이션 담당자입니다. 사용자의 키워드를 바탕으로, `1.`, `  1)` 등 일반 텍스트 형식의 번호 매기기와 줄바꿈을 명확히 사용한 '사내 공지문' 초안을 생성합니다. 응답은 'title', 'target', 'summary', 'details', 'contact' key를 포함하는 JSON 형식이어야 합니다. `#` 기호는 사용하지 마세요.", "user": f"핵심 키워드: '{keywords}'" },
        "공문": { "system": "당신은 대외 문서를 담당하는 총무팀 직원입니다. 사용자의 키워드를 바탕으로, '- 아 래 -' 형식과 `1.`, `  1)` 등 일반 텍스트 형식의 번호 매기기를 사용하여 격식에 맞는 '공문' 초안을 생성합니다. 응답은 'sender_org', 'receiver', 'cc', 'title', 'body', 'sender_name' key를 포함하는 JSON 형식이어야 합니다. `#` 기호는 사용하지 마세요.", "user": f"핵심 키워드: '{keywords}'" },
        "비즈니스 이메일": { "system": "당신은 비즈니스 커뮤니케이션 전문가입니다. 사용자의 키워드를 바탕으로, 줄바꿈과 가독성을 고려한 전문적인 '비즈니스 이메일' 초안을 생성합니다. 응답은 'to', 'cc', 'subject', 'intro', 'body', 'closing' key를 포함하는 JSON 형식이어야 합니다. `#` 기호는 사용하지 마세요.", "user": f"핵심 키워드: '{keywords}'" }
    }

    try:
        response = client.chat.completions.create(model="gpt-4o-mini", response_format={"type": "json_object"}, messages=[{"role": "system", "content": prompts[doc_type]["system"]}, {"role": "user", "content": prompts[doc_type]["user"]}], temperature=0.7)
        return json.loads(response.choices[0].message.content)
    except Exception as e:
        st.error(f"AI 생성 중 오류가 발생했습니다: {e}")
        return None

# --- 텍스트 및 문서 변환 함수들 ---
def clean_text(text):
    """AI가 생성한 텍스트에서 불필요한 마크다운 기호를 제거하고 정리합니다."""
    if not isinstance(text, str):
        return ""
    # `#`로 시작하는 제목 형식 제거
    text = re.sub(r'^\s*#+\s*', '', text, flags=re.MULTILINE)
    # `-`, `*` 로 시작하는 목록 기호는 유지하되, 일관성을 위해 `-`로 통일
    text = re.sub(r'^\s*\*\s*', '  - ', text, flags=re.MULTILINE)
    return text

def text_to_html(text):
    """정리된 텍스트를 HTML 형식으로 변환합니다."""
    return clean_text(text).replace('\n', '<br>')

def generate_pdf(html_content):
    font_css = CSS(string="@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@400;700&display=swap');")
    return HTML(string=html_content).write_pdf(stylesheets=[font_css])

def generate_docx(draft_data, doc_type):
    doc = Document()
    # (문서 생성 로직에 clean_text 함수 적용)
    if doc_type == '품의서':
        doc.add_heading(draft_data.get('title', '제목 없음'), level=1).alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph(clean_text(draft_data.get('purpose', '')))
        doc.add_paragraph("- 아 래 -").alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        doc.add_heading("1. 상세 내역", level=2)
        if "items" in draft_data and draft_data["items"]:
            df = pd.DataFrame(draft_data["items"])
            if not df.empty:
                table = doc.add_table(rows=1, cols=len(df.columns), style='Table Grid')
                hdr_cells = table.rows[0].cells
                for i, col_name in enumerate(df.columns):
                    hdr_cells[i].text = col_name
                for index, row in df.iterrows():
                    row_cells = table.add_row().cells
                    for i, col_name in enumerate(df.columns):
                        row_cells[i].text = str(row[col_name])
        elif "body" in draft_data:
            doc.add_paragraph(clean_text(draft_data.get('body', '')))

        doc.add_heading("2. 비고", level=2)
        doc.add_paragraph(clean_text(draft_data.get('remarks', '')))
        p_end = doc.add_paragraph("끝.")
        p_end.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    elif doc_type == '공지문':
        doc.add_heading(draft_data.get('title', '제목 없음'), level=1).alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph(f"대상: {draft_data.get('target', '')}")
        doc.add_paragraph(f"핵심 요약: {draft_data.get('summary', '')}")
        doc.add_paragraph("-" * 20)
        doc.add_paragraph(clean_text(draft_data.get('details', '')))
        doc.add_paragraph(f"\n문의: {draft_data.get('contact', '')}")
    
    # (공문, 이메일도 동일하게 clean_text 적용)
    elif doc_type == '공문':
        doc.add_heading("공 식 문 서", level=1).alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph(f"발신: {draft_data.get('sender_org', '')}")
        doc.add_paragraph(f"수신: {draft_data.get('receiver', '')}")
        doc.add_paragraph(f"참조: {draft_data.get('cc', '')}")
        doc.add_paragraph("-" * 20)
        doc.add_paragraph(f"제목: {draft_data.get('title', '')}")
        doc.add_paragraph(clean_text(draft_data.get('body', '')))
        p = doc.add_paragraph(f"\n\n{draft_data.get('sender_name', '')}")
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    elif doc_type == '비즈니스 이메일':
        doc.add_paragraph(f"받는 사람: {draft_data.get('to', '')}")
        doc.add_paragraph(f"참조: {draft_data.get('cc', '')}")
        doc.add_paragraph(f"제목: {draft_data.get('subject', '')}")
        doc.add_paragraph("-" * 20)
        doc.add_paragraph(clean_text(draft_data.get('intro', '')))
        doc.add_paragraph(clean_text(draft_data.get('body', '')))
        doc.add_paragraph(clean_text(draft_data.get('closing', '')))

    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# --- 기본 앱 설정 ---
st.set_page_config(page_title="문서 작성 도우미", layout="wide")
env = Environment(loader=FileSystemLoader('.'))
def load_template(template_name): return env.get_template(template_name)
def generate_html(template, context): return template.render(context)

st.sidebar.title("📑 문서 종류 선택")
doc_type = st.sidebar.radio("작성할 문서의 종류를 선택하세요.", ('품의서', '공지문', '공문', '비즈니스 이메일'), key="doc_type_selector")

draft_key = f"draft_{doc_type}"
html_key = f"html_{doc_type}"

if draft_key not in st.session_state: st.session_state[draft_key] = {}
if html_key not in st.session_state: st.session_state[html_key] = ""

st.title(f"✍️ AI {doc_type} 자동 생성")
st.markdown(f"핵심 키워드를 입력하면, AI가 **'{doc_type}'** 초안 전체를 자동으로 작성해줍니다.")
st.divider()

keyword_examples = { "품의서": "표 생성 예시: '영업팀 태블릿 5대 구매' / 줄글 생성 예시: '사내 휴게공간 개선 건의'", "공지문": "예: 10월 전사 워크숍, 제주도, 1박 2일", "공문": "예: A사에 신제품 기술 자료 요청", "비즈니스 이메일": "예: 박부장님께, 4분기 회의 일정 조율 요청" }
keywords = st.text_input("핵심 키워드", placeholder=keyword_examples.get(doc_type, ""))

col1, col2 = st.columns([3, 1])
with col1:
    if st.button(f"AI로 {doc_type} 전체 생성하기", type="primary", use_container_width=True):
        if keywords:
            with st.spinner(f"AI가 {doc_type} 전체를 작성 중입니다..."):
                ai_result = generate_ai_draft(doc_type, keywords)
                if ai_result:
                    st.session_state[draft_key] = ai_result
                    st.session_state[html_key] = ""
                    st.success("AI가 문서 초안을 모두 작성했습니다. 아래 내용을 확인하고 수정하세요.")
        else:
            st.warning("핵심 키워드를 입력해주세요.")
with col2:
    if st.button("새 문서 작성 (양식 초기화)"):
        st.session_state[draft_key] = {}
        st.session_state[html_key] = ""
        st.rerun()

st.divider()

draft = st.session_state.get(draft_key, {})

if draft:
    preview_button = False
    if doc_type == '품의서':
        p_data = draft
        p_data["title"] = st.text_input("제목", value=p_data.get("title", ""))
        p_data["purpose"] = st.text_area("목적 및 개요", value=clean_text(p_data.get("purpose", "")), height=100)
        if "items" in p_data and p_data["items"]:
            p_data["df"] = pd.DataFrame(p_data.get("items", []))
            st.subheader("상세 내역 (표)")
            p_data["df_edited"] = st.data_editor(p_data["df"], num_rows="dynamic")
            p_data["body_edited"] = ""
        else:
            st.subheader("상세 내용 (줄글)")
            p_data["body_edited"] = st.text_area("내용", value=clean_text(p_data.get("body", "")), height=200)
            p_data["df_edited"] = pd.DataFrame()
        p_data["remarks"] = st.text_area("비고 및 참고사항", value=clean_text(p_data.get("remarks", "")), height=150)
        preview_button = st.button("미리보기 생성", use_container_width=True)
    
    elif doc_type == '공지문':
        g_data = draft
        g_data["title"] = st.text_input("제목", value=g_data.get("title", ""))
        g_data["target"] = st.text_input("대상", value=g_data.get("target", ""))
        g_data["summary"] = st.text_area("핵심 요약", value=clean_text(g_data.get("summary", "")), height=100)
        g_data["details"] = st.text_area("상세 내용", value=clean_text(g_data.get("details", "")), height=200)
        g_data["contact"] = st.text_input("문의처", value=g_data.get("contact", ""))
        preview_button = st.button("미리보기 생성", use_container_width=True)

    elif doc_type == '공문':
        gm_data = draft
        gm_data["sender_org"] = st.text_input("발신 기관명", value=gm_data.get("sender_org", ""))
        gm_data["receiver"] = st.text_input("수신", value=gm_data.get("receiver", ""))
        gm_data["cc"] = st.text_input("참조", value=gm_data.get("cc", ""))
        gm_data["title"] = st.text_input("제목", value=gm_data.get("title", ""))
        gm_data["body"] = st.text_area("내용", value=clean_text(gm_data.get("body", "")), height=250)
        gm_data["sender_name"] = st.text_input("발신 명의", value=gm_data.get("sender_name", ""))
        preview_button = st.button("미리보기 생성", use_container_width=True)

    elif doc_type == '비즈니스 이메일':
        e_data = draft
        e_data["to"] = st.text_input("받는 사람", value=e_data.get("to", ""))
        e_data["cc"] = st.text_input("참조", value=e_data.get("cc", ""))
        e_data["subject"] = st.text_input("제목", value=e_data.get("subject", ""))
        e_data["intro"] = st.text_area("도입", value=clean_text(e_data.get("intro", "")), height=100)
        e_data["body"] = st.text_area("본론", value=clean_text(e_data.get("body", "")), height=150)
        e_data["closing"] = st.text_area("결론", value=clean_text(e_data.get("closing", "")), height=100)
        with st.expander("내 서명 정보 입력/수정"):
            e_data["signature_name"] = st.text_input("이름", value="홍길동")
            e_data["signature_title"] = st.text_input("직책", value="대리")
            e_data["signature_team"] = st.text_input("부서/팀", value="마케팅팀")
        preview_button = st.button("이메일 본문 생성", use_container_width=True)
    
    if preview_button:
        if doc_type == '품의서':
            context = { "title": p_data["title"], "purpose": text_to_html(p_data["purpose"]), "remarks": text_to_html(p_data["remarks"]), "generation_date": datetime.now().strftime('%Y-%m-%d') }
            if not p_data["df_edited"].empty:
                context["table_headers"] = list(p_data["df_edited"].columns)
                context["items"] = p_data["df_edited"].to_dict('records')
            else:
                context["body"] = text_to_html(p_data["body_edited"])
            template = load_template('pumui_template_final.html')
            st.session_state[html_key] = generate_html(template, context)
        
        elif doc_type == '공지문':
            context = { "title": g_data["title"], "target": g_data["target"], "summary": text_to_html(g_data["summary"]), "details": text_to_html(g_data["details"]), "contact": g_data["contact"] }
            context["generation_date"] = datetime.now().strftime('%Y. %m. %d.')
            template = load_template('gongji_template.html')
            st.session_state[html_key] = generate_html(template, context)

        elif doc_type == '공문':
            context = { "sender_org": gm_data["sender_org"], "receiver": gm_data["receiver"], "cc": gm_data["cc"], "title": gm_data["title"], "body": text_to_html(gm_data["body"]), "sender_name": gm_data["sender_name"] }
            context["generation_date"] = datetime.now().strftime('%Y. %m. %d.')
            template = load_template('gongmun_template.html')
            st.session_state[html_key] = generate_html(template, context)

        elif doc_type == '비즈니스 이메일':
            e_data["signature_company"] = "주식회사 몬쉘코리아"
            context = { "to": e_data["to"], "cc": e_data["cc"], "subject": e_data["subject"], "intro": text_to_html(e_data["intro"]), "body": text_to_html(e_data["body"]), "closing": text_to_html(e_data["closing"]), "signature_name": e_data["signature_name"], "signature_title": e_data["signature_title"], "signature_team": e_data["signature_team"], "signature_company": e_data["signature_company"] }
            template = load_template('email_template_final.html')
            st.session_state[html_key] = generate_html(template, context)
        
        st.rerun()

if st.session_state[html_key]:
    st.divider()
    st.subheader("📄 최종 미리보기")
    components.html(st.session_state[html_key], height=600, scrolling=True)

    if doc_type == "비즈니스 이메일":
        st.subheader("📋 복사할 HTML 코드")
        st.code(st.session_state[html_key], language='html')
    else:
        st.divider()
        col1, col2 = st.columns(2)
        with col1:
            pdf_output = generate_pdf(st.session_state[html_key])
            title_for_file = draft.get("title", "document")
            st.download_button(label="📥 PDF 파일로 다운로드", data=pdf_output, file_name=f"{title_for_file}.pdf", mime="application/pdf", use_container_width=True)
        with col2:
            docx_output = generate_docx(draft, doc_type)
            st.download_button(label="📄 Word 파일로 다운로드", data=docx_output, file_name=f"{title_for_file}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)

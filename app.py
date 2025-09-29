import streamlit as st
import pandas as pd
from jinja2 import Environment, FileSystemLoader
from weasyprint import HTML, CSS
from datetime import datetime
import streamlit.components.v1 as components
from openai import OpenAI
import json
import markdown
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io

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
            1.  **번호 매기기 상세 규칙:** 본문 항목 구분 시 `1. 첫째 수준`, `  1) 둘째 수준`, `    (1) 셋째 수준` 의 위계질서와 들여쓰기를 마크다운 문법에 맞춰 완벽하게 준수합니다. 이 규칙을 벗어난 번호 매기기는 절대로 사용하지 마세요.
            2.  **가독성:** 의미 단위로 명확하게 줄을 바꾸고, 문장은 간결하게 작성합니다.
            3.  **내용:** 결론을 먼저 제시하고, 이유나 상세 설명을 뒤에 붙이는 두괄식 구성을 선호합니다.
            4.  **종결:** 본문이 끝나면 "**끝.**" 표시를 사용합니다.
            5.  **출력 형식:** 본문(`body`) 내용은 반드시 마크다운(Markdown) 형식으로 작성해야 합니다.

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
                "body": "#### 1. 현상황\\n1) 제품 판매가격 대비 매출 저조로 인해 소극적인 운영이 불가피함.\\n2) 중간관리자 로스부담액 과다로 인해 매장 내 제품 구색이 떨어짐.\\n\\n#### 2. 조정 방안\\n- 기본 로스율 조정: **3% → 5%**\\n- 단, 강남점은 매출금액과 운영기간을 반영해 **4%**로 조정함.\\n\\n#### 3. 추후 대처 방안\\n1) SNS 마케팅을 통한 브랜드 인지도 향상\\n2) 브랜드 안정화 이후 로스율 재조정\\n\\n끝.",
                "remarks": "브랜드의 성공적인 시장 안착을 위한 한시적 조정임."
              }
              ```

            ### 최종 지시
            이제 사용자의 키워드를 분석하여, 위 규칙과 예시 스타일에 맞춰 'items'(표) 또는 'body'(줄글) 중 하나를 선택하여 품의서 초안 전체를 JSON 형식으로 생성해주세요. "title", "purpose", "remarks"는 항상 포함되어야 합니다.
            """,
            "user": f"핵심 키워드: '{keywords}'"
        },
        # (다른 문서 타입 프롬프트도 번호 매기기 규칙 강화)
        "공지문": { "system": "당신은 한국 기업의 사내 커뮤니케이션 담당자입니다. 사용자의 키워드를 바탕으로, `1.`, `  1)` 등 마크다운 형식의 번호 매기기와 줄바꿈을 명확히 사용한 '사내 공지문' 초안을 생성합니다. 응답은 'title', 'target', 'summary', 'details', 'contact' key를 포함하는 JSON 형식이어야 하며, 'details'는 마크다운 형식으로 작성해주세요.", "user": f"핵심 키워드: '{keywords}'" },
        "공문": { "system": "당신은 대외 문서를 담당하는 총무팀 직원입니다. 사용자의 키워드를 바탕으로, '- 아 래 -' 형식과 `1.`, `  1)` 등 마크다운 형식의 번호 매기기를 사용하여 격식에 맞는 '공문' 초안을 생성합니다. 응답은 'sender_org', 'receiver', 'cc', 'title', 'body', 'sender_name' key를 포함하는 JSON 형식이어야 하며, 'body'는 마크다운 형식으로 작성해주세요.", "user": f"핵심 키워드: '{keywords}'" },
        "비즈니스 이메일": { "system": "당신은 비즈니스 커뮤니케이션 전문가입니다. 사용자의 키워드를 바탕으로, 줄바꿈과 가독성을 고려한 전문적인 '비즈니스 이메일' 초안을 생성합니다. 응답은 'to', 'cc', 'subject', 'intro', 'body', 'closing' key를 포함하는 JSON 형식이어야 하며, 'body'는 마크다운 형식으로 작성해주세요.", "user": f"핵심 키워드: '{keywords}'" }
    }
    # ... (이전과 동일한 AI 호출 로직)
    try:
        response = client.chat.completions.create(model="gpt-4o-mini", response_format={"type": "json_object"}, messages=[{"role": "system", "content": prompts[doc_type]["system"]}, {"role": "user", "content": prompts[doc_type]["user"]}], temperature=0.7)
        return json.loads(response.choices[0].message.content)
    except Exception as e:
        st.error(f"AI 생성 중 오류가 발생했습니다: {e}")
        return None

# --- 문서 변환 함수들 ---
def md_to_html(text):
    return markdown.markdown(text, extensions=['fenced_code', 'tables'])

def generate_pdf(html_content):
    font_config = CSS(string="@font-face { font-family: 'Noto Sans KR'; src: url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@400;700&display=swap'); } body { font-family: 'Noto Sans KR', sans-serif; }")
    return HTML(string=html_content).write_pdf(stylesheets=[font_config])

def generate_docx(draft_data, doc_type):
    doc = Document()
    # (여기에 각 문서 타입별로 docx를 생성하는 상세 로직 추가)
    # 예시: 품의서
    if doc_type == '품의서':
        doc.add_heading(draft_data.get('title', '제목 없음'), level=1).alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph(draft_data.get('purpose', ''))
        doc.add_paragraph("- 아 래 -").alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # 표 또는 본문 처리
        if "items" in draft_data and draft_data["items"]:
            df = pd.DataFrame(draft_data["items"])
            table = doc.add_table(rows=1, cols=len(df.columns))
            table.style = 'Table Grid'
            hdr_cells = table.rows[0].cells
            for i, col_name in enumerate(df.columns):
                hdr_cells[i].text = col_name
            for index, row in df.iterrows():
                row_cells = table.add_row().cells
                for i, col_name in enumerate(df.columns):
                    row_cells[i].text = str(row[col_name])
        elif "body" in draft_data:
            doc.add_paragraph(draft_data.get('body', ''))

        doc.add_paragraph("비고").bold = True
        doc.add_paragraph(draft_data.get('remarks', ''))
        doc.add_paragraph("끝.").alignment = WD_ALIGN_PARAGRAPH.RIGHT

    # (다른 문서 타입에 대한 docx 생성 로직도 유사하게 추가)
    
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# --- 기본 앱 설정 ---
st.set_page_config(page_title="문서 작성 도우미", layout="wide")
env = Environment(loader=FileSystemLoader('.'))

def load_template(template_name):
    return env.get_template(template_name)
def generate_html(template, context):
    return template.render(context)

st.sidebar.title("📑 문서 종류 선택")
doc_type = st.sidebar.radio("작성할 문서의 종류를 선택하세요.", ('품의서', '공지문', '공문', '비즈니스 이메일'), key="doc_type_selector")

# --- 상태 관리 키 생성 ---
draft_key = f"draft_{doc_type}"
html_key = f"html_{doc_type}"

if draft_key not in st.session_state: st.session_state[draft_key] = {}
if html_key not in st.session_state: st.session_state[html_key] = ""

st.title(f"✍️ AI {doc_type} 자동 생성")
st.markdown(f"핵심 키워드를 입력하면, AI가 **'{doc_type}'** 초안 전체를 자동으로 작성해줍니다.")
st.divider()

keyword_examples = { "품의서": "표 생성 예시: '영업팀 태블릿 5대 구매' / 줄글 생성 예시: '사내 휴게공간 개선 건의'", "공지문": "예: 10월 전사 워크숍, 제주도, 1박 2일", "공문": "예: A사에 신제품 기술 자료 요청", "비즈니스 이메일": "예: 박부장님께, 4분기 회의 일정 조율 요청" }
keywords = st.text_input("핵심 키워드", placeholder=keyword_examples.get(doc_type, ""))

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

st.divider()

draft = st.session_state[draft_key]

if draft:
    # (각 문서 타입별 UI 생성 로직 - 이전과 유사하지만 draft_key 사용)
    if doc_type == '품의서':
        # ... UI 로직 ...
        if st.button("미리보기 생성", use_container_width=True):
            # ... context 생성 로직 ...
            template = load_template('pumui_template_final.html')
            st.session_state[html_key] = generate_html(template, context)
            
    # (다른 문서 타입 UI 로직)

# --- 미리보기 및 다운로드 ---
if st.session_state[html_key]:
    st.divider()
    st.subheader("📄 최종 미리보기")
    components.html(st.session_state[html_key], height=600, scrolling=True)

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("📥 PDF 다운로드")
        pdf_output = generate_pdf(st.session_state[html_key])
        title_for_file = draft.get("title", "document")
        st.download_button(label="PDF 파일로 다운로드", data=pdf_output, file_name=f"{title_for_file}.pdf", mime="application/pdf", use_container_width=True)
    with col2:
        st.subheader("📥 Word 파일 다운로드")
        docx_output = generate_docx(draft, doc_type) # draft 데이터를 직접 사용
        st.download_button(label="Word 파일로 다운로드", data=docx_output, file_name=f"{title_for_file}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)

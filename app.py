import streamlit as st
import pandas as pd
from jinja2 import Environment, FileSystemLoader
from weasyprint import HTML
import base64
from datetime import datetime

# --- 초기 설정 ---
st.set_page_config(page_title="문서 작성 도우미", layout="wide")

# Jinja2 템플릿 환경 설정
env = Environment(loader=FileSystemLoader('.'))

def load_template(template_name):
    """지정된 이름의 Jinja2 템플릿을 로드합니다."""
    return env.get_template(template_name)

def generate_html(template, context):
    """템플릿과 데이터를 결합하여 HTML을 생성합니다."""
    return template.render(context)

def generate_pdf(html_content):
    """HTML 내용을 PDF로 변환합니다."""
    return HTML(string=html_content).write_pdf()

# --- 사이드바 (템플릿 선택) ---
st.sidebar.title("📑 문서 종류 선택")
doc_type = st.sidebar.radio(
    "작성할 문서의 종류를 선택하세요.",
    ('품의서', '공지문 (준비 중)', '공문 (준비 중)', '비즈니스 이메일 (준비 중)'),
    label_visibility="collapsed"
)

# --- 메인 화면 ---
st.title("✍️ 지능형 문서 작성 도우미")
st.markdown("""
제공해주신 '문서 작성 매뉴얼'과 '품의서 샘플'을 기반으로 제작된 어플리케이션입니다.  
좌측 사이드바에서 문서 종류를 선택하고, 아래 양식에 내용을 입력하면 표준 서식의 문서가 자동으로 생성됩니다.
""")
st.divider()


# --- 품의서 작성 양식 ---
if doc_type == '품의서':
    st.header("품의서 작성")
    
    # 세션 상태를 사용하여 입력 데이터 유지
    if 'pumui_data' not in st.session_state:
        st.session_state.pumui_data = {
            "title": "선릉점 리뉴얼에 따른 상품 공급의 건",
            "purpose": "선정릉점 리뉴얼에 따른 상품 공급을 아래와 같이 진행하였기에 보고드리오니 검토 후 재가 부탁드립니다.",
            "remarks": "1. 대금결제방식\n  1) 라온 : 세금계산서 수취 후 10월 5일 결제\n  2) 카멜 : 법인카드 결제\n\n2. 특이사항\n  - 공급 물품에 10% 마진 설정, 배송/설치비에는 본사마진 없음",
            "items_df": pd.DataFrame([
                {"No": 1, "거래처": "라온", "품목": "35박스 냉동고", "매입금액": 1298000, "가맹공급금액": 1394800, "비고": "배송/설치비 포함"},
                {"No": 2, "거래처": "카멜", "품목": "DID 모니터", "매입금액": 1642000, "가맹공급금액": 1768200, "비고": "배송/설치비 포함"},
            ])
        }

    p_data = st.session_state.pumui_data

    # 1. 기본 정보 입력
    with st.container(border=True):
        st.subheader("1. 기본 정보")
        p_data["title"] = st.text_input(
            "제목",
            value=p_data["title"],
            help="문서의 핵심 내용이 한눈에 파악되도록 명확하게 작성하세요. (예: OOOO 진행의 건)"
        )
        p_data["purpose"] = st.text_area(
            "1. 목적 및 개요",
            value=p_data["purpose"],
            height=100,
            help="결재자가 '이 보고의 목적이 무엇인가?'라는 의문을 갖지 않도록 핵심 내용을 명료하게 작성하십시오. (문서작성매뉴얼.PDF 참고)"
        )

    # 2. 상세 내역 입력 (테이블)
    with st.container(border=True):
        st.subheader("2. 상세 내역")
        st.info("아래 표를 엑셀처럼 자유롭게 수정, 추가, 삭제할 수 있습니다.")
        
        # 사용자가 표를 직접 수정할 수 있는 st.data_editor 사용
        edited_df = st.data_editor(
            p_data["items_df"],
            num_rows="dynamic", # 행 추가/삭제 가능
            key="pumui_editor"
        )
        p_data["items_df"] = edited_df

    # 3. 추가 정보 입력
    with st.container(border=True):
        st.subheader("3. 비고 및 참고사항")
        p_data["remarks"] = st.text_area(
            "비고",
            value=p_data["remarks"],
            height=150,
            help="결제 조건, 계약 정보, 특이사항 등 의사결정에 필요한 추가 정보를 기입합니다."
        )

    # 4. 문서 생성 및 미리보기
    st.divider()
    if st.button("미리보기 생성 및 PDF 다운로드", type="primary", use_container_width=True):
        with st.spinner("문서를 생성 중입니다..."):
            # 테이블 데이터 가공
            items = p_data["items_df"].to_dict('records')
            
            # 총합계 계산
            total_purchase = p_data["items_df"]['매입금액'].sum()
            total_supply = p_data["items_df"]['가맹공급금액'].sum()

            # 템플릿에 전달할 데이터 (Context)
            context = {
                "title": p_data["title"],
                "purpose": p_data["purpose"].replace('\n', '<br>'),
                "items": items,
                "total_purchase": f"{total_purchase:,.0f}",
                "total_supply": f"{total_supply:,.0f}",
                "remarks": p_data["remarks"].replace('\n', '<br>'),
                "generation_date": datetime.now().strftime('%Y-%m-%d')
            }

            # HTML 생성
            template = load_template('pumui_template.html')
            html_output = generate_html(template, context)

            # PDF 생성
            pdf_output = generate_pdf(html_output)

            # 결과 표시
            st.success("🎉 문서 생성이 완료되었습니다!")
            
            # 미리보기
            with st.container(border=True):
                st.subheader("📄 문서 미리보기")
                st.markdown(html_output, unsafe_allow_html=True)
            
            # PDF 다운로드 버튼
            st.download_button(
                label="📥 PDF 파일 다운로드",
                data=pdf_output,
                file_name=f"{p_data['title']}.pdf",
                mime="application/pdf",
                use_container_width=True
            )
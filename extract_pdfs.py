#!/usr/bin/env python3
import PyPDF2
import json
import os

def extract_pdf_content(filename):
    """PDF 파일에서 텍스트 내용을 추출합니다."""
    try:
        with open(filename, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            text = ''
            for page in pdf_reader.pages:
                page_text = page.extract_text()
                if page_text:
                    text += page_text + '\n'
            return text.strip()
    except Exception as e:
        return f'Error reading {filename}: {str(e)}'

def main():
    """메인 함수"""
    # PDF 파일들 읽기
    manual_content = extract_pdf_content('문서작성메뉴얼.PDF')
    samples_content = extract_pdf_content('유제욱 품의서 모집.pdf')
    
    # 추출된 내용을 JSON 파일로 저장
    learned_content = {
        'manual': {
            'filename': '문서작성메뉴얼.PDF',
            'content': manual_content,
            'length': len(manual_content)
        },
        'samples': {
            'filename': '유제욱 품의서 모집.pdf', 
            'content': samples_content,
            'length': len(samples_content)
        },
        'learned_at': '2025-01-01'  # 학습 시점 기록
    }
    
    # learned_documents.json 파일로 저장
    with open('learned_documents.json', 'w', encoding='utf-8') as f:
        json.dump(learned_content, f, ensure_ascii=False, indent=2)
    
    print("PDF 학습 완료!")
    print(f"문서작성메뉴얼: {len(manual_content)} 문자")
    print(f"품의서 모집: {len(samples_content)} 문자")
    print("learned_documents.json 파일로 저장되었습니다.")
    
    # 미리보기 출력
    print("\n=== 문서작성메뉴얼 미리보기 ===")
    print(manual_content[:500] + "..." if len(manual_content) > 500 else manual_content)
    
    print("\n=== 품의서 모집 미리보기 ===") 
    print(samples_content[:500] + "..." if len(samples_content) > 500 else samples_content)

if __name__ == "__main__":
    main()

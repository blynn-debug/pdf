import os
import glob
import pdfplumber


def extract_text_from_pdf(pdf_path):
    """PDF 파일에서 텍스트를 추출하는 함수 (pdfplumber 사용)"""
    try:
        text = ""

        with pdfplumber.open(pdf_path) as pdf:
            # 모든 페이지에서 텍스트 추출
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:  # 텍스트가 있는 경우에만 추가
                    text += page_text + "\n\n"  # 페이지 구분을 위해 줄바꿈 추가

        return text.strip()  # 앞뒤 공백 제거

    except Exception as e:
        print(f"오류 발생 - {pdf_path}: {str(e)}")
        return None


def save_text_to_file(text, output_path):
    """추출된 텍스트를 파일로 저장하는 함수"""
    try:
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(text)
        print(f"저장 완료: {output_path}")

    except Exception as e:
        print(f"저장 오류 - {output_path}: {str(e)}")


def process_all_pdfs():
    """현재 디렉토리의 모든 PDF 파일을 처리하는 함수"""
    # 현재 디렉토리에서 모든 PDF 파일 찾기
    pdf_files = glob.glob("*.pdf")

    if not pdf_files:
        print("현재 디렉토리에 PDF 파일이 없습니다.")
        return

    print(f"총 {len(pdf_files)}개의 PDF 파일을 발견했습니다.")

    for pdf_file in pdf_files:
        print(f"\n처리 중: {pdf_file}")

        # PDF에서 텍스트 추출
        extracted_text = extract_text_from_pdf(pdf_file)

        if extracted_text:
            # 출력 파일명 생성 (확장자를 .txt로 변경)
            output_filename = os.path.splitext(pdf_file)[0] + ".txt"

            # 텍스트 파일로 저장
            save_text_to_file(extracted_text, output_filename)
        else:
            print(f"텍스트 추출 실패: {pdf_file}")


def extract_single_pdf(pdf_filename):
    """단일 PDF 파일을 처리하는 함수 (선택적 사용)"""
    if not os.path.exists(pdf_filename):
        print(f"파일을 찾을 수 없습니다: {pdf_filename}")
        return

    print(f"처리 중: {pdf_filename}")

    extracted_text = extract_text_from_pdf(pdf_filename)

    if extracted_text:
        output_filename = os.path.splitext(pdf_filename)[0] + ".txt"
        save_text_to_file(extracted_text, output_filename)
    else:
        print(f"텍스트 추출 실패: {pdf_filename}")


if __name__ == "__main__":
    print("PDF 스크래핑 시작... (pdfplumber 사용)")

    # 모든 PDF 파일 일괄 처리
    process_all_pdfs()

    # 또는 특정 파일만 처리하려면:
    # extract_single_pdf("example.pdf")

    print("\n모든 작업이 완료되었습니다.")
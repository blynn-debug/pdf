import os
import sys
import shutil
import re
from pathlib import Path
import tkinter as tk
from tkinter import filedialog
import glob

try:
    # COM 객체 캐시 모듈 제거 (문제 해결을 위함)
    gen_py_path = os.path.join(os.environ.get('LOCALAPPDATA', os.path.expanduser('~')),
                               'Temp', 'gen_py')
    if os.path.exists(gen_py_path):
        try:
            shutil.rmtree(gen_py_path)
            print(f"COM 캐시 폴더 삭제 완료: {gen_py_path}")
        except:
            print(f"COM 캐시 폴더 삭제 실패: {gen_py_path}")

    import win32com.client as win32

    WIN32COM_AVAILABLE = True
except ImportError:
    WIN32COM_AVAILABLE = False
    print("win32com이 설치되어 있지 않습니다. 설치해주세요: pip install pywin32")


def connect_to_excel():
    """
    이미 열려있는 엑셀에 연결하는 함수

    반환값:
    (excel_app, workbook) 튜플 - 성공 시
    (None, None) - 실패 시
    """
    try:
        # 기존 실행 중인 Excel 인스턴스 연결
        excel = win32.GetActiveObject("Excel.Application")
        print("기존 Excel 애플리케이션에 연결했습니다.")

        # 현재 활성화된 워크북 사용
        wb = excel.ActiveWorkbook
        if wb:
            print(f"현재 활성화된 워크북을 사용합니다: {wb.Name}")
            return excel, wb
        else:
            print("활성화된 워크북이 없습니다.")
            return None, None

    except Exception as e:
        print(f"Excel 연결 실패: {str(e)}")
        return None, None


def find_field_column(worksheet, field_name):
    """
    워크시트에서 특정 field가 포함된 열을 찾는 함수

    매개변수:
    worksheet: 검색할 워크시트 객체
    field_name: 찾을 field 이름 (예: "field 4", "field 6")

    반환값:
    열 번호 (int) - 찾은 경우, None - 못 찾은 경우
    """
    try:
        # 첫 번째 행부터 10번째 행까지 검색 (보통 헤더는 상단에 위치)
        for row in range(1, 11):
            # A열부터 Z열까지 검색
            for col in range(1, 27):
                cell_value = worksheet.Cells(row, col).Value
                if cell_value and field_name.lower() in str(cell_value).lower():
                    print(f"'{field_name}'을 {row}행 {col}열에서 찾았습니다.")
                    return col

        print(f"'{field_name}'을 찾을 수 없습니다.")
        return None
    except Exception as e:
        print(f"'{field_name}' 검색 중 오류: {str(e)}")
        return None


def find_worksheet(workbook, sheet_name):
    """
    워크북에서 특정 이름을 포함하는 시트를 찾는 함수

    매개변수:
    workbook: 워크북 객체
    sheet_name: 찾을 시트 이름

    반환값:
    워크시트 객체 - 찾은 경우, None - 못 찾은 경우
    """
    try:
        # 정확한 이름으로 먼저 시도
        return workbook.Worksheets(sheet_name)
    except:
        # 실패 시 부분 일치로 검색
        for i in range(1, workbook.Worksheets.Count + 1):
            sheet = workbook.Worksheets(i)
            if sheet_name in sheet.Name:
                return sheet
    return None


def get_field4_field6_mapping(worksheet, field4_col, field6_col):
    """
    설정순위 시트에서 field 4와 field 6의 매핑 데이터를 추출하는 함수

    매개변수:
    worksheet: 설정순위 워크시트 객체
    field4_col: field 4가 있는 열 번호
    field6_col: field 6이 있는 열 번호

    반환값:
    매핑 딕셔너리 {field4_value: field6_value}
    """
    mapping_data = {}

    try:
        # 2번째 행부터 시작 (헤더 제외)
        row = 2
        max_empty_rows = 20  # 연속으로 빈 셀이 20개 나오면 종료
        empty_row_count = 0

        while empty_row_count < max_empty_rows:
            field4_value = worksheet.Cells(row, field4_col).Value
            field6_value = worksheet.Cells(row, field6_col).Value

            if field4_value is None or str(field4_value).strip() == "":
                empty_row_count += 1
            else:
                empty_row_count = 0
                field4_str = str(field4_value).strip()
                field6_str = str(field6_value).strip() if field6_value else ""

                # 매핑 데이터 저장 (중복 제거)
                if field4_str and field4_str not in mapping_data:
                    mapping_data[field4_str] = field6_str
                    print(f"매핑 추가: {field4_str} → {field6_str}")

            row += 1

        print(f"\n총 {len(mapping_data)}개의 매핑을 생성했습니다.")
        return mapping_data

    except Exception as e:
        print(f"매핑 데이터 추출 중 오류: {str(e)}")
        return mapping_data


def extract_matching_key(filename, field4_values):
    """
    파일명에서 field4와 매칭되는 키를 찾는 함수

    매개변수:
    filename: txt 파일명 (예: SBI-A-002-01.txt)
    field4_values: field4 값들의 리스트

    반환값:
    매칭된 field4 값 또는 None
    """
    # 파일명에서 확장자 제거
    base_name = os.path.splitext(filename)[0]

    # 직접 매칭: 파일명이 field4로 시작하는 경우
    for field4 in field4_values:
        if base_name.startswith(field4):
            return field4

    # 파일명에서 마지막 -숫자 부분을 제거하고 매칭
    # 예: SBI-A-002-01 → SBI-A-002
    pattern = r'^(.*)-\d+$'
    match = re.match(pattern, base_name)
    if match:
        prefix = match.group(1)
        if prefix in field4_values:
            return prefix

    # field4 값들을 검사하여 파일명과 매칭되는 것 찾기
    for field4 in field4_values:
        # field4가 파일명에 포함된 경우
        if field4 in base_name:
            return field4

    return None


def process_txt_file(file_path, field4_value, field6_value):
    """
    txt 파일을 처리하여 차주명을 삽입하는 함수

    매개변수:
    file_path: 처리할 txt 파일 경로
    field4_value: 매칭된 field4 값
    field6_value: 차주명 (field6 값)

    반환값:
    성공 여부 (bool)
    """
    try:
        # 원본 파일 백업
        backup_path = file_path + '.backup'
        shutil.copy2(file_path, backup_path)

        # 파일 읽기
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()

        # 이미 차주명이 있는지 확인
        if content.startswith('#차주명'):
            print(f"  - {os.path.basename(file_path)}: 이미 차주명이 있습니다. 건너뜁니다.")
            os.remove(backup_path)  # 백업 파일 삭제
            return False

        # 차주명 추가
        new_content = f"#차주명 : {field6_value}\n{content}"

        # 파일 쓰기
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(new_content)

        print(f"  - {os.path.basename(file_path)}: 차주명 '{field6_value}' 삽입 완료")

        # 백업 파일 삭제 (성공 시)
        os.remove(backup_path)
        return True

    except Exception as e:
        print(f"  - {os.path.basename(file_path)}: 처리 중 오류 발생 - {str(e)}")
        # 오류 발생 시 백업 파일 복원
        if os.path.exists(backup_path):
            shutil.move(backup_path, file_path)
        return False


def update_rights_analysis_field8(workbook, mapping_data):
    """
    권리분석 시트의 field 8 열에 설정순위의 field 6 데이터를 입력하는 함수

    매개변수:
    workbook: 워크북 객체
    mapping_data: {field4: field6} 매핑 딕셔너리

    반환값:
    (성공_개수, 실패_개수) 튜플
    """
    try:
        # 권리분석 시트 찾기
        rights_sheet = find_worksheet(workbook, "권리분석")

        if rights_sheet is None:
            print("'권리분석' 시트를 찾을 수 없습니다.")
            return 0, 0

        print(f"'권리분석' 시트를 찾았습니다: {rights_sheet.Name}")

        # field 5와 field 8 열 찾기
        print("권리분석 시트에서 필드 검색 중...")
        field5_col = find_field_column(rights_sheet, "field 5")
        field8_col = find_field_column(rights_sheet, "field 8")

        if field5_col is None or field8_col is None:
            print("권리분석 시트에서 필요한 필드를 찾을 수 없습니다.")
            return 0, 0

        # 데이터 업데이트
        success_count = 0
        fail_count = 0
        row = 2  # 헤더 다음 행부터 시작
        max_empty_rows = 20
        empty_row_count = 0

        print("\n권리분석 시트 업데이트 시작...")

        while empty_row_count < max_empty_rows:
            field5_value = rights_sheet.Cells(row, field5_col).Value

            if field5_value is None or str(field5_value).strip() == "":
                empty_row_count += 1
            else:
                empty_row_count = 0
                field5_str = str(field5_value).strip()

                # 매핑 데이터에서 field5와 일치하는 field4 찾기
                if field5_str in mapping_data:
                    field6_value = mapping_data[field5_str]

                    try:
                        # field 8에 field 6 값 입력
                        rights_sheet.Cells(row, field8_col).Value = field6_value
                        print(f"  - {row}행: field 5 '{field5_str}' → field 8 '{field6_value}' 입력 완료")
                        success_count += 1
                    except Exception as e:
                        print(f"  - {row}행: 입력 실패 - {str(e)}")
                        fail_count += 1
                else:
                    print(f"  - {row}행: field 5 '{field5_str}'에 대한 매핑 데이터가 없습니다.")
                    fail_count += 1

            row += 1

        print(f"\n권리분석 시트 업데이트 완료: 성공 {success_count}개, 실패 {fail_count}개")
        return success_count, fail_count

    except Exception as e:
        print(f"권리분석 시트 업데이트 중 오류 발생: {str(e)}")
        return 0, 0


def select_folder():
    """
    폴더 선택 대화상자를 표시하는 함수

    반환값:
    선택된 폴더 경로 또는 None
    """
    root = tk.Tk()
    root.withdraw()  # 메인 윈도우 숨기기
    folder_path = filedialog.askdirectory(title="txt 파일이 있는 폴더를 선택하세요")
    root.destroy()
    return folder_path


# 메인 함수
if __name__ == "__main__":
    try:
        print("=" * 60)
        print("TXT 파일 일괄 처리 프로그램")
        print("=" * 60)
        print("txt 파일명과 엑셀 데이터를 매칭하여 차주명을 삽입합니다.")
        print("-" * 60)

        # win32com 사용 가능 여부 확인
        if not WIN32COM_AVAILABLE:
            print("win32com이 설치되어 있지 않습니다. pip install pywin32 명령으로 설치하세요.")
            sys.exit(1)

        # 이미 열린 Excel에 연결
        print("\n1. Excel 애플리케이션 연결 중...")
        excel, wb = connect_to_excel()

        if excel is None or wb is None:
            print("Excel 애플리케이션 연결에 실패했습니다.")
            print("Excel이 실행되어 있고 파일이 열려있는지 확인하세요.")
            sys.exit(1)

        # "설정순위" 시트 찾기
        print("\n2. '설정순위' 시트 검색 중...")
        settings_sheet = find_worksheet(wb, "설정순위")

        if settings_sheet is None:
            print("'설정순위' 시트를 찾을 수 없습니다.")
            sys.exit(1)
        else:
            print(f"'설정순위' 시트를 찾았습니다: {settings_sheet.Name}")

        # field 4와 field 6 열 찾기
        print("\n3. 필드 검색 중...")
        field4_col = find_field_column(settings_sheet, "field 4")
        field6_col = find_field_column(settings_sheet, "field 6")

        if field4_col is None or field6_col is None:
            print("필요한 필드를 찾을 수 없습니다.")
            sys.exit(1)

        # 매핑 데이터 추출
        print("\n4. 매핑 데이터 추출 중...")
        mapping_data = get_field4_field6_mapping(settings_sheet, field4_col, field6_col)

        if not mapping_data:
            print("추출된 매핑 데이터가 없습니다.")
            sys.exit(1)

        # txt 파일이 있는 폴더 설정
        print("\n5. txt 파일이 있는 폴더 설정...")
        current_dir = os.path.dirname(os.path.abspath(__file__))
        parent_dir = os.path.dirname(current_dir)  # PythonProject 폴더
        folder_path = os.path.join(parent_dir, "pdf_scrpy")

        if not os.path.exists(folder_path):
            print(f"폴더가 존재하지 않습니다: {folder_path}")
            print("폴더 선택 대화상자를 엽니다...")
            folder_path = select_folder()
            if not folder_path:
                print("폴더가 선택되지 않았습니다.")
                sys.exit(1)

        print(f"사용할 폴더: {folder_path}")

        # txt 파일 목록 가져오기
        txt_files = glob.glob(os.path.join(folder_path, "*.txt"))

        if not txt_files:
            print("선택한 폴더에 txt 파일이 없습니다.")
            sys.exit(1)

        print(f"\n6. 총 {len(txt_files)}개의 txt 파일을 찾았습니다.")

        # txt 파일 처리
        print("\n7. txt 파일 처리 중...")
        processed_count = 0
        skipped_count = 0
        error_count = 0

        field4_values = list(mapping_data.keys())

        for txt_file in txt_files:
            filename = os.path.basename(txt_file)

            # 파일명에서 매칭되는 field4 찾기
            matched_field4 = extract_matching_key(filename, field4_values)

            if matched_field4:
                field6_value = mapping_data[matched_field4]
                print(f"\n처리: {filename}")
                print(f"  - 매칭: {matched_field4} → {field6_value}")

                if process_txt_file(txt_file, matched_field4, field6_value):
                    processed_count += 1
                else:
                    skipped_count += 1
            else:
                print(f"\n{filename}: 매칭되는 데이터가 없습니다.")
                error_count += 1

        # 처리 결과 요약
        print("\n" + "=" * 60)
        print("처리 완료!")
        print(f"- 성공적으로 처리된 파일: {processed_count}개")
        print(f"- 건너뛴 파일 (이미 처리됨): {skipped_count}개")
        print(f"- 매칭 실패한 파일: {error_count}개")
        print(f"- 전체 파일: {len(txt_files)}개")
        print("=" * 60)

        print("\n8. 권리분석 시트 업데이트 중...")
        rights_success, rights_fail = update_rights_analysis_field8(wb, mapping_data)

        # 전체 작업 요약
        print("\n" + "=" * 60)
        print("전체 작업 요약:")
        print(f"[TXT 파일 처리]")
        print(f"- 성공: {processed_count}개")
        print(f"- 건너뜀: {skipped_count}개")
        print(f"- 실패: {error_count}개")
        print(f"\n[권리분석 시트 업데이트]")
        print(f"- 성공: {rights_success}개")
        print(f"- 실패: {rights_fail}개")
        print("=" * 60)

        # 변경사항 저장
        print("\n9. 엑셀 파일 저장 중...")
        try:
            wb.Save()
            print("엑셀 파일이 저장되었습니다.")
        except Exception as e:
            print(f"엑셀 파일 저장 실패: {str(e)}")


        print("\n프로그램을 종료합니다.")

    except Exception as e:
        print(f"\n프로그램 실행 중 예상치 못한 오류 발생: {str(e)}")
        import traceback

        traceback.print_exc()
        sys.exit(1)
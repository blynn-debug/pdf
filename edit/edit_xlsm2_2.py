import json
import os
import sys
import time
import shutil
from pathlib import Path
#edit_xlsm2_2
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


def excel_insert_data(sheet_name, cell_position, input_text):
    """
    엑셀 파일의 특정 시트, 특정 위치에 텍스트를 입력하는 함수

    매개변수:
    sheet_name (str): 입력할 시트 이름
    cell_position (str): 입력할 셀 위치 (예: 'A1', 'B2')
    input_text: 입력할 텍스트 또는 값

    반환값:
    성공 여부 (bool)
    """
    global excel, wb

    try:
        # 워크시트 가져오기
        try:
            ws = wb.Worksheets(sheet_name)
        except:
            print(f"오류: '{sheet_name}' 시트를 찾을 수 없습니다.")
            return False

        # 셀에 값 입력
        ws.Range(cell_position).Value = input_text
        print(f"'{sheet_name}' 시트의 {cell_position} 셀에 '{input_text}' 입력 완료")
        return True
    except Exception as e:
        print(f"오류 발생: {sheet_name} 시트의 {cell_position} 셀에 입력 실패 - {str(e)}")
        return False


def open_excel_file(file_path):
    """
    기존 엑셀 애플리케이션에 접근하거나 새로 열어 파일을 로드하는 함수

    매개변수:
    file_path (str): 엑셀 파일 경로

    반환값:
    성공 여부 (bool)
    """
    global excel, wb

    if not WIN32COM_AVAILABLE:
        print("win32com이 설치되어 있지 않습니다. pip install pywin32로 설치하세요.")
        return False

    try:
        # 파일이 존재하는지 확인
        if not os.path.exists(file_path):
            print(f"오류: 파일이 존재하지 않습니다: {file_path}")
            return False

        # 절대 경로로 변환
        abs_path = os.path.abspath(file_path)

        # 기존 실행 중인 Excel 인스턴스 찾기 시도
        try:
            excel = win32.GetActiveObject("Excel.Application")
            print("기존 Excel 애플리케이션에 연결했습니다.")
        except:
            # 실행 중인 Excel이 없으면 새로 시작
            excel = win32.DispatchEx('Excel.Application')
            print("새 Excel 애플리케이션을 시작했습니다.")

        # Excel 설정
        excel.DisplayAlerts = False  # 경고 메시지 표시 안 함

        # 이미 열려있는 워크북인지 확인
        is_open = False
        open_workbook = None

        for i in range(1, excel.Workbooks.Count + 1):
            if excel.Workbooks(i).FullName.lower() == abs_path.lower():
                is_open = True
                open_workbook = excel.Workbooks(i)
                print(f"'{file_path}' 파일이 이미 열려 있습니다.")
                break

        # 워크북 접근 또는 열기
        if is_open:
            wb = open_workbook
        else:
            wb = excel.Workbooks.Open(abs_path)
            print(f"엑셀 파일 '{file_path}'을 성공적으로 열었습니다.")

        # 시트 목록 출력 (디버깅 용도)
        print("시트 목록:")
        for i in range(1, wb.Sheets.Count + 1):
            sheet = wb.Sheets(i)
            print(f"  - {sheet.Name}")

        return True
    except Exception as e:
        print(f"엑셀 파일을 여는 중 오류가 발생했습니다: {str(e)}")
        return False


def save_excel_file(file_path=None):
    """
    열려있는 엑셀 파일의 변경사항을 저장하는 함수 (엑셀 유지)

    매개변수:
    file_path (str, optional): 저장할 파일 경로, 없으면 현재 열린 파일에 저장

    반환값:
    (bool, str): 성공 여부와 저장된 파일 경로
    """
    global wb

    try:
        # 저장
        wb.Save()
        saved_path = wb.FullName  # 현재 열린 파일의 전체 경로
        print(f"엑셀 파일 '{saved_path}'에 변경사항이 저장되었습니다.")
        return True, saved_path
    except Exception as e:
        print(f"엑셀 파일 저장 중 오류 발생: {str(e)}")

        # 저장 실패 시 다른 이름으로 저장 시도
        try:
            # file_path가 제공되지 않은 경우 현재 워크북의 경로 사용
            if file_path is None:
                temp_dir = os.path.dirname(wb.FullName) or '.'
            else:
                temp_dir = os.path.dirname(file_path) or '.'

            temp_file = os.path.join(temp_dir, f"temp_{int(time.time())}.xlsm")

            print(f"임시 파일로 저장: {temp_file}")
            wb.SaveAs(os.path.abspath(temp_file))
            print(f"임시 파일에 저장 성공: {temp_file}")
            return True, temp_file
        except Exception as e2:
            print(f"임시 파일 저장 중 오류 발생: {str(e2)}")
            return False, None


def process_json_to_excel(excel_file_path, json_file_path, target_sheet='입력시트'):
    """
    JSON 파일의 데이터를 읽어 엑셀 파일에 입력하는 함수

    매개변수:
    excel_file_path (str): 엑셀 파일 경로
    json_file_path (str): JSON 파일 경로
    target_sheet (str): 데이터를 입력할 시트 이름

    반환값:
    성공 여부 (bool)
    """
    global excel, wb

    # 엑셀 파일 열기
    if not open_excel_file(excel_file_path):
        print("엑셀 파일을 열지 못했습니다.")
        return False

    try:
        # JSON 파일 읽기
        with open(json_file_path, 'r', encoding='utf-8') as json_file:
            data = json.load(json_file)
        print(f"JSON 파일 '{json_file_path}'을 성공적으로 읽었습니다.")

        # JSON 데이터를 엑셀에 입력
        success_count = 0
        total_count = len(data)

        for key, value in data.items():
            location = value.get('위치', '')
            text = value.get('input_text', '')

            # 입력 호출 - 요청하신 3개 매개변수 함수 사용
            if excel_insert_data(target_sheet, location, text):
                success_count += 1

        # 결과 출력
        print(f"데이터 입력 결과: 총 {total_count}개 중 {success_count}개 성공")

        # 변경사항 저장 (엑셀 유지)
        save_success, saved_file = save_excel_file(excel_file_path)
        if save_success:
            print(f"데이터가 성공적으로 저장되었습니다: {saved_file}")
            print("엑셀을 종료하지 않고 유지합니다.")
        else:
            print("파일 저장에 실패했습니다.")

        return success_count == total_count

    except Exception as e:
        print(f"작업 중 오류가 발생했습니다: {str(e)}")
        return False



def process_bank_comparison(sheet_name):
    """
    K19 셀 값과 F28열부터 공란까지 비교하여 계산 및 입력하는 함수
    수식 참조가 아닌 실제 값을 사용하여 텍스트 형태의 수식으로 입력

    매개변수:
    sheet_name (str): 작업할 시트 이름

    반환값:
    bool: 성공 여부
    """
    global excel, wb
    if excel is None or wb is None:
        print("엑셀 또는 워크북이 초기화되지 않았습니다.")
        return False

    try:
        # 워크시트 가져오기
        try:
            ws = wb.Worksheets(sheet_name)
            print(f"'{sheet_name}' 시트를 성공적으로 가져왔습니다.")
        except Exception as e:
            print(f"오류: '{sheet_name}' 시트를 찾을 수 없습니다. 상세: {str(e)}")
            return False

        # K19 셀 값 가져오기 (비교 기준)
        k19_value = ws.Range("K19").Value
        print(f"K19 셀 값: {k19_value}")

        # 시작 행 설정
        current_row = 28
        same_values_amounts = []  # K19와 같은 값을 가진 행의 J열 실제 값들
        diff_values_amounts = []  # K19와 다른 값을 가진 행의 J열 실제 값들
        has_higher_priority = False  # K19와 다른 값이 더 선순위인지 여부

        # F열에 공란이 나올 때까지 비교
        while True:
            try:
                # F열 셀 값 읽기
                f_value = ws.Range(f"F{current_row}").Value
                print(f"F{current_row} 값: {f_value if f_value else '(공란)'}")

                # 공란인지 확인 (None 또는 빈 문자열)
                if f_value is None or f_value == "":
                    print(f"F{current_row}에서 공란을 발견했습니다. 비교를 중단합니다.")
                    break

                # J열 셀 값 읽기 (설정금액)
                j_value = ws.Range(f"J{current_row}").Value
                print(f"J{current_row} 값: {j_value}")

                # K19 셀 값과 비교
                if f_value == k19_value:
                    # 같은 경우
                    print(f"F{current_row} 값이 K19 값과 같습니다.")
                    same_values_amounts.append(str(j_value))
                else:
                    # 다른 경우
                    print(f"F{current_row} 값이 K19 값과 다릅니다.")
                    diff_values_amounts.append(str(j_value))

                    # Q열에 "대상외" 입력
                    excel_insert_data(sheet_name, f"Q{current_row}", "대상외")
                    print(f"Q{current_row}에 '대상외'를 입력했습니다.")

                    # 현재 행이 K19와 같은 값이 있는 첫 행보다 선순위인지 확인
                    if not same_values_amounts:  # 아직 같은 값이 없으면 선순위
                        has_higher_priority = True

                # 다음 행으로 이동
                current_row += 1

            except Exception as e:
                print(f"행 {current_row} 처리 중 오류 발생: {str(e)}")
                break

        # K19와 같은 값을 가진 행들의 J열 값들을 더하는 텍스트 수식 생성 및 E18에 입력
        if same_values_amounts:
            same_formula = "=" + "+".join(same_values_amounts)
            excel_insert_data(sheet_name, "E18", same_formula)
            print(f"E18에 텍스트 수식 '{same_formula}'을 입력했습니다.")

        # K19와 다른 값을 가진 행들 중 선순위가 있는 경우, 해당 행들의 J열 값들을 더하는 텍스트 수식 생성 및 E22에 입력
        if has_higher_priority and diff_values_amounts:
            diff_formula = "=" + "+".join(diff_values_amounts)
            excel_insert_data(sheet_name, "E22", diff_formula)
            print(f"E22에 텍스트 수식 '{diff_formula}'을 입력했습니다.")

        # 변경사항 저장
        save_excel_file()
        return True

    except Exception as e:
        print(f"OSB 비교 중 오류 발생: {str(e)}")
        return False


def mark_lower_priority_rows(sheet_name):
    """
    K19 셀 값과 일치하지 않으면서, K19와 일치하는 행보다 설정순위가 후순위인 행의 O열에 "Y"를 표시하는 함수

    매개변수:
    sheet_name (str): 작업할 시트 이름

    반환값:
    bool: 성공 여부
    """
    global excel, wb
    if excel is None or wb is None:
        print("엑셀 또는 워크북이 초기화되지 않았습니다.")
        return False

    try:
        # 워크시트 가져오기
        try:
            ws = wb.Worksheets(sheet_name)
            print(f"'{sheet_name}' 시트를 성공적으로 가져왔습니다.")
        except Exception as e:
            print(f"오류: '{sheet_name}' 시트를 찾을 수 없습니다. 상세: {str(e)}")
            return False

        # K19 셀 값 가져오기 (비교 기준)
        k19_value = ws.Range("K19").Value
        print(f"K19 셀 값: {k19_value}")

        # 시작 행 설정
        current_row = 28
        rows_data = []  # 각 행의 데이터를 저장할 리스트

        # F열에 공란이 나올 때까지 데이터 수집
        while True:
            try:
                # D열(설정순위) 셀 값 읽기
                d_value = ws.Range(f"D{current_row}").Value

                # F열(금융기관) 셀 값 읽기
                f_value = ws.Range(f"F{current_row}").Value

                # 공란인지 확인 (None 또는 빈 문자열)
                if f_value is None or f_value == "":
                    print(f"F{current_row}에서 공란을 발견했습니다. 비교를 중단합니다.")
                    break

                # 행 데이터 저장
                rows_data.append({
                    'row': current_row,
                    'priority': d_value,  # 설정순위
                    'matches_k19': f_value == k19_value  # K19와 일치 여부
                })

                print(f"F{current_row} 값: {f_value}, 설정순위: {d_value}, K19 일치: {f_value == k19_value}")

                # 다음 행으로 이동
                current_row += 1

            except Exception as e:
                print(f"행 {current_row} 처리 중 오류 발생: {str(e)}")
                break

        # K19와 일치하는 행 중 가장 높은 설정순위(가장 작은 숫자) 찾기
        matching_rows = [row for row in rows_data if row['matches_k19']]
        if matching_rows:
            highest_priority = min(row['priority'] for row in matching_rows)
            print(f"K19와 일치하는 행 중 가장 높은 설정순위: {highest_priority}")

            # K19와 일치하지 않으면서 설정순위가 더 낮은(숫자가 더 큰) 행에 O열에 "Y" 표시
            for row in rows_data:
                if not row['matches_k19'] and row['priority'] > highest_priority:
                    row_num = row['row']
                    print(f"행 {row_num}은 K19와 일치하지 않으면서 설정순위가 더 낮습니다. O{row_num}에 'Y'를 표시합니다.")
                    excel_insert_data(sheet_name, f"O{row_num}", "Y")
        else:
            print(f"K19 값과 일치하는 행이 없습니다.")

        # 변경사항 저장
        save_excel_file()
        return True

    except Exception as e:
        print(f"후순위 행 표시 중 오류 발생: {str(e)}")
        return False


def process_j_to_k_copy(sheet_name):
    """
    J28부터 시작하여 공란이 나올 때까지 J열의 내용을 K열로 복사하는 함수
    매개변수:
    sheet_name (str): 작업할 시트 이름
    반환값:
    bool: 성공 여부
    """
    global excel, wb
    if excel is None or wb is None:
        print("엑셀 또는 워크북이 초기화되지 않았습니다.")
        return False
    try:
        # 워크시트 가져오기
        try:
            ws = wb.Worksheets(sheet_name)
            print(f"'{sheet_name}' 시트를 성공적으로 가져왔습니다.")
        except Exception as e:
            print(f"오류: '{sheet_name}' 시트를 찾을 수 없습니다. 상세: {str(e)}")
            return False

        # 시작 행 설정
        current_row = 28

        # J열에 공란이 나올 때까지 복사
        while True:
            try:
                # J열 셀 값 읽기
                j_value = ws.Range(f"J{current_row}").Value
                print(f"J{current_row} 값: {j_value if j_value else '(공란)'}")

                # 공란인지 확인 (None 또는 빈 문자열)
                if j_value is None or j_value == "":
                    print(f"J{current_row}에서 공란을 발견했습니다. 복사를 중단합니다.")
                    break

                # J열 셀 값을 K열 동일 행에 복사
                excel_insert_data(sheet_name, f"K{current_row}", j_value)
                print(f"J{current_row}의 값을 K{current_row}에 복사했습니다.")

                # 다음 행으로 이동
                current_row += 1

            except Exception as e:
                print(f"행 {current_row} 처리 중 오류 발생: {str(e)}")
                break

        # 변경사항 저장
        save_excel_file()
        return True
    except Exception as e:
        print(f"J열에서 K열로 복사 중 오류 발생: {str(e)}")
        return False
# 전역 변수
excel = None
wb = None

# 프로그램의 메인 부분 변경
if __name__ == "__main__":
    try:
        # 파일 경로 설정
        excel_file = 'test_.xlsm'  # 매크로가 포함된 엑셀 파일
        json_file = 'test.json'
        sheet_name = '입력시트'
        file_path = './'

        print(f"프로그램 시작: Excel 파일 '{excel_file}', JSON 파일 '{json_file}', 시트 '{sheet_name}'")

        # win32com 사용 가능 여부 확인
        if WIN32COM_AVAILABLE:
            print("win32com이 설치되어 있습니다. ActiveX 컨트롤이 유지됩니다.")
        else:
            print("win32com이 설치되어 있지 않습니다. pip install pywin32 명령으로 설치하세요.")
            sys.exit(1)

        # JSON 데이터를 엑셀에 입력
        result = process_json_to_excel(excel_file, json_file, sheet_name)

        # J열 데이터를 K열로 복사
        print("J열 데이터 K열 복사 시작...")
        if process_j_to_k_copy(sheet_name):
            print("J열 데이터 K열 복사를 성공적으로 완료했습니다.")
        else:
            print("J열 데이터 K열 복사 중 오류가 발생했습니다.")

        print("대상외 작업 시작...")
        if process_bank_comparison(sheet_name):
            print("완료.  !!!채무자와 차주 꼭 추가 비교하기!!!!")
        else:
            print("오류발생.")


        print("후순위 작업 시작...")
        if mark_lower_priority_rows(sheet_name):
            print("완료.  !!!그래도 한번 더 보기!!!!")
            print("완료.  !!!채무자와 차주 꼭 추가 비교하기!!!!")
        else:
            print("오류발생.")

    except Exception as e:
        print(f"프로그램 실행 중 예상치 못한 오류 발생: {str(e)}")
        sys.exit(1)
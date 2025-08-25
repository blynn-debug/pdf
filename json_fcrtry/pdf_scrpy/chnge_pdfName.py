import os
import re
from pathlib import Path


def rename_pdf_files(directory_path="."):
    """
    PDF 파일명을 다음 규칙에 따라 변경합니다:
    - R-XXX -> XXX
    - R-XXXX-X-XX -> XXXX-XX-XX
    - WRB-XXX-XX-XX -> XXX-XX-XX
    - [한글...] 부분은 모두 제거
    """
    directory = Path(directory_path)
    renamed_count = 0
    skipped_count = 0

    for file_path in directory.glob("*.pdf"):
        original_name = file_path.stem

        # [대괄호] 이전 부분만 추출
        if '[' in original_name:
            new_name = original_name.split('[')[0].strip()
        else:
            new_name = original_name

        # R-로 시작하는 경우 처리
        if new_name.startswith("R-"):
            name_without_prefix = new_name[2:]
            parts = name_without_prefix.split("-")

            if len(parts) == 1:
                # R-007 -> 007
                new_name = parts[0]
            elif len(parts) == 3:
                # R-1252-1-01 -> 1252-01-01
                main_num = parts[0]
                mid_num = parts[1].zfill(2)
                end_num = parts[2].zfill(2)
                new_name = f"{main_num}-{mid_num}-{end_num}"

        # WRB-로 시작하는 경우 처리
        elif new_name.startswith("WRB-"):
            new_name = new_name[4:]  # WRB- 제거

        elif new_name.startswith("A-R-"):
            new_name = new_name[4:]  # WRB- 제거

        # 파일명이 변경된 경우에만 실제 파일명 변경
        if new_name != original_name:
            new_file_path = file_path.with_name(f"{new_name}.pdf")

            if new_file_path.exists():
                print(f"경고: '{new_file_path.name}' 파일이 이미 존재합니다. 건너뜁니다.")
                skipped_count += 1
            else:
                try:
                    file_path.rename(new_file_path)
                    print(f"변경: '{original_name}.pdf' -> '{new_name}.pdf'")
                    renamed_count += 1
                except Exception as e:
                    print(f"오류: '{original_name}.pdf' 변경 실패 - {e}")
                    skipped_count += 1
        else:
            print(f"유지: '{original_name}.pdf'")
            skipped_count += 1

    print(f"\n완료: {renamed_count}개 파일 변경, {skipped_count}개 파일 건너뜀")


# 실행 코드
if __name__ == "__main__":
    # pdf 폴더에서 실행
    rename_pdf_files('')
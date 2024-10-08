import os
import pandas as pd


# 폴더 경로 설정
OUTPUT_DIR = "output"  # 병합할 엑셀 파일들이 저장된 폴더
SAVE_DIR = "finished"


def merge_excel_files_with_blank_rows():
    excel_files = [f for f in os.listdir(OUTPUT_DIR) if f.endswith(".xlsx")]

    if not excel_files:
        print("병합할 엑셀 파일이 없습니다.")
        return

    # 첫 번째 파일 이름 가져오기
    first_file_name = os.path.splitext(excel_files[0])[0]
    merged_file = f"{first_file_name}_merged.xlsx"

    # 병합할 파일 데이터를 저장할 데이터프레임
    combined_df = pd.DataFrame()

    for i, excel_file in enumerate(excel_files):
        file_path = os.path.join(OUTPUT_DIR, excel_file)

        try:
            # 엑셀 파일 읽기
            df = pd.read_excel(file_path, header=None, engine='openpyxl')

        except Exception as e:
            print(f"{file_path} 파일을 읽는 중 오류 발생: {e}")
            continue

        # 데이터프레임 병합
        if not combined_df.empty:
            # 파일 구분을 위해 빈 행 추가
            combined_df = pd.concat([combined_df, pd.DataFrame(
                [[""] * df.shape[1]])], ignore_index=True)

        combined_df = pd.concat([combined_df, df], ignore_index=True)

    # 병합된 데이터프레임을 저장
    output_file_path = os.path.join(SAVE_DIR, merged_file)
    combined_df.to_excel(output_file_path, index=False, header=False)
    print(f"병합된 파일을 저장했습니다: {output_file_path}")


if __name__ == "__main__":
    merge_excel_files_with_blank_rows()

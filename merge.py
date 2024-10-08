import os
import pandas as pd

OUTPUT_DIR = "output"
SAVE_DIR = "finished"


def merge_excel_files_with_blank_rows():
    excel_files = [f for f in os.listdir(OUTPUT_DIR) if f.endswith(".xlsx")]

    if not excel_files:
        print("병합할 엑셀 파일이 없습니다.")
        return

    first_file_name = os.path.splitext(excel_files[0])[0]
    merged_file = f"{first_file_name}_merged.xlsx"
    combined_df = pd.DataFrame()

    for excel_file in excel_files:
        file_path = os.path.join(OUTPUT_DIR, excel_file)

        try:
            df = pd.read_excel(file_path, header=None, engine='openpyxl')
        except Exception as e:
            print(f"{file_path} 파일을 읽는 중 오류 발생: {e}")
            continue

        if not combined_df.empty:
            combined_df = pd.concat([combined_df, pd.DataFrame(
                [[""] * df.shape[1]])], ignore_index=True)

        combined_df = pd.concat([combined_df, df], ignore_index=True)

    output_file_path = os.path.join(SAVE_DIR, merged_file)
    combined_df.to_excel(output_file_path, index=False, header=False)
    print(f"병합된 파일을 저장했습니다: {output_file_path}")

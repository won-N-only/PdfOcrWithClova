"""
1. 파이선3 설치 후 cmd에 pip install requests pandas openpyxl
2. pdf.py와 같은 위치에 images, output폴더 생성 
3. run_ocr.bat 실행
"""

import os
import json
import uuid
import time
import requests

import pandas as pd

from dotenv import load_dotenv

# .env 파일에서 환경 변수 로드
load_dotenv()


# .env에 api 설정 값, api gateway url 넣으면 됨
CLOVA_API_URL = os.getenv("CLOVA_API_URL")  # 클로바 OCR API URL로 변경
CLOVA_API_KEY = os.getenv("CLOVA_API_KEY")  # 클로바 OCR API Key로 변경

# 폴더 경로 설정
IMAGE_DIR = "images"  # OCR 처리를 할 이미지 파일들이 저장된 폴더
OUTPUT_DIR = "output"  # 결과 엑셀 파일을 저장할 폴더

# 경로가 존재하지 않으면 생성
if not os.path.exists(OUTPUT_DIR):
    os.makedirs(OUTPUT_DIR)


def process_local_images():
    print("ocr을 시작합니다")
    # 지정된 디렉터리의 모든 이미지 파일 읽기
    image_files = [f for f in os.listdir(
        IMAGE_DIR) if os.path.isfile(os.path.join(IMAGE_DIR, f))]
    if not image_files:
        print(f" {IMAGE_DIR} 폴더에 이미지를 넣어주십쇼")
        return

    # 각 파일을 처리
    for image_file in image_files:
        # 이미지 파일 경로 설정
        image_path = os.path.join(IMAGE_DIR, image_file)

        # Clova OCR API 요청
        with open(image_path, 'rb') as img_file:
            headers = {
                "X-OCR-SECRET": CLOVA_API_KEY
            }
            files = {"file": img_file}
            data = {
                "version": "V2",  # V1 또는 V2
                "requestId": str(uuid.uuid4()),  # 임의의 고유 ID
                "timestamp": int(time.time() * 1000),  # 현재 타임스탬프 (밀리초 단위)
                "lang": "ko",  # 언어 설정 (한국어)
                "enableTableDetection": "true",  # 표 형식으로 받아오기
                "images": [
                    {
                        "format": "jpg",  # 파일 형식 (jpg, png 등)
                        "name": image_file  # 이미지 이름
                    }
                ]
            }
            response = requests.post(
                CLOVA_API_URL,
                headers=headers,
                files=files,
                data={"message": json.dumps(data)}
            )
        print(f"서버에 {image_file} 요청 했습니다")

        #  OCR 요청이 완료된 경우에만 처리
        if response.status_code == 200:
            ocr_data = response.json()
            save_table_data_to_excel(ocr_data, image_file)
            print(f"{image_file} 처리 성공")
        else:
            """
            에러 나면 status code 몇번인지, text 뭐 떴는지 보고
            https://api.ncloud-docs.com/docs/ai-application-service-ocr#%EC%9D%91%EB%8B%B5 참고하기
            """
            print(
                f"에러 발생 {image_file}: 상태코드 {response.status_code}, 상태문구 {response.text}")

    print("완료")


def save_table_data_to_excel(ocr_data, image_file):
    # JSON 데이터를 parsing하여 Python 딕셔너리로 변환
    try:
        table_data = ocr_data['images'][0]['tables'][0]['cells']
    except (KeyError, IndexError):
        print(f"표가 아닙니다 {image_file}")
        return

    # 최대 행과 열의 수를 구합니다.
    max_row = max(cell.get('rowIndex', 0) for cell in table_data) + 1
    max_col = max(cell.get('columnIndex', 0) for cell in table_data) + 1

    # 빈 테이블을 생성합니다.
    table = [[''] * max_col for _ in range(max_row)]

    # 셀 데이터를 데이터프레임에 삽입합니다.
    for cell in table_data:
        row_idx = cell.get('rowIndex', 0)
        col_idx = cell.get('columnIndex', 0)

        # 셀 내의 모든 텍스트를 결합합니다.
        text_lines = []
        for line in cell.get('cellTextLines', []):
            if 'inferText' in line:
                text_lines.append(line['inferText'])
            elif 'cellWords' in line:
                for word in line.get('cellWords', []):
                    if 'inferText' in word:
                        text_lines.append(word['inferText'])

        # 결합된 텍스트를 하나의 셀에 저장
        text = ' '.join(text_lines)
        table[row_idx][col_idx] = text

    # Pandas 데이터프레임으로 변환
    df = pd.DataFrame(table)

    # Excel 파일로 저장
    output_file = os.path.join(
        OUTPUT_DIR, f"{os.path.splitext(image_file)[0]}.xlsx")
    try:
        df.to_excel(output_file, index=False, header=False)
        print(f"Excel 파일로 저장 완료: {output_file}")
    except Exception as e:
        print(f"Excel 저장 중 오류 발생: {e}")
        return


if __name__ == "__main__":
    process_local_images()

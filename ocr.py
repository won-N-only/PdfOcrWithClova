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
IMAGE_DIR = "images"
OUTPUT_DIR = "output"

# 경로가 존재하지 않으면 생성
if not os.path.exists(OUTPUT_DIR):
    os.makedirs(OUTPUT_DIR)


def process_local_images():
    print("OCR을 시작합니다")
    image_files = [f for f in os.listdir(IMAGE_DIR) if os.path.isfile(os.path.join(IMAGE_DIR, f))]
    if not image_files:
        print(f"{IMAGE_DIR} 폴더에 이미지를 넣어주십쇼")
        return

    for image_file in image_files:
        image_path = os.path.join(IMAGE_DIR, image_file)

        # Clova OCR API 요청
        with open(image_path, 'rb') as img_file:
            headers = {"X-OCR-SECRET": CLOVA_API_KEY}
            files = {"file": img_file}
            data = {
                "version": "V2",
                "requestId": str(uuid.uuid4()),
                "timestamp": int(time.time() * 1000),
                "lang": "ko",
                "enableTableDetection": "true",
                "images": [{"format": "jpg", "name": image_file}]
            }
            response = requests.post(CLOVA_API_URL, headers=headers, files=files, data={"message": json.dumps(data)})
        
        print(f"서버에 {image_file} 요청 했습니다")

        if response.status_code == 200:
            ocr_data = response.json()
            save_table_data_to_excel(ocr_data, image_file)
            print(f"{image_file} 처리 성공")
        else:
            print(f"에러 발생 {image_file}: 상태코드 {response.status_code}, 상태문구 {response.text}")
    
    print("완료")


def save_table_data_to_excel(ocr_data, image_file):
    try:
        table_data = ocr_data['images'][0]['tables'][0]['cells']
    except (KeyError, IndexError):
        print(f"표가 아닙니다 {image_file}")
        return

    max_row = max(cell.get('rowIndex', 0) for cell in table_data) + 1
    max_col = max(cell.get('columnIndex', 0) for cell in table_data) + 1
    table = [[''] * max_col for _ in range(max_row)]

    for cell in table_data:
        row_idx = cell.get('rowIndex', 0)
        col_idx = cell.get('columnIndex', 0)
        text_lines = []
        for line in cell.get('cellTextLines', []):
            if 'inferText' in line:
                text_lines.append(line['inferText'])
            elif 'cellWords' in line:
                for word in line.get('cellWords', []):
                    if 'inferText' in word:
                        text_lines.append(word['inferText'])

        text = ' '.join(text_lines)
        table[row_idx][col_idx] = text

    df = pd.DataFrame(table)
    output_file = os.path.join(OUTPUT_DIR, f"{os.path.splitext(image_file)[0]}.xlsx")
    try:
        df.to_excel(output_file, index=False, header=False)
        print(f"Excel 파일로 저장 완료: {output_file}")
    except Exception as e:
        print(f"Excel 저장 중 오류 발생: {e}")

import os
import shutil

IMAGE_DIR = "images"
OUTPUT_DIR = "output"
SAVE_DIR = "finished"
FINISHED_IMAGES_DIR = os.path.join(SAVE_DIR, "images")
FINISHED_OUTPUT_DIR = os.path.join(SAVE_DIR, "output")


def move_files_to_finished():
    if not os.path.exists(FINISHED_IMAGES_DIR):
        os.makedirs(FINISHED_IMAGES_DIR)
    if not os.path.exists(FINISHED_OUTPUT_DIR):
        os.makedirs(FINISHED_OUTPUT_DIR)

    for file_name in os.listdir(IMAGE_DIR):
        source_file = os.path.join(IMAGE_DIR, file_name)
        destination_file = os.path.join(FINISHED_IMAGES_DIR, file_name)
        if os.path.isfile(source_file):
            shutil.move(source_file, destination_file)
            print(f"이미지 파일 이동: {source_file} -> {destination_file}")

    for file_name in os.listdir(OUTPUT_DIR):
        source_file = os.path.join(OUTPUT_DIR, file_name)
        destination_file = os.path.join(FINISHED_OUTPUT_DIR, file_name)
        if os.path.isfile(source_file):
            shutil.move(source_file, destination_file)
            print(f"엑셀 파일 이동: {source_file} -> {destination_file}")

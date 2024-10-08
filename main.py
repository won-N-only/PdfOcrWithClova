from ocr import process_local_images
from merge import merge_excel_files_with_blank_rows
from move_files import move_files_to_finished

if __name__ == "__main__":
    process_local_images()

    merge_excel_files_with_blank_rows()

    move_files_to_finished()

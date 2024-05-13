import sys
from pathlib import Path
wd = Path(__file__).parent.parent.resolve()
sys.path.append(str(wd))
import shutil
from ocr import OCR

image_folder = 'models/yolo/yolo/tables_with_texts'
output_folder_json = 'ocr/json_files'
output_folder_txt = 'ocr/tables_json_text'

ocr = OCR(image_folder, output_folder_json, output_folder_txt)
ocr.run()

try:
    folder_path = 'models/yolo/yolo'
    shutil.rmtree(folder_path)
    print('Folder and its content removed') # Folder and its content removed
except:
    print('Folder not deleted')

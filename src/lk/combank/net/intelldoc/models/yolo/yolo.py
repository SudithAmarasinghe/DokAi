from PIL import Image
import os
import shutil
import pandas as pd
import pytesseract
from ultralytics import YOLO
import time
import sys
from pathlib import Path
wd = Path(__file__).parent.parent.resolve()
sys.path.append(str(wd))
from model.model import Model

class Yolo(Model):
    def __init__(self, model_path, source_dir):
        super().__init__('yolo')
        self.model_path = model_path
        self.source_dir = source_dir
        self.load_model(model_path)

    def perform_object_detection(self, conf=0.5, iou=0.5):
        results = self.model.predict(source=self.source_dir, save=True, conf=conf, iou=iou)
        return results

    def extract_objects(self, results):
        for result_index, result in enumerate(results):
            if result.boxes:
                image_path = result.path
                image_name = os.path.basename(image_path)
                image_output_dir = os.path.join("models/yolo/yolo/prediction_results", image_name)
                os.makedirs(image_output_dir, exist_ok=True)
                for box_index, box in enumerate(result.boxes):
                    class_id = box.cls[0].item()
                    class_name = result.names[class_id]
                    class_output_dir = os.path.join(image_output_dir, class_name)
                    os.makedirs(class_output_dir, exist_ok=True)
                    cords = box.xyxy[0].tolist()
                    x1, y1, x2, y2 = cords
                    cropped_image = Image.open(image_path).crop((x1, y1, x2, y2))
                    output_filename = f"{image_name}_{box_index}.jpg"
                    output_path = os.path.join(class_output_dir, output_filename)
                    cropped_image.save(output_path)

                    with open(os.path.join(class_output_dir, "prediction_text.txt"), "w") as text_file:
                        text_file.write(f"Object type: {class_name}\n")
                        text_file.write(f"Coordinates: {cords}\n")
                        conf = box.conf[0].item()
                        text_file.write(f"Probability: {conf:.2f}\n")
                        text_file.write("---\n")

                    data = {"Class Name": class_name, "Coordinates": cords, "Probability": conf}
                    table = pd.DataFrame([data])
                    table.to_csv(os.path.join(class_output_dir, "prediction_table.csv"), index=False)


    def combine_objects(self, object_type, output_dir):
        os.makedirs(output_dir, exist_ok=True)
        for root, dirs, files in os.walk(output_dir, topdown=False):
            for file in files:
                os.remove(os.path.join(root, file))
            for dir in dirs:
                os.rmdir(os.path.join(root, dir))

        for subdir, dirs, files in os.walk("models/yolo/yolo/prediction_results"):
            for filename in files:
                if object_type in subdir:
                    filepath = os.path.join(subdir, filename)
                    if os.path.isfile(filepath):
                        new_filepath = os.path.join(output_dir, filename)
                        shutil.copyfile(filepath, new_filepath)

    def run_ocr(self, input_dir, output_dir):
        os.makedirs(output_dir, exist_ok=True)
        for filename in os.listdir(output_dir):
            file_path = os.path.join(output_dir, filename)
            try:
                if os.path.isfile(file_path):
                    os.unlink(file_path)
                elif os.path.isdir(file_path):
                    shutil.rmtree(file_path)
            except Exception as e:
                print(f"Failed to delete {file_path}. Reason: {e}")

        for filename in os.listdir(input_dir):
            input_file_path = os.path.join(input_dir, filename)
            try:
                with Image.open(input_file_path) as image:
                    text = pytesseract.image_to_string(image)
                    if text.strip():
                        output_file_path = os.path.join(output_dir, filename)
                        shutil.move(input_file_path, output_file_path)
                        self.save_data(text, output_file_path + ".json")
            except OSError:
                pass


    def run(self):
        results = self.perform_object_detection()
        self.extract_objects(results)
        self.combine_objects(object_type="table", output_dir="models/yolo/yolo/all_tables")
        self.combine_objects(object_type="figure", output_dir="models/yolo/yolo/all_figures")
        self.run_ocr(input_dir="models/yolo/yolo/all_tables", output_dir="models/yolo/yolo/tables_with_texts")
        self.run_ocr(input_dir="models/yolo/yolo/all_figures", output_dir="models/yolo/yolo/figures_with_texts")

     
        
            

         
        
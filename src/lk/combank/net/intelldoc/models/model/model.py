import os
import json
from ultralytics import YOLO
from langchain_community.llms import DeepInfra

class Model:
    def __init__(self, model_type):
        self.model_type = model_type
        # self.model = None

    def load_model(self, model_path):
        """
        Method to load the model based on model type.
        """
        if self.model_type == 'yolo':
            # Load YOLO model
            self.model = YOLO(model_path)
            print("YOLO model loaded successfully.")
        elif self.model_type == 'llama':
            # Load Llama model
            # Define your DeepInfra API token
            DEEPINFRA_API_TOKEN = "mvBGqNrv4z9LxttYNxv2BGDXqTXnrFVI"
            os.environ["DEEPINFRA_API_TOKEN"] = DEEPINFRA_API_TOKEN
            self.model = DeepInfra(model_id="meta-llama/Meta-Llama-3-70B-Instruct")
            self.model.model_kwargs = {
                "temperature": 0.7,
                "repetition_penalty": 1.2,
                "max_new_tokens": 6000,
                "top_p": 0.9,
                "do_sample": True
            }
            print("Llama model loaded successfully.")
        else:
            print("Invalid model type. Please specify 'yolo' or 'llama'.")
        
        return self.model

    def save_data(self, data, output_file_path):
        """
        Method to save data to a file.
        """
        with open(output_file_path, "w") as outfile:
            json.dump(data, outfile, indent=4)
        print(f"Data saved as: {output_file_path}")

class Yolo(Model):
    def __init__(self, model_path, source_dir):
        super().__init__('yolo')
        self.model_path = model_path
        self.source_dir = source_dir
        self.load_model(model_path)

    # Other methods specific to YOLO clas
class Llama(Model):
    def __init__(self):   
        super().__init__('llama')
        self.load_model(None)

    # Other methods specific to Llama class


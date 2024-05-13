import os
import json
import sys
from pathlib import Path

wd = Path(__file__).parent.parent.resolve()
sys.path.append(str(wd))

from llama_deep_infra import Llama

def main():
    # Define input and output directories
    input_directory = "ocr/json_files"
    output_directory = "models/llama/llama_output"

    # Create an instance of the Llama class
    llama = Llama()

    # Run the Llama model on the input files
    llama.run(input_directory, output_directory)

if __name__ == "__main__":
    main()

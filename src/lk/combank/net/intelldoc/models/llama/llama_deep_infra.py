import os
import json
from langchain_community.llms import DeepInfra
import sys
sys.path.insert(0, 'models/model')
from model import Model

class Llama(Model):
    def __init__(self):
        self.model_path = None
        self.model_type = "llama"
        DEEPINFRA_API_TOKEN = "mvBGqNrv4z9LxttYNxv2BGDXqTXnrFVI"
        os.environ["DEEPINFRA_API_TOKEN"] = DEEPINFRA_API_TOKEN
        self.model = self.load_model(self.model_path)

    def process_json_files(self, input_directory, output_directory):
        files = os.listdir(input_directory)
        json_files = [file for file in files if file.endswith(".json")]

        if not os.path.exists(output_directory):
            os.makedirs(output_directory)
        else:
            for existing_file in os.listdir(output_directory):
                file_path = os.path.join(output_directory, existing_file)
                os.remove(file_path)

        for json_file in json_files:
            input_file_path = os.path.join(input_directory, json_file)
            with open(input_file_path, "r") as file:
                json_content = file.read()

            input_text = json_content

            prompt = """### Instruction:
Create an accurate and meaningful json object from the given input.You're a helpful assistant. Give a json object by carefully understanding the given text input. 
      These text inputs contain tabular data from financial statement tables.  So the table data are read row by row. So  carefully understand columns and rows of the table and cell data inside. Then create the json object accordingly.
      The json output should have all the data from the input. Don't even leave any number. 
      Also identify whether a particular number represents a money amount, year, note number, etc. So give the number in the correct format. 
      For example, if it is a money amount, then use the standard money format.(ex:- 201,022.00)
      All the details from the input text should be inserted into the json object in the following standard template format. 
Here, understand the suitable names for the main table name and sub table names and replace them with correct name.
      So add  the relevant details and values to the template and give the json object. 


### Input:
""" + input_text + """

### Response:
"""

            output = self.model.invoke(prompt)
            prompt2 = """### Instruction:
You're a helpful assistant tasked with converting JSON objects into BSR format. The BSR format organizes data hierarchically, with key fields and associated years containing numeric values.

To recreate the JSON object in the BSR format, carefully study the provided input JSON, which may deviate significantly from the BSR format.

Identify Suitable Topic: Start by identifying the suitable topic, typically the first parent node in the JSON structure.

Identify Key-Value Pairs: Identify the key fields and their associated values. Note that these values are typically numeric and represent amounts.

Identify Years and Key Fields: Identify the immediate parent nodes of the key-value pairs, usually representing years.

Convert Values to Numeric Format: Convert the values of the year fields to float numbers. Ensure that the values are numeric without quotation marks, with two decimal places for precision.

Remove Unnecessary Nodes: Remove any unnecessary parent and child nodes from the final JSON output while adhering to the BSR format.

Example Input:
{
  "financial data": {
    "2023": {
      "Revenue": 100000,
      "Costs": 50000
    },
    "2022": {
      "Revenue": 200000,
      "Costs": 150000
    }
  }
}
Example Output:
{
  "financial data": {
    "Revenue": {
      "2023": 100000.00,
      "2022": 200000.00
    },
    "Costs": {
      "2023": 50000.00,
      "2022": 150000.00
    }
  }
}

Study the provided example input and its corresponding output, then proceed to convert JSON objects into the BSR format, ensuring accurate organization and numeric values for further calculations.
The final json object output should have the same parents and child structure as in the example output. 
So each leaf node should only have maximum of 2 parents.

Please ensure that the final JSON object does not include any lists.

Also if there is no specific year to be found, add any year in this list: [2021,2022,2023] instead of "year1","year2,"year3",.. and so on.
Please use any year for next to the leaf node values.Don't use any text other than the values in this list: [2021,2022,2023] as the keys to the leaf node values which shows the actual numerical value.

Important: Please ensure that the response is just a JSON object and no more texts other than the JSON object is added 
and has maximum of 2 parents for each child.You can remove unncessary parents if want.
For example, if there is json file like this : { A:{B:{C:{D:{F:100}}} }}, then you can just take the first parents A ,B and the F:100 to create the final json object.
You can remove C and D. 

Also make sure the final output is always a nested structured json object similar to the example output.

Extreamly Important: 
1. You need to correct all writing mistakes in the provided text. 
Ensure that all words or phrases are meaningful by correcting spelling mistakes, phrase mistakes, symbols, punctuations, and word mismatches.
Some words may have been combined together (e.g., "PropertyPlantAndEquipment"), and you need to separate them correctly (e.g., "Property, Plant And Equipment").
Make sure every word in the output is meaningful.

2. If there is no value to be found for a field, add 0 instead of Null.

3. Identify and correct the errors in the following sentences, including any instances where words are incorrectly glued together. 
      For example, if you find a word like 'stepaside,' correct it to 'step aside.' Be sure to consider both spelling and grammatical errors in your corrections.

### Input:
[Your provided text goes here]

### Response:

### Input:
""" + output + """

"""
            output2 = self.model.invoke(prompt2)

            while True:
                start_index = output2.find("{")
                end_index = output2.rfind("}") + 1

                json_data = output2[start_index:end_index]
                if json_data:
                    try:
                        json_object = json.loads(json_data)
                        output_file_path = os.path.join(output_directory, json_file)
                        self.save_data(json_object, output_file_path)
                        print(f"Generated JSON object saved as: {output_file_path}")
                        break  
                    except json.JSONDecodeError as e:
                        print(f"Error decoding JSON data: {e}")
                        break 
                else:
                    print("Error: json_data is empty")
                    break    

    def run(self, input_directory, output_directory):
        self.process_json_files(input_directory, output_directory)

    def save_data(self, data, output_file_path):
        """
        Method to save data to a file.
        """
        with open(output_file_path, "w") as outfile:
            json.dump(data, outfile, indent=4)
        print(f"Data saved as: {output_file_path}")

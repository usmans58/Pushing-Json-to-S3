import docx
from simplify_docx import simplify
import json
import os

# Read in a document
docx_file_path = os.path.join(os.getcwd(), "a.docx")
my_doc = docx.Document(docx_file_path)

# Coerce to JSON using the standard options
my_doc_as_json = simplify(my_doc)

# Define the path for the JSON file
json_file_path = os.path.join(os.getcwd(), "output.json")

# Write JSON data to a file, creating or overwriting it
with open(json_file_path, 'w') as json_file:
    json.dump(my_doc_as_json, json_file, indent=2)

print(f"JSON data has been written to {json_file_path}")

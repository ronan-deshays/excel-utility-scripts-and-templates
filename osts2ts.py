import os
import json

# Get a list of all files in the current directory
files = os.listdir()

# Filter the list to include only .osts files
osts_files = [file for file in files if file.endswith('.osts')]

output_folder_path = "C:/Users/%USERNAME%/source/repos/GitHub/ronan-deshays/excel-utility-scripts-and-templates" # replace with your target path

# Loop through all .osts files
for osts_file_name in osts_files:
    # Remove the file extension to get the base name
    base_name = os.path.splitext(osts_file_name)[0]

    # Open the JSON file
    with open(osts_file_name, 'r') as json_file:
        data = json.load(json_file)

    # Extract the value of the "body" field
    body_value = data.get('body', '')

    # Write the value to a new text file
    with open(output_folder_path + "/" + base_name + '.ts', 'w') as text_file:
        text_file.write(body_value)


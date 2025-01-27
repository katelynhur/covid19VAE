import os
import pandas as pd
import re

# Folder containing the text files
input_folder = './vaxafe_output/'
output_folder = './excel_output/'

# Get the directory of the current script
script_directory = os.path.dirname(os.path.abspath(__file__))

# Set the working directory to the script's directory
os.chdir(script_directory)

# Ensure output folder exists
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

def process_vaccine_data(file_content):
    """Extracts vaccine data from text and returns a dictionary of dataframes for each vaccine."""
    vaccine_data = {}
    lines = file_content.strip().split('\n')
    
    current_vaccine = None
    current_data = []
    headers = None

    for line in lines:
        line = line.strip()
        if line.startswith("Below are the results of statistically significant AEs associated"):
            if current_vaccine and current_data:
                # Save the current vaccine's data
                df = pd.DataFrame(current_data[1:], columns=current_data[0])
                vaccine_data[current_vaccine] = df

            # Extract the new vaccine name and reset data
            current_vaccine = line.split(':')[1].strip()
            current_data = []
        elif 'AE _ Case No _ PRR' in line:
            headers = [x.strip() for x in line.split('_')]
            current_data.append(headers)
        elif "_" in line:
            parts = line.split("_")
            # Append the row data
            current_data.append(parts)

    # Save the last vaccine's data
    if current_vaccine and current_data:
        df = pd.DataFrame(current_data[1:], columns=current_data[0])
        vaccine_data[current_vaccine] = df

    return vaccine_data

def clean_vaccine_name(vaccine_name):
    """Cleans the vaccine name by removing specific prefixes and suffixes."""
    # Remove 'COVID19 (COVID19 (" from the beginning and '))' from the end
    cleaned_name = re.sub(r'^COVID19 \(COVID19 \("?', '', vaccine_name)  # Remove prefix
    cleaned_name = re.sub(r'\)\)?$', '', cleaned_name)  # Remove suffix
    return cleaned_name

def save_to_excel(file_name, vaccine_data, output_folder):
    """Saves the extracted vaccine data to an Excel file with multiple sheets."""
    output_path = os.path.join(output_folder, file_name.replace('.txt', '.xlsx'))
    with pd.ExcelWriter(output_path) as writer:
        for vaccine, df in vaccine_data.items():
            # Clean the vaccine name
            cleaned_vaccine_name = clean_vaccine_name(vaccine)
            # Each vaccine gets its own sheet
            df.to_excel(writer, sheet_name=cleaned_vaccine_name[:30], index=False)  # sheet_name max length is 31 characters

# Process each text file in the input folder
for file_name in os.listdir(input_folder):
    if file_name.endswith('.txt'):
        input_path = os.path.join(input_folder, file_name)
        
        # Read the content of the file
        with open(input_path, 'r') as f:
            file_content = f.read()

        # Process the file content to extract vaccine data
        vaccine_data = process_vaccine_data(file_content)

        # Save the extracted data into an Excel file
        save_to_excel(file_name, vaccine_data, output_folder)

print(f"Processing complete. Excel files saved in {output_folder}")

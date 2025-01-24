import os
import shutil
import pandas as pd
from PyPDF2 import PdfMerger

script_directory = os.path.dirname(os.path.abspath(__file__))
working_folder = os.path.abspath(os.path.join(script_directory, '..'))

def create_dictionaries(df):
    """
    Create a dictionary for each header and its associated files.
    """
    header_dicts = {}
    for column in df.columns:
        header = df.iloc[0, column]  # Header as the key
        files_to_merge = df.iloc[1:, column].dropna().tolist()  # Files below as the list
        header_dicts[header] = files_to_merge
    print("Created Dictionaries:")  # Debugging statement
    for header, files in header_dicts.items():
        print(f"{header}: {files}")
    return header_dicts

def process_dictionaries(header_dicts, output_folder):
    """
    Process each dictionary: try to find and merge files, save merged output.
    """
    os.makedirs(output_folder, exist_ok=True)
    missing_files_by_header = {}  # To track missing files for each header
    
    for header, files in header_dicts.items():
        merged_file_path = os.path.join(output_folder, header)
        missing_files = []  # Track missing files for the current header
        
        if len(files) == 1:
            # If there's only one file, copy it to the output folder
            single_file_path = eval(files[0]) if 'os.path.join' in files[0] else files[0]
            if os.path.exists(single_file_path):
                shutil.copy(single_file_path, merged_file_path)
                print(f"Copied {single_file_path} to {merged_file_path}")
            else:
                print(f"File {single_file_path} does not exist. Skipping.")
                missing_files.append(single_file_path)
        elif len(files) > 1:
            # If there are multiple files, merge them
            merger = PdfMerger()
            try:
                for file_path in files:
                    evaluated_path = eval(file_path) if 'os.path.join' in file_path else file_path
                    if os.path.exists(evaluated_path):
                        merger.append(evaluated_path)
                    else:
                        print(f"File {evaluated_path} does not exist. Skipping.")
                        missing_files.append(evaluated_path)
                # Write the merged file
                if not missing_files:
                    merger.write(merged_file_path)
                    print(f"\n********\nFile successfully merged \n{os.path.join(*merged_file_path.split(os.sep)[-2:])}\n*************\n")
            except Exception as e:
                print(f"Error while merging files for {header}: {e}")
            finally:
                merger.close()
        else:
            print(f"No files found for header {header}. Skipping.")
        
        # Add missing files to the tracking dictionary
        if missing_files:
            missing_files_by_header[header] = missing_files
    
    # Provide summary feedback
    if not missing_files_by_header:
        print("\n******\n Not a single file is missing.\n*********+")
    else:
        print("\nMissing files:")
        for header, files in missing_files_by_header.items():
            print(f"{header}:")
            for file in files:
                print(f"  - {os.path.join(*file.split(os.sep)[-2:])}")
                #print(f"  - {file}")

def main():
    # Define paths
    excel_file = os.path.join(working_folder, 'Cartas.xlsx')
    output_folder = os.path.join(working_folder, 'Híbridos')
    
    # Load the Excel file and read the 'Hybrids' sheet
    try:
        df = pd.read_excel(excel_file, sheet_name='Hybrids', header=None)
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return
    
    # Create dictionaries and process
    header_dicts = create_dictionaries(df)
    process_dictionaries(header_dicts, output_folder)

if __name__ == "__main__":
    main()
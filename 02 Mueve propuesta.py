import os
import shutil
import pandas as pd

def clear_move_directories(move_paths):
    """
    Clears all files in the specified directories.
    
    Args:
    - move_paths (list): List of unique directory paths to clear.
    """
    for move_path in move_paths:
        if os.path.exists(move_path):
            for file_name in os.listdir(move_path):
                file_path = os.path.join(move_path, file_name)
                try:
                    if os.path.isfile(file_path) or os.path.islink(file_path):
                        os.unlink(file_path)  # Remove file or symlink
                        print(f"Removed: {file_path}")
                    elif os.path.isdir(file_path):
                        shutil.rmtree(file_path)  # Remove directory
                        print(f"Removed directory: {file_path}")
                except Exception as e:
                    print(f"Failed to remove {file_path}: {e}")
        else:
            os.makedirs(move_path)  # Create the directory if it doesn't exist
            print(f"Created directory: {move_path}")

def audit_copy(input_data):
    """
    Audits file presence in the source directory and copies to the specified destination.
    If a file is missing or data is invalid, it is logged in the missingfiles list.
    """
    missingfiles = []
    
    for index, row in input_data.iterrows():
        file_name = row['Nombre de archivo']
        source_dir = row['Source']
        destination_dir = row['Move']
        
        # Ensure file_name and source_dir are strings
        if pd.isna(file_name) or pd.isna(source_dir) or pd.isna(destination_dir):
            print(f"Skipping row {index} due to missing data: {row}")
            missingfiles.append({'Nombre de archivo': file_name, 'Source': source_dir})
            continue
        
        file_name = str(file_name)
        source_dir = str(source_dir)
        destination_dir = str(destination_dir)
        
        source_path = os.path.join(source_dir, file_name)
        destination_path = os.path.join(destination_dir, file_name)
        
        if os.path.exists(source_path):
            # Create destination directory if it doesn't exist
            os.makedirs(destination_dir, exist_ok=True)
            # Copy file
            shutil.copy2(source_path, destination_path)
            print(f"{file_name} from {source_dir} was copied to {destination_dir}")
        else:
            print(f"File not found: {source_path}")
            missingfiles.append({'Nombre de archivo': file_name, 'Source': source_dir})
    
    return missingfiles

def main():
    # Path to the Excel file
    excel_path = r'.\Adjudicación Dabigatrán.xlsx'
    
    # Load the Excel file as a dataframe from the sheet named 'Core'
    input_data = pd.read_excel(
        excel_path, 
        sheet_name='Parametrización',  # Specify the sheet name
        usecols=['Nombre de archivo', 'Source', 'Move']  # Columns to load
    )
    
    # Check if required columns exist
    required_columns = {'Nombre de archivo', 'Source', 'Move'}
    if required_columns.issubset(input_data.columns):
        input_data = input_data.dropna(subset=required_columns)
        input_data = input_data[(input_data['Nombre de archivo'].str.strip() != '') &
                                (input_data['Source'].str.strip() != '') &
                                (input_data['Move'].str.strip() != '')]        
        # Clear previous files in 'Move' directories
        unique_moves = input_data['Move'].dropna().unique()
        clear_move_directories(unique_moves)
        
        # Audit and copy files
        missingfiles = audit_copy(input_data)
        if missingfiles:
            print("\nMissing files:")
            for item in missingfiles:
                print(f"File: {item['Nombre de archivo']} from: {item['Source']}")
        else: 
            print("\n*************\nSuccess! \n*************\n") 
    else:
        print("The required columns ['Nombre de archivo', 'Source', 'Move'] are missing in the input data.")

if __name__ == "__main__":
    main()

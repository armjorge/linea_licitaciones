import os
import sys
import re
import PyPDF2
import pdfplumber

# Define script paths
script_directory = os.path.dirname(os.path.abspath(__file__))
working_folder = os.path.abspath(os.path.join(script_directory, '..'))
function_library = os.path.abspath(os.path.join(script_directory, 'Library'))
sys.path.append(function_library)  # Add the library folder to the path.
import pandas as pd

def extract_text_from_pdf_pypdf(pdf_path):
    """Extracts text from a PDF file using PyPDF2 (fallback)."""
    print(f"Extracting text with PyPDF2 from: {pdf_path}")  # Debug
    text = ""
    try:
        with open(pdf_path, "rb") as file:
            reader = PyPDF2.PdfReader(file)
            for page_num, page in enumerate(reader.pages):
                page_text = page.extract_text()
                if page_text:
                    text += page_text + "\n"
                else:
                    print(f"Warning: No text found on page {page_num} in {pdf_path}")  # Debug
    except Exception as e:
        print(f"Error reading {pdf_path} with PyPDF2: {e}")  # Debug
    return text

def extract_text_from_pdf_plumber(pdf_path):
    """Extracts text from a PDF file using pdfplumber."""
    print(f"Extracting text with pdfplumber from: {pdf_path}")  # Debug
    text = ""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    text += page_text + "\n"
                else:
                    print(f"Warning: No text extracted from page {pdf.pages.index(page)}")  # Debug
    except Exception as e:
        print(f"Error reading {pdf_path} with pdfplumber: {e}")  # Debug
    return text

def extract_text_with_fallback(pdf_path):
    """Tries pdfplumber first; if it fails, falls back to PyPDF2."""
    text = extract_text_from_pdf_plumber(pdf_path)  # Try pdfplumber first

    if not text.strip():  # If empty, try PyPDF2
        print("Fallback to PyPDF2 as pdfplumber failed.")  # Debug
        text = extract_text_from_pdf_pypdf(pdf_path)

    return text

def fix_broken_lines(text):
    """Fixes broken lines in extracted text."""
    return text.replace("\n", " ")

def extract_data_from_textPYPDF(text, search_dict):
    """Extracts data from text using search dictionary."""
    extracted_data = []

    print(f"Searching for structured data within {{}}...")  # Debug
    matches = re.findall(r"\{(.*?)\}", text)

    if not matches:
        print("No structured data found in {} format.")  # Debug

    for match in matches:
        fixed_text = fix_broken_lines(match)
        print(f"Found structured line: {fixed_text}")  # Debug

        values = {}
        elements = search_dict.split(", ")
        for element in elements:
            pattern = rf"{element}:\s*([^,}}]+)"
            match = re.search(pattern, fixed_text)
            if match:
                values[element] = match.group(1).strip()
        
        if values:
            extracted_data.append(values)

    return extracted_data

def get_dicts(pdf_files, search_dict):
    """Processes PDFs and extracts structured data."""
    dict_vals = []

    print(f"Loading expected fields: {search_dict}")  # Debug
    print(f"Starting the extraction process...")  # Debug

    for pdf_file in pdf_files:
        pdf_path = os.path.join(working_folder, 'Requisitos', pdf_file)
        if not os.path.exists(pdf_path):
            print(f"Error: File not found {pdf_path}")  # Debug
            continue

        print(f"Processing file: {pdf_file}")  # Debug
        text = extract_text_with_fallback(pdf_path)  # Use pdfplumber first, PyPDF2 as fallback
        extracted_data = extract_data_from_textPYPDF(text, search_dict)
        
        if extracted_data:
            print(f"Extracted data from {pdf_file}: {extracted_data}")  # Debug
        else:
            print(f"No data extracted from {pdf_file}")  # Debug

        dict_vals.extend(extracted_data)

    return dict_vals

def dictionary_to_excel(dictionary, output_path):
    """Converts a list of dictionaries to an Excel file."""
    
    # Create a DataFrame from the list of dictionaries
    output_df = pd.DataFrame(dictionary)
    
    # Check if the DataFrame is not empty
    if not output_df.empty:
        print("✅ DataFrame created successfully!")

    # Define the output file path
    filename = "Extracción.xlsx"
    file_path = os.path.join(output_path, filename)

    # Save to Excel
    output_df.to_excel(file_path, index=False)

    print(f"✅ Excel file saved successfully at: {file_path}")

def main():
    pdf_folder = os.path.join(working_folder, 'Requisitos')
    
    if not os.path.exists(pdf_folder):
        print(f"Error: Folder {pdf_folder} does not exist.")  # Debug
        return

    print(f"PDF folder found: {pdf_folder}")  # Debug

    pdf_files = [file for file in os.listdir(pdf_folder) if file.endswith(".pdf")]
    
    if not pdf_files:
        print("No PDF files found in the folder.")  # Debug
        return

    print(f"Found PDF files: {pdf_files}")  # Debug

    search_dict = "Área, Tipo, Nombre"
    requisitos = get_dicts(pdf_files, search_dict)
    dictionary_to_excel(requisitos, pdf_folder)
    print(f"Final extracted data: {requisitos}")  # Debug

if __name__ == "__main__":
    main()
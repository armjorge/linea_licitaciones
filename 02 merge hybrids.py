import os
import shutil
import pandas as pd
from PyPDF2 import PdfMerger, PdfWriter
from dotenv import load_dotenv

import glob 



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
    # Definición de rutas
    load_dotenv()
    working_folder = os.getenv("working_folder")
    excel_file = os.path.join(working_folder, 'files_flow.xlsx')
    output_folder = os.path.join(working_folder, 'File management', 'hibridos PDF_PDF')
    folder_prefix = os.path.join(working_folder, 'File management')

    try:
        # Cargamos sin cabecera para tratar la primera fila como el nombre del archivo final
        df = pd.read_excel(excel_file, sheet_name='PDF_PDF', header=None)
    except Exception as e:
        print(f"❌ Error al leer la hoja PDF_PDF: {e}")
        return

    # --- 1. Preparación de la Carpeta de Salida ---
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    
    # Verificar extensiones no permitidas en la carpeta destino
    non_pdf_files = [f for f in os.listdir(output_folder) if not f.lower().endswith('.pdf')]
    if non_pdf_files:
        print(f"⚠️ Alerta: Se encontraron archivos no PDF en la carpeta de salida: {non_pdf_files}")
        print("Por favor, límpiala antes de continuar.")
        return

    # Limpiar PDFs existentes
    for f in glob.glob(os.path.join(output_folder, "*.pdf")):
        os.remove(f)

    # --- 2. Procesamiento por Columna ---
    for col in df.columns:
        serie = df[col].dropna()
        if serie.empty:
            continue
        
        final_name = serie.iloc[0]  # El primer registro es el nombre del archivo final
        input_records = serie.iloc[1:]  # El resto son las rutas relativas
        
        file_paths_to_merge = []
        missing_files = []

        for record in input_records:
            # Separamos 'folder, filename'
            try:
                parts = [p.strip() for p in record.split(',')]
                if len(parts) == 2:
                    subfolder, filename = parts
                    full_path = os.path.join(folder_prefix, subfolder, filename)
                    
                    if os.path.exists(full_path):
                        file_paths_to_merge.append(full_path)
                    else:
                        missing_files.append(full_path)
                else:
                    print(f"⚠️ Formato incorrecto en registro: '{record}'. Debe ser 'Carpeta, archivo.pdf'")
            except Exception as e:
                print(f"Error procesando registro {record}: {e}")

        # --- 3. Validación y Mezclado ---
        if missing_files:
            print(f"\n❌ Error: Faltan archivos para generar '{final_name}':")
            for m in missing_files:
                print(f"   - No existe: {m}")
            continue # Salta a la siguiente columna

        # Si todos los archivos de la serie existen, procedemos al merge
        if file_paths_to_merge:
            writer = PdfWriter()
            try:
                for path in file_paths_to_merge:
                    writer.append(path)
                
                dest_path = os.path.join(output_folder, final_name)
                with open(dest_path, "wb") as f:
                    writer.write(f)
                
                print(f"✅ Híbrido generado: {final_name} (Piezas: {len(file_paths_to_merge)})")
            except Exception as e:
                print(f"❌ Error al fusionar {final_name}: {e}")

    print(f"\n✨ Proceso de hibridación completado en: {output_folder}")

if __name__ == "__main__":
    main()

if __name__ == "__main__":
    main()
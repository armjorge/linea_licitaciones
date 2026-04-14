import os
from PyPDF2 import PdfReader, PdfWriter
#import xlsxwriter
import sys
from dotenv import load_dotenv
load_dotenv()
import unicodedata
import re
import pandas as pd
import glob 

working_folder = os.getenv("working_folder")

def clean_filename(text, index):
    # 1. Quitar acentos y normalizar
    text = unicodedata.normalize('NFKD', text).encode('ascii', 'ignore').decode('utf-8')
    text = text.upper()
    
    # 2. Abreviaturas clave para licitaciones
    replacements = {
        'MANIFESTACION': 'MANIF',
        'ARTICULOS': 'ART',
        'OBLIGACIONES FISCALES': 'FISCAL',
        'CUMPLIMIENTO': 'CUMP',
        'PERSONALIDAD JURIDICA': 'PERS_JUR',
        'CONFLICTO DE INTERESES': 'NO_CONF_INT'
    }
    for word, rep in replacements.items():
        text = text.replace(word, rep)

    # 3. Limpiar caracteres no permitidos (deja letras, números y espacios)
    text = re.sub(r'[^A-Z0-9\s]', '', text)
    
    # 4. Acortar: tomamos las primeras 4 palabras y agregamos el índice para orden
    words = text.split()
    short_name = "_".join(words[:4])
    
    # Retornamos formato: 01_Nombre_Corto.pdf
    return f"{str(index).zfill(2)}_{short_name.capitalize()}.pdf"

def main():
    # Step 1: Get user input for PDF file name (without extension)
    pdf_name = 'IM_abril'  # Example PDF name
    pdf_path = os.path.join(working_folder, 'DOC Templates', f'{pdf_name}.pdf')  # Use an f-string to inject pdf_name
    print(f"PDF path: {pdf_path}")
    # Step 2: Verify the PDF file exists
    if not os.path.isfile(pdf_path):
        print(f"The file '{pdf_path}' does not exist.")
        return

    # Step 3: Load PDF and get bookmarks
    pdf = PdfReader(pdf_path)
    bookmarks = pdf.outline
    # --- NUEVO: Lista para recolectar datos ---
    extracted_data = []

    # Iteramos con enumerate para tener un índice (0, 1, 2...)
    for i, item in enumerate(bookmarks):
        # Nota: Algunos PDFs tienen sub-bookmarks en listas; esto asume estructura plana
        if isinstance(item, dict): 
            title = item.get('/Title')
            if title:
                print(f"\t📌 Título encontrado: {title}")
                
                # Aplicamos la limpieza
                # i+1 para que el primer archivo sea el 01 y no el 00
                sanitized = clean_filename(title, i + 1)
                
                # Guardamos en nuestra lista de diccionarios
                extracted_data.append({
                    'word_headers': title,
                    'sanitized_name': sanitized
                })

    # --- NUEVO: Parte de generación de Excel ---
    if extracted_data:
        df = pd.DataFrame(extracted_data)
        output_path = os.path.join(working_folder, 'DOC Templates', 'headers_files.xlsx')
        
        try:
            df.to_excel(output_path, index=False)
            print(f"\n✅ Excel generado con {len(df)} registros.")
            print(f"📍 Ubicación: {output_path}")
        except Exception as e:
            print(f"❌ Error al salvar Excel: {e}")
    else:
        print("⚠️ No se encontraron bookmarks para procesar.")
        return

    
    orchestrate_bookmarks_split(pdf, bookmarks)

    # # Ensure the output folder exists
    # output_folder = os.path.join(working_folder,'output')
    # if not os.path.exists(output_folder):
    #     os.makedirs(output_folder)

    # # Step 6: Call the split function with the new filenames
    # split_pdf_by_bookmarks(pdf_path, output_folder, user_bookmark_names)
    # print("PDF split by bookmarks and saved with the specified names.")


def orchestrate_bookmarks_split(pdf_reader, bookmarks):
    # 1. Cargar el Excel de flujo
    split_frame_path = os.path.join(working_folder, 'files_flow.xlsx')
    if not os.path.exists(split_frame_path):
        print(f"❌ No se encontró el archivo: {split_frame_path}")
        return

    df_split = pd.read_excel(split_frame_path, sheet_name='Parametrización')
    
    # 2. Obtener los marcadores físicos (Gold Standard)
    # Creamos un diccionario {Título: Página_Inicio} para facilitar la búsqueda
    physical_bookmarks = {}
    for item in bookmarks:
        if isinstance(item, dict):
            title = item.get('/Title')
            page_num = pdf_reader.get_destination_page_number(item)
            physical_bookmarks[title] = page_num

    # 3. Validaciones de NaN y consistencia
    # Filtrar filas donde falte información crítica
    mask_missing = df_split['Word header'].isna() & df_split['letter_name'].notna()
    mask_inverse = df_split['letter_name'].isna() & df_split['Word header'].notna()
    
    if mask_missing.any() or mask_inverse.any():
        print("\n⚠️  ADVERTENCIA: Se encontraron filas con datos incompletos en el Excel:")
        print(df_split[mask_missing | mask_inverse][['Source', 'letter_name', 'Word header']])
        print("\nPor favor, corrige estas filas y vuelve a ejecutar.")
        return

    # 4. Verificar si los 'Word header' del Excel existen en el PDF
    loaded_headers = df_split['Word header'].dropna().tolist()
    missing_in_pdf = [h for h in loaded_headers if h not in physical_bookmarks]

    if missing_in_pdf:
        print("\n" + "!"*50)
        print("🛑 ERROR DE CONCORDANCIA DETECTADO")
        print("Los siguientes encabezados del Excel NO existen en el PDF:")
        for m in missing_in_pdf:
            print(f"   - {m}")
        print("\nAcción requerida: Asegúrate de que el nombre en Word coincida exactamente con el Excel.")
        print("!"*50)
        return

    # 5. Si todo está bien, procedemos al split
    print("\n✅ Validación exitosa. Iniciando segmentación de PDF...")
    split_pdf_by_bookmarks(pdf_reader, bookmarks, df_split, physical_bookmarks)

def split_pdf_by_bookmarks(pdf_reader, bookmarks, df_split, physical_map):
    # Configurar carpetas
    output_folder = os.path.join(working_folder, 'File management', 'output')
    
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
        print(f"📁 Carpeta de salida creada: {output_folder}")
    else:
        print("✅ Carpeta de salida detectada.")

    # Limpieza: Borrar PDFs existentes
    files_to_delete = glob.glob(os.path.join(output_folder, "*.pdf"))
    for f in files_to_delete:
        os.remove(f)
    if files_to_delete: print(f"🗑️ Se eliminaron {len(files_to_delete)} archivos PDF previos.")

    # Crear lista de todos los marcadores físicos ordenados por página para saber dónde termina cada uno
    sorted_physical = sorted(physical_map.items(), key=lambda x: x[1])
    
    # Procesar solo lo que el usuario pidió en el Excel
    for _, row in df_split.dropna(subset=['Word header']).iterrows():
        header_name = row['Word header']
        letter_name = row['letter_name']
        
        start_page = physical_map[header_name]
        
        # El final es la página donde empieza el siguiente marcador físico del PDF
        end_page = len(pdf_reader.pages) # Por defecto hasta el final
        for i, (title, page) in enumerate(sorted_physical):
            if title == header_name and i + 1 < len(sorted_physical):
                end_page = sorted_physical[i+1][1]
                break

        # Extraer y guardar
        writer = PdfWriter()
        for page_idx in range(start_page, end_page):
            writer.add_page(pdf_reader.pages[page_idx])

        output_filename = f"{letter_name}.pdf" if not str(letter_name).endswith('.pdf') else letter_name
        dest_path = os.path.join(output_folder, output_filename)

        with open(dest_path, "wb") as f:
            writer.write(f)
        
        print(f"📄 Generado: {output_filename} (Págs {start_page+1}-{end_page})")

    print(f"\n✨ Proceso terminado. Archivos listos en: {output_folder}")# Run the main function
if __name__ == "__main__":
    main()

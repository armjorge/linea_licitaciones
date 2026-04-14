
import os
import shutil
import pandas as pd
from dotenv import load_dotenv

def clear_move_directories(move_paths):
    """
    Clears all files in the specified directories.
    
    Args:
    - move_paths (list): List of unique directory paths to clear.
    """
    for move_path in move_paths:
        move_path = eval(move_path)  # Convert to path using os.path.join
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

def audit_copy(input_data, working_folder):
    """
    Audits file presence in the source directory and copies to the specified destination.
    If a file is missing or data is invalid, it is logged in the missingfiles list.
    """
    missingfiles = []
    
    for index, row in input_data.iterrows():
        file_name = row['Nombre de archivo']
        source_dir = eval(row['Source'])  # Convert using os.path.join
        destination_dir = eval(row['Move'])  # Convert using os.path.join
        
        # Ensure file_name and source_dir are strings
        if pd.isna(file_name) or pd.isna(source_dir) or pd.isna(destination_dir):
            print(f"Skipping row {index} due to missing data: {row}")
            missingfiles.append({'Nombre de archivo': file_name, 'Source': source_dir})
            continue
        
        file_name = str(file_name)
        
        source_path = os.path.join(source_dir, file_name)
        destination_path = os.path.join(destination_dir, file_name)
        
        if os.path.exists(source_path):
            # Create destination directory if it doesn't exist
            os.makedirs(destination_dir, exist_ok=True)
            # Copy file
            shutil.copy2(source_path, destination_path)
            print(f"{file_name} from \\{os.path.basename(source_dir)} was copied to \\{os.path.basename(destination_dir)}")

        else:
            print(f"File not found: {os.path.basename(source_path)}")
            missingfiles.append({'Nombre de archivo': file_name, 'Source': source_dir})
    return missingfiles

def main():
    load_dotenv()
    working_folder = os.getenv("working_folder")
    file_manager = os.path.join(working_folder, 'File management')
    final_proposal = os.getenv('final_proposal')
    excel_file = os.path.join(working_folder, 'files_flow.xlsx')

    try:
        df = pd.read_excel(excel_file, sheet_name='Parametrización')
        df = df[['File_name', 'Source', 'Source name', 'Move']]
    except Exception as e:
        print(f"❌ Error al leer la hoja Parametrización: {e}")
        return

    # --- 1. Validación de Restricción de Datos ---
    # Si 'Source name' es NaN, esperamos que las demás también lo sean
    inconsistent_rows = df[df['Source name'].isna() & df[['File_name', 'Source', 'Move']].notna().any(axis=1)]
    
    if not inconsistent_rows.empty:
        print("\n⚠️ ADVERTENCIA: Filas con datos inconsistentes (Source name vacío pero otras columnas con datos):")
        print(inconsistent_rows)
        print("\nPor favor, soluciona esto en el Excel antes de continuar.")
        return

    # Limpiar filas vacías basadas en Source name
    df = df.dropna(subset=['Source name'])

    # --- 2. Verificación de existencia de archivos fuente ---
    missing_sources = []
    tasks = []

    for _, row in df.iterrows():
        source_path = os.path.join(file_manager, str(row['Source']), str(row['Source name']))
        
        # Procesar destino (manejar el caso de "CARPETA, ARCHIVO.pdf")
        move_folder = str(row['Move'])
        file_name_raw = str(row['File_name'])
        
        if ',' in file_name_raw:
            sub_folder, clean_file_name = [x.strip() for x in file_name_raw.split(',')]
            target_path = os.path.join(final_proposal, move_folder, sub_folder, clean_file_name)
        else:
            target_path = os.path.join(final_proposal, move_folder, file_name_raw)

        if os.path.exists(source_path):
            tasks.append((source_path, target_path))
        else:
            missing_sources.append(source_path)

    if missing_sources:
        print("\n❌ ERROR: Los siguientes archivos fuente no se encontraron:")
        for m in missing_sources:
            print(f"   - {m}")
        return

    # --- 3. Limpiar carpeta Final Proposal y Copiar ---
    print(f"\n🧹 Limpiando carpeta de propuesta final: {final_proposal}")
    if os.path.exists(final_proposal):
        # Borramos el contenido para asegurar una carga limpia
        shutil.rmtree(final_proposal)
    
    os.makedirs(final_proposal, exist_ok=True)

    print(f"🚀 Iniciando copiado de {len(tasks)} archivos...")
    
    for src, dst in tasks:
        # Crear subcarpetas si no existen (ej. LEGAL, o carpetas por CLAVE)
        os.makedirs(os.path.dirname(dst), exist_ok=True)
        # Copiamos preservando metadatos
        shutil.copy2(src, dst)
        print(f"✅ Copiado: {os.path.basename(dst)} -> {os.path.relpath(dst, final_proposal)}")

    print(f"\n✨ PROCESO COMPLETADO EXITOSAMENTE ✨")
    print(f"La propuesta está lista en: {final_proposal}")


if __name__ == "__main__":
    main()
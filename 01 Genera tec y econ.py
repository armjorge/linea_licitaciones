
import os
import shutil
import pandas as pd
from dotenv import load_dotenv
import subprocess

tech_map = { # Left the current columns in the existing dataframe, right the expected column name in the target dataframe
    'PARTIDA': 'NO. PARTIDA',
    'CLAVE': 'CLAVE',
    'DESCRIPCIÓN': 'DESCRIPCIÓN',
    'DENOMINACIÓN GENERICA': 'DENOMINACIÓN GENERICA',
    'PRESENTACIÓN': 'PRESENTACIÓN',
    'CANTIDAD MAX OFERTADA': 'CANTIDAD MAX OFERTADA',
    'MARCA O DENOMINACIÓN DISTINTIVA': 'MARCA O DENOMINACIÓN DISTINTIVA',
    'FABRICANTE': 'FABRICANTE',
    'PAÍS DE ORIGEN': 'PAÍS DE ORIGEN',
    'NUMERO DE REGISTRO SANITARIO (CUANDO APLIQUE)': 'NUMERO DE REGISTRO SANITARIO (CUANDO APLIQUE)'
    }

eco_map = {  # Left the current columns in the existing dataframe, right the expected column name in the target dataframe
    'PARTIDA': 'PARTIDA',
    'CUCOP': 'CUCOP',
    'CLAVE': 'Clave',
    'DESCRIPCIÓN': 'Descripción',
    'Unidad de Medida': 'Unidad de Medida',
    'Cantidad': 'Cantidad',
    'PRESENTACIÓN': 'Tipo de presentación',
    'CANTIDAD MAX OFERTADA': 'Cantidad ofertada',
    'PAÍS DE ORIGEN': 'País de Origen',
    'Precio Unitario': 'Precio Unitario',
    'Total': 'Importe Máximo',
    'IVA': 'IVA',
    'Total': 'Total'
}

def generate_tec_eco(df, col_map):
    # 1. Obtener las columnas que esperamos encontrar (las llaves del diccionario)
    expected_cols = list(col_map.keys())
    
    # 2. Verificar cuáles columnas faltan
    missing_cols = [col for col in expected_cols if col not in df.columns]
    
    if missing_cols:
        print("\n" + "!"*40)
        print("🛑 ERROR: Faltan columnas en el Excel")
        print("No se encontraron las siguientes columnas requeridas:")
        for col in missing_cols:
            print(f"   - {col}")
        print("\nColumnas disponibles actualmente en tu hoja:")
        print(list(df.columns))
        print("!"*40)
        # Retornamos el df original o None para evitar que el script truene más adelante
        return None 
    
    # 3. Si todas existen, filtramos solo las que nos interesan en el orden del mapa
    df_filtered = df[expected_cols].copy()
    
    # 4. Renombrar las columnas según los valores del mapa
    df_filtered = df_filtered.rename(columns=col_map)
    
    print(f"\n✅ Mapeo exitoso: Se procesaron {len(expected_cols)} columnas correctamente.")
    return df_filtered


def main():
    load_dotenv()
    working_folder = os.getenv("working_folder")
    excel_file = os.path.join(working_folder, 'files_flow.xlsx')
    file_manager = os.path.join(working_folder, 'File management')
    excel_final = os.path.join(file_manager, 'tecnica_economica_copy.xlsx')
    word_path = os.path.join(working_folder, 'DOC Templates', 'IM_abril.docx')
    #final_proposal = os.getenv('final_proposal')
    df = pd.read_excel(excel_file, sheet_name='Tecnica_economica')

    df_tech = generate_tec_eco(df, tech_map)
    #print(df_tech.head())

    df_eco = generate_tec_eco(df, eco_map)
    #print(df_eco.head())
    # --- Guardar en Excel (Sobrescribiendo datos previos) ---
    if df_tech is not None and df_eco is not None:
        try:
            # ExcelWriter con modo 'w' (write) elimina el archivo anterior si existe
            with pd.ExcelWriter(excel_final, engine='openpyxl', mode='w') as writer:
                df_tech.to_excel(writer, sheet_name='df_tech', index=False)
                df_eco.to_excel(writer, sheet_name='df_eco', index=False)
            
            print(f"✅ Excel generado exitosamente en: {excel_final}")

            # --- Abrir archivos y notificar ---
            # En Windows usamos os.startfile
            os.startfile(excel_final)
            
            if os.path.exists(word_path):
                os.startfile(word_path)
            else:
                print(f"⚠️ No se pudo encontrar el Word en: {word_path}")

            print("\n" + "="*60)
            print("🚀 ACCIÓN REQUERIDA:")
            print("Replace the existing data economic and technical with the open excel file data")
            print("="*60)

        except Exception as e:
            print(f"❌ Error al guardar o abrir archivos: {e}")

if __name__ == "__main__":
    main()
# ['PARTIDA', 'CUCOP', 'CLAVE', 'DESCRIPCIÓN', 'Unidad de Medida',
#    'Cantidad', 'PRESENTACIÓN', 'DENOMINACIÓN GENERICA',
#    'CANTIDAD MAX OFERTADA', 'MARCA O DENOMINACIÓN DISTINTIVA',
#    'Importe Máximo', 'PAÍS DE ORIGEN',
#    'NUMERO DE REGISTRO SANITARIO (CUANDO APLIQUE)', 'FABRICANTE',
#    'Precio Unitario', 'IVA', 'Total']
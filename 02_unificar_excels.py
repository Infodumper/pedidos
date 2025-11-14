import pandas as pd
import os
import re
from typing import Optional

# === CONFIGURACIÓN ===
# Carpeta de donde se leerán los archivos Excel
FOLDER_CAMPAIGN = r"C:\PerlaNegra\11 NACHO ADMINISTRATIVO\Minipedido\C1025"
# Nombre del archivo Excel de salida
OUTPUT_FILE = f"{FOLDER_CAMPAIGN}_Unificado.xlsx"

# Columnas finales esperadas en el DataFrame unificado (Orden final y nombres exactos)
COLUMNS_ORDER = [
    "Nro", "Cliente", "U. Ent.", "Falt.", "U. Ped", "P.V.P.", 
    "Ofertas", "Extras", "Costo Rev.", "Bonif.", "Lider"
]
# Nombres que se asignarán a las primeras 10 columnas del Excel, por POSICIÓN FIJA.
# CLAVE: Se salta el índice 2 y 3 para corregir el desplazamiento.
FIXED_COLUMN_NAMES = {
    0: "Nro",
    1: "Cliente",
    4: "U. Ent.",
    5: "Falt.",
    6: "U. Ped",
    7: "P.V.P.",
    8: "Ofertas",
    9: "Extras",
    10: "Costo Rev.",
    11: "Bonif.",
}


def normalize_string(text: str) -> str:
    """Limpia el string para buscar referencias robustas."""
    if not isinstance(text, str):
        return ""
    text = str(text).lower().strip()
    text = text.replace('á', 'a').replace('é', 'e').replace('í', 'i').replace('ó', 'o').replace('ú', 'u')
    text = re.sub(r'[.,°\#\/\-\s]', '', text) 
    return text

def extract_lider_number(df_raw: pd.DataFrame) -> Optional[str]:
    """Extrae el número de líder de la Columna B adyacente a 'Líder :'."""
    try:
        col_a = df_raw.iloc[:, 0].astype(str).str.upper()
        # Buscar la fila que contiene 'LÍDER'
        lider_row_index_list = col_a[col_a.apply(lambda x: 'LÍDER' in x)].index.tolist()
        
        if not lider_row_index_list:
             return None

        lider_row_index = lider_row_index_list[0]

        # Obtener el valor de la columna B (índice 1)
        if df_raw.shape[1] > 1:
            lider_nro = str(df_raw.iloc[lider_row_index, 1]).strip()
            return lider_nro if lider_nro else None
        else:
            return None

    except Exception as e:
        print(f"     Error al extraer líder: {e}")
        return None

def process_file(file_path: str) -> Optional[pd.DataFrame]:
    """Carga, procesa un archivo Excel forzando los nombres de columna por posición."""
    file_name = os.path.basename(file_path)
    print(f"-> Procesando: {file_name}")

    try:
        # 1. Lectura Inicial (sin encabezados, solo para buscar la posición)
        df_raw = pd.read_excel(file_path, header=None, dtype=str).fillna("")
    except Exception as e:
        print(f"   ❌ Error leyendo el archivo: {e}")
        return None

    # --- 1. Extracción del Número de Líder ---
    lider_nro = extract_lider_number(df_raw)
    
    # --- 2. Identificación de la Fila de Encabezado (donde está 'N° Cli.') ---
    header_row_index = -1
    
    for i in range(min(20, len(df_raw))): 
        row = df_raw.iloc[i].astype(str).str.strip().tolist()
        # Buscamos la fila que contiene 'Nro'
        if any(normalize_string(c) in ['ncli', 'ncli'] for c in row):
            header_row_index = i
            break
    
    if header_row_index == -1:
        print(f"   Advertencia: No se encontró la fila de encabezado ('N° Cli.' o 'Nº Cli.'). Saltando.")
        return None

    # --- 3. Re-leer el archivo, saltando las dos filas de encabezado ---
    # La data real comienza TRES filas después del inicio (header_row_index + 2)
    data_start_row_index = header_row_index + 2 

    if data_start_row_index >= len(df_raw):
        print("   Advertencia: La fila de datos está fuera de los límites del archivo. Saltando.")
        return None

    try:
        # header=None fuerza a Pandas a usar columnas numeradas (0, 1, 2, ...)
        df_data = pd.read_excel(file_path, header=None, skiprows=data_start_row_index, dtype=str).fillna("")
    except Exception as e:
        print(f"   ❌ Error re-leyendo el archivo con el nuevo desplazamiento: {e}")
        return None
    
    # --- 4. Asignación Forzada de Nombres de Columna ---
    # Asignar los nombres fijos a las columnas correctas, ignorando las columnas 2 y 3.
    renames = {}
    for col_index, new_name in FIXED_COLUMN_NAMES.items():
        if col_index < df_data.shape[1]:
            renames[col_index] = new_name
            
    # El DataFrame se lee con headers numéricos (0, 1, 2, ...), lo renombramos
    df_data.rename(columns=renames, inplace=True)
    
    # --- 5. Filtrado de registros de clientes ---
    if 'Nro' not in df_data.columns:
        print("   Advertencia: La columna 'Nro' no fue encontrada. Posiblemente el formato no es el esperado. Saltando.")
        return None
        
    df_clean = df_data[
        df_data['Nro'].astype(str).str.strip().str.isdigit() & 
        (df_data['Nro'].astype(str).str.strip().str.len() >= 4)
    ].copy()
    
    if df_clean.empty:
        print(f"   Advertencia: No se encontraron registros válidos de clientes.")
        return None

    # --- 6. Agregar la columna 'Lider' y seleccionar/reordenar ---
    df_clean['Lider'] = lider_nro
    
    final_cols = [col for col in COLUMNS_ORDER if col in df_clean.columns]
    df_final = df_clean[final_cols]
    
    print(f"   Columnas mapeadas: {final_cols}")
    print(f"   Filas de clientes extraídas: {len(df_final)}")
    return df_final


def main():
    if not os.path.exists(FOLDER_CAMPAIGN):
        print(f"❌ Error: La carpeta '{FOLDER_CAMPAIGN}' no existe. Créala y coloque los archivos .xlsx dentro.")
        return

    file_list = [
        os.path.join(FOLDER_CAMPAIGN, f) 
        for f in os.listdir(FOLDER_CAMPAIGN) 
        if f.lower().endswith(".xlsx") and not f.startswith('~')
    ]

    if not file_list:
        print(f"⚠️ No se encontraron archivos .xlsx en la carpeta '{FOLDER_CAMPAIGN}'.")
        return

    print(f"Se encontraron {len(file_list)} archivos para unificar.")
    
    all_data = []
    
    for file_path in file_list:
        df = process_file(file_path)
        if df is not None:
            all_data.append(df)
            
    if not all_data:
        print("\n❌ Error: No se pudo extraer información de ningún archivo.")
        return

    # 1. Concatenar todos los DataFrames
    df_unified = pd.concat(all_data, ignore_index=True)
    
    # 2. Filtrar y reordenar el DataFrame final
    existing_cols_in_order = [col for col in COLUMNS_ORDER if col in df_unified.columns]
    
    try:
        df_unified = df_unified[existing_cols_in_order]
    except KeyError as e:
        print(f"\n❌ ERROR CRÍTICO al reordenar columnas: Las columnas mapeadas no coinciden. {e}")
        return
    
    # --- Guardar el archivo unificado ---
    try:
        df_unified.to_excel(OUTPUT_FILE, index=False)
        print("\n=============================================")
        print(f"✅ UNIFICACIÓN EXITOSA")
        print(f"Columnas finales: {existing_cols_in_order}")
        print(f"Total de registros de clientes: {len(df_unified)}")
        print(f"Archivo guardado en: {os.path.abspath(OUTPUT_FILE)}")
        print("=============================================")
    except Exception as e:
        print(f"\n❌ Error al guardar el archivo '{OUTPUT_FILE}': {e}")


if __name__ == "__main__":
    main()
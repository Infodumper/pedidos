import os
import win32com.client as win32
from typing import List

# --- DIRECTORIO DE TRABAJO ---
# ⚠️ CAMBIA ESTA RUTA POR LA QUE NECESITES
TARGET_DIRECTORY = r"C:\PerlaNegra\11 NACHO ADMINISTRATIVO\Minipedido\C1025" 
# La 'r' (raw string) asegura que las barras invertidas se traten correctamente.
# -----------------------------

def find_xls_files(root_dir: str) -> List[str]:
    """Busca recursivamente archivos .xls dentro del directorio raíz."""
    xls_files = []
    for dirpath, _, filenames in os.walk(root_dir):
        for filename in filenames:
            # Incluir solo archivos .xls y excluir archivos temporales de Excel (~$)
            if filename.lower().endswith(".xls") and not filename.startswith('~'):
                full_path = os.path.join(dirpath, filename)
                xls_files.append(full_path)
    return xls_files

def convert_xls_to_xlsx(xls_path: str, excel_app, target_dir: str) -> bool:
    """
    Convierte un archivo .xls a .xlsx usando la aplicación de Excel, 
    guardándolo en el directorio destino especificado.
    """
    try:
        # 1. Calcular la nueva ruta de archivo .xlsx en el directorio destino
        # Usamos os.path.basename para obtener solo el nombre del archivo original.
        filename = os.path.basename(xls_path)
        base_name = filename[:-4] # Quitar .xls
        xlsx_filename = base_name + ".xlsx"
        xlsx_path = os.path.join(target_dir, xlsx_filename)
        
        # 2. Abrir el archivo .xls (la ruta original)
        print(f"   Abriendo: {filename}...")
        workbook = excel_app.Workbooks.Open(xls_path)
        
        # 3. Guardar como formato .xlsx (FileFormat=51)
        # 51 es el código para el formato xlOpenXMLWorkbook (xlsx)
        print(f"   Guardando como: {xlsx_filename} en {target_dir}...")
        workbook.SaveAs(xlsx_path, FileFormat=51)
        
        # 4. Cerrar el libro
        workbook.Close(SaveChanges=False) # No guardar cambios en el .xls original
        
        print(f"   ✅ Convertido con éxito.")
        return True
    
    except Exception as e:
        print(f"   ❌ ERROR al procesar {os.path.basename(xls_path)}: {e}")
        return False

def main():
    # Establecer el directorio raíz para la BÚSQUEDA de archivos .xls
    root_directory = TARGET_DIRECTORY
    
    # El directorio para GUARDAR los archivos .xlsx es el mismo
    save_directory = TARGET_DIRECTORY
    
    # Asegurarse de que el directorio exista
    if not os.path.isdir(root_directory):
        print(f"---")
        print(f"❌ ERROR: El directorio '{root_directory}' no existe.")
        return

    print(f"Buscando archivos .xls en: {root_directory} (y subcarpetas)...")
    xls_files = find_xls_files(root_directory)
    
    if not xls_files:
        print("---")
        print("⚠️ No se encontraron archivos .xls para convertir.")
        return

    print(f"Total de archivos .xls encontrados: {len(xls_files)}")
    print("Iniciando aplicación de Microsoft Excel...")

    # Inicializar la aplicación de Excel
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False 
    excel.DisplayAlerts = False 

    try:
        count = 0
        for xls_file in xls_files:
            # Pasamos el directorio donde debe GUARDAR el nuevo archivo
            if convert_xls_to_xlsx(xls_file, excel, save_directory):
                count += 1
        
        print("\n--- RESUMEN ---")
        print(f"Conversión finalizada. Archivos convertidos con éxito: {count} de {len(xls_files)}")
        print(f"Los nuevos archivos .xlsx se encuentran en: {save_directory}")

    finally:
        # Es crucial cerrar la aplicación de Excel al finalizar
        excel.Quit()
        print("Aplicación de Excel cerrada.")

if __name__ == "__main__":
    main()
import pandas as pd
import mysql.connector
import os
import re
import getpass # Importar getpass para una entrada de contraseña segura
from typing import Optional, Tuple

# ====================================================================
# === CONFIGURACIÓN ÚNICA A MODIFICAR ===
# ====================================================================

# 🚨 CAMPO MODIFICADO: Ahora contiene la ruta COMPLETA del archivo unificado
archivo_entrada = r"C:\PerlaNegra\11 NACHO ADMINISTRATIVO\Minipedido\C1025_Unificado.xlsx" 

# ====================================================================
# === SOLICITAR CREDENCIALES AL USUARIO ===
# ====================================================================

print("--- Credenciales de MySQL ---")
# Solicitar HOST y DATABASE (que suelen ser fijos)
# db_host = input("Ingrese el Host de la base de datos (ej. localhost): ").strip()
# db_name = input("Ingrese el nombre de la base de datos (ej. gerencia): ").strip()
db_host = 'localhost'
db_name = 'gerencia'

# Solicitar usuario y contraseña
db_user = input("Ingrese el Usuario de MySQL: ").strip()
# getpass oculta la entrada del usuario para la contraseña
db_password = getpass.getpass("Ingrese la Contraseña de MySQL: ") 
print("-----------------------------\n")

# Configuración de la base de datos dinámica
DB_CONFIG = {
    "host": db_host,
    "user": db_user,
    "password": db_password,
    "database": db_name
}
# ====================================================================

# Columnas del Excel de origen que se mapearán a la tabla 'pedidos'
EXCEL_COLUMNS_TO_EXTRACT = [
    "Nro",          # Nro Cliente
    "U. Ped",       # Unidades Pedidas (Cantidad)
    "Falt.",        # Unidades Faltantes
    "P.V.P.",       # Precio Venta Público
    "Costo Rev.",   # Costo de Revista
]

# Columnas de la tabla MySQL gerencia.pedidos
MYSQL_PEDIDOS_COLUMNS = [
    "Campaña", "Nro", "Unidades", "Faltantes", "PVP", "Costo_Rev"
]

# ====================================================================
# === FUNCIONES AUXILIARES ===
# ====================================================================

def extract_campania(file_path: str) -> Optional[str]:
    """
    Extrae la Campaña (CmmAA) del nombre del archivo (ej. C1025) a partir de la ruta completa.
    Busca 'C' seguido de 4 dígitos dentro del nombre del archivo.
    """
    if not isinstance(file_path, str):
        return None
        
    # 1. Obtener el nombre del archivo (ej. C1025_Unificado.xlsx)
    file_name = os.path.basename(file_path)
    
    # 2. Buscar 'C' seguido de 4 dígitos en el nombre del archivo
    match = re.search(r"(C\d{4})", file_name, re.IGNORECASE)
    
    if match:
        return match.group(1).upper()
        
    return None

def clean_monetary_value(value: str) -> Optional[float]:
    """Limpia valores monetarios (quita $, espacios, comas) y los convierte a float."""
    if not isinstance(value, str):
        return None
    
    clean_str = value.replace('$', '').replace(' ', '').replace('.', '').replace(',', '.').strip()
    
    # Intenta convertir a float, si falla retorna None
    try:
        return float(clean_str)
    except ValueError:
        return None

# ====================================================================
# === FUNCIÓN PRINCIPAL ===
# ====================================================================

def main():
    # 1. Obtener Campaña y verificar archivo
    campania = extract_campania(archivo_entrada)
    if not campania:
        print(f"❌ Error: No se pudo determinar la Campaña (CmmAA) del nombre del archivo '{os.path.basename(archivo_entrada)}'.")
        print("   Asegúrate de que el nombre del archivo contenga el patrón CmmAA (ej. C1025).")
        return
        
    if not os.path.exists(archivo_entrada):
        print(f"❌ Error: No se encontró el archivo de entrada en la ruta: {archivo_entrada}")
        return

    print(f"✅ Leyendo archivo: {archivo_entrada}")
    print(f"   Campaña detectada: {campania}")

    # 2. Leer el archivo Excel unificado
    try:
        df_out = pd.read_excel(archivo_entrada, dtype=str).fillna("")
        
        # Validar que las columnas necesarias existan
        missing_cols = [col for col in EXCEL_COLUMNS_TO_EXTRACT if col not in df_out.columns]
        if missing_cols:
            print(f"❌ Error: El Excel unificado NO contiene las siguientes columnas: {missing_cols}")
            print("   Revisa el script de unificación y la estructura de tu Excel.")
            return

    except Exception as e:
        print(f"❌ Error al leer o procesar el archivo Excel: {e}")
        return

    # 3. Preparación de datos para la base de datos
    pedidos_a_insertar = []
    
    for _, row in df_out.iterrows():
        # Extracción y limpieza
        nro = str(row["Nro"]).strip()
        
        # Saltamos filas sin Nro de cliente válido
        if not nro.isdigit() or len(nro) < 4:
            continue
            
        unidades = int(str(row["U. Ped"]).strip() or 0)
        faltantes = int(str(row["Falt."]).strip() or 0)
        
        # Limpieza de valores monetarios
        pvp = clean_monetary_value(row["P.V.P."])
        costo_rev = clean_monetary_value(row["Costo Rev."])
        
        # El orden de la tupla DEBE coincidir con MYSQL_PEDIDOS_COLUMNS
        pedidos_a_insertar.append((
            campania,
            nro,
            unidades,
            faltantes,
            pvp,
            costo_rev
        ))

    if not pedidos_a_insertar:
        print("⚠️ No se encontraron registros de pedidos válidos para insertar.")
        return
        
    print(f"   Pedidos válidos para cargar: {len(pedidos_a_insertar)}")

    # 4. Conexión y Carga a MySQL
    conn = None
    try:
        # Intento de conexión con las credenciales ingresadas
        conn = mysql.connector.connect(**DB_CONFIG)
        cursor = conn.cursor()
        print("✅ Conexión a MySQL establecida con éxito.")

        # Crear/Asegurar tabla pedidos
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS pedidos (
            idPedidos INT AUTO_INCREMENT PRIMARY KEY,
            Campaña VARCHAR(5) NOT NULL,
            Nro VARCHAR(6) NOT NULL,
            Unidades INT NULL,
            Faltantes INT NULL,
            PVP DECIMAL(10,2) NULL,
            Costo_Rev DECIMAL(10,2) NULL,
            UNIQUE KEY uk_campana_nro (Campaña, Nro)  -- Clave única para UPDATE
        ) CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci
        """)
        
        # SQL para INSERT O UPDATE (si ya existe la Campaña y el Nro)
        sql_pedidos = f"""
        INSERT INTO pedidos ({', '.join(MYSQL_PEDIDOS_COLUMNS)})
        VALUES (%s, %s, %s, %s, %s, %s)
        ON DUPLICATE KEY UPDATE
            Unidades = VALUES(Unidades),
            Faltantes = VALUES(Faltantes),
            PVP = VALUES(PVP),
            Costo_Rev = VALUES(Costo_Rev)
        """

        # Ejecución masiva
        cursor.executemany(sql_pedidos, pedidos_a_insertar)
        conn.commit()

        print("\n--- Carga de Pedidos Terminada ---")
        print(f"   Registros insertados/actualizados en 'pedidos': {cursor.rowcount}")
        print("----------------------------------")

    except mysql.connector.Error as err:
        print(f"\n❌ Error de base de datos o conexión: {err}")
        if conn and conn.is_connected():
            conn.rollback()
    except Exception as e:
        print(f"\n❌ Error inesperado: {e}")
    finally:
        if 'cursor' in locals() and cursor:
            cursor.close()
        if conn and conn.is_connected():
            conn.close()
            print("Conexión a MySQL cerrada.")

if __name__ == "__main__":
    main()
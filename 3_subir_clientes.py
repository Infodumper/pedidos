import pandas as pd
import mysql.connector
import os
import getpass # Importar getpass para una entrada de contraseña segura

# ====================================================================
# === CONFIGURACIÓN ÚNICA A MODIFICAR ===
# ====================================================================

# 🚨 ÚNICO CAMPO A CAMBIAR: Nombre del archivo unificado (debe estar en la misma carpeta)
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
# === INICIO DEL SCRIPT ===
# ====================================================================

# Comprobar si el archivo existe
if not os.path.exists(archivo_entrada):
    print(f"❌ Error: No se encontró el archivo de entrada: {archivo_entrada}")
    exit()

print(f"✅ Leyendo archivo: {archivo_entrada}")

# === PASO 1: Leer el archivo Excel unificado ===
try:
    # Leer el DataFrame unificado (asumimos que ya tiene las columnas correctas)
    df_out = pd.read_excel(archivo_entrada, dtype=str).fillna("")
    print(f"   Filas detectadas en el Excel: {len(df_out)}")
    
    # Asegurarse de tener las columnas clave para el proceso
    if 'Nro' not in df_out.columns or 'Cliente' not in df_out.columns or 'Lider' not in df_out.columns:
        print("❌ Error: El Excel no contiene las columnas 'Nro', 'Cliente' o 'Lider'.")
        exit()

except Exception as e:
    print(f"❌ Error al leer o procesar el archivo Excel: {e}")
    exit()


# === PASO 2: Subir SOLAMENTE a la tabla MySQL 'clientes' (Lógica: Evitar si Nro ya existe) ===
if len(df_out) == 0:
    print("⚠️ No se detectaron registros de clientes.")
else:
    conn = None # Inicializar conexión a None
    try:
        # Intento de conexión con las credenciales ingresadas
        conn = mysql.connector.connect(**DB_CONFIG)
        cursor = conn.cursor()
        print("✅ Conexión a MySQL establecida con éxito.")

        # 1. Crear/Asegurar tabla clientes (Nro es crucial que sea UNIQUE)
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS clientes (
            idCliente INT AUTO_INCREMENT PRIMARY KEY,
            Nro VARCHAR(6) UNIQUE COMMENT 'Nº Cliente',  -- Nro debe ser UNIQUE
            Cliente VARCHAR(255) COMMENT 'Nombre de Cliente',
            Lider VARCHAR(20) COMMENT 'Nº Líder'
        ) CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci
        """)
        
        # 2. Sentencias SQL para la lógica de Clientes
        sql_select_nro = "SELECT Nro FROM clientes WHERE Nro = %s"
        sql_insert_cliente = "INSERT INTO clientes (Nro, Cliente, Lider) VALUES (%s, %s, %s)"

        insertados_clientes = 0
        clientes_saltados = 0

        for _, row in df_out.iterrows():
            # Limpieza y conversión de datos
            nro = str(row["Nro"]).strip()
            cliente = str(row["Cliente"]).strip()
            # Si el Nro es vacío, saltamos la fila (puede ser ruido)
            if not nro.isdigit() or len(nro) < 4:
                continue 
                
            nro_lider = str(row["Lider"]).strip() if row["Lider"] else None 

            # LÓGICA DE RESTRICCIÓN: Buscar si el Nro ya existe
            cursor.execute(sql_select_nro, (nro,))
            result = cursor.fetchone()
            
            if result:
                # El Nro existe en la base de datos -> SALTAR ESTE REGISTRO
                clientes_saltados += 1
                continue 
            else:
                # El Nro no existe -> Insertar nuevo registro
                cursor.execute(sql_insert_cliente, (nro, cliente, nro_lider))
                insertados_clientes += 1

        conn.commit()
        print("\n--- Carga de Clientes Terminada ---")
        print(f"   Clientes insertados (Nro nuevo): {insertados_clientes}")
        print(f"   Clientes saltados (Nro preexistente): {clientes_saltados}")
        print("-----------------------------------")

    except mysql.connector.Error as err:
        # Captura errores de conexión (p. ej., credenciales incorrectas)
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
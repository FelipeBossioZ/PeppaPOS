import pandas as pd
import sqlite3
import os

# --- CONFIGURACIÓN ---
EXCEL_FILE = "Backup_24102025_211211.xlsm"
DB_FILE = "pos_database.db"
SHEET_PRODUCTOS = "Hoja1"
SHEET_PROVEEDORES = "Proveedores"
SHEET_COMPRAS = "Compras"

def migrate_data():
    """
    Función principal para leer datos de un archivo Excel y migrarlos a una base de datos SQLite.
    """
    print("Iniciando la migración de datos...")

    # --- 1. VALIDACIÓN DEL ARCHIVO EXCEL ---
    if not os.path.exists(EXCEL_FILE):
        print(f"ERROR: No se encontró el archivo Excel '{EXCEL_FILE}'. Abortando migración.")
        return

    print(f"Archivo Excel '{EXCEL_FILE}' encontrado.")

    # --- 2. CONEXIÓN A LA BASE DE DATOS ---
    # Eliminar la base de datos anterior si existe para empezar de cero.
    if os.path.exists(DB_FILE):
        os.remove(DB_FILE)
        print(f"Base de datos anterior '{DB_FILE}' eliminada.")

    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    print(f"Base de datos '{DB_FILE}' creada y lista para recibir datos.")

    # --- 3. MIGRACIÓN DE PRODUCTOS ---
    try:
        print(f"Leyendo la hoja '{SHEET_PRODUCTOS}'...")
        # Leemos solo las primeras 16 columnas (A:P) para evitar leer celdas vacías lejanas.
        df_productos = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_PRODUCTOS, skiprows=2, usecols='A:P')

        # Renombrar columnas para que sean más limpias para la base de datos
        df_productos.columns = [
            'codigo', 'nombre', 'unidad', 'costo', 'stock', 'precio_venta',
            'fecha_creacion', 'col8', 'col9', 'col10', 'col11', 'col12',
            'ultimo_ajuste', 'motivo_ajuste', 'ultima_compra', 'detalle_compra'
        ]

        # Seleccionar solo las columnas que nos interesan
        df_productos = df_productos[['codigo', 'nombre', 'unidad', 'costo', 'stock', 'precio_venta', 'fecha_creacion']]

        # Limpieza de datos
        df_productos.dropna(subset=['codigo'], inplace=True) # Eliminar filas sin código
        df_productos.drop_duplicates(subset=['codigo'], keep='last', inplace=True) # Eliminar códigos duplicados

        # Rellenar valores nulos de forma segura por columna
        df_productos['nombre'] = df_productos['nombre'].fillna('SIN NOMBRE')
        df_productos['unidad'] = df_productos['unidad'].fillna('UND')

        # --- Limpieza de fechas a prueba de errores ---
        # Convertir a datetime, los errores se convierten en NaT (Not a Time)
        dates = pd.to_datetime(df_productos['fecha_creacion'], errors='coerce')
        # Crear una nueva lista formateando fechas válidas y usando None para las inválidas
        df_productos['fecha_creacion'] = [d.strftime('%Y-%m-%d %H:%M:%S') if pd.notna(d) else None for d in dates]

        # Rellenar columnas numéricas
        numeric_cols = ['costo', 'stock', 'precio_venta']
        for col in numeric_cols:
            df_productos[col] = pd.to_numeric(df_productos[col], errors='coerce').fillna(0)

        # Convertir tipos de datos de forma segura
        df_productos = df_productos.astype({
            'costo': 'float',
            'stock': 'int',
            'precio_venta': 'float'
        })


        # Crear la tabla de productos en SQLite
        cursor.execute('''
        CREATE TABLE productos (
            codigo TEXT PRIMARY KEY,
            nombre TEXT NOT NULL,
            unidad TEXT,
            costo REAL DEFAULT 0,
            stock INTEGER DEFAULT 0,
            precio_venta REAL DEFAULT 0,
            fecha_creacion TEXT
        )
        ''')

        # Insertar los datos en la tabla
        df_productos.to_sql('productos', conn, if_exists='append', index=False)
        print(f"OK: {len(df_productos)} productos migrados a la tabla 'productos'.")

    except Exception as e:
        print(f"ERROR al migrar productos: {e}")

    # --- 4. MIGRACIÓN DE PROVEEDORES ---
    try:
        print(f"Leyendo la hoja '{SHEET_PROVEEDORES}'...")
        df_proveedores = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_PROVEEDORES, skiprows=0)
        df_proveedores.columns = ['id', 'nit', 'nombre', 'telefono', 'direccion', 'email', 'fecha_registro']
        df_proveedores = df_proveedores[['nit', 'nombre', 'telefono', 'direccion', 'email']]
        df_proveedores.dropna(subset=['nombre'], inplace=True)

        # Eliminar duplicados por 'nit', conservando el último registro
        df_proveedores.drop_duplicates(subset=['nit'], keep='last', inplace=True)

        cursor.execute('''
        CREATE TABLE proveedores (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nit TEXT UNIQUE,
            nombre TEXT NOT NULL,
            telefono TEXT,
            direccion TEXT,
            email TEXT
        )
        ''')
        df_proveedores.to_sql('proveedores', conn, if_exists='append', index=False)
        print(f"OK: {len(df_proveedores)} proveedores migrados a la tabla 'proveedores'.")

    except Exception as e:
        print(f"ERROR al migrar proveedores: {e}")

    # --- 5. MIGRACIÓN DE COMPRAS ---
    try:
        print(f"Leyendo la hoja '{SHEET_COMPRAS}'...")
        df_compras = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_COMPRAS)

        # Columnas originales: "Fecha", "Hora", "No.Compra", "Proveedor", "Código", "Producto", "Cantidad", ...
        df_compras = df_compras.iloc[:, :16] # Tomar solo las columnas relevantes (ahora 16 con IVA)
        df_compras.columns = [
            'fecha', 'hora', 'num_compra', 'proveedor', 'codigo_producto', 'nombre_producto',
            'cantidad', 'precio_unitario', 'subtotal', 'valor_iva', 'lote', 'descuento_porc',
            'total_con_desc', 'unidad', 'precio_venta', 'ultimo_costo'
        ]
        df_compras.dropna(subset=['num_compra'], inplace=True)

        cursor.execute('''
        CREATE TABLE compras (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            fecha TEXT,
            hora TEXT,
            num_compra INTEGER,
            proveedor TEXT,
            codigo_producto TEXT,
            cantidad INTEGER,
            precio_unitario REAL,
            subtotal REAL,
            valor_iva REAL,
            lote TEXT,
            descuento_porc REAL,
            total_con_desc REAL
        )
        ''')

        # Seleccionar solo las columnas que coinciden con la tabla
        df_compras_final = df_compras[[
            'fecha', 'hora', 'num_compra', 'proveedor', 'codigo_producto',
            'cantidad', 'precio_unitario', 'subtotal', 'valor_iva', 'lote', 'descuento_porc', 'total_con_desc'
        ]]

        df_compras_final.to_sql('compras', conn, if_exists='append', index=False)
        print(f"OK: {len(df_compras)} registros de compras migrados a la tabla 'compras'.")

    except Exception as e:
        print(f"ERROR al migrar compras: {e}")

    # --- CERRAR CONEXIÓN ---
    conn.commit()
    conn.close()
    print("\nMigración de datos completada.")
    print(f"Todos los datos han sido guardados en '{DB_FILE}'.")

if __name__ == "__main__":
    migrate_data()

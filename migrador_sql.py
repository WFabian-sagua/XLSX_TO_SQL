# Carga de Datos desde Excel a SQLite
# Nombre: Fabian Sagua
# Correo electrónico: datacountai@proton.me
# Compañía: DATA COUNT AI

import os
import pandas as pd
import sqlite3

# Nombre del archivo Excel
excel_filename = 'tbl_bnk.xlsx'

# Rutas para los archivos y carpetas
excel_file_path = excel_filename
db_folder_path = 'db'
db_file_path = os.path.join(db_folder_path, 'datos_bnk.db')

# Crear la carpeta si no existe
os.makedirs(db_folder_path, exist_ok=True)

# Cargar los datos desde el archivo Excel
df = pd.read_excel(excel_file_path, sheet_name='bnk')

# Conectar a la base de datos SQLite
conn = sqlite3.connect(db_file_path)
cursor = conn.cursor()

# Crear la tabla si no existe
create_table_query = '''
CREATE TABLE IF NOT EXISTS bnk_data (
    mayor TEXT,
    mayor_nombre TEXT,
    subcta TEXT,
    subcta_nombre TEXT,
    expediente TEXT,
    periodo TEXT,
    mes TEXT,
    doc_tipo INTEGER,
    doc_numero TEXT,
    fecha TEXT,
    meta_nombre TEXT,
    importe REAL
);
'''
cursor.execute(create_table_query)

# Insertar los datos en la tabla
insert_query = '''
INSERT INTO bnk_data VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);
'''

for _, row in df.iterrows():
    values = (
        row['mayor'], row['mayor_nombre'], row['subcta'], row['subcta_nombre'],
        row['expediente'], row['periodo'], row['mes'], row['doc_tipo'],
        row['doc_numero'], row['fecha'], row['meta_nombre'], row['importe']
    )
    cursor.execute(insert_query, values)

# Guardar los cambios y cerrar la conexión
conn.commit()
conn.close()

print("Datos insertados correctamente en la base de datos SQLite.")

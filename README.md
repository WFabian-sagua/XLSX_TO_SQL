# Explicación del Código: Carga de Datos desde Excel a SQLite
## Autor
- Nombre: Fabian Sagua
- Correo electrónico: datacountai@proton.me
- Compañía: DATA COUNT AI

En un mundo donde la tecnología avanza a pasos agigantados, sostengo firmemente la creencia de que la contabilidad y los profesionales contables deben adaptarse y evolucionar. La habilidad para desarrollar tecnología que fortalezca la contabilidad actual resulta esencial. En este contexto, he decidido emplear un sencillo archivo de Excel con datos contables, ampliamente utilizados en análisis contables. Aunque este ejemplo se centra en un único archivo de Excel, considero que brinda un punto de partida valioso para ilustrar cómo los datos pueden ser transferidos a una base de datos SQL. Estoy convencido de que, en el ámbito de la contabilidad gubernamental y privada en mi país, esta evolución se presenta como un pilar fundamental para aprovechar plenamente las oportunidades que la tecnología nos brinda.

Es importante destacar que la destreza en el manejo de SQL y Python para el análisis de enormes conjuntos de datos supera en beneficios a la continuación del uso de Excel y VBA. La capacidad de realizar análisis más profundos y eficientes es un diferenciador clave en la toma de decisiones informadas en un entorno cada vez más impulsado por la tecnología.

### Importación de Biblitecas
En esta sección, importamos las bibliotecas necesarias: os para interactuar con el sistema operativo, pandas para el análisis de datos y sqlite3 para trabajar con bases de datos SQLite, tal como se muestra a continuación:

```python
# Importación de Bibliotecas
import os
import pandas as pd
import sqlite3
```

### Nombre del archivo Excel
Aquí, establecemos el nombre del archivo Excel que vamos a utilizar para cargar los datos.
```python
# Nombre del archivo Excel
excel_filename = 'tbl_bnk.xlsx'
```

### Definiendo rutas del archivo
Definimos las rutas de archivos y carpetas necesarias. excel_file_path es la ruta del archivo Excel, db_folder_path es el nombre de la carpeta donde guardaremos la base de datos y db_file_path es la ruta completa al archivo de la base de datos, creada usando os.path.join().
```python
# Rutas para los archivos y carpetas
excel_file_path = excel_filename
db_folder_path = 'db'
db_file_path = os.path.join(db_folder_path, 'datos_bnk.db')
```
### Crear carpeta db
Aseguramos que la carpeta donde guardaremos la base de datos (db_folder_path) exista. Si no existe, la creamos.
```python
# Crear la carpeta si no existe
os.makedirs(db_folder_path, exist_ok=True)
```
### Cargar datos desde Excel
Utilizamos Pandas para cargar los datos desde el archivo Excel (excel_file_path). Específicamente, cargamos la hoja llamada 'bnk'.
```python
# Cargar los datos desde el archivo Excel
df = pd.read_excel(excel_file_path, sheet_name='bnk')
```

### Crear conexión a la base de datos SQLite
Creamos una conexión a la base de datos SQLite utilizando sqlite3.connect() y creamos un cursor para ejecutar consultas en la base de datos.
```python
# Conectar a la base de datos SQLite
conn = sqlite3.connect(db_file_path)
cursor = conn.cursor()
```

# Crear tabla bnk_data
Creamos una consulta SQL para crear una tabla llamada bnk_data en la base de datos si aún no existe. Ejecutamos esta consulta utilizando cursor.execute().
```python
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
```

### Inserción de datos a la tabla bnk_data
Definimos una consulta SQL de inserción para agregar datos a la tabla bnk_data. Usamos placeholders (?) para los valores que insertaremos posteriormente.

En esta parte del código, estamos iterando a través de cada fila del DataFrame df, que contiene los datos cargados desde el archivo Excel. Para cada fila, estamos extrayendo los valores de las columnas correspondientes, como 'mayor', 'mayor_nombre', 'subcta', etc. Estos valores se almacenan en la tupla values.

Luego, utilizamos el cursor (cursor) para ejecutar la consulta de inserción (insert_query) en la base de datos SQLite. La consulta de inserción tiene marcadores de posición (?) para los valores que vamos a insertar. Los valores de la tupla values se pasan a la consulta utilizando la función execute() del cursor.
```python
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
```

### Guardar cambios y cerrando conexión a la base de datos SQLite
Después de completar todas las iteraciones y ejecutar las inserciones, llamamos a conn.commit() para guardar los cambios en la base de datos. Finalmente, cerramos la conexión a la base de datos utilizando conn.close().
```python
# Guardar los cambios y cerrar la conexión
conn.commit()
conn.close()
```
Finalmente, imprimimos un mensaje para confirmar que los datos se han insertado correctamente en la base de datos SQLite.
```python
print("Datos insertados correctamente en la base de datos SQLite.")
```
En resumen, este fragmento del código, desde que realizamos la inserción de datos desde Excel a SQLite se encarga de recorrer cada fila de los datos cargados desde el archivo Excel y realizar inserciones en la base de datos SQLite con la información correspondiente.

¡Espero que esta explicación te sea útil!

__"Somos presos de lo que nos enseñaron, hasta que decidimos aprender por nuestra propia cuenta."__ 😄👨‍💻
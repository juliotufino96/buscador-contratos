import os
import glob
import re
import unicodedata
import pandas as pd
import gdown

FOLDER_ID = "1IenfFVfGPxVyEjBaK7M_1JtAKbBGaWlf"
RUTA_BASE = "Datos_Descargados"
ruta_parquet = os.path.join(RUTA_BASE, "historico.parquet")
ARCHIVO_SALIDA = "datos_procesados.parquet"

def normalizar_para_busqueda(texto):
    if pd.isna(texto): return ""
    texto = str(texto).lower()
    texto = unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('utf-8')
    return re.sub(r'[.,]', '', texto).strip()

def limpiar_y_agrupar_proveedor(texto):
    if pd.isna(texto): return ""
    texto = str(texto).upper()
    texto = unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('utf-8')
    texto = re.sub(r'[.,-]', ' ', texto) 
    texto = re.sub(r'\s+', ' ', texto).strip()
    patron_legal = r'\b(S\s*A\s*DE\s*C\s*V|S\s*DE\s*R\s*L\s*DE\s*C\s*V|S\s*DE\s*R\s*L|S\s*A\s*P\s*I\s*DE\s*C\s*V|S\s*A\s*B\s*DE\s*C\s*V|S\s*A|S\s*C|R\s*L\s*DE\s*C\s*V|S\s*P\s*R\s*DE\s*R\s*L|A\s*C)\b'
    texto = re.sub(patron_legal, '', texto).strip()
    return re.sub(r'\s+', ' ', texto).strip()

def limpiar_institucion(texto):
    if pd.isna(texto): return ""
    texto = str(texto).upper()
    texto = unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('utf-8')
    texto = re.sub(r'[.,]', '', texto)
    return texto.strip()

def leer_csv_seguro(ruta):
    codificaciones = ['utf-8-sig', 'latin-1', 'windows-1252']
    for cod in codificaciones:
        try:
            return pd.read_csv(ruta, encoding=cod, dtype=str)
        except Exception:
            continue
    return None

def ejecutar_etl():
    print("1. Descargando archivos desde Google Drive...")
    url = f'https://drive.google.com/drive/folders/{FOLDER_ID}'
    gdown.download_folder(url, output=RUTA_BASE, quiet=True, use_cookies=False)

    # Agregamos 'Descripción del contrato' antes de 'Proveedor o contratista'
    columnas_finales = [
        "Nombre del archivo", "Orden de gobierno", "Clave Ramo",
        "Siglas de la Institución", "Institución", "Número de procedimiento",
        "Núm. del contrato", "Fecha de inicio del contrato",
        "Fecha de fin del contrato", "Importe",
        "Descripción del contrato",
        "Proveedor o contratista", "Dirección del anuncio"
    ]

    df_list = []
    if os.path.exists(ruta_parquet):
        df_list.append(pd.read_parquet(ruta_parquet))

    archivos_csv = sorted(glob.glob(os.path.join(RUTA_BASE, "*.csv")))
    for csv in archivos_csv:
        df_temp = leer_csv_seguro(csv)
        if df_temp is not None:
            df_temp = df_temp.rename(columns={
                'Importe DRC': 'Importe', 
                'Número del procedimiento': 'Número de procedimiento', 
                'Importe del contrato': 'Importe',
                'Descripción del procedimiento': 'Descripción del contrato',
                'Descripción': 'Descripción del contrato'
            })
            df_temp['Nombre del archivo'] = os.path.basename(csv)
            df_list.append(df_temp)

    if df_list:
        print("Consolidando y aplicando transformaciones avanzadas...")
        df_maestro = pd.concat(df_list, ignore_index=True).reindex(columns=columnas_finales)
        
        # Limpiezas precalculadas
        df_maestro['Proveedor_Limpio'] = df_maestro['Proveedor o contratista'].apply(normalizar_para_busqueda)
        df_maestro['Proveedor_Agrupado'] = df_maestro['Proveedor o contratista'].apply(limpiar_y_agrupar_proveedor)
        df_maestro['Institución'] = df_maestro['Institución'].apply(limpiar_institucion)
        df_maestro['Orden de gobierno'] = df_maestro['Orden de gobierno'].fillna('No especificado')
        df_maestro['Clave Ramo'] = df_maestro['Clave Ramo'].fillna('Vacío')
        df_maestro['Descripción del contrato'] = df_maestro['Descripción del contrato'].fillna('Sin descripción')
        df_maestro['Año'] = df_maestro['Nombre del archivo'].astype(str).str.extract(r'(\d{4})').fillna('Sin Año')
        
        # Limpieza de importe
        df_maestro['Importe'] = pd.to_numeric(df_maestro['Importe'].astype(str).str.replace(',', '', regex=False), errors='coerce').fillna(0)

        print(f"Guardando {len(df_maestro):,} registros en {ARCHIVO_SALIDA}...")
        df_maestro.to_parquet(ARCHIVO_SALIDA, compression="zstd")
        print("¡ETL finalizado con éxito!")

if __name__ == "__main__":
    ejecutar_etl()

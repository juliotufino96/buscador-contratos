import os
import glob
import re
import unicodedata
import pandas as pd
import gdown

FOLDER_ID = "1IenfFVfGPxVyEjBaK7M_1JtAKbBGaWlf"
RUTA_BASE = "Datos_Descargados"
ruta_parquet = os.path.join(RUTA_BASE, "historico.parquet")

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
        # IGNORAMOS el archivo de exclusiones para que no se mezcle con los contratos
        nombre_csv = os.path.basename(csv).lower().replace(" ", "")
        if nombre_csv.startswith("exclusiones_apf"):
            continue
            
        df_temp = leer_csv_seguro(csv)
        if df_temp is not None:
            # Únicamente se renombran los campos de Importe y Procedimiento
            df_temp = df_temp.rename(columns={
                'Importe DRC': 'Importe', 
                'Número del procedimiento': 'Número de procedimiento', 
                'Importe del contrato': 'Importe'
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

        # =========================================================================
        # NUEVO PROCESAMIENTO: CAMBIAR APF A OTRAS BASADO EN EXCLUSIONES_APF.CSV
        # =========================================================================
        ruta_exclusiones = None
        for archivo in os.listdir(RUTA_BASE):
            nombre_limpio = archivo.lower().strip().replace(" ", "")
            if nombre_limpio.startswith("exclusiones_apf") and nombre_limpio.endswith(".csv"):
                ruta_exclusiones = os.path.join(RUTA_BASE, archivo)
                break
                
        if ruta_exclusiones:
            print(f"Aplicando reglas de exclusión usando el archivo: {os.path.basename(ruta_exclusiones)}")
            df_excl = leer_csv_seguro(ruta_exclusiones)
            
            if df_excl is not None:
                # Buscamos la columna "Institución" de forma inteligente (ignora acentos y espacios)
                columna_inst = None
                for col in df_excl.columns:
                    if normalizar_para_busqueda(col) == "institucion":
                        columna_inst = col
                        break

                if columna_inst:
                    # 1. Normalizamos las instituciones del CSV usando la misma función
                    instituciones_excluidas = set(df_excl[columna_inst].apply(limpiar_institucion).dropna())
                    
                    # 2. Creamos la condición: Es APF y la institución está en el CSV
                    mascara = (df_maestro['Orden de gobierno'] == 'APF') & (df_maestro['Institución'].isin(instituciones_excluidas))
                    
                    # 3. Aplicamos el cambio y contamos cuántos fueron
                    registros_modificados = mascara.sum()
                    df_maestro.loc[mascara, 'Orden de gobierno'] = 'AUTÓNOMOS'
                    
                    print(f"✅ Se actualizaron {registros_modificados:,} registros de 'APF' a 'AUTÓNOMOS'.")
                else:
                    print(f"⚠️ El archivo se encontró, pero no tiene la columna Institución. Columnas encontradas: {list(df_excl.columns)}")
        else:
            print("⚠️ No se encontró ningún archivo parecido a Exclusiones_APF.csv en Google Drive.")
        # =========================================================================

        CARPETA_SALIDA = "datos_parquet"
        os.makedirs(CARPETA_SALIDA, exist_ok=True)
        
        print(f"Guardando {len(df_maestro):,} registros particionados por año en la carpeta '{CARPETA_SALIDA}'...")
        
        # Agrupamos por año y guardamos un archivo por cada año
        for anio, df_anio in df_maestro.groupby('Año'):
            nombre_archivo = f"datos_{anio}.parquet"
            ruta_guardado = os.path.join(CARPETA_SALIDA, nombre_archivo)
            df_anio.to_parquet(ruta_guardado, compression="zstd")
        print("¡ETL finalizado con éxito!")

if __name__ == "__main__":
    ejecutar_etl()

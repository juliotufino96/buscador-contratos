import streamlit as st
import pandas as pd
import os
import glob
import re
import unicodedata
import io
import gdown
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
from openpyxl.utils import get_column_letter

# ================= CONFIGURACIÓN =================
st.set_page_config(page_title="Buscador de Contratos", page_icon="🏢", layout="centered")

# ID de tu carpeta pública de Drive (¡Cámbialo por el tuyo!)
FOLDER_ID = "1IenfFVfGPxVyEjBaK7M_1JtAKbBGaWlf"
RUTA_BASE = "Datos_Descargados"
ruta_parquet = os.path.join(RUTA_BASE, "historico.parquet")

# ================= FUNCIONES DE LIMPIEZA =================
def normalizar_para_busqueda(texto):
    if pd.isna(texto): return ""
    texto = str(texto).lower()
    texto = unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('utf-8')
    return re.sub(r'[.,]', '', texto).strip()

def leer_csv_seguro(ruta):
    codificaciones = ['utf-8-sig', 'latin-1', 'windows-1252']
    for cod in codificaciones:
        try:
            return pd.read_csv(ruta, encoding=cod, dtype=str)
        except Exception:
            continue
    return None

# Función para formatear el Excel (sin cambios de tu código original)
def formatear_y_escribir_tabla_integrada(ws, pt_count, pt_sum, nombre_proveedor, start_row=2):
    guinda_fill = PatternFill(start_color='9B2247', end_color='9B2247', fill_type='solid')
    verde_fill = PatternFill(start_color='1E5B4F', end_color='1E5B4F', fill_type='solid')
    white_font = Font(color='FFFFFF', bold=True)
    align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    border_thick = Side(style='thick')
    border_thin = Side(style='thin')

    columnas_periodos = list(pt_count.columns[2:])
    num_rows = len(pt_count)
    total_columnas = 2 + len(columnas_periodos) * 2
    row_titulo, row_super_header, row_sub_header = start_row, start_row + 1, start_row + 2

    ws.cell(row=row_titulo, column=1, value=f"{nombre_proveedor}")
    ws.merge_cells(start_row=row_titulo, start_column=1, end_row=row_titulo, end_column=total_columnas)
    for c in range(1, total_columnas + 1):
        cell = ws.cell(row=row_titulo, column=c)
        cell.fill, cell.font, cell.alignment = guinda_fill, white_font, align_center

    for c_idx, nombre_col in enumerate(["Clave Ramo", "Institución"], 1):
        ws.cell(row=row_super_header, column=c_idx, value=nombre_col)
        ws.merge_cells(start_row=row_super_header, start_column=c_idx, end_row=row_sub_header, end_column=c_idx)
        for r_idx in [row_super_header, row_sub_header]:
            cell = ws.cell(row=r_idx, column=c_idx)
            cell.fill, cell.font, cell.alignment = verde_fill, white_font, align_center

    col_actual = 3
    for periodo in columnas_periodos:
        ws.cell(row=row_super_header, column=col_actual, value=periodo)
        ws.merge_cells(start_row=row_super_header, start_column=col_actual, end_row=row_super_header, end_column=col_actual + 1)
        for c in range(col_actual, col_actual + 2):
            cell = ws.cell(row=row_super_header, column=c)
            cell.fill, cell.font, cell.alignment = guinda_fill, white_font, align_center
        for idx, sub_val in enumerate(["Número de procedimiento", "Importe (Millones)"]):
            sc = ws.cell(row=row_sub_header, column=col_actual + idx, value=sub_val)
            sc.fill, sc.font, sc.alignment = guinda_fill, white_font, align_center
        col_actual += 2

    data_start_row = row_sub_header + 1
    for i in range(num_rows):
        r_idx = data_start_row + i
        row_count, row_sum = pt_count.iloc[i], pt_sum.iloc[i]
        es_total = (str(row_count['Clave Ramo']) == 'Total General')

        c_ramo = ws.cell(row=r_idx, column=1, value=row_count['Clave Ramo'])
        c_inst = ws.cell(row=r_idx, column=2, value=row_count['Institución'])

        fill_color = guinda_fill if es_total else verde_fill
        for c in [c_ramo, c_inst]: c.fill, c.font = fill_color, white_font

        col_data = 3
        for periodo in columnas_periodos:
            c_cnt = ws.cell(row=r_idx, column=col_data, value=row_count[periodo])
            c_mnt = ws.cell(row=r_idx, column=col_data + 1, value=row_sum[periodo])
            c_cnt.number_format, c_mnt.number_format = '#,##0', '"$"#,##0.00'
            if es_total:
                for c in [c_cnt, c_mnt]: c.fill, c.font = guinda_fill, white_font
            col_data += 2

    end_row = data_start_row + num_rows - 1
    for r in range(start_row, end_row + 1):
        for c in range(1, total_columnas + 1):
            cell = ws.cell(row=r, column=c)
            t = border_thick if r == start_row or r == data_start_row else border_thin
            b = border_thick if r == row_sub_header or r == end_row else border_thin
            if data_start_row <= r < end_row and pt_count.iloc[r - data_start_row]['Clave Ramo'] != pt_count.iloc[r - data_start_row + 1]['Clave Ramo']:
                b = border_thin
            cell.border = Border(top=t, bottom=b, left=border_thick if c == 1 else border_thin, right=border_thick if c == total_columnas else border_thin)
    return end_row, total_columnas

# ================= INTERFAZ WEB STREAMLIT =================
st.title("🏢 Plataforma de Consulta de Contratos")
st.markdown("---")

# Variables globales en sesión para no perder datos al recargar botones
if 'df_maestro' not in st.session_state:
    st.session_state.df_maestro = pd.DataFrame()

# SECCIÓN 1: SINCRONIZAR DATOS
st.subheader("1. Actualizar Datos desde la Nube")
if st.button("Descargar e Integrar Archivos"):
    with st.spinner("Descargando archivos desde Google Drive y procesando..."):
        # Descargar carpeta pública
        url = f'https://drive.google.com/drive/folders/{FOLDER_ID}'
        gdown.download_folder(url, output=RUTA_BASE, quiet=True, use_cookies=False)
        
        columnas_finales = [
            "Nombre del archivo", "Orden de gobierno", "Clave Ramo",
            "Siglas de la Institución", "Institución", "Número de procedimiento",
            "Núm. del contrato", "Fecha de inicio del contrato",
            "Fecha de fin del contrato", "Importe",
            "Proveedor o contratista", "Dirección del anuncio"
        ]
        
        df_list = []
        if os.path.exists(ruta_parquet):
            df_list.append(pd.read_parquet(ruta_parquet))
            
        archivos_csv = sorted(glob.glob(os.path.join(RUTA_BASE, "*.csv")))
        for csv in archivos_csv:
            df_temp = leer_csv_seguro(csv)
            if df_temp is not None:
                df_temp = df_temp.rename(columns={'Importe DRC': 'Importe', 'Número del procedimiento': 'Número de procedimiento', 'Importe del contrato': 'Importe'})
                df_temp['Nombre del archivo'] = os.path.basename(csv)
                df_list.append(df_temp)

        if df_list:
            df_maestro = pd.concat(df_list, ignore_index=True).reindex(columns=columnas_finales)
            df_maestro['Proveedor_Limpio'] = df_maestro['Proveedor o contratista'].apply(normalizar_para_busqueda)
            st.session_state.df_maestro = df_maestro
            st.success(f"¡Base de datos lista! {len(df_maestro):,} registros cargados.")
        else:
            st.error("No se encontraron archivos en la carpeta.")

st.markdown("---")

# SECCIÓN 2: BUSCAR Y DESCARGAR EXCEL
st.subheader("2. Buscar Proveedor y Generar Reporte")
termino = st.text_input("Nombre del proveedor:", placeholder="Ej. LIVERPOOL")

if st.button("Buscar Contratos"):
    if st.session_state.df_maestro.empty:
        st.warning("⚠️ Primero debes hacer clic en 'Descargar e Integrar Archivos'.")
    elif not termino:
        st.warning("⚠️ Ingresa el nombre de un proveedor.")
    else:
        with st.spinner(f"Buscando '{termino}'..."):
            df = st.session_state.df_maestro
            termino_norm = normalizar_para_busqueda(termino)
            df_filtrado = df[df['Proveedor_Limpio'].str.contains(termino_norm, regex=False, na=False)].copy()

            if df_filtrado.empty:
                st.error("❌ No se encontraron contratos para este proveedor.")
            else:
                st.success(f"✅ {len(df_filtrado)} contratos encontrados. Preparando Excel...")
                
                # Procesamiento de datos
                df_exportar = df_filtrado.copy()
                proveedor_estandar = df_exportar['Proveedor o contratista'].mode()[0]
                df_exportar['Proveedor o contratista'] = proveedor_estandar
                df_exportar['Importe'] = pd.to_numeric(df_exportar['Importe'].astype(str).str.replace(',', '', regex=False), errors='coerce').fillna(0)
                df_exportar.insert(df_exportar.columns.get_loc('Importe') + 1, 'Importe (en millones)', df_exportar['Importe'] / 1000000.0)
                df_exportar['Año'] = df_exportar['Nombre del archivo'].astype(str).str.extract(r'(\d{4})').fillna('Sin Año')
                df_exportar['Clave Ramo'] = df_exportar['Clave Ramo'].fillna('Vacío')
                df_exportar['Ramo_Sort'] = df_exportar['Clave Ramo'].apply(lambda x: float(x) if str(x).replace('.','',1).isdigit() else 9999)
                df_exportar['Cuenta'] = 1

                pt_args = {'data': df_exportar, 'index': ['Ramo_Sort', 'Clave Ramo', 'Institución'], 'columns': ['Año'], 'margins': True, 'margins_name': 'Total General', 'fill_value': 0}
                pt_count = pd.pivot_table(values='Cuenta', aggfunc='sum', **pt_args).reset_index()
                pt_sum = pd.pivot_table(values='Importe (en millones)', aggfunc='sum', **pt_args).reset_index()

                for pt in [pt_count, pt_sum]:
                    pt.loc[pt['Ramo_Sort'] == 'Total General', 'Clave Ramo'] = 'Total General'
                    pt.drop(columns=['Ramo_Sort'], inplace=True)

                for col in ['Proveedor_Limpio', 'Ramo_Sort', 'Cuenta']: 
                    df_exportar.drop(columns=[col], errors='ignore', inplace=True)

                # Generar el Excel en Memoria (BytesIO)
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_exportar.to_excel(writer, sheet_name='Contratos', index=False)
                    ws_base = writer.sheets['Contratos']
                    idx_imp, idx_mill = df_exportar.columns.get_loc("Importe") + 1, df_exportar.columns.get_loc("Importe (en millones)") + 1
                    for r in range(2, len(df_exportar) + 2):
                        ws_base.cell(row=r, column=idx_imp).number_format = '"$"#,##0.00'
                        ws_base.cell(row=r, column=idx_mill).number_format = '"$"#,##0.00'

                    ws_dinamica = writer.book.create_sheet('Tabla Dinámica')
                    end_row, col_total = formatear_y_escribir_tabla_integrada(ws_dinamica, pt_count, pt_sum, proveedor_estandar, 2)
                    ws_dinamica.column_dimensions['A'].width, ws_dinamica.column_dimensions['B'].width = 18, 45
                    for col in range(3, col_total + 1): 
                        ws_dinamica.column_dimensions[get_column_letter(col)].width = 24
                
                # Botón de Descarga real de Streamlit
                st.download_button(
                    label=f"⬇️ Descargar Reporte: {proveedor_estandar}",
                    data=output.getvalue(),
                    file_name=f"Contratos_{re.sub(r'[\\/*?:<>|]', '', termino)}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
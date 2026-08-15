import streamlit as st
import pandas as pd
import os
import glob
import re
import unicodedata
import io
import gdown
import datetime
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
from openpyxl.utils import get_column_letter

# ================= CONFIGURACIÓN =================
st.set_page_config(page_title="Buscador de Contratos", page_icon="🏢", layout="centered")

# ID de tu carpeta pública de Drive
FOLDER_ID = "1IenfFVfGPxVyEjBaK7M_1JtAKbBGaWlf"
RUTA_BASE = "Datos_Descargados"
ruta_parquet = os.path.join(RUTA_BASE, "historico.parquet")

# ================= FUNCIONES DE LIMPIEZA =================
def normalizar_para_busqueda(texto):
    if pd.isna(texto): return ""
    texto = str(texto).lower()
    texto = unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('utf-8')
    return re.sub(r'[.,]', '', texto).strip()

def limpiar_institucion(texto):
    """ Estandariza el nombre de la institución a mayúsculas, sin acentos ni signos de puntuación """
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

def create_pivot_with_subtotals(df, values_col, agg_func):
    """ Crea una tabla pivote jerárquica con subtotales por Orden de Gobierno """
    # 1. Datos base
    base = pd.pivot_table(df, values=values_col, index=['Orden de gobierno', 'Ramo_Sort', 'Clave Ramo', 'Institución'], columns=['Año'], aggfunc=agg_func, fill_value=0).reset_index()
    base['Is_Subtotal'] = False
    base['Is_Total'] = False

    # 2. Subtotales por Orden de Gobierno
    sub = pd.pivot_table(df, values=values_col, index=['Orden de gobierno'], columns=['Año'], aggfunc=agg_func, fill_value=0).reset_index()
    sub['Ramo_Sort'] = 99999  # Forzar al final del grupo
    sub['Clave Ramo'] = 'Total ' + sub['Orden de gobierno']
    sub['Institución'] = ''
    sub['Is_Subtotal'] = True
    sub['Is_Total'] = False

    # 3. Total General
    tot = pd.pivot_table(df, values=values_col, index=lambda x: 'Total General', columns=['Año'], aggfunc=agg_func, fill_value=0).reset_index()
    tot.rename(columns={'index': 'Orden de gobierno'}, inplace=True)
    tot['Orden de gobierno'] = 'Total General'
    tot['Ramo_Sort'] = 999999 # Forzar hasta el fondo
    tot['Clave Ramo'] = ''
    tot['Institución'] = ''
    tot['Is_Subtotal'] = False
    tot['Is_Total'] = True

    # Integrar y ordenar
    res = pd.concat([base, sub, tot], ignore_index=True).fillna(0)
    res['Orden_Sort'] = res['Orden de gobierno'].apply(lambda x: 'ZZZZ' if x == 'Total General' else x)
    res.sort_values(by=['Orden_Sort', 'Ramo_Sort', 'Clave Ramo', 'Institución'], inplace=True)
    res.drop(columns=['Orden_Sort'], inplace=True)
    return res

# Función para formatear el Excel 
def formatear_y_escribir_tabla_integrada(ws, pt_count, pt_sum, nombre_proveedor, start_row=2):
    guinda_fill = PatternFill(start_color='9B2247', end_color='9B2247', fill_type='solid')
    verde_fill = PatternFill(start_color='1E5B4F', end_color='1E5B4F', fill_type='solid')
    white_font = Font(color='FFFFFF', bold=True)
    align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    border_thick = Side(style='thick')
    border_thin = Side(style='thin')

    columnas_periodos = [c for c in pt_count.columns if str(c).isdigit() or str(c) == 'Sin Año']
    num_rows = len(pt_count)
    total_columnas = 3 + len(columnas_periodos) * 2  # Actualizado para 3 columnas estáticas
    row_titulo, row_super_header, row_sub_header = start_row, start_row + 1, start_row + 2

    # 1. Título General
    ws.cell(row=row_titulo, column=1, value=f"{nombre_proveedor}")
    ws.merge_cells(start_row=row_titulo, start_column=1, end_row=row_titulo, end_column=total_columnas)
    for c in range(1, total_columnas + 1):
        cell = ws.cell(row=row_titulo, column=c)
        cell.fill, cell.font, cell.alignment = guinda_fill, white_font, align_center

    # 2. Encabezados de Columnas Fijas
    for c_idx, nombre_col in enumerate(["Orden de gobierno", "Clave Ramo", "Institución"], 1):
        ws.cell(row=row_super_header, column=c_idx, value=nombre_col)
        ws.merge_cells(start_row=row_super_header, start_column=c_idx, end_row=row_sub_header, end_column=c_idx)
        for r_idx in [row_super_header, row_sub_header]:
            cell = ws.cell(row=r_idx, column=c_idx)
            cell.fill, cell.font, cell.alignment = verde_fill, white_font, align_center

    # 3. Encabezados de Periodos
    col_actual = 4
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

    # 4. Datos
    data_start_row = row_sub_header + 1
    for i in range(num_rows):
        r_idx = data_start_row + i
        row_count, row_sum = pt_count.iloc[i], pt_sum.iloc[i]
        
        es_total = row_count.get('Is_Total', False)
        es_subtotal = row_count.get('Is_Subtotal', False)

        c_ord = ws.cell(row=r_idx, column=1, value=row_count['Orden de gobierno'])
        c_ramo = ws.cell(row=r_idx, column=2, value=row_count['Clave Ramo'])
        c_inst = ws.cell(row=r_idx, column=3, value=row_count['Institución'])

        # Colorear fila completa de guinda si es un Subtotal o Total General
        fill_color = guinda_fill if (es_total or es_subtotal) else verde_fill
        for c in [c_ord, c_ramo, c_inst]: c.fill, c.font = fill_color, white_font

        col_data = 4
        for periodo in columnas_periodos:
            c_cnt = ws.cell(row=r_idx, column=col_data, value=row_count[periodo])
            c_mnt = ws.cell(row=r_idx, column=col_data + 1, value=row_sum[periodo])
            c_cnt.number_format, c_mnt.number_format = '#,##0', '"$"#,##0.00'
            if es_total or es_subtotal:
                for c in [c_cnt, c_mnt]: c.fill, c.font = fill_color, white_font
            col_data += 2

    # 5. Bordes
    end_row = data_start_row + num_rows - 1
    for r in range(start_row, end_row + 1):
        for c in range(1, total_columnas + 1):
            cell = ws.cell(row=r, column=c)
            t = border_thick if r == start_row or r == data_start_row else border_thin
            b = border_thick if r == row_sub_header or r == end_row else border_thin
            
            if data_start_row <= r < end_row:
                b = border_thin
                row_current = pt_count.iloc[r - data_start_row]
                row_next = pt_count.iloc[r - data_start_row + 1]
                # Borde grueso para delimitar saltos de agrupaciones
                if row_current['Orden de gobierno'] != row_next['Orden de gobierno'] or row_current.get('Is_Subtotal'):
                    b = border_thick

            cell.border = Border(top=t, bottom=b, left=border_thick if c == 1 else border_thin, right=border_thick if c == total_columnas else border_thin)
    return end_row, total_columnas

# ================= INTERFAZ WEB STREAMLIT =================
st.title("🏢 Consulta de Contratos")
st.markdown("---")

if 'df_maestro' not in st.session_state:
    st.session_state.df_maestro = pd.DataFrame()

# SECCIÓN 1: SINCRONIZAR DATOS
st.subheader("1. Actualizar Datos desde la Nube")
if st.button("Actualizar"):
    with st.spinner("Descargando archivos desde Google Drive y procesando..."):
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
        st.warning("⚠️ Primero debes hacer clic en 'Actualizar'.")
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
                
                # Limpieza de Institución e imputación de Orden de Gobierno
                df_exportar['Institución'] = df_exportar['Institución'].apply(limpiar_institucion)
                df_exportar['Orden de gobierno'] = df_exportar['Orden de gobierno'].fillna('No especificado')
                df_exportar['Clave Ramo'] = df_exportar['Clave Ramo'].fillna('Vacío')
                df_exportar['Ramo_Sort'] = df_exportar['Clave Ramo'].apply(lambda x: float(x) if str(x).replace('.','',1).isdigit() else 9999)
                df_exportar['Cuenta'] = 1

                # Tablas dinámicas jerarquizadas
                pt_count = create_pivot_with_subtotals(df_exportar, 'Cuenta', 'sum')
                pt_sum = create_pivot_with_subtotals(df_exportar, 'Importe (en millones)', 'sum')

                # Columnas sobrantes a eliminar antes del volcado final
                for col in ['Proveedor_Limpio', 'Ramo_Sort', 'Cuenta']: 
                    df_exportar.drop(columns=[col], errors='ignore', inplace=True)

                # Generar el Excel en Memoria
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    # 1. Crear Hoja de Detalle
                    df_exportar.to_excel(writer, sheet_name='Detalle Contratos', index=False)
                    ws_base = writer.sheets['Detalle Contratos']
                    idx_imp, idx_mill = df_exportar.columns.get_loc("Importe") + 1, df_exportar.columns.get_loc("Importe (en millones)") + 1
                    for r in range(2, len(df_exportar) + 2):
                        ws_base.cell(row=r, column=idx_imp).number_format = '"$"#,##0.00'
                        ws_base.cell(row=r, column=idx_mill).number_format = '"$"#,##0.00'

                    # 2. Crear Hoja Resumen
                    ws_dinamica = writer.book.create_sheet('Resumen')
                    end_row, col_total = formatear_y_escribir_tabla_integrada(ws_dinamica, pt_count, pt_sum, proveedor_estandar, 2)
                    ws_dinamica.column_dimensions['A'].width = 18 # Orden de gobierno
                    ws_dinamica.column_dimensions['B'].width = 18 # Clave Ramo
                    ws_dinamica.column_dimensions['C'].width = 45 # Institución
                    for col in range(4, col_total + 1): 
                        ws_dinamica.column_dimensions[get_column_letter(col)].width = 24
                    
                    # 3. Invertir el orden (Mover Resumen al principio)
                    writer.book.move_sheet("Resumen", offset=-1)
                
                # Botón de Descarga
                st.download_button(
                    label=f"⬇️ Descargar Reporte: {proveedor_estandar}",
                    data=output.getvalue(),
                    file_name=f"Contratos_{re.sub(r'[\\/*?:<>|]', '', termino)}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

# ================= DISCLAIMER =================
año_actual = datetime.datetime.now().year
st.markdown("<br><br>", unsafe_allow_html=True)
st.markdown("---")
st.markdown(
    f"<p style='text-align: center; font-size: 11px; color: #666;'>"
    f"Fuente: La información fue generada con base en los datos abiertos de los contratos registrados en la plataforma Compras MX "
    f"durante el periodo 2019-{año_actual}, para consulta pública. Disponibles en: "
    f"<a href='https://comprasmx.buengobierno.gob.mx/datos-abiertos' target='_blank'>https://comprasmx.buengobierno.gob.mx/datos-abiertos</a></p>",
    unsafe_allow_html=True
)

import streamlit as st
import pandas as pd
import re
from io import BytesIO

# Configuración básica de la página
st.set_page_config(page_title="Control Pagos Ombligo", layout="centered")

def limpiar_rut(rut):
    if pd.isna(rut):
        return ""
    # Convertir a texto y mayúsculas
    rut = str(rut).upper()
    # Dejar solo números y la letra K
    return re.sub(r'[^0-9K]', '', rut)

st.title("Control de Pagos - Ombligo Gen 25")
st.markdown("Sube los documentos para cruzar los RUTs y obtener la lista de correos para los recordatorios.")

# Carga de archivos desde la interfaz
file_inscripciones = st.file_uploader("1. Inscripciones Ombligo Gen 25", type=['xlsx'])
file_externo = st.file_uploader("2. Pago cuota externo", type=['xlsx'])
file_interno = st.file_uploader("3. Pago cuota interno", type=['xlsx'])

if st.button("Procesar y Comparar") and file_inscripciones and file_externo and file_interno:
    with st.spinner("Cruzando datos..."):
        # Leer los Excels
        df_ins = pd.read_excel(file_inscripciones)
        df_ext = pd.read_excel(file_externo)
        df_int = pd.read_excel(file_interno)
        
        # Encontrar las columnas de RUT dinámicamente por si cambia un poco el nombre en el Forms
        col_rut_ins = [c for c in df_ins.columns if 'Rut' in c or 'RUT' in str(c).upper()][0]
        col_rut_ext = [c for c in df_ext.columns if 'Rut' in c or 'RUT' in str(c).upper()][0]
        col_rut_int = [c for c in df_int.columns if 'Rut' in c or 'RUT' in str(c).upper()][0]
        
        # Limpiar y estandarizar RUTs
        df_ins['RUT_Limpio'] = df_ins[col_rut_ins].apply(limpiar_rut)
        df_ext['RUT_Limpio'] = df_ext[col_rut_ext].apply(limpiar_rut)
        df_int['RUT_Limpio'] = df_int[col_rut_int].apply(limpiar_rut)
        
        # Unir ambas listas de pagos
        pagos_ext = set(df_ext['RUT_Limpio'][df_ext['RUT_Limpio'] != ""])
        pagos_int = set(df_int['RUT_Limpio'][df_int['RUT_Limpio'] != ""])
        todos_pagos = pagos_ext.union(pagos_int)
        
        # Filtrar a los inscritos que NO están en la lista de todos los pagos
        df_pendientes = df_ins[~df_ins['RUT_Limpio'].isin(todos_pagos)].copy()
        
        # Armar la tabla final con ID y Correo
        resultado = df_pendientes[['Id', 'Nombre completo', col_rut_ins, 'Mail personal']].copy()
        resultado.columns = ['ID', 'Nombre Completo', 'RUT', 'Correo Electrónico']
        resultado.dropna(subset=['Nombre Completo'], inplace=True)
        
        st.warning(f"⚠️ Se encontraron {len(resultado)} personas que aún no pagan.")
        st.dataframe(resultado)
        
        # Generar el Excel para descargar
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            resultado.to_excel(writer, index=False, sheet_name='Pendientes')
        processed_data = output.getvalue()
        
        st.download_button(
            label="Descargar Excel de Pendientes",
            data=processed_data,
            file_name="Pendientes_Pago_Ombligo.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

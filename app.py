import streamlit as st
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="Generador de Correos - Pagos", layout="centered")

def limpiar_rut(rut):
    if pd.isna(rut):
        return ""
    rut = str(rut).upper()
    return re.sub(r'[^0-9K]', '', rut)

st.title("Generador de Recordatorios de Pago")
st.markdown("Sube la base de inscritos y **un** formulario de pago (interno o externo) para obtener la lista de correos de quienes faltan en ese formulario en particular.")

# Solo pedimos 2 archivos
file_inscripciones = st.file_uploader("1. Base de Inscripciones Ombligo", type=['xlsx'])
file_pago = st.file_uploader("2. Formulario de Pago (Interno o Externo)", type=['xlsx'])

if st.button("Generar Lista de Correos") and file_inscripciones and file_pago:
    with st.spinner("Cruzando datos..."):
        try:
            # Leer los Excels
            df_ins = pd.read_excel(file_inscripciones)
            df_pago = pd.read_excel(file_pago)
            
            # Encontrar las columnas de RUT
            col_rut_ins = [c for c in df_ins.columns if 'Rut' in c or 'RUT' in str(c).upper()][0]
            col_rut_pago = [c for c in df_pago.columns if 'Rut' in c or 'RUT' in str(c).upper()][0]
            
            # Limpiar y estandarizar RUTs
            df_ins['RUT_Limpio'] = df_ins[col_rut_ins].apply(limpiar_rut)
            df_pago['RUT_Limpio'] = df_pago[col_rut_pago].apply(limpiar_rut)
            
            # Obtener lista de RUTs que ya pagaron en el formulario subido
            pagos_hechos = set(df_pago['RUT_Limpio'][df_pago['RUT_Limpio'] != ""])
            
            # Filtrar a los inscritos que NO están en este formulario
            df_pendientes = df_ins[~df_ins['RUT_Limpio'].isin(pagos_hechos)].copy()
            
            # Identificar la columna de correo (por si varía el nombre)
            col_correo = 'Mail personal' if 'Mail personal' in df_ins.columns else 'Correo electrónico'
            
            # Armar la tabla final
            resultado = df_pendientes[['Id', 'Nombre completo', col_rut_ins, col_correo]].copy()
            resultado.columns = ['ID', 'Nombre Completo', 'RUT', 'Correo Electrónico']
            resultado.dropna(subset=['Nombre Completo'], inplace=True)
            
            st.warning(f"⚠️ Se encontraron {len(resultado)} personas en la base que NO están en este formulario de pago.")
            st.dataframe(resultado)
            
            # Generar el Excel para descargar
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                resultado.to_excel(writer, index=False, sheet_name='Pendientes')
            
            st.download_button(
                label="Descargar Excel de Pendientes",
                data=output.getvalue(),
                file_name="Recordatorios_Pago.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            st.error(f"Ocurrió un error al procesar los archivos: {e}")

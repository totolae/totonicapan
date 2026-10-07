import streamlit as st

from config import DEPARTAMENTO, MUNICIPIOS_OPCIONES, EXCEL_MAPPINGS
from extraccion import procesar_pdf
from clasificacion import clasificar_factura
from reporte_excel import ReporteMAGA, PlantillaInvalidaError

# --- TRUCO CSS PARA TRADUCIR LA INTERFAZ A ESPAÑOL ---
st.markdown("""
    <style>
        div[data-testid="stFileUploader"] label p {
            font-size: 40px !important;
        }
    </style>
""", unsafe_allow_html=True)

# --- WEB UI ---
st.title(f"🇬🇹 MAGA: Procesador de Facturas por la LAE: {DEPARTAMENTO}")

selected_municipio = st.selectbox(
    label='1. Seleccione el Municipio de las facturas',
    options=["-- Seleccionar municipio --"] + list(MUNICIPIOS_OPCIONES.keys()),
    help="Todas las facturas que suba deben corresponder a este municipio"
)

uploaded_pdfs = st.file_uploader(label='2. Seleccione sus Facturas (PDFs)', type='pdf', accept_multiple_files=True)
uploaded_xlsx = st.file_uploader(label='3. Seleccione su Archivo de Excel', type='xlsx')

municipio_valido = selected_municipio != "-- Seleccionar municipio --"

if municipio_valido:
    st.info(f"📍 Municipio seleccionado: **{selected_municipio}**. Asegúrese de que todas las facturas correspondan a este municipio.")
else:
    st.warning("⚠️ Por favor seleccione un municipio antes de iniciar el proceso.")

if st.button("INICIAR PROCESO") and uploaded_pdfs and uploaded_xlsx and municipio_valido:
    try:
        user_m_id = MUNICIPIOS_OPCIONES[selected_municipio]
        user_m_name = selected_municipio

        # Validates the Excel template before reading any PDFs
        reporte = ReporteMAGA(uploaded_xlsx.read(), EXCEL_MAPPINGS)

        facturas_procesadas = []   # Paso 2: these will be saved to the database
        skipped_non_standard = []
        progress_bar = st.progress(0)

        for i, pdf_file in enumerate(uploaded_pdfs):
            factura = procesar_pdf(pdf_file, pdf_file.name)

            if factura is None:
                skipped_non_standard.append(pdf_file.name)
            else:
                clasificar_factura(factura)
                reporte.agregar_factura(factura, user_m_id, user_m_name)
                facturas_procesadas.append(factura)

            progress_bar.progress((i + 1) / len(uploaded_pdfs))

        xlsx_salida = reporte.finalizar()
        new_count = len(facturas_procesadas)
        unmatched_count = reporte.unmatched_count

        success_msg = f"¡Proceso completado! {new_count} facturas procesadas y agregadas al Excel con éxito."
        if unmatched_count > 0:
            success_msg += f"""\n\n⚠️ {unmatched_count} items sin clasificar encontrados. Están en la tercera hoja del archivo de Excel, 'Items sin Clasificar', para revisión manual.
                            Los totales de esos productos no fueron agregados a la cantidad de la primera hoja"""
        st.success(success_msg)

        if skipped_non_standard:
            warning_msg = f"⚠️ **{len(skipped_non_standard)} factura(s) no estándar fueron ignoradas** (proformas, cotizaciones, u otros formatos no oficiales). Estas deben procesarse manualmente:\n\n"
            for pdf_name in skipped_non_standard:
                warning_msg += f"- {pdf_name}\n"
            st.warning(warning_msg)

        st.download_button("Descargar Reporte Final", data=xlsx_salida,
                           file_name="Reporte_MAGA_Actualizado.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except PlantillaInvalidaError as e:
        st.error(str(e))
    except Exception as e:
        st.error(f"Error crítico detectado: {e}")

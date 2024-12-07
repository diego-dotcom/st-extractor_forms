import streamlit as st
import requests
import pandas as pd
from io import BytesIO
import base64
from openpyxl.utils import get_column_letter


def process_files(api_url, uploaded_files):
    try:
        # DataFrame para consolidar los datos
        consolidated_df = pd.DataFrame()

        for uploaded_file in uploaded_files:
            # Enviar el archivo al endpoint de la API
            response = requests.post(
                api_url,
                files={'file': (uploaded_file.name,
                                uploaded_file, 'application/pdf')}
            )

            if response.status_code == 200:
                response_data = response.json()
                informacion = response_data.get('informacion', {})

                # Nombre del período (clave identificadora)
                periodo = informacion.get(
                    'Periodo', uploaded_file.name.split('/')[-1].split('.')[0]
                ).replace("/", "-")

                # Convertir datos a DataFrame temporal
                df_temp = pd.DataFrame(
                    list(informacion.items()), columns=['Dato', periodo])

                # Consolidar con el DataFrame general
                if consolidated_df.empty:
                    consolidated_df = df_temp
                else:
                    consolidated_df = pd.merge(
                        consolidated_df, df_temp, on='Dato', how='outer'
                    )
            else:
                raise Exception(
                    f"Error en la API: {response.status_code} {response.text}"
                )

        return consolidated_df

    except Exception as e:
        raise Exception(f"Error procesando archivos: {e}")


def adjust_column_width(writer, sheet_name):
    worksheet = writer.sheets[sheet_name]
    for col_num, column_cells in enumerate(worksheet.columns, start=1):
        max_length = 0
        for cell in column_cells:
            try:
                if cell.value is not None:
                    max_length = max(max_length, len(str(cell.value)))
            except Exception:
                pass
        adjusted_width = max_length + 2
        worksheet.column_dimensions[get_column_letter(
            col_num)].width = adjusted_width


def download_excel(dataframe):
    output_buffer = BytesIO()
    with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
        # Escribir el DataFrame en el archivo Excel
        dataframe.to_excel(writer, index=False, sheet_name="Consolidado")

        # Ajustar ancho de columnas
        adjust_column_width(writer, "Consolidado")

    output_buffer.seek(0)
    b64 = base64.b64encode(output_buffer.getvalue()).decode()
    href = f'<a href="data:application/octet-stream;base64,{b64}" download="resultados_consolidados.xlsx">Descargar Excel</a>'
    return href


def main():
    st.title("App Extractor de Datos")

    st.write(
        "Seleccione un tipo de formulario y cargue sus archivos PDF. Al finalizar, podrá descargar el archivo Excel consolidado."
    )

    # Opciones de API
    api_urls = {
        "F731 / F810": "https://api-extractor-formularios-ddjj.mrbot.com.ar/extraerF731/",
        "F931": "https://api-extractor-formularios-ddjj.mrbot.com.ar/extraerF931/",
        "F2002": "https://api-extractor-formularios-ddjj.mrbot.com.ar/extraerF2002/",
    }

    # Seleccionar el endpoint
    option = st.selectbox(
        "Seleccionar tipo de formulario", list(api_urls.keys()))
    api_url = api_urls[option]

    # Subir archivos
    uploaded_files = st.file_uploader(
        "Subir archivos PDF", accept_multiple_files=True, type=["pdf"]
    )

    if st.button("Procesar Archivos") and uploaded_files:
        try:
            # Procesar los archivos y obtener el DataFrame consolidado
            consolidated_df = process_files(api_url, uploaded_files)

            # Mostrar el DataFrame en Streamlit
            st.write("Vista previa del Excel:")
            st.dataframe(consolidated_df)

            # Generar enlace para descargar el Excel
            st.markdown(download_excel(consolidated_df),
                        unsafe_allow_html=True)

        except Exception as e:
            st.error(f"Error al procesar los archivos: {e}")

    # Créditos
    st.markdown("---")
    st.markdown("Desarrollado por Diego Mendizábal y Agustín Bustos Piasentini")
    st.markdown(
        "Todo el crédito para la [API](https://api-extractor-formularios-ddjj.mrbot.com.ar/docs) de Agustín")
    st.markdown(
        "Cafecito ☕️ [https://cafecito.app/abustos](https://cafecito.app/abustos)")


if __name__ == "__main__":
    main()

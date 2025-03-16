import streamlit as st
import pandas as pd

# Configurar la página en modo ancho
st.set_page_config(page_title="Gestor de Casos (BlueStars)", layout="wide")

# Fondo negro y texto blanco + bordes para tabla
st.markdown("""
<style>
body { background-color: black; color: white; }
.stApp { background-color: black; color: white; }
[data-testid="stSidebar"] { background-color: #222222; }
h1, h2, h3, h4, h5, h6, label, p, div, span { color: white !important; }
.table-cell {
    border: 1px solid white;
    padding: 5px;
    text-align: center;
}
.table-header {
    border: 2px solid white;
    padding: 8px;
    font-weight: bold;
    background-color: #444;
    text-align: center;
}
</style>
""", unsafe_allow_html=True)


def main():
    st.title("Gestor de Casos (BlueStars)")

    # Subir el archivo .xlsm
    uploaded_file = st.file_uploader("Sube el archivo Excel (.xlsm)", type=["xlsm"])
    if uploaded_file:
        try:
            # Leer la hoja ARMADRE
            df = pd.read_excel(uploaded_file, sheet_name='ARMADRE', engine='openpyxl')
            st.success("Archivo cargado correctamente")

            # ================================
            # Extraer datos usando índices:
            # ================================
            df['CASO'] = df.iloc[:, 16].astype(str)
            df['NUNC'] = df.iloc[:, 16].apply(lambda x: str(x).split('-')[1] if '-' in str(x) else '')
            df['ID'] = df.iloc[:, 4].astype(str)
            df['NUMERO DEL ID'] = df['ID'].apply(lambda x: x.split('/')[0] if '/' in x else x)
            df['TIPO DE EMP'] = df.iloc[:, 7].astype(str)
            df['EMPs'] = df.iloc[:, 10].astype(str)

            lista_casos = df['CASO'].dropna().unique().tolist()
            if not lista_casos:
                st.warning("No se encontraron valores en la columna Q para generar la lista de casos.")
                return

            caso_seleccionado = st.selectbox("Selecciona un CASO:", lista_casos)
            if caso_seleccionado:
                st.subheader(f"Información del CASO: {caso_seleccionado}")
                df_filtrado = df[df['CASO'] == caso_seleccionado].copy()
                df_filtrado.reset_index(drop=True, inplace=True)

                envase_options = ["TTG", "TTR", "TTL", "TTV", "FP", "BP"]

                # Mostrar encabezados con bordes
                columnas_mostrar = ['CASO', 'ID', 'NUMERO DEL ID', 'TIPO DE EMP', 'NUNC', 'EMPs']
                num_columnas = len(columnas_mostrar) + 1  # +1 para 'TIPO ENVASE'

                header_cols = st.columns(num_columnas)
                for idx, col_name in enumerate(columnas_mostrar):
                    header_cols[idx].markdown(f'<div class="table-header">{col_name}</div>', unsafe_allow_html=True)
                header_cols[-1].markdown(f'<div class="table-header">TIPO ENVASE</div>', unsafe_allow_html=True)

                tipo_envase_seleccionado = []

                # Mostrar filas con cuadrícula
                for idx, row in df_filtrado.iterrows():
                    row_cols = st.columns(num_columnas)
                    for j, col_name in enumerate(columnas_mostrar):
                        row_cols[j].markdown(f'<div class="table-cell">{row[col_name]}</div>', unsafe_allow_html=True)

                    selected_envase = row_cols[-1].selectbox(
                        " ",
                        envase_options,
                        key=f"envase_fila_{idx}"
                    )
                    tipo_envase_seleccionado.append(selected_envase)

                # Agregar columna final
                df_filtrado['TIPO ENVASE'] = tipo_envase_seleccionado

                # Mostrar DataFrame final
                st.write("### Resultado final con Tipo de Envase seleccionado:")
                columnas_finales = [
                    'CASO', 'ID', 'NUMERO DEL ID', 'TIPO DE EMP', 'NUNC', 'EMPs', 'TIPO ENVASE'
                ]
                st.dataframe(df_filtrado[columnas_finales], use_container_width=True)

                # ✅ Opción para descargar como Excel
                output = df_filtrado[columnas_finales].to_excel(index=False, engine='openpyxl')
                st.download_button(
                    label="Descargar tabla como Excel",
                    data=output,
                    file_name=f"CASO_{caso_seleccionado}_con_envases.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        except Exception as e:
            st.error(f"Error al leer el archivo: {e}")


if __name__ == "__main__":
    main()

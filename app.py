import streamlit as st
import pandas as pd

# Configurar la página en modo ancho
st.set_page_config(page_title="Gestor de Casos (BlueStars)", layout="wide")

# Fondo negro y texto blanco
st.markdown("""
<style>
body { background-color: black; color: white; }
.stApp { background-color: black; color: white; }
[data-testid="stSidebar"] { background-color: #222222; }
h1, h2, h3, h4, h5, h6, label, p, div, span {
  color: white !important;
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
            # Columna Q (índice 16): CASO (todo el valor)
            df['CASO'] = df.iloc[:, 16].astype(str)

            # También sacamos NUNC (lo que hay después del '-')
            df['NUNC'] = df.iloc[:, 16].apply(
                lambda x: str(x).split('-')[1] if '-' in str(x) else ''
            )

            # Columna E (índice 4): ID completo
            df['ID'] = df.iloc[:, 4].astype(str)
            # NÚMERO DEL ID (antes de '/')
            df['NUMERO DEL ID'] = df['ID'].apply(
                lambda x: x.split('/')[0] if '/' in x else x
            )

            # Columna H (índice 7): TIPO DE EMP
            df['TIPO DE EMP'] = df.iloc[:, 7].astype(str)

            # Columna K (índice 10): EMPs
            df['EMPs'] = df.iloc[:, 10].astype(str)

            # Lista de CASOS
            lista_casos = df['CASO'].dropna().unique().tolist()
            if not lista_casos:
                st.warning("No se encontraron valores en la columna Q para generar la lista de casos.")
                return

            # Seleccionar un CASO
            caso_seleccionado = st.selectbox("Selecciona un CASO:", lista_casos)
            if caso_seleccionado:
                st.subheader(f"Información del CASO: {caso_seleccionado}")
                # Filtrar el DataFrame
                df_filtrado = df[df['CASO'] == caso_seleccionado].copy()

                # Reseteamos el índice para iterar más fácilmente
                df_filtrado.reset_index(drop=True, inplace=True)

                # Opciones de envase
                envase_options = ["TTG", "TTR", "TTL", "TTV", "FP", "BP"]

                # Mostrar tabla con selectbox integrado en cada fila
                st.write("### Tabla de datos con selección de tipo de envase:")
                columnas_mostrar = [
                    'CASO',
                    'ID',
                    'NUMERO DEL ID',
                    'TIPO DE EMP',
                    'NUNC',
                    'EMPs'
                ]

                # Encabezados
                cols = st.columns(len(columnas_mostrar) + 1)  # +1 para el selectbox de 'TIPO ENVASE'
                for idx, col_name in enumerate(columnas_mostrar):
                    cols[idx].markdown(f"**{col_name}**")
                cols[-1].markdown("**TIPO ENVASE**")

                # Fila por fila con selectbox
                tipo_envase_seleccionado = []  # Para guardar lo que el usuario selecciona
                for idx, row in df_filtrado.iterrows():
                    cols = st.columns(len(columnas_mostrar) + 1)
                    for j, col_name in enumerate(columnas_mostrar):
                        cols[j].markdown(str(row[col_name]))
                    
                    # Selectbox en la misma fila
                    selected_envase = cols[-1].selectbox(
                        " ",
                        envase_options,
                        key=f"envase_fila_{idx}"
                    )
                    tipo_envase_seleccionado.append(selected_envase)

                # Agregar la columna "TIPO ENVASE" al DataFrame filtrado
                df_filtrado['TIPO ENVASE'] = tipo_envase_seleccionado

                # Mostrar DataFrame final
                st.write("### Resultado final con Tipo de Envase seleccionado:")
                columnas_finales = [
                    'CASO',
                    'ID',
                    'NUMERO DEL ID',
                    'TIPO DE EMP',
                    'NUNC',
                    'EMPs',
                    'TIPO ENVASE'
                ]
                st.dataframe(df_filtrado[columnas_finales], use_container_width=True)

        except Exception as e:
            st.error(f"Error al leer el archivo: {e}")


if __name__ == "__main__":
    main()

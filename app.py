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

                st.write("### Selecciona un tipo de envase para cada fila:")
                envase_options = ["TTG", "TTR", "TTL", "TTV", "FP", "BP"]

                # Creamos un selectbox por cada fila
                for idx, row in df_filtrado.iterrows():
                    # Si ya elegimos algo antes para esta fila, lo recordamos; si no, usamos la primera opción
                    default_value = st.session_state.get(f"envase_{idx}", envase_options[0])
                    chosen_envase = st.selectbox(
                        f"Tipo de Envase (Fila {idx+1})",
                        envase_options,
                        key=f"envase_select_{idx}",
                        index=envase_options.index(default_value) if default_value in envase_options else 0
                    )
                    # Guardamos la selección en session_state para persistir
                    st.session_state[f"envase_{idx}"] = chosen_envase

                # Ahora creamos la columna "TIPO ENVASE" en df_filtrado
                df_filtrado["TIPO ENVASE"] = [
                    st.session_state.get(f"envase_{idx}", envase_options[0]) for idx in df_filtrado.index
                ]

                # Columnas finales a mostrar
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

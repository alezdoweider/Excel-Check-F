import streamlit as st
import pandas as pd

# Configurar la página en modo ancho
st.set_page_config(page_title="Gestor de Casos (BlueStars)", layout="wide")

# Fondo azul cielo, texto verde oliva y estilos para la tabla con autoajuste y bordes
st.markdown("""
<style>
body { 
    background-color: #87CEEB;  /* Azul cielo */
    color: #808000;  /* Verde oliva */
}
.stApp { 
    background-color: #87CEEB;  /* Azul cielo */
    color: #808000;  /* Verde oliva */
}
[data-testid="stSidebar"] { 
    background-color: #222222; 
}
h1, h2, h3, h4, h5, h6, label, p, div, span {
    color: #808000 !important;  /* Verde oliva */
}

/* Estilos para la cuadrícula de la tabla */
.table-cell {
    border: 1px solid #808000;  /* Bordes en verde oliva */
    padding: 5px;
    text-align: center;
    word-wrap: break-word;
    white-space: normal;
}
.table-header {
    border: 2px solid #808000;  /* Bordes en verde oliva */
    padding: 8px;
    font-weight: bold;
    background-color: #444;  /* Fondo oscuro para las cabeceras */
    text-align: center;
    word-wrap: break-word;
    white-space: normal;
}
</style>
""", unsafe_allow_html=True)

def main():
    st.title("Gestor de Casos (BlueStars)")

    # Subir el archivo .xlsm
    uploaded_file = st.file_uploader("Sube el archivo Excel (.xlsm)", type=["xlsm"])
    if uploaded_file:
        try:
            # Leer la hoja "ARMADRE" usando openpyxl
            df = pd.read_excel(uploaded_file, sheet_name='ARMADRE', engine='openpyxl')
            st.success("Archivo cargado correctamente")

            # Extraer datos usando índices numéricos:
            # Columna Q (índice 16): 
            # - CASO: solo la parte antes del guion
            # - NUNC: la parte después del guion (si existe)
            df['CASO'] = df.iloc[:, 16].astype(str).apply(lambda x: x.split('-')[0])
            df['NUNC'] = df.iloc[:, 16].apply(lambda x: str(x).split('-')[1] if '-' in str(x) else '')

            # Columna E (índice 4): ID completo
            df['ID'] = df.iloc[:, 4].astype(str)

            # Columna F (índice 5): Nro. ID (convertido a número)
            df['Nro. ID'] = pd.to_numeric(df.iloc[:, 5], errors='coerce')

            # Columna H (índice 7): TIPO DE EMP
            df['TIPO DE EMP'] = df.iloc[:, 7].astype(str)

            # Columna K (índice 10): NOMBRE
            df['NOMBRE'] = df.iloc[:, 10].astype(str)

            # Lista de CASOS únicos (según la columna 'CASO')
            lista_casos = df['CASO'].dropna().unique().tolist()
            if not lista_casos:
                st.warning("No se encontraron valores en la columna Q para generar la lista de casos.")
                return

            # Seleccionar un CASO
            caso_seleccionado = st.selectbox("Selecciona un CASO:", lista_casos)
            if caso_seleccionado:
                st.subheader(f"Información del CASO: {caso_seleccionado}")
                df_filtrado = df[df['CASO'] == caso_seleccionado].copy()
                df_filtrado.reset_index(drop=True, inplace=True)

                # Opciones de envase disponibles
                envase_options = ["TTG", "TTR", "TTL", "TTV", "FP", "BP"]

                # Definir encabezados de la tabla (más la columna de TIPO ENVASE)
                columnas_mostrar = ['CASO', 'ID', 'Nro. ID', 'TIPO DE EMP', 'NUNC', 'NOMBRE']
                num_columnas = len(columnas_mostrar) + 1  # +1 para "TIPO ENVASE"

                # Mostrar encabezados con bordes
                header_cols = st.columns(num_columnas)
                for idx, col_name in enumerate(columnas_mostrar):
                    header_cols[idx].markdown(f'<div class="table-header">{col_name}</div>', unsafe_allow_html=True)
                header_cols[-1].markdown(f'<div class="table-header">TIPO ENVASE</div>', unsafe_allow_html=True)

                # Mostrar cada fila con su selectbox integrado en la última columna
                tipo_envase_seleccionado = []
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

                # Agregar la columna "TIPO ENVASE" con las selecciones del usuario
                df_filtrado['TIPO ENVASE'] = tipo_envase_seleccionado

                # Mostrar DataFrame final con todas las columnas deseadas
                columnas_finales = [
                    'CASO',
                    'ID',
                    'Nro. ID',
                    'TIPO DE EMP',
                    'NUNC',
                    'NOMBRE',
                    'TIPO ENVASE'
                ]
                
                # Mostrar la tabla final con un ajuste para la columna 'Nro. ID' (índice 5)
                st.write("### Resultado final con Tipo de Envase seleccionado:")
                st.write(df_filtrado[columnas_finales].to_string(index=False))

                # Opción para descargar el DataFrame final como Excel
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

# Cargar el archivo Excel
archivo = 'BlueStars.xlsm'  # Aquí pones la ruta de tu archivo
wb = openpyxl.load_workbook(archivo)

# Seleccionar las hojas HT y LCH
hoja_HT = wb['HT']
hoja_LCH = wb['LCH']

# Aquí defines el valor de AÑO, MES, DÍA y CUSTODIO (estos pueden ser ingresados por la interfaz)
AÑO = 2025
MES = 1
DIA = 3
CUSTODIO = "Juan Pérez"  # Ejemplo de valor para CUSTODIO

# Formatear MES y DIA para tener siempre dos dígitos
MES = str(MES).zfill(2)  # Asegura que MES tenga dos dígitos (ejemplo: 01)
DIA = str(DIA).zfill(2)  # Asegura que DIA tenga dos dígitos (ejemplo: 03)

# Asignar los valores de AÑO, MES y DÍA a las celdas correspondientes en la hoja HT
hoja_HT['AA7'] = AÑO  # Celda AA7 para el AÑO
hoja_HT['AE7'] = MES  # Celda AE7 para el MES
hoja_HT['AG7'] = DIA  # Celda AG7 para el DÍA

# Asignar el valor de CUSTODIO en las celdas correspondientes en HT y LCH
hoja_HT['AE11'] = CUSTODIO  # Celda AE11 para CUSTODIO en la hoja HT
hoja_LCH['Q6'] = CUSTODIO   # Celda Q6 para CUSTODIO en la hoja LCH

# Guardar los cambios en el archivo
wb.save(archivo)

print("Los valores han sido guardados exitosamente.")

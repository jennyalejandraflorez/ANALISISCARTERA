
# =========================
# PROGRAMA: Análisis de Cartera - EDEQ
# AUTOR: Jenny Alejandra Flórez
# DESCRIPCIÓN: Aplicación en Streamlit para análisis interactivo de cartera
# =========================

# =========================
# LIBRERÍAS
# =========================
import streamlit as st
import pandas as pd
import io
import plotly.express as px
import plotly.io as pio
import tempfile

# =========================
# CONFIGURACIÓN DE PÁGINA
# =========================
st.set_page_config(page_title="Análisis de Cartera - EDEQ", layout="wide")

# =========================
# ENCABEZADO CON LOGO Y AUTOR
# =========================
col_logo, col_title = st.columns([1, 4])
with col_logo:
    st.image("logoEDEQ.png", width=180)
with col_title:
    st.title("📊 Análisis de Cartera - EDEQ")
    st.markdown("Bienvenido al Sistema Interactivo de Análisis de cartera.")
    st.caption("Autor del programa: **Jenny Alejandra Flórez**")

# =========================
# CARGAR ARCHIVO EXCEL
# =========================
archivo = st.file_uploader("📂 Seleccione el archivo Excel", type=["xlsx"])

# =========================
# FUNCIÓN PARA LEER Y LIMPIAR DATOS
# =========================
@st.cache_data
def cargar_datos(archivo):
    df = pd.read_excel(archivo, skiprows=5, header=0)
    df.columns = df.columns.str.strip()

    # ✅ Reemplazos para CLASE_SERVICIO
    reemplazos_servicio = {
        '1': 'Residencial', '2': 'Comercial', '3': 'Industrial', '4': 'Oficial',
        '5': 'Alumbrado Público', '6': 'Especial', '7': 'Provisional',
        '10': 'Área Común', '11': 'Autoconsumo'
    }
    df['CLASE_SERVICIO'] = df['CLASE_SERVICIO'].astype(str).str.strip().replace(reemplazos_servicio)

    # ✅ Reemplazos para EDAD
    reemplazos_edad = {
        '0': 'Corriente', '1': '1 a 30 Días', '2': '31 a 60 Días', '3': '61 a 90 Días',
        '4': '91 a 120 Días', '5': '121 a 150 Días', '6': '151 a 180 Días'
    }
    df['EDAD'] = df['EDAD'].astype(str).str.strip().replace(reemplazos_edad)
    df['EDAD'] = df['EDAD'].apply(lambda x: '> a 180 Días' if x not in reemplazos_edad.values() else x)

    # ✅ Reemplazos para CODIGO_DEL_SERVICIO
    reemplazos_codigo = {'601': 'Energía', '603': 'Somos', '609': 'Terceros'}
    df['CODIGO_DEL_SERVICIO'] = df['CODIGO_DEL_SERVICIO'].astype(str).str.strip().replace(reemplazos_codigo)

    return df

# =========================
# FUNCIÓN PARA FORMATEAR VALORES
# =========================
def formato_pesos(valor):
    """Convierte un número en formato pesos con separador de miles usando punto."""
    return f"${valor:,.0f}".replace(",", ".")

# =========================
# PROCESAMIENTO DE DATOS
# =========================
if archivo:
    df = cargar_datos(archivo)

    # =========================
    # TABS PARA ORGANIZAR VISTAS
    # =========================
    tab1, tab2 = st.tabs(["📋 Análisis Principal", "📈 Informes"])

    # =========================
    # TAB 1: ANÁLISIS PRINCIPAL
    # =========================
    with tab1:
        st.sidebar.header("🔍 Filtros")
        municipio = st.sidebar.multiselect("Municipio", sorted(df["DESCRIPCION_MUNICIPIO"].dropna().unique()))
        ciclo = st.sidebar.multiselect("Ciclo", sorted(df["CICLO"].dropna().unique()))
        clase_servicio = st.sidebar.multiselect("Clase Servicio", sorted(df["CLASE_SERVICIO"].dropna().unique()))
        edad = st.sidebar.multiselect("Edad", sorted(df["EDAD"].dropna().unique()))
        servicio = st.sidebar.multiselect("Servicio", sorted(df["CODIGO_DEL_SERVICIO"].dropna().unique()))
        niu = st.sidebar.selectbox("NIU (opcional)", [""] + sorted(df["NIU"].dropna().astype(str).unique()))
        aplicar = st.sidebar.button("✅ Aplicar filtros")

        if aplicar:
            df_filtrado = df.copy()
            if municipio: df_filtrado = df_filtrado[df_filtrado["DESCRIPCION_MUNICIPIO"].isin(municipio)]
            if ciclo: df_filtrado = df_filtrado[df_filtrado["CICLO"].isin(ciclo)]
            if clase_servicio: df_filtrado = df_filtrado[df_filtrado["CLASE_SERVICIO"].isin(clase_servicio)]
            if edad: df_filtrado = df_filtrado[df_filtrado["EDAD"].isin(edad)]
            if servicio: df_filtrado = df_filtrado[df_filtrado["CODIGO_DEL_SERVICIO"].isin(servicio)]
            if niu: df_filtrado = df_filtrado[df_filtrado["NIU"].astype(str) == niu]

            # Agrupar por NIU
            agrupado = df_filtrado.groupby(['NIU']).agg({
                'SALDO_CARTERA': 'sum',
                'NOMBRE': 'first',
                'DIRECCION': 'first',
                'DESCRIPCION_MUNICIPIO': 'first',
                'CICLO': 'first',
                'EDAD': 'first',
                'CODIGO_DEL_SERVICIO': 'first'
            }).reset_index()

            agrupado['SALDO_CARTERA'] = agrupado['SALDO_CARTERA'].apply(formato_pesos)

            columnas_ordenadas = ["NIU", "NOMBRE", "DIRECCION", "DESCRIPCION_MUNICIPIO", "CICLO", "SALDO_CARTERA", "EDAD", "CODIGO_DEL_SERVICIO"]
            agrupado = agrupado[[col for col in columnas_ordenadas if col in agrupado.columns]]

            st.subheader("📈 Resultados")
            # ✅ Alinear SOLO la columna SALDO_CARTERA a la derecha
            styled_df = agrupado.style.set_properties(subset=['SALDO_CARTERA'], **{'text-align': 'right'})
            st.write(styled_df.to_html(), unsafe_allow_html=True)

            # Exportar a Excel
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                agrupado.to_excel(writer, index=False, sheet_name='Resultados')
            st.download_button("📥 Descargar Excel", data=buffer.getvalue(),
                               file_name="analisis_cartera.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # =========================
    # TAB 2: INFORMES
    # =========================
    with tab2:
        st.header("📊 Informes de Cartera")

        filtro_servicio = st.selectbox("Seleccione Servicio", sorted(df["CODIGO_DEL_SERVICIO"].unique()))
        df_filtrado_servicio = df[df["CODIGO_DEL_SERVICIO"] == filtro_servicio]

        # KPIs con formato
        saldo_total = df_filtrado_servicio["SALDO_CARTERA"].sum()
        clientes_unicos = df_filtrado_servicio["NIU"].nunique()
        promedio = saldo_total / clientes_unicos if clientes_unicos > 0 else 0

        col_kpi1, col_kpi2, col_kpi3 = st.columns(3)
        col_kpi1.markdown(f"**Saldo Total:** {formato_pesos(saldo_total)}")
        col_kpi2.markdown(f"**Clientes Únicos:** {clientes_unicos:,}".replace(",", "."))
        col_kpi3.markdown(f"**Promedio por Cliente:** {formato_pesos(promedio)}")

        # Informe 1:(Gráfico existente)
        st.subheader("USUARIOS POR EDADES")
        cartera_por_edad = df_filtrado_servicio.groupby('EDAD').agg({'SALDO_CARTERA': 'sum', 'NIU': 'nunique'}).reset_index()
        cartera_por_edad['SALDO_CARTERA_MILES'] = cartera_por_edad['SALDO_CARTERA'] / 1000

        fig1 = px.bar(cartera_por_edad, x='EDAD', y='SALDO_CARTERA_MILES', text='NIU',
                      title=f"Cartera por Edad ({filtro_servicio})",
                      labels={'SALDO_CARTERA_MILES': 'Saldo (Miles de $)', 'EDAD': 'Edad'},
                      color='EDAD', template='plotly_dark')
        fig1.update_traces(texttemplate='%{text} clientes', textposition='outside')
        st.plotly_chart(fig1, use_container_width=True)

        # =========================
        # NUEVO INFORME: Cartera por EDAD de Mora con valores y exportación a PDF
        # =========================
        st.subheader("SALDO POR EDADES")

        grafico_por_edad = df_filtrado_servicio.groupby('EDAD').agg({'SALDO_CARTERA': 'sum'}).reset_index()
        grafico_por_edad['SALDO_FORMATO'] = grafico_por_edad['SALDO_CARTERA'].apply(formato_pesos)
        total_cartera = grafico_por_edad['SALDO_CARTERA'].sum()

        fig_edad = px.bar(
            grafico_por_edad,
            x='EDAD',
            y='SALDO_CARTERA',
            text='SALDO_FORMATO',  # ✅ Mostrar valores sobre las barras
            title=f"Cartera por EDAD de Mora (Total: {formato_pesos(total_cartera)})",
            labels={'EDAD': 'Edad', 'SALDO_CARTERA': 'Saldo de Cartera ($)'},
            color='EDAD',
            template='plotly_dark'
        )

        fig_edad.update_traces(textposition='outside')
        st.plotly_chart(fig_edad, use_container_width=True)

        # Exportar gráfico a PDF
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmpfile:
            pio.write_image(fig_edad, tmpfile.name, format="pdf")
            pdf_bytes = tmpfile.read()

        st.download_button(
            label="📥 Descargar Gráfico en PDF",
            data=pdf_bytes,
            file_name="cartera_por_edad_de_mora.pdf",
            mime="application/pdf"
        )

        # Informe 2: Top 10 cartera
        st.subheader(f"TOP 10 USUARIOS CARTERA - {filtro_servicio}")
        top10 = df_filtrado_servicio.nlargest(10, 'SALDO_CARTERA')[['NIU', 'NOMBRE', 'SALDO_CARTERA']]
        top10['SALDO_CARTERA'] = top10['SALDO_CARTERA'].apply(formato_pesos)
        st.table(top10.reset_index(drop=True))

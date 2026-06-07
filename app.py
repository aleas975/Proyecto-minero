import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

# Configuración de la página
st.set_page_config(page_title="Dashboard Forecast vs Budget", layout="wide", initial_sidebar_state="expanded")

st.title("📊 Visualización Avanzada: Control de Gestión y Presupuesto")
st.markdown("Plataforma interactiva para el seguimiento del Forecast 5+7 frente a los Budgets consolidados.")

# Panel lateral para cargar datos
with st.sidebar:
    st.header("📂 Carga de Datos")
    uploaded_file = st.file_uploader("Sube tu archivo (Excel o CSV)", type=["csv", "xlsx"])

if uploaded_file:
    # 1. Lectura robusta del archivo
    try:
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file, encoding='latin-1')
        else:
            df = pd.read_excel(uploaded_file)

        # Limpieza y estandarización de nomenclaturas de negocio
        df = df.replace({'Operaciones': 'Trabajador', 'OPERACIONES': 'TRABAJADOR'}, regex=True)

        st.sidebar.success("Base de datos procesada correctamente.")

        # 2. Filtros Interactivos
        st.sidebar.markdown("### ⚙️ Filtros")

        # Buscar columnas categóricas comunes en tus archivos (ej. Classif, Gerencia, VP)
        col_categoria = [col for col in ['Classif', 'Categoría', 'Item', 'Desc Item'] if col in df.columns]

        if col_categoria:
            categoria_seleccionada = st.sidebar.selectbox("Seleccionar Categoría / Clasificación",
                                                          ["Todas"] + list(df[col_categoria[0]].unique()))
            if categoria_seleccionada != "Todas":
                df = df[df[col_categoria[0]] == categoria_seleccionada]

        # 3. Cálculo de KPIs principales (Buscando columnas clave de tus Excels)
        col_forecast = [c for c in df.columns if 'Forecast FY' in c or 'Suma de Forecast' in c]
        col_budget = [c for c in df.columns if 'Budget FY' in c or 'Suma de Budget' in c]
        col_var = [c for c in df.columns if 'Var' in c or 'Varianza' in c]

        if col_forecast and col_budget:
            total_forecast = df[col_forecast[0]].sum()
            total_budget = df[col_budget[0]].sum()
            total_var = df[col_var[0]].sum() if col_var else (total_forecast - total_budget)

            st.markdown("### 📈 KPIs Globales del Periodo")
            kpi1, kpi2, kpi3 = st.columns(3)
            kpi1.metric("Proyección (Forecast)", f"$ {total_forecast:,.0f}")
            kpi2.metric("Presupuesto (Budget)", f"$ {total_budget:,.0f}")
            kpi3.metric("Varianza (Desviación)", f"$ {total_var:,.0f}", delta_color="inverse")
            st.divider()

        # 4. Visualización 1: Comparativa Categorías (Barras)
        if col_categoria and col_forecast and col_budget:
            st.subheader("Análisis de Varianza por Ítem / Categoría")

            # Agrupar datos para el gráfico
            df_grouped = df.groupby(col_categoria[0])[[col_forecast[0], col_budget[0]]].sum().reset_index()

            fig_bar = go.Figure()
            fig_bar.add_trace(
                go.Bar(x=df_grouped[col_categoria[0]], y=df_grouped[col_budget[0]], name='Budget Aprobado',
                       marker_color='#1f77b4'))
            fig_bar.add_trace(
                go.Bar(x=df_grouped[col_categoria[0]], y=df_grouped[col_forecast[0]], name='Forecast Proyectado',
                       marker_color='#ff7f0e'))

            fig_bar.update_layout(barmode='group', xaxis_title="Categoría", yaxis_title="Monto ($)")
            st.plotly_chart(fig_bar, use_container_width=True)

        # 5. Visualización 2: Tendencia Mensual (Líneas)
        # Identificar columnas de meses (Jan-26, Feb-26, etc.)
        meses_cols = [c for c in df.columns if any(
            mes in c for mes in ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'])]

        if meses_cols:
            st.subheader("Curva de Ejecución Mensual (Forecast 5+7)")
            df_meses = df[meses_cols].sum().reset_index()
            df_meses.columns = ['Mes', 'Gasto Proyectado']

            fig_line = px.line(df_meses, x='Mes', y='Gasto Proyectado', markers=True,
                               title="Distribución del Gasto a lo largo del año")

            # Añadir una línea plana promedio del budget si existe
            if col_budget:
                promedio_mensual = total_budget / 12
                fig_line.add_hline(y=promedio_mensual, line_dash="dot", annotation_text="Budget Promedio Mensual",
                                   annotation_position="bottom right", line_color="red")

            fig_line.update_traces(line=dict(width=3))
            st.plotly_chart(fig_line, use_container_width=True)

        # 6. Tabla de datos sin procesar
        st.subheader("Datos Detallados")
        st.dataframe(df, use_container_width=True)

    except Exception as e:
        st.error(f"Error procesando el archivo. Asegúrate de que el formato coincida. Detalle técnico: {e}")

else:
    st.info("👈 Por favor, carga un archivo Excel o CSV desde el panel lateral para iniciar el análisis visual.")
import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px

st.set_page_config(page_title="Control de Gestión: Forecast vs Budget", layout="wide")

st.title("📊 Control de Gestión Minera: Forecast 5+7 vs Budget")
st.markdown("Análisis granular de variaciones presupuestarias indexado por Centro de Costo (CC).")

# 1. Carga del archivo Excel
with st.sidebar:
    st.header("📂 Ingesta de Datos")
    uploaded_file = st.file_uploader("Sube el archivo Excel del Proyecto", type=["xlsx"])

if uploaded_file:
    try:
        # Leer estructura completa de hojas sin cargar datos en memoria aún
        xls = pd.ExcelFile(uploaded_file)
        hojas_disponibles = xls.sheet_names

        st.sidebar.markdown("### 🗺️ Mapeo de Estructura")
        hoja_f = st.sidebar.selectbox("Selecciona la hoja de FORECAST", hojas_disponibles, index=0)
        hoja_b = st.sidebar.selectbox("Selecciona la hoja de BUDGET", hojas_disponibles,
                                      index=min(1, len(hojas_disponibles) - 1))

        # Cargar hojas seleccionadas
        df_forecast = pd.read_excel(uploaded_file, sheet_name=hoja_f)
        df_budget = pd.read_excel(uploaded_file, sheet_name=hoja_b)

        # Limpieza de nombres de columnas (eliminar espacios ocultos)
        df_forecast.columns = df_forecast.columns.str.strip()
        df_budget.columns = df_budget.columns.str.strip()

        # Estandarización de nomenclaturas de portales (regla de negocio invisible)
        for df in [df_forecast, df_budget]:
            if 'Gerencia' in df.columns:
                df['Gerencia'] = df['Gerencia'].astype(str).str.strip().replace(
                    {'Operaciones': 'Trabajador', 'OPERACIONES': 'TRABAJADOR'})

        # 2. Definición de dimensiones temporales (Meses del periodo)
        meses = ['Jan-26', 'Feb-26', 'Mar-26', 'Apr-26', 'May-26', 'Jun-26', 'Jul-26', 'Aug-26', 'Sep-26', 'Oct-26',
                 'Nov-26', 'Dec-26']

        # Validar presencia de columnas críticas
        if 'CC' in df_forecast.columns and 'CC' in df_budget.columns:

            # 3. Preparación de DataFrames para cruce mensual
            # Agrupar por CC para asegurar consolidación limpia antes del merge
            columnas_dimensiones = ['CC', 'Classif', 'Gerencia', 'Desc Item']
            dims_f = [c for c in columnas_dimensiones if c in df_forecast.columns]
            dims_b = [c for c in columnas_dimensiones if c in df_budget.columns]

            df_f_melt = df_forecast.groupby(dims_f)[meses].sum().reset_index()
            df_b_melt = df_budget.groupby(['CC'])[meses].sum().reset_index()

            # Renombrar meses en budget para evitar colisiones
            meses_b_mapping = {m: f"{m}_Bud" for m in meses}
            df_b_melt = df_b_melt.rename(columns=meses_b_mapping)

            # Cruce relacional por llave única (CC)
            df_merged = pd.merge(df_f_melt, df_b_melt, on='CC', how='inner')

            # 4. Panel de Filtros en Barra Lateral
            st.sidebar.markdown("### ⚙️ Filtros Operacionales")

            filtro_classif = "Todas"
            if 'Classif' in df_merged.columns:
                filtro_classif = st.sidebar.selectbox("Clasificación de Costo (Classif)",
                                                      ["Todas"] + list(df_merged['Classif'].dropna().unique()))
                if filtro_classif != "Todas":
                    df_merged = df_merged[df_merged['Classif'] == filtro_classif]

            filtro_gerencia = "Todas"
            if 'Gerencia' in df_merged.columns:
                filtro_gerencia = st.sidebar.selectbox("Gerencia",
                                                       ["Todas"] + list(df_merged['Gerencia'].dropna().unique()))
                if filtro_gerencia != "Todas":
                    df_merged = df_merged[df_merged['Gerencia'] == filtro_gerencia]

            # 5. Cómputo de Series Mensuales Agregadas
            valores_forecast = [df_merged[m].sum() for m in meses]
            valores_budget = [df_merged[f"{m}_Bud"].sum() for m in meses]
            varianzas_mensuales = [f - b for f, b in zip(valores_forecast, valores_budget)]

            # Acumulados
            sum_forecast_fy = sum(valores_forecast)
            sum_budget_fy = sum(valores_budget)
            varianza_total = sum_forecast_fy - sum_budget_fy

            # 6. Despliegue de Indicadores Financieros (KPIs)
            st.markdown(f"### 🎯 Desempeño Financiero Comercial — Vista: {filtro_classif} / Gerencia: {filtro_gerencia}")
            kpi1, kpi2, kpi3 = st.columns(3)
            kpi1.metric("Forecast Anual Proyectado (5+7)", f"$ {sum_forecast_fy:,.2f}")
            kpi2.metric("Budget Original Aprobado", f"$ {sum_budget_fy:,.2f}")
            kpi3.metric("Varianza Neta del Periodo", f"$ {varianza_total:,.2f}", delta_color="inverse")

            st.divider()

            # 7. Visualización Interactiva Multi-Eje (Plotly)
            st.subheader("📈 Curva de Desviación Presupuestaria y Tendencia Mensual")

            fig = go.Figure()
            # Columnas barras de presupuesto original
            fig.add_trace(go.Bar(x=meses, y=valores_budget, name="Budget", marker_color="#1f77b4", opacity=0.65))
            # Columnas barras de forecast proyectado
            fig.add_trace(
                go.Bar(x=meses, y=valores_forecast, name="Forecast 5+7", marker_color="#ff7f0e", opacity=0.85))
            # Línea de comportamiento de varianza mensual
            fig.add_trace(
                go.Scatter(x=meses, y=varianzas_mensuales, name="Varianza Mensual (F - B)", mode="lines+markers",
                           line=dict(color="#d62728", width=3, dash="solid"), yaxis="y2"))

            fig.update_layout(
                barmode="group",
                hovermode="x unified",
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                yaxis=dict(title="Montos Financieros ($)"),
                yaxis2=dict(title="Varianza Neta por Mes ($)", overlaying="y", side="right", showgrid=False),
                margin=dict(l=40, r=40, t=40, b=40)
            )
            st.plotly_chart(fig, use_container_width=True)

            # 8. Desglose Detallado por Ítem de Gasto
            st.subheader("📋 Matriz de Desviación por Centro de Costo")
            df_tabla = df_merged.copy()
            # Calcular varianza agregada por fila para visualización en tabla
            df_tabla['Varianza Total'] = 0.0
            for m in meses:
                df_tabla['Varianza Total'] += (df_tabla[m] - df_tabla[f"{m}_Bud"])

            columnas_visibles = [c for c in ['CC', 'Classif', 'Gerencia', 'Desc Item'] if c in df_tabla.columns] + [
                'Varianza Total']
            st.dataframe(df_tabla[columnas_visibles].sort_values(by='Varianza Total', ascending=True),
                         use_container_width=True)

        else:
            st.error(
                "Error estructural: No se encontró la columna clave 'CC' en ambas hojas para realizar la unión relacional.")

    except Exception as e:
        st.error(f"Error crítico de procesamiento: {str(e)}")
else:
    st.info(
        "👈 Esperando carga de la planilla de control de gestión minera (.xlsx) para inicializar el pipeline analítico.")
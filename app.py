import streamlit as st
import pandas as pd
import plotly.graph_objects as go

st.set_page_config(page_title="Control de Gestión: Forecast vs Budget", layout="wide")

st.title("📊 Control de Gestión Minera: Forecast vs Budget")
st.markdown("Análisis ejecutivo con detección dinámica de meses por año fiscal (ej. Jan-24, Jan-26).")

# 1. Carga del archivo Excel
with st.sidebar:
    st.header("📂 Ingesta de Datos")
    uploaded_file = st.file_uploader("Sube el archivo Excel del Proyecto", type=["xlsx"])

if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)
        hojas_disponibles = xls.sheet_names

        st.sidebar.markdown("### 🗺️ Mapeo de Estructura")
        hoja_f = st.sidebar.selectbox("Selecciona la hoja de FORECAST", hojas_disponibles, index=0)
        hoja_b = st.sidebar.selectbox("Selecciona la hoja de BUDGET", hojas_disponibles,
                                      index=min(1, len(hojas_disponibles) - 1))

        df_forecast = pd.read_excel(uploaded_file, sheet_name=hoja_f)
        df_budget = pd.read_excel(uploaded_file, sheet_name=hoja_b)

        df_forecast.columns = df_forecast.columns.str.strip()
        df_budget.columns = df_budget.columns.str.strip()

        # Estandarización de nomenclaturas de portales de reporte
        for df in [df_forecast, df_budget]:
            if 'Gerencia' in df.columns:
                df['Gerencia'] = df['Gerencia'].astype(str).str.strip().replace(
                    {'Operaciones': 'Trabajador', 'OPERACIONES': 'TRABAJADOR'})

        # DETECCIÓN DINÁMICA DE COLUMNAS DE MESES
        abrev_meses = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']

        meses_f_map = {}
        for m in abrev_meses:
            match = [c for c in df_forecast.columns if c.startswith(m)]
            if match:
                meses_f_map[m] = match[0]

        meses_b_map = {}
        for m in abrev_meses:
            match = [c for c in df_budget.columns if c.startswith(m)]
            if match:
                meses_b_map[m] = match[0]

        meses_validos = [m for m in abrev_meses if m in meses_f_map and m in meses_b_map]

        if not meses_validos:
            st.error("No se detectaron columnas de meses válidas (ej. Jan-24, Jan-26) en las hojas seleccionadas.")
        else:
            columnas_meses_f = [meses_f_map[m] for m in meses_validos]
            columnas_meses_b = [meses_b_map[m] for m in meses_validos]

            if 'CC' in df_forecast.columns and 'CC' in df_budget.columns:
                columnas_dimensiones = ['CC', 'Classif', 'Gerencia', 'Desc Item']
                dims_f = [c for c in columnas_dimensiones if c in df_forecast.columns]

                df_f_melt = df_forecast.groupby(dims_f)[columnas_meses_f].sum().reset_index()
                df_b_melt = df_budget.groupby(['CC'])[columnas_meses_b].sum().reset_index()

                # Renombrar temporalmente el budget para el cruce limpio
                meses_b_rename = {meses_b_map[m]: f"{m}_Bud" for m in meses_validos}
                df_b_melt = df_b_melt.rename(columns=meses_b_rename)

                df_merged = pd.merge(df_f_melt, df_b_melt, on='CC', how='inner')

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

                # Cómputo usando el mapa dinámico detectado
                valores_forecast = [df_merged[meses_f_map[m]].sum() for m in meses_validos]
                valores_budget = [df_merged[f"{m}_Bud"].sum() for m in meses_validos]

                # Variación clásica: Forecast - Budget
                varianzas_mensuales = [f - b for f, b in zip(valores_forecast, valores_budget)]

                sum_forecast_m = sum(valores_forecast) / 1_000_000
                sum_budget_m = sum(valores_budget) / 1_000_000
                varianza_total_m = sum_forecast_m - sum_budget_m

                # Despliegue de KPIs en Millones (M$)
                st.markdown(f"### 🎯 Desempeño Financiero Faena — Vista: {filtro_classif} / Gerencia: {filtro_gerencia}")
                kpi1, kpi2, kpi3 = st.columns(3)

                kpi1.metric("Forecast Anual Proyectado (5+7)", f"$ {sum_forecast_m:,.2f} M")
                kpi2.metric("Budget Original Aprobado", f"$ {sum_budget_m:,.2f} M")

                if varianza_total_m >= 0:
                    kpi3.metric("Varianza Total", f"$ {varianza_total_m:,.2f} M",
                                delta=f"$ {varianza_total_m:,.2f} M (Por sobre Budget)", delta_color="inverse")
                else:
                    kpi3.metric("Varianza Total", f"$ {varianza_total_m:,.2f} M",
                                delta=f"$ {abs(varianza_total_m):,.2f} M (Bajo Budget)", delta_color="normal")

                st.divider()

                # Gráfico interactivo adaptivo
                st.subheader("📈 Curva de Tendencia Mensual y Simetría de Variaciones")

                fig = go.Figure()
                fig.add_trace(
                    go.Bar(x=meses_validos, y=valores_budget, name="Budget", marker_color="#1f77b4", opacity=0.65))
                fig.add_trace(go.Bar(x=meses_validos, y=valores_forecast, name="Forecast 5+7", marker_color="#ff7f0e",
                                     opacity=0.85))
                fig.add_trace(
                    go.Scatter(x=meses_validos, y=varianzas_mensuales, name="Varianza Mensual (Forecast - Budget)",
                               mode="lines+markers",
                               line=dict(color="#d62728" if varianza_total_m >= 0 else "#2ca02c", width=3), yaxis="y2"))

                fig.update_layout(
                    barmode="group",
                    hovermode="x unified",
                    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                    yaxis=dict(title="Montos Absolutos ($)", showgrid=True),
                    yaxis2=dict(title="Varianza Neta (Forecast - Budget) ($)", overlaying="y", side="right",
                                showgrid=False),
                    margin=dict(l=40, r=40, t=40, b=40)
                )
                st.plotly_chart(fig, use_container_width=True)

                # Tabla detallada ordenada por desviación descendente
                st.subheader("📋 Matriz de Desviación por Centro de Costo")
                df_tabla = df_merged.copy()
                df_tabla['Varianza Total'] = 0.0
                for m in meses_validos:
                    df_tabla['Varianza Total'] += (df_tabla[meses_f_map[m]] - df_tabla[f"{m}_Bud"])

                columnas_visibles = [c for c in ['CC', 'Classif', 'Gerencia', 'Desc Item'] if c in df_tabla.columns] + [
                    'Varianza Total']
                st.dataframe(df_tabla[columnas_visibles].sort_values(by='Varianza Total', ascending=False),
                             use_container_width=True)

            else:
                st.error(
                    "Error estructural: No se encontró la columna clave 'CC' en ambas hojas para realizar la unión relacional.")

    except Exception as e:
        st.error(f"Error crítico de procesamiento: {str(e)}")
else:
    st.info(
        "👈 Esperando carga de la planilla de control de gestión minera (.xlsx) para inicializar el pipeline analítico.")
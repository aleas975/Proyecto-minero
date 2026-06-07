import streamlit as st
import pandas as pd
import plotly.graph_objects as go

st.set_page_config(page_title="Control de Gestión: Forecast vs Budget", layout="wide")

st.title("📊 Control de Gestión Minera: Forecast vs Budget")
st.markdown("Plataforma corporativa para el análisis y visualización de desviaciones presupuestarias.")

# 1. Carga del archivo Excel
with st.sidebar:
    st.header("📂 Ingesta de Datos")
    uploaded_file = st.file_uploader("Sube el archivo Excel del Proyecto", type=["xlsx"])

if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)
        hojas_disponibles = xls.sheet_names

        st.sidebar.markdown("### 🗺️ Mapeo de Estructura")
        # Se corrige el orden y los índices para resolver la inversión de casillas
        hoja_b = st.sidebar.selectbox("Selecciona la hoja de BUDGET", hojas_disponibles, index=0)
        hoja_f = st.sidebar.selectbox("Selecciona la hoja de FORECAST", hojas_disponibles,
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

        # Detección dinámica de columnas de meses
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

                # Selector interactivo de fecha/mes de corte para análisis "Hasta Hoy" (YTD dinámico)
                st.sidebar.markdown("### 📅 Análisis Temporal")
                mes_corte = st.sidebar.selectbox("Seleccionar mes de corte (Acumulado hasta hoy):", meses_validos,
                                                 index=len(meses_validos) - 1)

                # Filtrar meses hasta el mes seleccionado
                idx_corte = meses_validos.index(mes_corte)
                meses_hasta_hoy = meses_validos[:idx_corte + 1]

                # Cómputo usando el mapa dinámico detectado
                valores_forecast = [df_merged[meses_f_map[m]].sum() for m in meses_validos]
                valores_budget = [df_merged[f"{m}_Bud"].sum() for m in meses_validos]
                varianzas_mensuales = [f - b for f, b in zip(valores_forecast, valores_budget)]

                # Cómputo acumulado hasta el mes de corte seleccionado ("Hasta hoy")
                valores_forecast_hoy = [df_merged[meses_f_map[m]].sum() for m in meses_hasta_hoy]
                valores_budget_hoy = [df_merged[f"{m}_Bud"].sum() for m in meses_hasta_hoy]
                sum_forecast_hoy_m = sum(valores_forecast_hoy) / 1_000_000
                sum_budget_hoy_m = sum(valores_budget_hoy) / 1_000_000
                varianza_hoy_m = sum_forecast_hoy_m - sum_budget_hoy_m

                # Totales anuales completos (FY)
                sum_forecast_m = sum(valores_forecast) / 1_000_000
                sum_budget_m = sum(valores_budget) / 1_000_000
                varianza_total_m = sum_forecast_m - sum_budget_m

                # 2. Despliegue de Indicadores Financieros (KPIs)
                st.markdown(f"### 🎯 Desempeño Financiero — Vista: {filtro_classif} / Gerencia: {filtro_gerencia}")

                # Fila 1: Totales Anuales Completos (FY)
                st.markdown("#### Totales Anuales Completos (Full Year)")
                col1, col2, col3 = st.columns(3)
                col1.metric("Forecast Anual (FY)", f"$ {sum_forecast_m:,.2f} M")
                col2.metric("Budget Anual Approved", f"$ {sum_budget_m:,.2f} M")
                col3.metric("Varianza Total Anual", f"$ {varianza_total_m:,.2f} M",
                            delta=f"$ {varianza_total_m:,.2f} M",
                            delta_color="inverse" if varianza_total_m >= 0 else "normal")

                st.divider()

                # Fila 2: Análisis de Acumulados "Hasta Hoy" según fecha seleccionada
                st.markdown(f"#### Acumulado Temporal Dinámico (Ene a {mes_corte})")
                col_hoy1, col_hoy2, col_hoy3 = st.columns(3)
                col_hoy1.metric(f"Forecast Acumulado ({mes_corte})", f"$ {sum_forecast_hoy_m:,.2f} M")
                col_hoy2.metric(f"Budget Acumulado ({mes_corte})", f"$ {sum_budget_hoy_m:,.2f} M")
                col_hoy3.metric(f"Varianza Acumulada (Hasta {mes_corte})", f"$ {varianza_hoy_m:,.2f} M",
                                delta=f"$ {varianza_hoy_m:,.2f} M",
                                delta_color="inverse" if varianza_hoy_m >= 0 else "normal")

                st.divider()

                # 3. Pestañas para Gráficos para ordenar la visualización en detalle
                st.subheader("📊 Paneles de Visualización Avanzada")
                tab1, tab2, tab3 = st.tabs(
                    ["Comparativa Combinada", "Comportamiento Individual", "Gráfico de Cascada (Variación)"])

                with tab1:
                    st.markdown("##### Curva de Tendencia Mensual y Simetría de Variaciones")
                    fig = go.Figure()
                    fig.add_trace(
                        go.Bar(x=meses_validos, y=valores_budget, name="Budget", marker_color="#1f77b4", opacity=0.65))
                    fig.add_trace(
                        go.Bar(x=meses_validos, y=valores_forecast, name="Forecast 5+7", marker_color="#ff7f0e",
                               opacity=0.85))
                    fig.add_trace(
                        go.Scatter(x=meses_validos, y=varianzas_mensuales, name="Varianza (Forecast - Budget)",
                                   mode="lines+markers",
                                   line=dict(color="#d62728", width=3), yaxis="y2"))

                    fig.update_layout(
                        barmode="group", hovermode="x unified",
                        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                        yaxis=dict(title="Montos Absolutos ($)", showgrid=True),
                        yaxis2=dict(title="Varianza Neta (Forecast - Budget) ($)", overlaying="y", side="right",
                                    showgrid=False),
                        margin=dict(l=40, r=40, t=40, b=40)
                    )
                    st.plotly_chart(fig, use_container_width=True)

                with tab2:
                    st.markdown("##### Comportamiento Mensual por Separado")
                    col_g1, col_g2 = st.columns(2)

                    with col_g1:
                        fig_f = go.Figure()
                        fig_f.add_trace(go.Scatter(x=meses_validos, y=valores_forecast, mode="lines+markers+text",
                                                   name="Forecast 5+7", line=dict(color="#ff7f0e", width=3)))
                        fig_f.update_layout(title="Tendencia Mensual Pura: Forecast 5+7", xaxis_title="Meses",
                                            yaxis_title="Monto ($)")
                        st.plotly_chart(fig_f, use_container_width=True)

                    with col_g2:
                        fig_b = go.Figure()
                        fig_b.add_trace(go.Scatter(x=meses_validos, y=valores_budget, mode="lines+markers+text",
                                                   name="Budget", line=dict(color="#1f77b4", width=3)))
                        fig_b.update_layout(title="Tendencia Mensual Pura: Budget Aprobado", xaxis_title="Meses",
                                            yaxis_title="Monto ($)")
                        st.plotly_chart(fig_b, use_container_width=True)

                with tab3:
                    st.markdown("##### Gráfico de Cascada (Waterfall): Impacto de las Variaciones Mensuales")

                    # Estructurar el gráfico de cascada
                    fig_waterfall = go.Figure(go.Waterfall(
                        name="Variación",
                        orientation="v",
                        x=meses_validos,
                        textposition="outside",
                        text=[f"${v / 1_000_000:.2f}M" for v in varianzas_mensuales],
                        y=varianzas_mensuales,
                        connector=dict(line=dict(color="rgb(63, 63, 63)", width=1.5)),
                        decreasing=dict(marker=dict(color="#2ca02c")),  # Verde si reduce gasto vs budget (ahorro)
                        increasing=dict(marker=dict(color="#d62728"))  # Rojo si aumenta gasto vs budget (sobregasto)
                    ))

                    fig_waterfall.update_layout(
                        title="Evolución Acumulada de la Varianza (Forecast - Budget)",
                        xaxis_title="Meses del Periodo",
                        yaxis_title="Delta Financiero ($)",
                        showlegend=False
                    )
                    st.plotly_chart(fig_waterfall, use_container_width=True)

                # 4. Tabla detallada
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
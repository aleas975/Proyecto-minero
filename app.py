import streamlit as st
import pandas as pd
import plotly.graph_objects as go

st.set_page_config(page_title="Control de Gestión: Forecast vs Budget", layout="wide")

st.title("📊 Control de Gestión Minera: Forecast vs Budget")
st.markdown("Plataforma corporativa para el análisis y visualización de desviaciones presupuestarias.")

with st.sidebar:
    st.header("📂 Ingesta de Datos")
    uploaded_file = st.file_uploader("Sube el archivo Excel del Proyecto", type=["xlsx"])

if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)
        hojas_disponibles = xls.sheet_names

        st.sidebar.markdown("### 🗺️ Mapeo de Estructura")
        hoja_b = st.sidebar.selectbox("Selecciona la hoja de BUDGET", hojas_disponibles, index=0)
        hoja_f = st.sidebar.selectbox("Selecciona la hoja de FORECAST", hojas_disponibles,
                                      index=min(1, len(hojas_disponibles) - 1))

        df_forecast = pd.read_excel(uploaded_file, sheet_name=hoja_f)
        df_budget = pd.read_excel(uploaded_file, sheet_name=hoja_b)

        df_forecast.columns = df_forecast.columns.str.strip()
        df_budget.columns = df_budget.columns.str.strip()

        for df in [df_forecast, df_budget]:
            if 'Gerencia' in df.columns:
                df['Gerencia'] = df['Gerencia'].astype(str).str.strip().replace(
                    {'Operaciones': 'Trabajador', 'OPERACIONES': 'TRABAJADOR'})

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

            col_fy_budget = None
            if columnas_meses_b:
                nombre_mes_ejemplo = columnas_meses_b[0]
                partes = nombre_mes_ejemplo.split('-')
                if len(partes) > 1:
                    anio = partes[-1].strip()
                    posible_col = f"FY{anio}"
                    if posible_col in df_budget.columns:
                        col_fy_budget = posible_col

            if not col_fy_budget:
                cols_fy = [c for c in df_budget.columns if c.startswith('FY')]
                if cols_fy:
                    col_fy_budget = cols_fy[0]

            if 'CC' in df_forecast.columns and 'CC' in df_budget.columns:
                columnas_dimensiones = ['CC', 'Classif', 'Gerencia', 'Desc Item']
                dims_f = [c for c in columnas_dimensiones if c in df_forecast.columns]

                # CORRECCIÓN: Se copia la lista sin incluir 'CC' para evitar el error "cannot insert CC"
                cols_b_extra = columnas_meses_b.copy()
                if col_fy_budget and col_fy_budget in df_budget.columns:
                    cols_b_extra.append(col_fy_budget)

                df_f_melt = df_forecast.groupby(dims_f)[columnas_meses_f].sum().reset_index()
                df_b_melt = df_budget.groupby('CC')[cols_b_extra].sum().reset_index()

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

                valores_forecast_full = [df_merged[meses_f_map[m]].sum() for m in meses_validos]
                valores_budget_full_list = [df_merged[f"{m}_Bud"].sum() for m in meses_validos]

                if col_fy_budget and filtro_classif == "Todas" and filtro_gerencia == "Todas":
                    sum_budget_m = df_merged[col_fy_budget].sum() / 1_000_000
                else:
                    sum_budget_m = sum(valores_budget_full_list) / 1_000_000

                sum_forecast_m = sum(valores_forecast_full) / 1_000_000
                varianza_total_m = sum_forecast_m - sum_budget_m
                varianzas_mensuales_full = [f - b for f, b in zip(valores_forecast_full, valores_budget_full_list)]

                tab_anual, tab_temporal = st.tabs(
                    ["📅 Análisis Año Completo (Full Year)", "⏳ Análisis Temporal Interactivo"])

                with tab_anual:
                    st.markdown(f"### 🎯 Desempeño Financiero Anual Consolidado")
                    col1, col2, col3 = st.columns(3)
                    col1.metric("Forecast Anual (FY)", f"$ {sum_forecast_m:,.2f} M")
                    col2.metric("Budget Anual Oficial", f"$ {sum_budget_m:,.2f} M")
                    col3.metric("Varianza Total Anual", f"$ {varianza_total_m:,.2f} M",
                                delta=f"$ {varianza_total_m:,.2f} M",
                                delta_color="inverse" if varianza_total_m >= 0 else "normal")

                    st.divider()

                    st.subheader("📊 Gráfico Comparativo Principal")
                    fig_comb = go.Figure()
                    fig_comb.add_trace(
                        go.Bar(x=meses_validos, y=valores_budget_full_list, name="Budget", marker_color="#1f77b4",
                               opacity=0.65))
                    fig_comb.add_trace(
                        go.Bar(x=meses_validos, y=valores_forecast_full, name="Forecast 5+7", marker_color="#ff7f0e",
                               opacity=0.85))
                    fig_comb.add_trace(
                        go.Scatter(x=meses_validos, y=varianzas_mensuales_full, name="Varianza (Forecast - Budget)",
                                   mode="lines+markers",
                                   line=dict(color="#d62728", width=3), yaxis="y2"))
                    fig_comb.update_layout(
                        barmode="group", hovermode="x unified",
                        title="Distribución de Volúmenes y Variación del Periodo",
                        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                        yaxis=dict(title="Montos Absolutos ($)", showgrid=True),
                        yaxis2=dict(title="Varianza Neta ($)", overlaying="y", side="right", showgrid=False)
                    )
                    st.plotly_chart(fig_comb, use_container_width=True)

                with tab_temporal:
                    st.markdown("### ⏳ Simulación Temporal Acumulada (YTD Dinámico)")

                    mes_corte = st.selectbox("Seleccionar mes de corte (Gasto acumulado hasta hoy):", meses_validos,
                                             index=len(meses_validos) - 1)

                    idx_corte = meses_validos.index(mes_corte)
                    meses_hasta_hoy = meses_validos[:idx_corte + 1]

                    valores_forecast_hoy = [df_merged[meses_f_map[m]].sum() for m in meses_hasta_hoy]
                    valores_budget_hoy = [df_merged[f"{m}_Bud"].sum() for m in meses_hasta_hoy]
                    varianzas_mensuales_hoy = [f - b for f, b in zip(valores_forecast_hoy, valores_budget_hoy)]

                    sum_forecast_hoy_m = sum(valores_forecast_hoy) / 1_000_000
                    sum_budget_hoy_m = sum(valores_budget_hoy) / 1_000_000
                    varianza_hoy_m = sum_forecast_hoy_m - sum_budget_hoy_m

                    col_hoy1, col_hoy2, col_hoy3 = st.columns(3)
                    col_hoy1.metric(f"Forecast Acumulado (Ene-{mes_corte})", f"$ {sum_forecast_hoy_m:,.2f} M")
                    col_hoy2.metric(f"Budget Acumulado (Ene-{mes_corte})", f"$ {sum_budget_hoy_m:,.2f} M")
                    col_hoy3.metric(f"Varianza Acumulada (Hasta {mes_corte})", f"$ {varianza_hoy_m:,.2f} M",
                                    delta=f"$ {varianza_hoy_m:,.2f} M",
                                    delta_color="inverse" if varianza_hoy_m >= 0 else "normal")

                    st.divider()

                    st.subheader(f"📊 Desglose Visual Filtrado (Ene a {mes_corte})")
                    gt_col1, gt_col2 = st.columns([2, 1])

                    with gt_col1:
                        fig_comb_hoy = go.Figure()
                        fig_comb_hoy.add_trace(
                            go.Bar(x=meses_hasta_hoy, y=valores_budget_hoy, name="Budget", marker_color="#1f77b4",
                                   opacity=0.65))
                        fig_comb_hoy.add_trace(go.Bar(x=meses_hasta_hoy, y=valores_forecast_hoy, name="Forecast 5+7",
                                                      marker_color="#ff7f0e", opacity=0.85))
                        fig_comb_hoy.add_trace(
                            go.Scatter(x=meses_hasta_hoy, y=varianzas_mensuales_hoy, name="Varianza (F - B)",
                                       mode="lines+markers",
                                       line=dict(color="#d62728", width=3), yaxis="y2"))
                        fig_comb_hoy.update_layout(
                            barmode="group", hovermode="x unified",
                            title=f"Desviación de Líneas de Base hasta {mes_corte}",
                            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                            yaxis=dict(title="Montos Absolutos ($)"),
                            yaxis2=dict(title="Varianza Neta ($)", overlaying="y", side="right", showgrid=False)
                        )
                        st.plotly_chart(fig_comb_hoy, use_container_width=True)

                    with gt_col2:
                        fig_water_hoy = go.Figure(go.Waterfall(
                            name="Variación Acumulada", orientation="v", x=meses_hasta_hoy, textposition="outside",
                            text=[f"${v / 1_000_000:.2f}M" for v in varianzas_mensuales_hoy], y=varianzas_mensuales_hoy,
                            connector=dict(line=dict(color="rgb(63, 63, 63)", width=1.5)),
                            decreasing=dict(marker=dict(color="#2ca02c")), increasing=dict(marker=dict(color="#d62728"))
                        ))
                        fig_water_hoy.update_layout(title=f"Gráfico de Cascada Acumulado ({mes_corte})",
                                                    showlegend=False)
                        st.plotly_chart(fig_water_hoy, use_container_width=True)

                    st.markdown("##### Comportamiento de Tendencias Individuales")
                    col_sep1, col_sep2 = st.columns(2)
                    with col_sep1:
                        fig_f_hoy = go.Figure(
                            go.Scatter(x=meses_hasta_hoy, y=valores_forecast_hoy, mode="lines+markers", name="Forecast",
                                       line=dict(color="#ff7f0e", width=3)))
                        fig_f_hoy.update_layout(title="Curva de Ejecución Pura: Forecast", xaxis_title="Meses",
                                                yaxis_title="Monto ($)")
                        st.plotly_chart(fig_f_hoy, use_container_width=True)
                    with col_sep2:
                        fig_b_hoy = go.Figure(
                            go.Scatter(x=meses_hasta_hoy, y=valores_budget_hoy, mode="lines+markers", name="Budget",
                                       line=dict(color="#1f77b4", width=3)))
                        fig_b_hoy.update_layout(title="Curva de Ejecución Pura: Budget", xaxis_title="Meses",
                                                yaxis_title="Monto ($)")
                        st.plotly_chart(fig_b_hoy, use_container_width=True)

                st.subheader("📋 Matriz Detallada por Centro de Costo")
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
import streamlit as st
import pandas as pd
import plotly.graph_objects as go

st.set_page_config(page_title="Control de Gestión: Forecast vs Budget", layout="wide")

# =====================================================================
# SECCIÓN 1: CONFIGURACIÓN INICIAL Y CARGA DE DATOS
# =====================================================================
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

        # Estandarización de nomenclaturas de áreas operativas
        for df in [df_forecast, df_budget]:
            if 'Gerencia' in df.columns:
                df['Gerencia'] = df['Gerencia'].astype(str).str.strip().replace(
                    {'Operaciones': 'Trabajador', 'OPERACIONES': 'TRABAJADOR'})

        # =====================================================================
        # SECCIÓN 2: PROCESAMIENTO ANALÍTICO Y UNIÓN DE BASES (OUTER JOIN)
        # =====================================================================
        abrev_meses = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']

        meses_f_map = {m: [c for c in df_forecast.columns if c.startswith(m)][0] for m in abrev_meses if
                       [c for c in df_forecast.columns if c.startswith(m)]}
        meses_b_map = {m: [c for c in df_budget.columns if c.startswith(m)][0] for m in abrev_meses if
                       [c for c in df_budget.columns if c.startswith(m)]}
        meses_validos = [m for m in abrev_meses if m in meses_f_map and m in meses_b_map]

        if not meses_validos:
            st.error("No se detectaron columnas de meses válidas en las hojas seleccionadas.")
        else:
            columnas_meses_f = [meses_f_map[m] for m in meses_validos]
            columnas_meses_b = [meses_b_map[m] for m in meses_validos]

            # CORRECCIÓN DE DETECCIÓN EXACTA DE COLUMNA FY (EVITA DESPASES AL 2028)
            anio_detectado = "".join([char for char in columnas_meses_b[0] if char.isdigit()])
            if len(anio_detectado) > 2:
                anio_detectado = anio_detectado[-2:]

            col_fy_budget = f"FY{anio_detectado}"
            if col_fy_budget not in df_budget.columns:
                col_fy_budget = f"FY20{anio_detectado}"
                if col_fy_budget not in df_budget.columns:
                    cols_fy_fallback = [c for c in df_budget.columns if c.startswith('FY')]
                    col_fy_budget = cols_fy_fallback[0] if cols_fy_fallback else None

            # Extracción de todas las columnas plurianuales de presupuesto disponibles
            cols_fy_todas = sorted([c for c in df_budget.columns if c.startswith('FY')])

            if 'CC' in df_forecast.columns and 'CC' in df_budget.columns:
                llaves_cruce = ['CC']
                if 'Classif' in df_forecast.columns and 'Classif' in df_budget.columns:
                    llaves_cruce.append('Classif')
                if 'Gerencia' in df_forecast.columns and 'Gerencia' in df_budget.columns:
                    llaves_cruce.append('Gerencia')
                if 'Desc Item' in df_forecast.columns and 'Desc Item' in df_budget.columns:
                    llaves_cruce.append('Desc Item')

                cols_b_extra = list(set(columnas_meses_b.copy() + cols_fy_todas))

                df_f_melt = df_forecast.groupby(llaves_cruce)[columnas_meses_f].sum().reset_index()
                df_b_melt = df_budget.groupby(llaves_cruce)[cols_b_extra].sum().reset_index()

                meses_b_rename = {meses_b_map[m]: f"{m}_Bud" for m in meses_validos}
                df_b_melt = df_b_melt.rename(columns=meses_b_rename)

                df_merged = pd.merge(df_f_melt, df_b_melt, on=llaves_cruce, how='outer')

                columnas_numericas = columnas_meses_f + [f"{m}_Bud" for m in meses_validos] + cols_fy_todas
                df_merged[columnas_numericas] = df_merged[columnas_numericas].fillna(0)

                st.sidebar.markdown("### ⚙️ Filtros Operacionales")
                filtro_classif = st.sidebar.selectbox("Clasificación de Costo (Classif)",
                                                      ["Todas"] + list(df_merged['Classif'].dropna().unique()))
                if filtro_classif != "Todas":
                    df_merged = df_merged[df_merged['Classif'] == filtro_classif]

                filtro_gerencia = st.sidebar.selectbox("Gerencia",
                                                       ["Todas"] + list(df_merged['Gerencia'].dropna().unique()))
                if filtro_gerencia != "Todas":
                    df_merged = df_merged[df_merged['Gerencia'] == filtro_gerencia]

                # =====================================================================
                # SECCIÓN 3: CÁLCULOS GLOBALES (FULL YEAR)
                # =====================================================================
                valores_forecast_full = [df_merged[meses_f_map[m]].sum() for m in meses_validos]
                valores_budget_full_list = [df_merged[f"{m}_Bud"].sum() for m in meses_validos]

                if col_fy_budget and col_fy_budget in df_merged.columns:
                    sum_budget_m = df_merged[col_fy_budget].sum() / 1_000_000
                else:
                    sum_budget_m = sum(valores_budget_full_list) / 1_000_000

                sum_forecast_m = sum(valores_forecast_full) / 1_000_000
                varianza_total_m = sum_forecast_m - sum_budget_m
                varianzas_mensuales_full = [f - b for f, b in zip(valores_forecast_full, valores_budget_full_list)]

                tab_anual, tab_temporal, tab_multi_presupuesto = st.tabs([
                    "📅 Análisis Año Completo (Full Year)",
                    "⏳ Análisis Temporal Interactivo",
                    "📈 Tendencias Plurianuales (Budgets)"
                ])

                # =====================================================================
                # SECCIÓN 4: INTERFAZ - PESTAÑA ANUAL (FULL YEAR)
                # =====================================================================
                with tab_anual:
                    st.markdown(f"### 🎯 Desempeño Financiero Anual Consolidado")
                    col1, col2, col3 = st.columns(3)
                    col1.metric("Forecast Anual (FY)", f"$ {sum_forecast_m:,.2f} M")
                    col2.metric(f"Budget Anual Oficial Corregido ({col_fy_budget if col_fy_budget else ''})",
                                f"$ {sum_budget_m:,.2f} M")
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
                        yaxis2=dict(title="Varianza Neta ($)", overlaying="y", side="right", showgrid=False),
                        height=500
                    )
                    st.plotly_chart(fig_comb, use_container_width=True)

                # =====================================================================
                # SECCIÓN 5: INTERFAZ - PESTAÑA TEMPORAL (YTD DINÁMICO)
                # =====================================================================
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

                    st.subheader(f"📊 Impacto Acumulado de Variaciones (Ene a {mes_corte})")
                    fig_water_hoy = go.Figure(go.Waterfall(
                        name="Variación Acumulada", orientation="v", x=meses_hasta_hoy, textposition="outside",
                        text=[f"${v / 1_000_000:.2f}M" for v in varianzas_mensuales_hoy], y=varianzas_mensuales_hoy,
                        connector=dict(line=dict(color="rgb(63, 63, 63)", width=1.5)),
                        decreasing=dict(marker=dict(color="#2ca02c")), increasing=dict(marker=dict(color="#d62728"))
                    ))
                    fig_water_hoy.update_layout(title=f"Gráfico de Cascada Acumulado ({mes_corte})", showlegend=False,
                                                height=550)
                    st.plotly_chart(fig_water_hoy, use_container_width=True)

                    st.divider()

                    st.markdown("##### Comportamiento de Tendencias Individuales")
                    col_sep1, col_sep2 = st.columns(2)
                    with col_sep1:
                        fig_f_hoy = go.Figure(
                            go.Scatter(x=meses_hasta_hoy, y=valores_forecast_hoy, mode="lines+markers", name="Forecast",
                                       line=dict(color="#ff7f0e", width=3)))
                        fig_f_hoy.update_layout(title="Curva de Ejecución Pura: Forecast", xaxis_title="Meses",
                                                yaxis_title="Monto ($)", height=450)
                        st.plotly_chart(fig_f_hoy, use_container_width=True)
                    with col_sep2:
                        fig_b_hoy = go.Figure(
                            go.Scatter(x=meses_hasta_hoy, y=valores_budget_hoy, mode="lines+markers", name="Budget",
                                       line=dict(color="#1f77b4", width=3)))
                        fig_b_hoy.update_layout(title="Curva de Ejecución Pura: Budget", xaxis_title="Meses",
                                                yaxis_title="Monto ($)", height=450)
                        st.plotly_chart(fig_b_hoy, use_container_width=True)

                        # =====================================================================
                        # SECCIÓN 6: INTERFAZ - PESTAÑA MULTI-PRESUPUESTO (NUEVA ANÁLISIS)
                        # =====================================================================
                        with tab_multi_presupuesto:
                            st.markdown("### 📈 Análisis de Evolución Presupuestaria Horizonte Plurianual")
                            st.markdown(
                                "Revisión de las proyecciones de inversión y puente de variaciones interanuales por categoría.")

                            # 1. Gráfico de tendencias histórico (Línea)
                            valores_multi_fy = [df_merged[c].sum() / 1_000_000 for c in cols_fy_todas]

                            fig_trend_fy = go.Figure()
                            fig_trend_fy.add_trace(go.Scatter(
                                x=cols_fy_todas, y=valores_multi_fy,
                                mode="lines+markers+text",
                                text=[f"${v:,.2f}M" for v in valores_multi_fy],
                                textposition="top center",
                                line=dict(color="#9467bd", width=4),
                                marker=dict(size=10, symbol="diamond")
                            ))
                            fig_trend_fy.update_layout(
                                title="Tendencia de Presupuestos Consolidados por Año de Ejercicio",
                                xaxis_title="Periodos Presupuestarios",
                                yaxis_title="Monto Total Asignado (M$)",
                                height=400
                            )
                            st.plotly_chart(fig_trend_fy, use_container_width=True)

                            st.divider()

                            # 2. NUEVO: Gráfico de Cascada Interanual (Puente Presupuestario)
                            st.subheader("🌊 Variación Detallada entre Presupuestos (Puente Interanual)")

                            if len(cols_fy_todas) >= 2:
                                # Selectores de años a comparar
                                col_sel1, col_sel2 = st.columns(2)
                                fy_base = col_sel1.selectbox("Seleccione Presupuesto Base (Año Inicial):",
                                                             cols_fy_todas, index=0)
                                fy_target = col_sel2.selectbox("Seleccione Presupuesto Destino (Año Final):",
                                                               cols_fy_todas, index=1)

                                # Extracción de totales
                                total_base = df_merged[fy_base].sum() / 1_000_000
                                total_target = df_merged[fy_target].sum() / 1_000_000

                                # Agrupación de la varianza por Clasificación de Costo (Classif)
                                df_puente = df_merged.groupby('Classif')[[fy_base, fy_target]].sum().reset_index()
                                df_puente['Delta'] = (df_puente[fy_target] - df_puente[fy_base]) / 1_000_000

                                # Ordenar por impacto para una visualización de cascada más limpia
                                df_puente = df_puente.sort_values(by='Delta', ascending=False)

                                categorias_puente = df_puente['Classif'].tolist()
                                deltas_puente = df_puente['Delta'].tolist()

                                # Vectores de construcción estructural para Plotly Waterfall
                                x_water = [fy_base] + categorias_puente + [fy_target]
                                y_water = [total_base] + deltas_puente + [
                                    0]  # 0 es el valor técnico de Plotly para cerrar la columna "total"
                                measure_water = ["absolute"] + ["relative"] * len(categorias_puente) + ["total"]

                                # Formateo de etiquetas de texto con signos explícitos
                                text_water = [f"${total_base:,.2f}M"] + [f"${d:,.2f}M" if d < 0 else f"+${d:,.2f}M" for
                                                                         d in deltas_puente] + [
                                                 f"${total_target:,.2f}M"]

                                fig_bridge = go.Figure(go.Waterfall(
                                    name="Puente Presupuestario",
                                    orientation="v",
                                    measure=measure_water,
                                    x=x_water,
                                    textposition="outside",
                                    text=text_water,
                                    y=y_water,
                                    connector=dict(line=dict(color="rgb(63, 63, 63)", width=1.5, dash="dot")),
                                    decreasing=dict(marker=dict(color="#2ca02c")),
                                    # Verde: Reducción de presupuesto interanual
                                    increasing=dict(marker=dict(color="#d62728")),
                                    # Rojo: Aumento de presupuesto interanual
                                    totals=dict(marker=dict(color="#1f77b4"))  # Azul: Barras base y destino
                                ))

                                fig_bridge.update_layout(
                                    title=f"Explicación de Variaciones por Categoría: {fy_base} → {fy_target}",
                                    showlegend=False,
                                    height=550,
                                    margin=dict(t=50, b=50),
                                    yaxis=dict(title="Monto (M$)")
                                )
                                st.plotly_chart(fig_bridge, use_container_width=True)
                            else:
                                st.info(
                                    "Se requiere información de al menos dos años presupuestarios (ej. FY24 y FY25) para generar el puente interanual.")

                # =====================================================================
                # SECCIÓN 7: MATRIZ DE DATOS (TABLA DETALLADA)
                # =====================================================================
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
                st.error("Error estructural: No se encontró la columna clave 'CC' en ambas hojas.")

    except Exception as e:
        st.error(f"Error crítico de procesamiento: {str(e)}")
else:
    st.info(
        "👈 Esperando carga de la planilla de control de gestión minera (.xlsx) para inicializar el pipeline analítico.")
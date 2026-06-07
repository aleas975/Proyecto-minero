import streamlit as st
import pandas as pd
import numpy as np
from sklearn.preprocessing import PolynomialFeatures
from sklearn.linear_model import LinearRegression
import plotly.express as px
import plotly.graph_objects as go

# Configuración de página
st.set_page_config(page_title="Visualizador Forecast 5+7 - Minería", layout="wide")

st.title("📊 Plataforma de Visualización: Forecast 5+7 vs Budget")
st.markdown("Análisis predictivo no lineal para control de gestión en faena minera.")

# Carga de datos (Simulando la lectura de tus hojas 'Forecast 5+7' y 'Tabla de análisis')
uploaded_file = st.file_uploader("Cargar consolidado de datos (CSV/Excel)", type=["csv", "xlsx"])

if uploaded_file is not None:
    # Lógica asumiendo que subes el CSV de la tabla de análisis
    df = pd.read_csv(uploaded_file)
    
    st.subheader("Análisis de Varianza por Categoría")
    # Filtro por categoría (Labor, Fuel, Maintenance, etc.)
    if 'Categoría' in df.columns:
        categoria = st.selectbox("Seleccione Categoría de Gasto", df['Categoría'].unique())
        df_cat = df[df['Categoría'] == categoria]
    else:
        df_cat = df # Fallback si no está la columna
        
    # Visualización de KPIs
    col1, col2, col3 = st.columns(3)
    col1.metric("Gasto Real (YTD Ene-May)", "$ 120.5M") # Valores de ejemplo a reemplazar con variables del df
    col2.metric("Budget Aprobado (FY)", "$ 285.0M")
    col3.metric("Varianza Proyectada", "$ -15.2M", delta_color="inverse")

    st.markdown("---")
    
    # Simulación de Proyección No Lineal (Polinómica grado 2)
    st.subheader("📈 Curva de Proyección No Lineal (Junio - Diciembre)")
    
    # Datos simulados basados en meses (1 al 12)
    meses = np.array(range(1, 13)).reshape(-1, 1)
    # Supongamos que estos son los gastos reales de Ene-May
    gastos_reales = np.array([20, 22, 25, 24, 28]) 
    
    # Modelo polinómico
    poly = PolynomialFeatures(degree=2)
    meses_poly = poly.fit_transform(meses[:5])
    modelo = LinearRegression().fit(meses_poly, gastos_reales)
    
    # Proyectando el año completo
    meses_completos_poly = poly.transform(meses)
    proyeccion = modelo.predict(meses_completos_poly)
    
    # Gráfico Plotly
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=['Ene', 'Feb', 'Mar', 'Abr', 'May'], y=gastos_reales, mode='lines+markers', name='YTD Real', line=dict(color='blue', width=3)))
    fig.add_trace(go.Scatter(x=['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'], y=proyeccion, mode='lines', name='Forecast No Lineal', line=dict(color='orange', dash='dash')))
    # Línea plana de Budget para comparación
    fig.add_trace(go.Scatter(x=['Ene', 'Dic'], y=[24, 24], mode='lines', name='Budget Promedio', line=dict(color='red', width=2)))
    
    st.plotly_chart(fig, use_container_width=True)
    
    st.info("El modelo polinómico captura la tendencia al alza del consumo, redistribuyendo la desviación presupuestaria en los meses de menor carga operativa (Q4).")
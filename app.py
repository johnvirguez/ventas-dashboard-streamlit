import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

st.set_page_config(page_title="Ventas — Dashboard & Proyección H1 2026", layout="wide")
st.title("Dashboard de Ventas (Excel) + Proyección H1 2026")

st.markdown(
    """
**Qué hace esta app**
- Cargas un Excel de ventas.
- Normaliza columnas y calcula reportes por **Año**, **Trimestre (Q)**, **Comercial**, **SKU**.
- Genera una **proyección H1 2026** con base en tendencia mensual histórica (método simple y transparente).
"""
)

# =========================
# 1) Carga del archivo
# =========================
st.sidebar.header("1) Cargar archivo")
uploaded = st.sidebar.file_uploader("Sube el Excel de ventas (.xlsx)", type=["xlsx"])

st.sidebar.header("2) Configuración")
st.sidebar.caption("Ajusta nombres de columnas si tu Excel usa otros títulos.")

# Nombres esperados (puedes mapear)
col_fecha = st.sidebar.text_input("Columna de Fecha", value="Fecha")
col_comercial = st.sidebar.text_input("Columna Comercial", value="Comercial")
col_sku = st.sidebar.text_input("Columna SKU", value="SKU")
col_ventas = st.sidebar.text_input("Columna Ventas (valor)", value="Ventas")

st.sidebar.header("3) Proyección")
metodo = st.sidebar.selectbox(
    "Método de proyección",
    [
        "Tendencia lineal sobre ventas mensuales (global)",
        "Promedio móvil 6 meses (global)"
    ]
)

# Helper: convertir a periodo trimestral
def quarter_label(dt: pd.Series) -> pd.Series:
    return "Q" + dt.dt.quarter.astype(str)

def read_excel(file) -> pd.DataFrame:
    # Lee el primer sheet por defecto
    return pd.read_excel(file, engine="openpyxl")

def validate_and_prepare(df: pd.DataFrame) -> pd.DataFrame:
    missing = [c for c in [col_fecha, col_comercial, col_sku, col_ventas] if c not in df.columns]
    if missing:
        raise ValueError(f"Faltan columnas en el Excel: {missing}. Ajusta los nombres en la barra lateral.")
    out = df[[col_fecha, col_comercial, col_sku, col_ventas]].copy()

    # Fecha
    out[col_fecha] = pd.to_datetime(out[col_fecha], errors="coerce")
    out = out.dropna(subset=[col_fecha])

    # Ventas numéricas
    out[col_ventas] = pd.to_numeric(out[col_ventas], errors="coerce").fillna(0)

    # Limpieza texto
    out[col_comercial] = out[col_comercial].astype(str).str.strip()
    out[col_sku] = out[col_sku].astype(str).str.strip()

    # Atributos de tiempo
    out["Año"] = out[col_fecha].dt.year
    out["Mes"] = out[col_fecha].dt.to_period("M").astype(str)
    out["Trimestre"] = quarter_label(out[col_fecha])
    out["Año-Q"] = out["Año"].astype(str) + "-" + out["Trimestre"]

    return out

def kpi_cards(df: pd.DataFrame):
    total = df[col_ventas].sum()
    n_rows = len(df)
    n_com = df[col_comercial].nunique()
    n_sku = df[col_sku].nunique()
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Ventas totales", f"{total:,.2f}")
    c2.metric("Registros", f"{n_rows:,}")
    c3.metric("Comerciales", f"{n_com:,}")
    c4.metric("SKUs", f"{n_sku:,}")

def plot_bar(series: pd.Series, title: str, xlabel: str, ylabel: str):
    fig, ax = plt.subplots(figsize=(9, 3.5))
    ax.bar(series.index.astype(str), series.values)
    ax.set_title(title)
    ax.set_xlabel(xlabel)
    ax.set_ylabel(ylabel)
    ax.grid(True, linestyle="--", alpha=0.4)
    plt.xticks(rotation=45, ha="right")
    st.pyplot(fig)

def plot_line(x, y, title, xlabel, ylabel):
    fig, ax = plt.subplots(figsize=(9, 3.5))
    ax.plot(x, y, marker="o", linewidth=1)
    ax.set_title(title)
    ax.set_xlabel(xlabel)
    ax.set_ylabel(ylabel)
    ax.grid(True, linestyle="--", alpha=0.4)
    plt.xticks(rotation=45, ha="right")
    st.pyplot(fig)

def forecast_h1_2026(monthly: pd.Series, method: str) -> pd.DataFrame:
    """
    monthly: Series index = 'YYYY-MM', values = ventas
    Retorna DataFrame con proyección 2026-01 .. 2026-06
    """
    # Asegurar orden temporal
    monthly = monthly.copy()
    monthly.index = pd.PeriodIndex(monthly.index, freq="M")
    monthly = monthly.sort_index()

    # Crear eje numérico para regresión
    y = monthly.values.astype(float)
    x = np.arange(len(y))

    future_periods = pd.period_range("2026-01", "2026-06", freq="M")
    n_future = len(future_periods)

    if method.startswith("Tendencia lineal"):
        # y = a*x + b
        a, b = np.polyfit(x, y, 1)
        x_future = np.arange(len(y), len(y) + n_future)
        y_future = a * x_future + b
        y_future = np.maximum(y_future, 0)  # no negativos
    else:
        # Promedio móvil 6 meses
        window = 6 if len(y) >= 6 else max(1, len(y))
        avg = pd.Series(y).rolling(window).mean().iloc[-1]
        y_future = np.array([max(avg, 0)] * n_future)

    return pd.DataFrame(
        {
            "Mes": [p.strftime("%Y-%m") for p in future_periods],
            "Proyección": y_future.round(2),
        }
    )

# =========================
# 2) Procesamiento y reportes
# =========================
if not uploaded:
    st.info("Carga un Excel para iniciar. Asegúrate de tener columnas de Fecha, Comercial, SKU y Ventas.")
    st.stop()

try:
    raw = read_excel(uploaded)
    df = validate_and_prepare(raw)
except Exception as e:
    st.error(f"Error leyendo/validando el archivo: {e}")
    st.stop()

kpi_cards(df)

st.write("---")
st.subheader("Vista previa de datos normalizados")
st.dataframe(df.head(50), height=260)

# Filtros
st.sidebar.header("4) Filtros")
years = sorted(df["Año"].unique().tolist())
selected_years = st.sidebar.multiselect("Años", years, default=years)
selected_comerciales = st.sidebar.multiselect("Comerciales", sorted(df[col_comercial].unique()), default=[])
selected_skus = st.sidebar.multiselect("SKUs", sorted(df[col_sku].unique()), default=[])

df_f = df[df["Año"].isin(selected_years)]
if selected_comerciales:
    df_f = df_f[df_f[col_comercial].isin(selected_comerciales)]
if selected_skus:
    df_f = df_f[df_f[col_sku].isin(selected_skus)]

st.write("---")
st.subheader("Reportes")

tab1, tab2, tab3, tab4, tab5 = st.tabs(["Por Año", "Por Q", "Por Comercial", "Por SKU", "Proyección H1 2026"])

with tab1:
    g = df_f.groupby("Año")[col_ventas].sum().sort_index()
    st.dataframe(g.reset_index().rename(columns={col_ventas: "Ventas"}))
    plot_bar(g, "Ventas por Año", "Año", "Ventas")

with tab2:
    g = df_f.groupby("Año-Q")[col_ventas].sum()
    # ordenar por año y trimestre
    order = sorted(g.index, key=lambda s: (int(s.split("-")[0]), int(s.split("Q")[1])))
    g = g.loc[order]
    st.dataframe(g.reset_index().rename(columns={"Año-Q": "Periodo", col_ventas: "Ventas"}))
    plot_bar(g, "Ventas por Trimestre (Año-Q)", "Periodo", "Ventas")

with tab3:
    g = df_f.groupby(col_comercial)[col_ventas].sum().sort_values(ascending=False).head(30)
    st.dataframe(g.reset_index().rename(columns={col_comercial: "Comercial", col_ventas: "Ventas"}))
    plot_bar(g, "Top 30 comerciales por ventas", "Comercial", "Ventas")

with tab4:
    g = df_f.groupby(col_sku)[col_ventas].sum().sort_values(ascending=False).head(30)
    st.dataframe(g.reset_index().rename(columns={col_sku: "SKU", col_ventas: "Ventas"}))
    plot_bar(g, "Top 30 SKU por ventas", "SKU", "Ventas")

with tab5:
    st.markdown("### Base histórica mensual (global)")
    monthly = df_f.groupby("Mes")[col_ventas].sum()
    monthly = monthly.sort_index()
    st.dataframe(monthly.reset_index().rename(columns={"Mes": "Mes", col_ventas: "Ventas"}))
    plot_line(monthly.index.astype(str), monthly.values, "Ventas mensuales (histórico)", "Mes", "Ventas")

    st.markdown("### Proyección H1 2026")
    proj = forecast_h1_2026(monthly, metodo)
    st.dataframe(proj)

    fig, ax = plt.subplots(figsize=(9, 3.5))
    ax.plot(proj["Mes"], proj["Proyección"], marker="o", linewidth=1)
    ax.set_title(f"Proyección H1 2026 — {metodo}")
    ax.set_xlabel("Mes")
    ax.set_ylabel("Ventas proyectadas")
    ax.grid(True, linestyle="--", alpha=0.4)
    plt.xticks(rotation=45, ha="right")
    st.pyplot(fig)

    st.metric("Total proyectado H1 2026", f"{proj['Proyección'].sum():,.2f}")

st.write("---")
st.subheader("Descargas")
csv = df_f.to_csv(index=False).encode("utf-8")
st.download_button("Descargar datos filtrados (CSV)", data=csv, file_name="ventas_filtradas.csv", mime="text/csv")

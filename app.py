import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

# =========================
# Configuración de la página
# =========================
st.set_page_config(
    page_title="Ventas — Dashboard & Proyección H1 2026",
    layout="wide"
)

st.title("Dashboard de Ventas (Excel) + Proyección H1 2026")

st.markdown(
    """
**Qué hace esta app**
- Cargas un Excel de ventas (puede tener varias hojas).
- Seleccionas la hoja correcta y mapeas columnas (Fecha, Comercial, SKU, Ventas).
- Genera reportes por **Año**, **Trimestre (Q)**, **Comercial**, **SKU**.
- Genera una **proyección H1 2026** basada en el histórico mensual (métodos simples y explicables).
"""
)

# =========================
# Sidebar: carga y configuración
# =========================
st.sidebar.header("1) Cargar archivo")
uploaded = st.sidebar.file_uploader("Sube el Excel de ventas (.xlsx)", type=["xlsx"])

st.sidebar.header("2) Selección de hoja y columnas")
st.sidebar.caption("Si tu Excel tiene varias hojas, selecciona la que contiene la tabla de ventas.")

# =========================
# Helpers
# =========================
def quarter_label(dt: pd.Series) -> pd.Series:
    return "Q" + dt.dt.quarter.astype(str)

def safe_to_datetime(s: pd.Series) -> pd.Series:
    return pd.to_datetime(s, errors="coerce")

def safe_to_numeric(s: pd.Series) -> pd.Series:
    # Permite números con separadores; si llega texto tipo "1,234.56" lo convierte
    if s.dtype == object:
        s = s.astype(str).str.replace(",", "", regex=False)
    return pd.to_numeric(s, errors="coerce")

def find_best_sheet(xls: pd.ExcelFile) -> str:
    """
    Heurística: elige la primera hoja que contenga al menos alguna combinación razonable
    de columnas típicas (Fecha, Comercial, SKU, Ventas o sinónimos).
    Si no encuentra, devuelve la primera hoja.
    """
    candidates = [
        {"Fecha", "Comercial", "SKU"},
        {"fecha", "comercial", "sku"},
        {"Fecha", "Vendedor", "SKU"},
        {"Fecha", "Sales Rep", "SKU"},
    ]
    for sname in xls.sheet_names:
        try:
            temp = pd.read_excel(xls, sheet_name=sname, nrows=5)
            cols = set(map(str, temp.columns))
            cols_lower = set(c.lower() for c in cols)
            # Si tiene alguno de los sets clave, la marcamos
            if any(cset.issubset(cols) for cset in candidates) or any(
                {c.lower() for c in cset}.issubset(cols_lower) for cset in candidates
            ):
                return sname
        except Exception:
            continue
    return xls.sheet_names[0]

def plot_bar(series: pd.Series, title: str, xlabel: str, ylabel: str):
    fig, ax = plt.subplots(figsize=(10, 3.8))
    ax.bar(series.index.astype(str), series.values)
    ax.set_title(title)
    ax.set_xlabel(xlabel)
    ax.set_ylabel(ylabel)
    ax.grid(True, linestyle="--", alpha=0.4)
    plt.xticks(rotation=45, ha="right")
    st.pyplot(fig)

def plot_line(x, y, title: str, xlabel: str, ylabel: str):
    fig, ax = plt.subplots(figsize=(10, 3.8))
    ax.plot(x, y, marker="o", linewidth=1)
    ax.set_title(title)
    ax.set_xlabel(xlabel)
    ax.set_ylabel(ylabel)
    ax.grid(True, linestyle="--", alpha=0.4)
    plt.xticks(rotation=45, ha="right")
    st.pyplot(fig)

def kpi_cards(df: pd.DataFrame, col_ventas: str, col_comercial: str, col_sku: str):
    total = df[col_ventas].sum()
    n_rows = len(df)
    n_com = df[col_comercial].nunique()
    n_sku = df[col_sku].nunique()
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Ventas totales", f"{total:,.2f}")
    c2.metric("Registros", f"{n_rows:,}")
    c3.metric("Comerciales", f"{n_com:,}")
    c4.metric("SKUs", f"{n_sku:,}")

def forecast_h1_2026(monthly: pd.Series, method: str) -> pd.DataFrame:
    """
    monthly: Series index = PeriodIndex mensual ('YYYY-MM'), values = ventas
    Retorna DataFrame con proyección 2026-01 .. 2026-06
    """
    monthly = monthly.copy()
    if not isinstance(monthly.index, pd.PeriodIndex):
        monthly.index = pd.PeriodIndex(monthly.index, freq="M")
    monthly = monthly.sort_index()

    y = monthly.values.astype(float)
    x = np.arange(len(y))

    future_periods = pd.period_range("2026-01", "2026-06", freq="M")
    n_future = len(future_periods)

    if len(y) == 0:
        y_future = np.zeros(n_future)
    elif method.startswith("Tendencia lineal"):
        # Si hay muy pocos puntos, usamos promedio como fallback
        if len(y) < 3:
            avg = float(np.mean(y))
            y_future = np.array([max(avg, 0)] * n_future)
        else:
            a, b = np.polyfit(x, y, 1)
            x_future = np.arange(len(y), len(y) + n_future)
            y_future = a * x_future + b
            y_future = np.maximum(y_future, 0)
    else:
        window = 6 if len(y) >= 6 else max(1, len(y))
        avg = float(pd.Series(y).rolling(window).mean().iloc[-1])
        y_future = np.array([max(avg, 0)] * n_future)

    return pd.DataFrame(
        {"Mes": [p.strftime("%Y-%m") for p in future_periods], "Proyección": np.round(y_future, 2)}
    )

def validate_and_prepare(df: pd.DataFrame, col_fecha: str, col_comercial: str, col_sku: str, col_ventas: str) -> pd.DataFrame:
    missing = [c for c in [col_fecha, col_comercial, col_sku, col_ventas] if c not in df.columns]
    if missing:
        raise ValueError(
            f"Faltan columnas en la hoja seleccionada: {missing}. "
            f"Ajusta los nombres en la barra lateral o selecciona otra hoja."
        )

    out = df[[col_fecha, col_comercial, col_sku, col_ventas]].copy()

    out[col_fecha] = safe_to_datetime(out[col_fecha])
    out = out.dropna(subset=[col_fecha])

    out[col_ventas] = safe_to_numeric(out[col_ventas]).fillna(0)

    out[col_comercial] = out[col_comercial].astype(str).str.strip()
    out[col_sku] = out[col_sku].astype(str).str.strip()

    out["Año"] = out[col_fecha].dt.year
    out["Mes"] = out[col_fecha].dt.to_period("M")  # PeriodIndex-like values
    out["Trimestre"] = quarter_label(out[col_fecha])
    out["Año-Q"] = out["Año"].astype(str) + "-" + out["Trimestre"]

    return out

# =========================
# Main: ejecutar solo si hay archivo
# =========================
if not uploaded:
    st.info("Carga un Excel para iniciar. Debe incluir columnas de Fecha, Comercial, SKU y Ventas.")
    st.stop()

# Leer el Excel y seleccionar hoja
try:
    xls = pd.ExcelFile(uploaded, engine="openpyxl")
except Exception as e:
    st.error(f"No pude abrir el archivo como Excel (.xlsx). Detalle: {e}")
    st.stop()

best_sheet = find_best_sheet(xls)
sheet = st.sidebar.selectbox("Hoja del Excel", xls.sheet_names, index=xls.sheet_names.index(best_sheet))

# Leer un preview para sugerir mapeo de columnas
try:
    preview_df = pd.read_excel(xls, sheet_name=sheet, nrows=50)
except Exception as e:
    st.error(f"No pude leer la hoja '{sheet}'. Detalle: {e}")
    st.stop()

cols = list(map(str, preview_df.columns))

def suggest(default: str, options: list[str]) -> str:
    # Si el default existe, úsalo; si no, intenta encontrar parecido por contains
    if default in options:
        return default
    low = [c.lower() for c in options]
    d = default.lower()
    # heurística: buscar columnas que contengan palabras clave
    keywords = {
        "fecha": ["fecha", "date"],
        "comercial": ["comercial", "vendedor", "sales rep", "seller", "ejecutivo"],
        "sku": ["sku", "producto", "item", "code"],
        "ventas": ["ventas", "venta", "total", "amount", "revenue", "venta total (usd)"],
    }
    if d in keywords:
        for kw in keywords[d]:
            for i, c in enumerate(low):
                if kw in c:
                    return options[i]
    return options[0] if options else default

# Mapeo de columnas (selectbox para evitar errores de escritura)
col_fecha = st.sidebar.selectbox("Columna de Fecha", cols, index=cols.index(suggest("Fecha", cols)) if cols else 0)
col_comercial = st.sidebar.selectbox("Columna Comercial", cols, index=cols.index(suggest("Comercial", cols)) if cols else 0)
col_sku = st.sidebar.selectbox("Columna SKU", cols, index=cols.index(suggest("SKU", cols)) if cols else 0)
col_ventas = st.sidebar.selectbox("Columna Ventas (valor)", cols, index=cols.index(suggest("Ventas", cols)) if cols else 0)

st.sidebar.header("3) Proyección")
metodo = st.sidebar.selectbox(
    "Método de proyección",
    [
        "Tendencia lineal sobre ventas mensuales (global)",
        "Promedio móvil 6 meses (global)"
    ]
)

# Procesar datos completos (no solo preview)
try:
    raw = pd.read_excel(xls, sheet_name=sheet)
    df = validate_and_prepare(raw, col_fecha, col_comercial, col_sku, col_ventas)
except Exception as e:
    st.error(f"Error validando datos: {e}")
    st.stop()

# =========================
# KPI + preview
# =========================
kpi_cards(df, col_ventas, col_comercial, col_sku)

st.write("---")
st.subheader("Vista previa de datos normalizados")
st.dataframe(df.head(50), height=260)

# =========================
# Filtros
# =========================
st.sidebar.header("4) Filtros")

years = sorted(df["Año"].unique().tolist())
selected_years = st.sidebar.multiselect("Años", years, default=years)

selected_comerciales = st.sidebar.multiselect(
    "Comerciales",
    sorted(df[col_comercial].unique().tolist()),
    default=[]
)

selected_skus = st.sidebar.multiselect(
    "SKUs",
    sorted(df[col_sku].unique().tolist()),
    default=[]
)

df_f = df[df["Año"].isin(selected_years)]
if selected_comerciales:
    df_f = df_f[df_f[col_comercial].isin(selected_comerciales)]
if selected_skus:
    df_f = df_f[df_f[col_sku].isin(selected_skus)]

# =========================
# Reportes
# =========================
st.write("---")
st.subheader("Reportes")

tab1, tab2, tab3, tab4, tab5 = st.tabs(["Por Año", "Por Q", "Por Comercial", "Por SKU", "Proyección H1 2026"])

with tab1:
    g = df_f.groupby("Año")[col_ventas].sum().sort_index()
    st.dataframe(g.reset_index().rename(columns={col_ventas: "Ventas"}))
    plot_bar(g, "Ventas por Año", "Año", "Ventas")

with tab2:
    g = df_f.groupby("Año-Q")[col_ventas].sum()
    # ordenar por año y trimestre (AÑO-Qn)
    order = sorted(g.index, key=lambda s: (int(str(s).split("-")[0]), int(str(s).split("Q")[1])))
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

    # Mostrar tabla mensual
    df_monthly = monthly.reset_index()
    df_monthly["Mes"] = df_monthly["Mes"].astype(str)
    df_monthly = df_monthly.rename(columns={col_ventas: "Ventas"})
    st.dataframe(df_monthly, height=260)

    plot_line(df_monthly["Mes"], df_monthly["Ventas"].values, "Ventas mensuales (histórico)", "Mes", "Ventas")

    st.markdown("### Proyección H1 2026")
    proj = forecast_h1_2026(monthly, metodo)
    st.dataframe(proj, height=260)

    fig, ax = plt.subplots(figsize=(10, 3.8))
    ax.plot(proj["Mes"], proj["Proyección"], marker="o", linewidth=1)
    ax.set_title(f"Proyección H1 2026 — {metodo}")
    ax.set_xlabel("Mes")
    ax.set_ylabel("Ventas proyectadas")
    ax.grid(True, linestyle="--", alpha=0.4)
    plt.xticks(rotation=45, ha="right")
    st.pyplot(fig)

    st.metric("Total proyectado H1 2026", f"{proj['Proyección'].sum():,.2f}")

# =========================
# Descargas
# =========================
st.write("---")
st.subheader("Descargas")

csv = df_f.drop(columns=["Mes"]).copy()
# Convertir Mes a string legible antes de descargar (si aún existiera)
if "Mes" in csv.columns:
    csv["Mes"] = csv["Mes"].astype(str)

csv_bytes = df_f.assign(Mes=df_f["Mes"].astype(str)).to_csv(index=False).encode("utf-8")
st.download_button(
    "Descargar datos filtrados (CSV)",
    data=csv_bytes,
    file_name="ventas_filtradas.csv",
    mime="text/csv"
)

st.caption(
    "Notas: 1) Selecciona la hoja correcta del Excel. 2) Mapea las columnas según tu archivo "
    "(por ejemplo, 'Venta Total (USD)' para ventas)."
)

import io
from datetime import datetime

import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns

# Excel
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.chart import BarChart, Reference
from openpyxl.utils import get_column_letter

# PDF
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer,
                                Table, TableStyle, Image, PageBreak)

# ---------- CONFIGURACIÓN ----------
st.set_page_config(
    page_title="Dashboard Comercial",
    page_icon="📊",
    layout="wide"
)

# ---------- CSS PERSONALIZADO ----------
st.markdown("""
<style>
    .block-container {padding-top: 2rem;}

    /* Tarjetas de KPI */
    [data-testid="stMetric"] {
        background-color: #FFFFFF;
        border-left: 6px solid #1F77B4;
        padding: 15px 20px;
        border-radius: 12px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.10);
    }
    [data-testid="stMetricLabel"] {font-size: 16px; color: #6B7280;}
    [data-testid="stMetricValue"] {font-size: 32px; color: #1F77B4;}

    /* Encabezado */
    .header {
        background: linear-gradient(90deg, #1F77B4, #4FA3D9);
        padding: 25px;
        border-radius: 12px;
        color: white;
        margin-bottom: 20px;
    }
    .header h1 {margin: 0; color: white;}
    .header p {margin: 0; opacity: 0.9;}

    /* Tarjetas de resumen */
    .insight {
        background: #FFFFFF;
        border-left: 6px solid #4FA3D9;
        padding: 12px 18px;
        border-radius: 10px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.08);
        margin-bottom: 10px;
        color: #1F2937;
    }
</style>
""", unsafe_allow_html=True)


# =====================================================
# FUNCIONES DE EXPORTACIÓN
# =====================================================
# >>> INICIO EXPORTACION
def fig_a_png(fig):
    """Convierte una figura de matplotlib en bytes PNG."""
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150, bbox_inches="tight")
    buf.seek(0)
    return buf.getvalue()


def generar_excel(df, kpis, res_prod, res_vend, pred_df):
    """Genera un archivo Excel con varias hojas y devuelve sus bytes."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl", datetime_format="DD/MM/YYYY") as writer:
        resumen = pd.DataFrame({
            "Indicador": ["Ventas Totales (S/)", "Cantidad Vendida", "Ticket Promedio (S/)"],
            "Valor": [kpis["ventas_totales"], kpis["cantidad_total"], round(kpis["ticket_promedio"], 2)],
        })
        resumen.to_excel(writer, sheet_name="Resumen", index=False)
        df.to_excel(writer, sheet_name="Datos", index=False)
        res_prod.reset_index().to_excel(writer, sheet_name="Por Producto", index=False)
        res_vend.reset_index().to_excel(writer, sheet_name="Por Vendedor", index=False)
        if pred_df is not None:
            pred_df.to_excel(writer, sheet_name="Predicción", index=False)

        # Estilos
        azul = PatternFill("solid", fgColor="1F77B4")
        for ws in writer.book.worksheets:
            for cell in ws[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = azul
                cell.alignment = Alignment(horizontal="center")
            ws.freeze_panes = "A2"
            for col in ws.columns:
                ancho = max(len(str(c.value)) if c.value is not None else 0 for c in col)
                ws.column_dimensions[get_column_letter(col[0].column)].width = min(ancho + 4, 40)

        # Gráfico de barras nativo de Excel
        ws = writer.book["Por Producto"]
        if ws.max_row >= 2:
            chart = BarChart()
            chart.title = "Ventas por Producto"
            chart.y_axis.title = "Ventas (S/)"
            chart.add_data(Reference(ws, min_col=2, min_row=1, max_row=ws.max_row), titles_from_data=True)
            chart.set_categories(Reference(ws, min_col=1, min_row=2, max_row=ws.max_row))
            chart.height, chart.width = 8, 14
            ws.add_chart(chart, "G2")
    return buf.getvalue()


def _tabla_pdf(data, col_widths=None):
    """Tabla con estilo para el PDF."""
    t = Table(data, colWidths=col_widths, repeatRows=1)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F77B4")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#D1D5DB")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F3F6FA")]),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    return t


def _imagen_pdf(png_bytes, ancho_cm=16, alto_cm=9):
    return Image(io.BytesIO(png_bytes), width=ancho_cm * cm, height=alto_cm * cm, kind="proportional")


def generar_pdf(df, kpis, hallazgos, filtros_txt, png_prod, png_vend, png_pred,
                res_prod, res_vend, pred_df):
    """Genera un reporte PDF y devuelve sus bytes."""
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=2 * cm, rightMargin=2 * cm,
                            topMargin=2 * cm, bottomMargin=2 * cm,
                            title="Reporte Comercial de Ventas")
    estilos = getSampleStyleSheet()
    h1, h2, normal = estilos["Title"], estilos["Heading2"], estilos["BodyText"]
    h1.textColor = colors.HexColor("#1F77B4")
    h2.textColor = colors.HexColor("#1F77B4")

    s = []
    s.append(Paragraph("Dashboard Comercial de Ventas", h1))
    s.append(Paragraph(f"Reporte generado el {datetime.now():%d/%m/%Y %H:%M}", normal))
    s.append(Paragraph(f"<b>Filtros aplicados:</b> {filtros_txt}", normal))
    s.append(Spacer(1, 0.5 * cm))

    # KPIs
    s.append(Paragraph("Indicadores clave", h2))
    s.append(_tabla_pdf([
        ["Ventas Totales", "Cantidad Vendida", "Ticket Promedio"],
        [f"S/ {kpis['ventas_totales']:,.0f}", f"{kpis['cantidad_total']:,}", f"S/ {kpis['ticket_promedio']:,.2f}"],
    ]))
    s.append(Spacer(1, 0.4 * cm))

    # Hallazgos
    s.append(Paragraph("Hallazgos clave", h2))
    for h in hallazgos:
        s.append(Paragraph(f"&bull; {h}", normal))
    s.append(Spacer(1, 0.4 * cm))

    # Gráficos
    s.append(Paragraph("Ventas por producto", h2))
    s.append(_imagen_pdf(png_prod, 14, 8))
    s.append(PageBreak())
    s.append(Paragraph("Participación por vendedor", h2))
    s.append(_imagen_pdf(png_vend, 12, 9))

    # Resúmenes
    s.append(Paragraph("Resumen por producto", h2))
    s.append(_tabla_pdf([["Producto", "Ventas (S/)", "Cantidad", "Ticket prom. (S/)"]] +
                        [[i, f"{r.Ventas:,.0f}", f"{r.Cantidad:,}", f"{r.Ticket_Promedio:,.2f}"]
                         for i, r in res_prod.iterrows()]))
    s.append(Spacer(1, 0.4 * cm))
    s.append(Paragraph("Resumen por vendedor", h2))
    s.append(_tabla_pdf([["Vendedor", "Ventas (S/)", "Cantidad", "Ticket prom. (S/)"]] +
                        [[i, f"{r.Ventas:,.0f}", f"{r.Cantidad:,}", f"{r.Ticket_Promedio:,.2f}"]
                         for i, r in res_vend.iterrows()]))

    # Predicción
    if pred_df is not None and png_pred is not None:
        s.append(PageBreak())
        s.append(Paragraph("Predicción de ventas", h2))
        s.append(Paragraph("Estimación basada en una tendencia lineal de las ventas diarias. "
                           "Es una referencia, no una garantía.", normal))
        s.append(_imagen_pdf(png_pred, 16, 7.5))
        s.append(_tabla_pdf([["Fecha", "Ventas estimadas (S/)"]] +
                            [[f"{r.Fecha:%d/%m/%Y}", f"{r.Ventas:,.2f}"] for r in pred_df.itertuples()],
                            col_widths=[5 * cm, 6 * cm]))

    # Datos
    s.append(PageBreak())
    s.append(Paragraph("Detalle de ventas", h2))
    s.append(_tabla_pdf([["Fecha", "Producto", "Vendedor", "Ventas (S/)", "Cantidad"]] +
                        [[f"{r.Fecha:%d/%m/%Y}", r.Producto, r.Vendedor, f"{r.Ventas:,.0f}", r.Cantidad]
                         for r in df.itertuples()]))

    doc.build(s)
    return buf.getvalue()
# <<< FIN EXPORTACION


# ---------- ENCABEZADO ----------
st.markdown("""
<div class="header">
    <h1>📊 Dashboard Comercial de Ventas</h1>
    <p>Análisis empresarial con Python + Streamlit</p>
</div>
""", unsafe_allow_html=True)

# ---------- CARGAR DATOS ----------
df = pd.read_csv("ventas.csv")
df["Fecha"] = pd.to_datetime(df["Fecha"])

# ---------- SIDEBAR: FILTROS ----------
st.sidebar.header("🔎 Filtros")

vendedores = st.sidebar.multiselect(
    "Vendedor",
    options=df["Vendedor"].unique(),
    default=df["Vendedor"].unique()
)
productos = st.sidebar.multiselect(
    "Producto",
    options=df["Producto"].unique(),
    default=df["Producto"].unique()
)

df = df[df["Vendedor"].isin(vendedores) & df["Producto"].isin(productos)]

if df.empty:
    st.warning("No hay datos con los filtros seleccionados.")
    st.stop()

# ---------- CÁLCULOS GENERALES ----------
ventas_totales = df["Ventas"].sum()
cantidad_total = df["Cantidad"].sum()
ticket_promedio = df["Ventas"].mean()

ventas_producto = df.groupby("Producto")["Ventas"].sum().sort_values(ascending=False)
ventas_vendedor = df.groupby("Vendedor")["Ventas"].sum().sort_values(ascending=False)
ventas_dia = df.groupby("Fecha")["Ventas"].sum()

res_prod = df.groupby("Producto").agg(
    Ventas=("Ventas", "sum"),
    Cantidad=("Cantidad", "sum"),
    Ticket_Promedio=("Ventas", "mean")
).round(2).sort_values("Ventas", ascending=False)

res_vend = df.groupby("Vendedor").agg(
    Ventas=("Ventas", "sum"),
    Cantidad=("Cantidad", "sum"),
    Ticket_Promedio=("Ventas", "mean")
).round(2).sort_values("Ventas", ascending=False)

mejor_producto = ventas_producto.index[0]
mejor_vendedor = ventas_vendedor.index[0]
mejor_dia = ventas_dia.idxmax()
pct_producto = ventas_producto.iloc[0] / ventas_totales * 100
pct_vendedor = ventas_vendedor.iloc[0] / ventas_totales * 100

# Variables que se llenan en las pestañas y se usan al exportar
png_prod = png_vend = png_pred = None
pred_df = None

sns.set_style("whitegrid")

# ---------- PESTAÑAS ----------
tab_resumen, tab_pred, tab_datos, tab_export = st.tabs(
    ["📌 Resumen", "🔮 Predicciones", "📋 Datos", "📤 Exportar"]
)

# =====================================================
# PESTAÑA 1: RESUMEN
# =====================================================
with tab_resumen:
    col1, col2, col3 = st.columns(3)
    col1.metric("💰 Ventas Totales", f"S/ {ventas_totales:,.0f}")
    col2.metric("📦 Cantidad Vendida", f"{cantidad_total:,}")
    col3.metric("🎫 Ticket Promedio", f"S/ {ticket_promedio:,.2f}")

    st.markdown("---")

    st.subheader("🧠 Hallazgos clave")
    h1, h2, h3 = st.columns(3)
    h1.markdown(
        f'<div class="insight">🏆 <b>Producto líder</b><br>{mejor_producto} '
        f'(S/ {ventas_producto.iloc[0]:,.0f} · {pct_producto:.1f}% del total)</div>',
        unsafe_allow_html=True)
    h2.markdown(
        f'<div class="insight">👤 <b>Mejor vendedor</b><br>{mejor_vendedor} '
        f'(S/ {ventas_vendedor.iloc[0]:,.0f} · {pct_vendedor:.1f}% del total)</div>',
        unsafe_allow_html=True)
    h3.markdown(
        f'<div class="insight">📅 <b>Mejor día</b><br>{mejor_dia:%d/%m/%Y} '
        f'(S/ {ventas_dia.max():,.0f})</div>',
        unsafe_allow_html=True)

    st.markdown("---")

    col_a, col_b = st.columns(2)

    with col_a:
        st.subheader("Ventas por Producto")
        fig1, ax1 = plt.subplots(figsize=(7, 5))
        sns.barplot(
            x=ventas_producto.index,
            y=ventas_producto.values,
            hue=ventas_producto.index,
            palette="Blues_r",
            legend=False,
            ax=ax1
        )
        ax1.set_xlabel("")
        ax1.set_ylabel("Ventas (S/)")
        ax1.set_title("Ingresos por Producto")
        sns.despine()
        st.pyplot(fig1)
        png_prod = fig_a_png(fig1)

    with col_b:
        st.subheader("Participación por Vendedor")
        fig2, ax2 = plt.subplots(figsize=(7, 5))
        ax2.pie(
            ventas_vendedor.values,
            labels=ventas_vendedor.index,
            autopct="%1.1f%%",
            startangle=90,
            colors=sns.color_palette("Blues", len(ventas_vendedor)),
            wedgeprops={"edgecolor": "white", "linewidth": 2}
        )
        ax2.set_title("Participación de Ventas")
        st.pyplot(fig2)
        png_vend = fig_a_png(fig2)

    st.subheader("📈 Evolución de Ventas")
    st.line_chart(ventas_dia)

# =====================================================
# PESTAÑA 2: PREDICCIONES
# =====================================================
with tab_pred:
    st.subheader("🔮 Predicción de ventas")
    st.caption("Modelo de tendencia lineal sobre las ventas diarias. "
               "Es una estimación simple: con pocos datos sirve como referencia, no como garantía.")

    dias_pred = st.slider("Días a predecir", min_value=3, max_value=30, value=7)

    rango = pd.date_range(ventas_dia.index.min(), ventas_dia.index.max(), freq="D")
    serie = ventas_dia.reindex(rango, fill_value=0)

    if len(serie) < 3:
        st.info("Se necesitan al menos 3 días de datos para generar una predicción. "
                "Ajusta los filtros del panel lateral.")
    else:
        x = np.arange(len(serie))
        y = serie.values
        pendiente, intercepto = np.polyfit(x, y, 1)

        y_aj = pendiente * x + intercepto
        ss_res = np.sum((y - y_aj) ** 2)
        ss_tot = np.sum((y - y.mean()) ** 2)
        r2 = 1 - ss_res / ss_tot if ss_tot > 0 else 0

        x_fut = np.arange(len(serie), len(serie) + dias_pred)
        fechas_fut = pd.date_range(serie.index[-1] + pd.Timedelta(days=1), periods=dias_pred, freq="D")
        y_fut = np.clip(pendiente * x_fut + intercepto, 0, None)
        pred = pd.Series(y_fut, index=fechas_fut)

        prom_actual = y.mean()
        prom_pred = pred.mean()
        crecimiento = (prom_pred - prom_actual) / prom_actual * 100 if prom_actual > 0 else 0

        p1, p2, p3 = st.columns(3)
        p1.metric(f"💵 Ventas esperadas ({dias_pred} días)", f"S/ {pred.sum():,.0f}")
        p2.metric("📊 Promedio diario esperado", f"S/ {prom_pred:,.0f}",
                  f"{crecimiento:+.1f}% vs. actual")
        tendencia = "📈 Creciente" if pendiente > 0 else ("📉 Decreciente" if pendiente < 0 else "➡️ Estable")
        p3.metric("Tendencia", tendencia, f"{pendiente:+,.1f} S/ por día")

        fig3, ax3 = plt.subplots(figsize=(10, 4.5))
        ax3.plot(serie.index, serie.values, marker="o", color="#1F77B4", label="Ventas reales")
        ax3.plot(serie.index, y_aj, linestyle=":", color="#6B7280", label="Tendencia ajustada")
        ax3.plot(pred.index, pred.values, marker="o", linestyle="--", color="#E4572E", label="Predicción")
        ax3.axvline(serie.index[-1], color="#9CA3AF", linewidth=1)
        ax3.set_ylabel("Ventas (S/)")
        ax3.set_title("Ventas reales vs. predicción")
        ax3.legend()
        fig3.autofmt_xdate()
        sns.despine()
        st.pyplot(fig3)
        png_pred = fig_a_png(fig3)

        if r2 < 0.3:
            st.warning(f"Ajuste bajo (R² = {r2:.2f}): las ventas varían mucho de un día a otro, "
                       "así que la tendencia es poco confiable.")
        else:
            st.success(f"Ajuste del modelo: R² = {r2:.2f}")

        pred_df = pd.DataFrame({"Fecha": fechas_fut, "Ventas": y_fut.round(2)})

        with st.expander("📋 Ver tabla de predicción"):
            st.dataframe(
                pd.DataFrame({"Fecha": fechas_fut.strftime("%d/%m/%Y"),
                              "Ventas estimadas (S/)": y_fut.round(2)}),
                hide_index=True
            )

# =====================================================
# PESTAÑA 3: DATOS
# =====================================================
with tab_datos:
    st.subheader("📋 Datos de ventas")
    st.caption(f"{len(df)} registros con los filtros actuales")
    st.dataframe(df, hide_index=True)

    st.markdown("---")
    d1, d2 = st.columns(2)
    with d1:
        st.subheader("Resumen por Producto")
        st.dataframe(res_prod)
    with d2:
        st.subheader("Resumen por Vendedor")
        st.dataframe(res_vend)

# =====================================================
# PESTAÑA 4: EXPORTAR
# =====================================================
with tab_export:
    st.subheader("📤 Exportar reporte")
    st.caption("Los archivos se generan con los filtros que tengas aplicados en el panel lateral.")

    kpis = {
        "ventas_totales": ventas_totales,
        "cantidad_total": int(cantidad_total),
        "ticket_promedio": ticket_promedio,
    }
    hallazgos = [
        f"Producto lider: {mejor_producto} (S/ {ventas_producto.iloc[0]:,.0f}, {pct_producto:.1f}% del total).",
        f"Mejor vendedor: {mejor_vendedor} (S/ {ventas_vendedor.iloc[0]:,.0f}, {pct_vendedor:.1f}% del total).",
        f"Mejor dia de ventas: {mejor_dia:%d/%m/%Y} (S/ {ventas_dia.max():,.0f}).",
    ]
    filtros_txt = f"Vendedores: {', '.join(vendedores)} | Productos: {', '.join(productos)}"
    sello = datetime.now().strftime("%Y%m%d_%H%M")

    excel_bytes = generar_excel(df, kpis, res_prod, res_vend, pred_df)
    pdf_bytes = generar_pdf(df, kpis, hallazgos, filtros_txt, png_prod, png_vend, png_pred,
                            res_prod, res_vend, pred_df)

    e1, e2, e3 = st.columns(3)
    with e1:
        st.markdown("#### 📗 Excel")
        st.write("Hojas: Resumen, Datos, Por Producto (con gráfico), Por Vendedor y Predicción.")
        st.download_button(
            "⬇️ Descargar Excel (.xlsx)",
            data=excel_bytes,
            file_name=f"reporte_ventas_{sello}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    with e2:
        st.markdown("#### 📕 PDF")
        st.write("Reporte con KPIs, hallazgos, gráficos, resúmenes, predicción y detalle.")
        st.download_button(
            "⬇️ Descargar PDF",
            data=pdf_bytes,
            file_name=f"reporte_ventas_{sello}.pdf",
            mime="application/pdf",
            use_container_width=True
        )
    with e3:
        st.markdown("#### 📄 CSV")
        st.write("Solo los datos filtrados, en formato plano.")
        st.download_button(
            "⬇️ Descargar CSV",
            data=df.to_csv(index=False).encode("utf-8"),
            file_name=f"ventas_filtradas_{sello}.csv",
            mime="text/csv",
            use_container_width=True
        )

# ---------- PIE ----------
st.caption("Dashboard desarrollado con Python + Streamlit · SENATI")
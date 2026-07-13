"""
Panel MRP - Histórico y Pronóstico de Demanda | Enaex S.A.
==========================================================

Panel de visualización que reemplaza al reporte de Power BI. Para cada
material muestra, en un solo gráfico, sus ingresos y egresos mensuales
(barras) y el stock de cierre (línea), junto a su clasificación de demanda
y el pronóstico del próximo mes.

Ejecutar con:
    streamlit run app.py
"""

from __future__ import annotations

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from mrp import config
from mrp.pipeline import construir

# --------------------------------------------------------------------------
# CONFIGURACIÓN DE LA PÁGINA
# --------------------------------------------------------------------------
st.set_page_config(
    page_title="MRP · Demanda y Pronóstico | Enaex",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
)

# --------------------------------------------------------------------------
# ESTILO (sala de control: pizarra + acero, códigos en monoespaciada)
# --------------------------------------------------------------------------
st.markdown(
    """
    <style>
      :root {
        --tinta:  #1B2A3A;   /* pizarra   */
        --acero:  #14618A;   /* acento    */
        --niebla: #6B7A8D;   /* texto sec */
        --linea:  #E2E6EB;
      }
      .stApp { background: #F5F6F8; }
      /* Cabecera */
      .mrp-header {
        background: linear-gradient(100deg, #1B2A3A 0%, #14618A 100%);
        color: #FFFFFF; padding: 22px 28px; border-radius: 14px;
        margin-bottom: 18px;
      }
      .mrp-header h1 {
        font-size: 1.55rem; font-weight: 700; margin: 0; letter-spacing: .2px;
      }
      .mrp-header p { margin: 4px 0 0; opacity: .82; font-size: .92rem; }
      /* Tarjetas de métrica */
      .metric-card {
        background: #FFFFFF; border: 1px solid var(--linea);
        border-radius: 12px; padding: 14px 16px; height: 100%;
        border-left: 4px solid var(--acero);
      }
      .metric-card .lbl {
        font-size: .72rem; text-transform: uppercase; letter-spacing: .6px;
        color: var(--niebla); margin-bottom: 4px;
      }
      .metric-card .val {
        font-size: 1.35rem; font-weight: 700; color: var(--tinta);
        line-height: 1.15;
      }
      .metric-card .sub { font-size: .78rem; color: var(--niebla); }
      /* Códigos en monoespaciada */
      .mono { font-family: "SFMono-Regular", "Consolas", monospace; }
      /* Chips de clasificación */
      .chip {
        display: inline-block; padding: 3px 12px; border-radius: 999px;
        font-size: .82rem; font-weight: 600;
      }
      .chip-suave        { background:#E3F2FD; color:#1565C0; }
      .chip-erratica     { background:#FFF3E0; color:#E65100; }
      .chip-intermitente { background:#F3E5F5; color:#6A1B9A; }
      .chip-irregular    { background:#FBE9E7; color:#C62828; }
      .chip-sin          { background:#ECEFF1; color:#546E7A; }
      h2, h3 { color: var(--tinta); }
    </style>
    """,
    unsafe_allow_html=True,
)


# --------------------------------------------------------------------------
# CARGA DE DATOS (cacheada para no recalcular en cada interacción)
# --------------------------------------------------------------------------
@st.cache_data(show_spinner="Calculando demanda y pronósticos…")
def cargar():
    r = construir()
    return r.serie, r.clasificacion, r.resultado, r.tabla_final


def _chip(tipo: str) -> str:
    clase = {
        "Suave": "chip-suave", "Errática": "chip-erratica",
        "Intermitente": "chip-intermitente", "Irregular": "chip-irregular",
        "Sin demanda": "chip-sin",
    }.get(tipo, "chip-sin")
    return f'<span class="chip {clase}">{tipo}</span>'


def _card(lbl: str, val: str, sub: str = "") -> str:
    return (
        f'<div class="metric-card"><div class="lbl">{lbl}</div>'
        f'<div class="val">{val}</div><div class="sub">{sub}</div></div>'
    )


# --------------------------------------------------------------------------
# GRÁFICO COMBINADO: barras (entrada/demanda) + línea (stock) + pronóstico
# --------------------------------------------------------------------------
def grafico_material(serie_mat: pd.DataFrame, info: pd.Series | None) -> go.Figure:
    fig = go.Figure()

    fig.add_bar(
        x=serie_mat["FechaMes"], y=serie_mat["Entrada Mensual"],
        name="Ingreso de material", marker_color=config.COLOR_ENTRADA,
        hovertemplate="%{x|%b %Y}<br>Ingreso: %{y:.0f}<extra></extra>",
    )
    fig.add_bar(
        x=serie_mat["FechaMes"], y=serie_mat["Demanda Mensual"],
        name="Egreso por uso (demanda)", marker_color=config.COLOR_DEMANDA,
        hovertemplate="%{x|%b %Y}<br>Demanda: %{y:.0f}<extra></extra>",
    )

    stock = serie_mat.dropna(subset=["Stock de cierre"])
    if not stock.empty:
        fig.add_trace(go.Scatter(
            x=stock["FechaMes"], y=stock["Stock de cierre"],
            name="Stock de cierre", mode="lines+markers",
            line=dict(color=config.COLOR_STOCK, width=2.5),
            marker=dict(size=5), yaxis="y2",
            hovertemplate="%{x|%b %Y}<br>Stock: %{y:.0f}<extra></extra>",
        ))

    # Punto de pronóstico del próximo mes
    if info is not None and pd.notna(info.get("MesPronosticado")):
        fig.add_trace(go.Scatter(
            x=[info["MesPronosticado"]], y=[info["Pronostico"]],
            name="Pronóstico próximo mes", mode="markers",
            marker=dict(color="#F39C12", size=14, symbol="star",
                        line=dict(color="#B9770E", width=1)),
            hovertemplate="%{x|%b %Y}<br>Pronóstico: %{y:.1f}<extra></extra>",
        ))

    fig.update_layout(
        barmode="group",
        height=460,
        margin=dict(l=10, r=10, t=30, b=10),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0),
        plot_bgcolor="#FFFFFF", paper_bgcolor="#FFFFFF",
        hovermode="x unified",
        xaxis=dict(title="", showgrid=False),
        yaxis=dict(title="Cantidad (ingresos / egresos)",
                   showgrid=True, gridcolor="#EEF1F4"),
        yaxis2=dict(title="Stock de cierre", overlaying="y",
                    side="right", showgrid=False),
    )
    return fig


# --------------------------------------------------------------------------
# APP
# --------------------------------------------------------------------------
def main():
    st.markdown(
        '<div class="mrp-header"><h1>📦 MRP · Histórico y Pronóstico de Demanda</h1>'
        '<p>Enaex S.A. — Planificación de materiales · Datos de SAP HANA (MB51 + MB5B)</p></div>',
        unsafe_allow_html=True,
    )

    # --- Cargar datos (con manejo de error si faltan archivos) ---
    try:
        serie, clasif, resultado, tabla = cargar()
    except FileNotFoundError as e:
        st.error("No se encontraron los archivos de datos.")
        st.info(
            "Deja los Excel de SAP en `data/MB51/` y `data/MB5B/`, "
            "o genera datos de ejemplo ejecutando en la terminal:\n\n"
            "```\npython scripts/generar_datos_ejemplo.py\n```"
        )
        st.caption(f"Detalle técnico: {e}")
        st.stop()

    tab_panel, tab_ayuda = st.tabs(["📊  Panel", "📖  Cómo usar y actualizar"])

    # ======================================================================
    # TAB 1 — PANEL
    # ======================================================================
    with tab_panel:
        # ---- Filtros laterales ----
        with st.sidebar:
            st.markdown("### Filtros")
            centros = sorted(tabla["Centro"].dropna().unique().tolist())
            centro_sel = st.multiselect("Centro", centros, default=centros)

            tipos = ["Suave", "Errática", "Intermitente", "Irregular", "Sin demanda"]
            tipos_pres = [t for t in tipos if t in tabla["Tipo_demanda"].unique()]
            tipo_sel = st.multiselect("Clasificación de demanda",
                                      tipos_pres, default=tipos_pres)

            st.markdown("---")
            ultimo_mes = serie["FechaMes"].max()
            st.caption(f"Datos hasta: **{ultimo_mes:%B %Y}**")
            st.caption(f"Materiales cargados: **{tabla['Material'].nunique()}**")

        tabla_f = tabla[
            tabla["Centro"].isin(centro_sel)
            & tabla["Tipo_demanda"].isin(tipo_sel)
        ].copy()

        if tabla_f.empty:
            st.warning("Ningún material coincide con los filtros seleccionados.")
            st.stop()

        # ---- Buscador de material (escribible) ----
        tabla_f["etiqueta"] = (
            tabla_f["Material"].astype(str) + "  —  "
            + tabla_f["Descripción del material"].fillna("(sin descripción)")
        )
        etiquetas = tabla_f.sort_values("Material")["etiqueta"].tolist()

        st.markdown("#### Selecciona un material")
        st.caption("Escribe el código o parte de la descripción para filtrar.")
        etiqueta_sel = st.selectbox("Material", etiquetas, label_visibility="collapsed")
        material_sel = etiqueta_sel.split("  —  ")[0].strip()

        info = tabla_f[tabla_f["Material"] == material_sel].iloc[0]

        # ---- Tarjetas de información del material ----
        c1, c2, c3, c4, c5 = st.columns(5)
        with c1:
            st.markdown(_card("Material",
                              f'<span class="mono">{info["Material"]}</span>',
                              f'Centro {info["Centro"]}'),
                        unsafe_allow_html=True)
        with c2:
            st.markdown(_card("Clasificación", _chip(info["Tipo_demanda"]),
                              f'Método: {info["Metodo"]}'),
                        unsafe_allow_html=True)
        with c3:
            pron = info.get("Pronostico_redondeado")
            pron_txt = "—" if pd.isna(pron) else f"{int(pron)}"
            sub = ("próx. mes" if pd.notna(info.get("MesPronosticado"))
                   else "sin demanda")
            st.markdown(_card("Pronóstico", pron_txt, sub), unsafe_allow_html=True)
        with c4:
            ph = info.get("PR_Periodos_Hasta_Prox")
            ph_txt = "—" if pd.isna(ph) else f"{ph:g} meses"
            st.markdown(_card("Hasta próx. demanda", ph_txt,
                              "solo intermitentes"),
                        unsafe_allow_html=True)
        with c5:
            mi = info.get("PR_Media_Intervalo")
            mi_txt = "—" if pd.isna(mi) else f"{mi:g} meses"
            st.markdown(_card("Intervalo de demanda", mi_txt,
                              "promedio entre demandas"),
                        unsafe_allow_html=True)

        st.markdown("")

        # ---- Gráfico combinado ----
        st.markdown(f"#### Histórico de **{info['Descripción del material'] or material_sel}**")
        serie_mat = serie[
            (serie["Material"] == material_sel)
            & (serie["Centro"] == info["Centro"])
        ].sort_values("FechaMes")
        st.plotly_chart(grafico_material(serie_mat, info), use_container_width=True)

        # ---- Detalle ampliado ----
        with st.expander("Ver todos los datos de pronóstico de este material"):
            detalle = info.drop(labels=["etiqueta"], errors="ignore")
            st.dataframe(
                detalle.rename("Valor").to_frame(),
                use_container_width=True,
            )

        st.markdown("---")

        # ---- Tabla final con todos los materiales ----
        st.markdown("#### Tabla de materiales (según filtros)")
        cols_tabla = [
            "Material", "Descripción del material", "Centro", "Tipo_demanda",
            "Metodo", "Pronostico_redondeado", "MesPronosticado",
            "PR_Periodos_Hasta_Prox", "PR_Media_Intervalo",
            "PR_Tamano_Esperado", "PR_IC95_Inf_Tamano", "PR_IC95_Sup_Tamano",
        ]
        cols_tabla = [c for c in cols_tabla if c in tabla_f.columns]
        vista = tabla_f[cols_tabla].sort_values("Material")

        st.dataframe(
            vista, use_container_width=True, hide_index=True,
            column_config={
                "Material": "Material",
                "Descripción del material": "Descripción",
                "Pronostico_redondeado": st.column_config.NumberColumn("Pronóstico"),
                "MesPronosticado": st.column_config.DateColumn("Mes pronosticado",
                                                               format="MMM YYYY"),
                "PR_Periodos_Hasta_Prox": "Hasta próx. dem.",
                "PR_Media_Intervalo": "Intervalo medio",
                "PR_Tamano_Esperado": "Tamaño esperado",
            },
        )

        st.download_button(
            "⬇️  Descargar tabla (CSV)",
            data=vista.to_csv(index=False).encode("utf-8-sig"),
            file_name="mrp_materiales.csv",
            mime="text/csv",
        )

    # ======================================================================
    # TAB 2 — AYUDA
    # ======================================================================
    with tab_ayuda:
        _mostrar_ayuda()


def _mostrar_ayuda():
    st.markdown(
        """
### Qué muestra este panel

Para cada material del MRP verás su **historia mensual** y su **pronóstico**:

- **Barras azules** — ingresos de material (movimientos de entrada, clase 101).
- **Barras rojas** — egresos por uso, es decir la **demanda real** (consumos, clases 201 y 261).
- **Línea verde** — stock de cierre de cada mes (viene de MB5B).
- **Estrella naranja** — pronóstico de demanda del próximo mes.

Arriba, cinco tarjetas resumen la **clasificación** del material, el **método**
de pronóstico usado, el **valor pronosticado**, y —para demanda intermitente—
cuántos meses faltan hasta la próxima demanda y el intervalo promedio entre demandas.

---

### Cómo usarlo

1. En la barra lateral, filtra por **Centro** y por **clasificación de demanda** si lo necesitas.
2. En **"Selecciona un material"**, escribe el código o parte de la descripción; la lista se filtra sola.
3. Revisa las tarjetas y el gráfico. Pasa el cursor por el gráfico para ver los valores exactos.
4. Al final está la **tabla completa** de materiales, que puedes descargar en CSV.

---

### Cómo actualizar los datos (cada mes)

El panel se alimenta de dos descargas de **SAP HANA**:

| Transacción | Qué es | Cada cuánto | Dónde va |
|---|---|---|---|
| **MB51** | Movimientos de material (ingresos y consumos), histórico | Mensual (o acumulado) | `data/MB51/` |
| **MB5B** | Stock del mes por material | **Un archivo por mes** | `data/MB5B/` |

**Paso a paso:**

1. Descarga desde SAP el Excel de **MB51** con los movimientos (desde 01-01-2023 hasta hoy)
   y déjalo en la carpeta `data/MB51/`. Puedes reemplazar el anterior o dejar varios; se leen todos.
2. Descarga el **MB5B** del mes recién cerrado y déjalo en `data/MB5B/`
   **sin borrar los meses anteriores** (cada archivo es la foto de stock de un mes).
3. Vuelve al panel y pulsa **"Rerun"** (menú ⋮ arriba a la derecha) o refresca la página.
   El cálculo se rehace solo con los datos nuevos.

> **Importante sobre las columnas:** el panel busca las columnas por su nombre
> (Material, Centro, Clase de movimiento, Fecha contabiliz., Ctd.en UM entrada
> en MB51; Material, De fecha, Stock de cierre, etc. en MB5B). Si SAP cambia
> ligeramente un encabezado, el panel intenta reconocer variantes comunes; si
> aun así falla, avisará qué columna no encontró.

---

### Cómo se calcula el pronóstico

1. **Demanda mensual:** se suman los consumos (201/261) por material y mes.
2. **Clasificación ADI / CV²** (Syntetos & Boylan): según cada cuánto hay demanda
   (ADI) y cuánto varía su tamaño (CV²), el material queda como
   *Suave*, *Errática*, *Intermitente* o *Irregular*.
3. **Método según el tipo:**
   - *Suave* → **SES** (suavizamiento exponencial simple).
   - *Errática* → **COMBINADO** (promedio de SES + media móvil 3 + media móvil 6).
   - *Intermitente / Irregular* → **SBA** (Croston con corrección) y, además, un
     **Proceso de Renovación** que estima cuándo y cuánto será la próxima demanda.
4. **Resultado final:** se toma el último pronóstico de cada material.

Todos los parámetros (factor α, cortes de clasificación, horizonte) están en
`mrp/config.py` y se pueden ajustar sin tocar la lógica.
        """
    )


if __name__ == "__main__":
    main()

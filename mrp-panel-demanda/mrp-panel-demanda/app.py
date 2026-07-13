"""
Panel MRP - Histórico y Pronóstico de Demanda | Enaex S.A.
==========================================================

Panel de visualización que reemplaza al reporte de Power BI. Para cada
material muestra, en un solo gráfico, sus ingresos y egresos mensuales
(barras) y el stock de cierre (línea), junto a su clasificación de demanda,
el método usado, el pronóstico y el tiempo estimado hasta la próxima demanda.

Ejecutar con:
    streamlit run app.py
"""

from __future__ import annotations

from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from mrp import config, data_loading, github_store
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
        --tinta:  #1B2A3A; --acero: #14618A; --niebla: #6B7A8D; --linea: #E2E6EB;
      }
      .stApp { background: #F5F6F8; }
      .mrp-header {
        background: linear-gradient(100deg, #1B2A3A 0%, #14618A 100%);
        color: #FFFFFF; padding: 22px 28px; border-radius: 14px; margin-bottom: 18px;
      }
      .mrp-header h1 { font-size: 1.55rem; font-weight: 700; margin: 0; }
      .mrp-header p { margin: 4px 0 0; opacity: .82; font-size: .92rem; }
      .metric-card {
        background: #FFFFFF; border: 1px solid var(--linea); border-radius: 12px;
        padding: 14px 16px; height: 100%; border-left: 4px solid var(--acero);
      }
      .metric-card .lbl {
        font-size: .72rem; text-transform: uppercase; letter-spacing: .6px;
        color: var(--niebla); margin-bottom: 4px;
      }
      .metric-card .val { font-size: 1.3rem; font-weight: 700; color: var(--tinta); line-height: 1.15; }
      .metric-card .sub { font-size: .78rem; color: var(--niebla); }
      .mono { font-family: "SFMono-Regular", "Consolas", monospace; }
      .chip { display: inline-block; padding: 3px 12px; border-radius: 999px;
              font-size: .82rem; font-weight: 600; }
      .chip-constante    { background:#E3F2FD; color:#1565C0; }
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
# CARGA DE DATOS (cacheada)
# --------------------------------------------------------------------------
@st.cache_data(show_spinner="Calculando demanda y pronósticos…")
def cargar():
    """Lee de las carpetas data/MB51 y data/MB5B y ejecuta el pipeline."""
    r = construir()
    return r.serie, r.clasificacion, r.resultado, r.tabla_final


def _chip(tipo: str) -> str:
    clase = {
        "Constante": "chip-constante", "Errática": "chip-erratica",
        "Intermitente": "chip-intermitente", "Irregular": "chip-irregular",
        "Sin demanda": "chip-sin",
    }.get(tipo, "chip-sin")
    return f'<span class="chip {clase}">{tipo}</span>'


def _card(lbl: str, val: str, sub: str = "") -> str:
    return (f'<div class="metric-card"><div class="lbl">{lbl}</div>'
            f'<div class="val">{val}</div><div class="sub">{sub}</div></div>')


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
    if info is not None and pd.notna(info.get("MesPronosticado")):
        indet = str(info.get("Tiempo_hasta_demanda", "")) == "Indeterminado"
        if indet:
            # Demanda intermitente con pocos datos: no se sabe cuándo -> se marca con *
            fig.add_trace(go.Scatter(
                x=[info["MesPronosticado"]], y=[info["Pronostico"]],
                name="Pronóstico (fecha indeterminada) *", mode="markers",
                marker=dict(color="#8E44AD", size=16, symbol="asterisk",
                            line=dict(color="#8E44AD", width=2)),
                hovertemplate="Pronóstico: %{y:.1f}<br>(momento indeterminado)<extra></extra>",
            ))
        else:
            fig.add_trace(go.Scatter(
                x=[info["MesPronosticado"]], y=[info["Pronostico"]],
                name="Pronóstico próxima demanda ★", mode="markers",
                marker=dict(color="#F39C12", size=14, symbol="star",
                            line=dict(color="#B9770E", width=1)),
                hovertemplate="%{x|%b %Y}<br>Pronóstico: %{y:.1f}<extra></extra>",
            ))
    fig.update_layout(
        barmode="group", height=460, margin=dict(l=10, r=10, t=30, b=10),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0),
        plot_bgcolor="#FFFFFF", paper_bgcolor="#FFFFFF", hovermode="x unified",
        xaxis=dict(title="", showgrid=False),
        yaxis=dict(title="Cantidad (ingresos / egresos)", showgrid=True, gridcolor="#EEF1F4"),
        yaxis2=dict(title="Stock de cierre", overlaying="y", side="right", showgrid=False),
    )
    return fig


# --------------------------------------------------------------------------
# PESTAÑA 1 — PANEL
# --------------------------------------------------------------------------
def render_panel():
    # --- Diagnóstico: qué archivos ve la app en las carpetas (lectura en vivo) ---
    with st.expander("🔍 Ver qué archivos está leyendo la app (diagnóstico)"):
        diag = data_loading.archivos_detectados()
        st.markdown(f"**Carpeta MB51:** `{diag['mb51_ruta']}`")
        st.write(diag["mb51_archivos"])
        st.markdown(f"**Carpeta MB5B:** `{diag['mb5b_ruta']}`")
        st.write(diag["mb5b_archivos"])
        st.caption(
            "Si aquí ves tus archivos reales pero el panel muestra otros datos, "
            "es la caché: pulsa **🔄 Recargar datos** (barra lateral) o reinicia "
            "la app (Manage app → Reboot). Si ves archivos de ejemplo, bórralos "
            "del repositorio."
        )

    try:
        serie, clasif, resultado, tabla = cargar()
    except FileNotFoundError:
        st.info(
            "Todavía no hay datos cargados. Ve a la pestaña **➕ Agregar datos** "
            "y sube tu Excel de **MB51** y al menos uno de **MB5B**. "
            "También puedes generar datos de ejemplo con "
            "`python scripts/generar_datos_ejemplo.py`."
        )
        return

    # ---- Filtros (barra lateral) ----
    with st.sidebar:
        st.markdown("### Filtros")
        centros = sorted(tabla["Centro"].dropna().unique().tolist())
        centro_sel = st.multiselect("Centro", centros, default=centros)
        tipos = ["Constante", "Errática", "Intermitente", "Irregular", "Sin demanda"]
        tipos_pres = [t for t in tipos if t in tabla["Tipo_demanda"].unique()]
        tipo_sel = st.multiselect("Clasificación de demanda", tipos_pres, default=tipos_pres)
        st.markdown("---")
        st.caption(f"Datos hasta: **{serie['FechaMes'].max():%b %Y}**")
        st.caption(f"Materiales: **{tabla['Material'].nunique()}**")
        if st.button("🔄 Recargar datos"):
            st.cache_data.clear()
            st.rerun()

        # Imagen explicativa de los tipos de demanda (clic para ampliar)
        st.markdown("---")
        with st.expander("❓ ¿Qué significan los tipos de demanda?"):
            img = Path(__file__).parent / "assets" / "Tipos_de_demanda.png"
            if img.exists():
                st.image(str(img), use_container_width=True,
                         caption="Pasa el cursor y usa el ícono ⛶ para ampliar.")
            else:
                st.caption("(No se encontró la imagen Tipos_de_demanda.png en assets/.)")

    tabla_f = tabla[
        tabla["Centro"].isin(centro_sel) & tabla["Tipo_demanda"].isin(tipo_sel)
    ].copy()
    if tabla_f.empty:
        st.warning("Ningún material coincide con los filtros seleccionados.")
        return

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

    st.markdown(f"### {info['Descripción del material'] or material_sel}")

    # ---- Tarjetas ----
    c1, c2, c3, c4, c5 = st.columns(5)
    with c1:
        st.markdown(_card("Material", f'<span class="mono">{info["Material"]}</span>',
                          f'Centro {info["Centro"]}'), unsafe_allow_html=True)
    with c2:
        st.markdown(_card("Clasificación", _chip(info["Tipo_demanda"]),
                          f'Método: {info["Metodo"]}'), unsafe_allow_html=True)
    with c3:
        pron = info.get("Pronostico_redondeado")
        pron_txt = "—" if pd.isna(pron) else f"{int(pron)}"
        st.markdown(_card("Pronóstico", pron_txt, "próxima demanda"), unsafe_allow_html=True)
    with c4:
        st.markdown(_card("Tiempo hasta demanda",
                          str(info.get("Tiempo_hasta_demanda", "—")),
                          "estimado"), unsafe_allow_html=True)
    with c5:
        mi = info.get("PR_Media_Intervalo")
        mi_txt = "—" if pd.isna(mi) else f"{mi:g} meses"
        st.markdown(_card("Intervalo de demanda", mi_txt,
                          "promedio entre demandas"), unsafe_allow_html=True)

    st.markdown("")

    # ---- Gráfico combinado ----
    serie_mat = serie[
        (serie["Material"] == material_sel) & (serie["Centro"] == info["Centro"])
    ].sort_values("FechaMes")
    st.plotly_chart(grafico_material(serie_mat, info), use_container_width=True)

    # ---- Detalle ampliado ----
    with st.expander("Ver todos los datos de este material"):
        detalle = info.drop(labels=["etiqueta"], errors="ignore")
        st.dataframe(detalle.rename("Valor").to_frame(), use_container_width=True)

    st.markdown("---")

    # ---- Tabla final ----
    st.markdown("#### Tabla de materiales (según filtros)")
    cols_tabla = [
        "Material", "Descripción del material", "Centro", "Tipo_demanda", "Metodo",
        "Pronostico_redondeado", "Tiempo_hasta_demanda", "MesPronosticado",
        "PR_Periodos_Hasta_Prox", "PR_Media_Intervalo", "PR_Tamano_Esperado",
        "PR_IC95_Inf_Tamano", "PR_IC95_Sup_Tamano",
    ]
    cols_tabla = [c for c in cols_tabla if c in tabla_f.columns]
    vista = tabla_f[cols_tabla].sort_values("Material")
    st.dataframe(
        vista, use_container_width=True, hide_index=True,
        column_config={
            "Descripción del material": "Descripción",
            "Tipo_demanda": "Tipo de demanda",
            "Metodo": "Método",
            "Pronostico_redondeado": st.column_config.NumberColumn("Pronóstico"),
            "Tiempo_hasta_demanda": "Tiempo hasta demanda",
            "MesPronosticado": st.column_config.DateColumn("Mes pronosticado", format="MMM YYYY"),
            "PR_Periodos_Hasta_Prox": "Períodos hasta dem.",
            "PR_Media_Intervalo": "Intervalo medio (meses)",
            "PR_Tamano_Esperado": "Tamaño esperado",
        },
    )
    st.download_button(
        "⬇️  Descargar tabla (CSV)",
        data=vista.to_csv(index=False).encode("utf-8-sig"),
        file_name="mrp_materiales.csv", mime="text/csv",
    )


# --------------------------------------------------------------------------
# PESTAÑA 2 — AGREGAR DATOS
# --------------------------------------------------------------------------
def render_agregar_datos():
    st.markdown("#### Agregar datos nuevos")
    st.caption(
        "Aquí actualizas los Excel que alimentan el panel. "
        "**MB51** se descarga cada semana y reemplaza al anterior. "
        "**MB5B** se agrega uno nuevo cada mes."
    )

    st.warning(
        "📅 **Recordatorio:** al inicio de **cada mes** debes agregar el nuevo "
        "**MB5B** del mes recién cerrado. El **MB51** conviene actualizarlo cada semana."
    )

    # ---- Instrucciones de descarga desde SAP HANA ----
    with st.expander("📥 Cómo descargar los Excel desde SAP HANA (layouts)"):
        st.markdown(
            """
**MB51 — movimientos**
1. Entra a la transacción **MB51**.
2. Aplica el layout **`/CALCDEMANDA`** — *"MOV. PARA PRONOSTICO DE DEMANDA"*.
3. Exporta a Excel. Ese archivo es el que subes en **"1) MB51"**.

**MB5B — stock del mes**
1. Entra a la transacción **MB5B**.
2. Usa un layout que deje **exactamente estas columnas** (en este orden):

   `Material · Descripción del material · De fecha · A fecha · Stock inicial ·
   Total ctd.entrada mcía. · Total cantidades salida · Stock de cierre ·
   Unidad medida base · Stock especial`
3. Exporta a Excel (uno por mes) y súbelo en **"2) MB5B"**.
            """
        )

    # ---- Dónde se guardarán los archivos ----
    modo_github = github_store.disponible()
    if modo_github:
        st.success("🟢 Los archivos se guardarán **directamente en el repositorio de GitHub** "
                   "(quedan permanentes). La app se actualizará sola en ~1 minuto.")
    else:
        st.info("🟡 Guardado **local**: en la nube estos archivos son temporales. "
                "Para guardarlos en GitHub automáticamente, configura los *secrets* "
                "(ver README, sección 'Guardar en GitHub desde la app').")

    # ---- Contraseña ----
    autorizado = True
    if github_store.password_configurada():
        clave = st.text_input("🔒 Contraseña para agregar datos", type="password")
        autorizado = github_store.password_ok(clave)
        if not autorizado:
            st.caption("Ingresa la contraseña para habilitar la carga de archivos.")
    else:
        st.caption("⚠️ No hay contraseña configurada. Se recomienda agregar "
                   "`APP_PASSWORD` en los *secrets* (ver README).")

    # Estado actual de MB5B (protegido)
    try:
        estado = data_loading.estado_mb5b()
        if estado["meses"]:
            meses_txt = ", ".join(m.strftime("%b %Y") for m in estado["meses"])
            st.info(f"Meses de MB5B cargados ({estado['n_archivos']}): {meses_txt}")
        if estado["falta"] and estado["mes_faltante"] is not None:
            st.error(f"Parece que falta cargar el **MB5B de "
                     f"{estado['mes_faltante']:%B %Y}**. Súbelo abajo.")
    except Exception as e:
        st.caption(f"(No se pudo revisar el estado de MB5B: {e})")

    if not autorizado:
        return

    def _guardar(archivo, es_mb51: bool):
        contenido = archivo.getvalue()
        if modo_github:
            if es_mb51:
                github_store.guardar_mb51(archivo.name, contenido)
            else:
                github_store.agregar_mb5b(archivo.name, contenido)
            return "en GitHub (la app se actualizará sola en ~1 min)"
        else:
            if es_mb51:
                data_loading.reemplazar_mb51(archivo)
            else:
                data_loading.agregar_mb5b(archivo)
            return "localmente"

    col1, col2 = st.columns(2)
    with col1:
        st.markdown("##### 1) MB51 — movimientos (semanal)")
        st.caption("Al subirlo, **reemplaza** el MB51 anterior.")
        mb51_file = st.file_uploader("Sube el Excel de MB51", type=["xlsx", "xls"],
                                     key="up_mb51")
        if mb51_file is not None and st.button("Reemplazar MB51", type="primary"):
            try:
                donde = _guardar(mb51_file, es_mb51=True)
                st.cache_data.clear()
                st.success(f"MB51 guardado {donde}: {mb51_file.name}")
                if not modo_github:
                    st.rerun()
            except Exception as e:
                st.error(f"No se pudo guardar el MB51: {e}")

    with col2:
        st.markdown("##### 2) MB5B — stock del mes (mensual)")
        st.caption("Al subirlo, **se agrega** a los meses anteriores.")
        mb5b_file = st.file_uploader("Sube el Excel de MB5B del mes", type=["xlsx", "xls"],
                                     key="up_mb5b")
        if mb5b_file is not None and st.button("Agregar MB5B"):
            try:
                donde = _guardar(mb5b_file, es_mb51=False)
                st.cache_data.clear()
                st.success(f"MB5B guardado {donde}: {mb5b_file.name}")
                if not modo_github:
                    st.rerun()
            except Exception as e:
                st.error(f"No se pudo guardar el MB5B: {e}")

    st.markdown("---")
    st.caption(
        "ℹ️ **Nota para la versión en internet (Streamlit Cloud):** los archivos "
        "que subas aquí se guardan solo mientras la app está encendida y se "
        "pierden si la app se reinicia. Para que queden permanentes, súbelos al "
        "repositorio de GitHub (carpetas `data/MB51` y `data/MB5B`) o usa el panel "
        "en tu computador. Corriendo en tu PC, los archivos sí quedan guardados."
    )


# --------------------------------------------------------------------------
# PESTAÑA 3 — AYUDA
# --------------------------------------------------------------------------
def render_ayuda():
    st.markdown(
        """
### Qué muestra este panel

Para cada material verás su **historia mensual** y su **pronóstico**:

- **Barras azules** — ingresos de material (movimientos de entrada, clase 101).
- **Barras rojas** — egresos por uso, es decir la **demanda real** (consumos, clases 201 y 261).
- **Línea verde** — stock de cierre de cada mes (viene de MB5B).
- **Estrella naranja** — pronóstico de la próxima demanda.

Arriba, cinco tarjetas resumen el **material**, su **clasificación**, el **método**,
el **pronóstico**, el **tiempo hasta la próxima demanda** y el **intervalo** promedio.

---

### Cómo se elige el método y el "tiempo hasta demanda"

| Tipo de demanda | Método | Tiempo hasta la próxima demanda |
|---|---|---|
| **Constante** (antes "Suave") | **SES** | Días que faltan hasta el próximo mes |
| **Errática** | **COMBINADO** (SES + medias móviles 3 y 6) | Días que faltan hasta el próximo mes |
| **Intermitente / Irregular** con **menos de 4** demandas | **SBA** | **Indeterminado** |
| **Intermitente / Irregular** con **4 o más** demandas | **PR** (Proceso de Renovación) | **Días estimados** hasta la próxima demanda |
| **Sin demanda** | — | — |

La clasificación se hace con la metodología **ADI / CV²** (Syntetos & Boylan):
ADI mide cada cuánto hay demanda y CV² cuánto varía su tamaño.

---

### Cómo usar el panel

1. En la barra lateral, filtra por **Centro** y **clasificación** si lo necesitas.
2. En **"Selecciona un material"**, escribe el código o parte de la descripción.
3. Revisa las tarjetas y el gráfico (pasa el cursor para ver valores exactos).
4. Al final está la **tabla completa**, descargable en CSV.

---

### Cómo actualizar los datos

Ve a la pestaña **➕ Agregar datos**:

- **MB51** (movimientos): en SAP usa la transacción **MB51** con el layout
  **`/CALCDEMANDA`** ("MOV. PARA PRONOSTICO DE DEMANDA"), exporta a Excel y
  súbelo. Se descarga **cada semana** y **reemplaza** al anterior.
- **MB5B** (stock del mes): en SAP usa la transacción **MB5B** con un layout que
  deje estas columnas — Material, Descripción del material, De fecha, A fecha,
  Stock inicial, Total ctd.entrada mcía., Total cantidades salida, Stock de
  cierre, Unidad medida base, Stock especial. Se agrega **uno nuevo cada mes**.

Al subir un archivo, el panel recalcula todo automáticamente.

Todos los parámetros (α, cortes ADI/CV², umbral de PR, días por mes) están en
`mrp/config.py` y se pueden ajustar sin tocar la lógica.
        """
    )


# --------------------------------------------------------------------------
# APP
# --------------------------------------------------------------------------
def main():
    st.markdown(
        '<div class="mrp-header"><h1>📦 MRP · Histórico y Pronóstico de Demanda</h1>'
        '<p>Enaex S.A. — Planificación de materiales · Datos de SAP HANA (MB51 + MB5B)</p></div>',
        unsafe_allow_html=True,
    )
    tab_panel, tab_datos, tab_ayuda = st.tabs(
        ["📊  Panel", "➕  Agregar datos", "📖  Cómo usar"]
    )
    with tab_panel:
        try:
            render_panel()
        except Exception as e:
            st.error(f"Ocurrió un problema al mostrar el panel: {e}")
    with tab_datos:
        try:
            render_agregar_datos()
        except Exception as e:
            st.error(f"Ocurrió un problema en 'Agregar datos': {e}")
    with tab_ayuda:
        render_ayuda()


if __name__ == "__main__":
    main()

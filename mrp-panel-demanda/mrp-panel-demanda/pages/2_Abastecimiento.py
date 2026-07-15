"""
Panel 2 · Abastecimiento — seguimiento y control de materiales
==============================================================

Integra MRP semanal + MM60 + ME5A + ME2M + TAT + Demanda (panel 1).

Principio de diseño: NUNCA se pierden materiales. Todas las uniones son
'left join' sobre el MRP, así que un material sin solped, sin OC o sin TAT
igual aparece (con sus campos en blanco). Además hay pestañas separadas para
ver la gestión (validación / solped bloqueada / días) sin que se filtren filas.
"""

from __future__ import annotations

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from mrp import config, data_loading, github_store
from mrp.abastecimiento import construir_abastecimiento, estado_archivos

st.set_page_config(page_title="Abastecimiento | Enaex", page_icon="🚚", layout="wide")

st.markdown(
    """
    <style>
      .stApp { background: #F5F6F8; }
      .ab-header {
        background: linear-gradient(100deg, #1B2A3A 0%, #B9770E 100%);
        color: #fff; padding: 20px 26px; border-radius: 14px; margin-bottom: 16px;
      }
      .ab-header h1 { font-size: 1.5rem; margin: 0; }
      .ab-header p { margin: 4px 0 0; opacity: .85; font-size: .9rem; }
      .chip { display:inline-block; padding:3px 12px; border-radius:999px;
              font-size:.82rem; font-weight:600; }
      .e-oc     { background:#E3F2FD; color:#1565C0; }
      .e-solped { background:#E8F5E9; color:#2E7D32; }
      .e-bloq   { background:#FBE9E7; color:#C62828; }
      .e-valid  { background:#FFF3E0; color:#E65100; }
      .e-sin    { background:#ECEFF1; color:#546E7A; }
    </style>
    """,
    unsafe_allow_html=True,
)

COLOR_COND = {"Sobre Stock": "#2E86DE", "Stock OK": "#27AE60",
              "Bajo Stock": "#F39C12", "Quiebre Stock": "#E74C3C"}
COLOR_GEST = {"Con OC": "#1565C0", "Con Solped": "#2E7D32", "Solped bloqueada": "#C62828",
              "Validación": "#E65100", "Sin gestión": "#90A4AE"}
COLOR_CUMPLE = {"Cumple": "#27AE60", "Alerta": "#F39C12",
                "Urgente": "#E67E22", "No cumple": "#E74C3C",
                "Sin pronóstico": "#B0BEC5", "Sin stock dato": "#B0BEC5"}
COLOR_NAC = {"Nacional": "#2E86DE", "Internacional": "#8E44AD",
             "Otro": "#95A5A6", "Sin OC": "#CFD8DC"}


def _chip(estado: str) -> str:
    clase = {"Con OC": "e-oc", "Con Solped": "e-solped", "Solped bloqueada": "e-bloq",
             "Validación": "e-valid", "Sin gestión": "e-sin"}.get(estado, "e-sin")
    return f'<span class="chip {clase}">{estado}</span>'


@st.cache_data(show_spinner="Integrando MRP + MM60 + ME5A + ME2M + TAT + Demanda…")
def cargar():
    r = construir_abastecimiento()
    return r.tabla, r.kpis


def barras(conteo: dict, colores: dict, titulo: str, horizontal=False) -> go.Figure:
    conteo = {k: v for k, v in conteo.items() if pd.notna(k)}
    if horizontal:
        fig = go.Figure(go.Bar(
            y=list(conteo.keys()), x=list(conteo.values()), orientation="h",
            marker_color=[colores.get(k, "#607D8B") for k in conteo],
            text=list(conteo.values()), textposition="outside"))
    else:
        fig = go.Figure(go.Bar(
            x=list(conteo.keys()), y=list(conteo.values()),
            marker_color=[colores.get(k, "#607D8B") for k in conteo],
            text=list(conteo.values()), textposition="outside"))
    fig.update_layout(title=titulo, height=300, margin=dict(l=10, r=10, t=40, b=10),
                      plot_bgcolor="#fff", paper_bgcolor="#fff",
                      xaxis=dict(showgrid=False), yaxis=dict(showgrid=True, gridcolor="#EEF1F4"))
    return fig


st.markdown(
    '<div class="ab-header"><h1>🚚 Abastecimiento · Seguimiento y control de materiales</h1>'
    '<p>Enaex S.A. — Solped, OC, disponibilidad, TAT y demanda · MRP + MM60 + ME5A + ME2M + TAT</p></div>',
    unsafe_allow_html=True,
)

# ------------------------------------------------- Estado de los Excel cargados
def _mostrar_estado_archivos(expandido=False):
    est = estado_archivos()
    faltan = [e for e in est if not e["cargado"]]
    titulo = ("📁 Archivos cargados — todo listo" if not faltan
              else f"📁 Archivos: faltan {len(faltan)} por subir")
    with st.expander(titulo, expanded=expandido):
        st.caption("Estos son los Excel que alimentan este panel. "
                   "Se suben en la pestaña **➕ Actualizar datos**.")
        filas = []
        for e in est:
            filas.append({
                "Excel": e["nombre"] + (" (obligatorio)" if e["obligatorio"] else ""),
                "Estado": "✅ Cargado" if e["cargado"] else "❌ Falta",
                "Carpeta": f"data/{e['carpeta']}",
                "Archivo": ", ".join(e["archivos"]) if e["archivos"] else "—",
                "Para qué sirve": e["para_que"],
            })
        st.dataframe(pd.DataFrame(filas), use_container_width=True, hide_index=True)
        if faltan:
            st.warning(
                "Sin los archivos que faltan, el panel igual funciona, pero esas "
                "columnas quedarán en blanco (por ejemplo, sin ME5A no hay días de solped)."
            )


try:
    tabla, kpis = cargar()
except FileNotFoundError:
    st.error("Falta el **MRP semanal**, que es la base de este panel.")
    _mostrar_estado_archivos(expandido=True)
    st.info(
        "**Qué hacer:** sube el Excel `Planificacion_Simpl…xlsx` a la carpeta "
        "`data/MRP` (o desde la pestaña ➕ Actualizar datos de este panel). "
        "Luego pulsa **Manage app → Reboot** o el botón 🔄 Recargar datos."
    )
    st.stop()
except Exception as e:
    st.error(f"No se pudieron integrar los datos: {e}")
    _mostrar_estado_archivos(expandido=True)
    st.stop()

_mostrar_estado_archivos()

# ---------------------------------------------------------------- Filtros
with st.sidebar:
    st.markdown("### Filtros")

    def _multi(col, etiqueta):
        if col not in tabla.columns:
            return None
        opciones = sorted(str(x) for x in tabla[col].dropna().unique() if str(x).strip())
        return st.multiselect(etiqueta, opciones, default=opciones)

    f_centro = _multi("Centro", "Centro")
    f_area = _multi("Area", "Área")
    f_crit = _multi("Criticidad texto", "Criticidad")
    f_cond = _multi("Condicion Stock", "Condición de stock")
    f_gest = _multi("Estado gestión", "Estado de gestión")
    f_oc = _multi("Estado OC", "Estado de la OC")
    f_nac = _multi("Nacionalidad", "Nacionalidad de la OC")
    f_rec = _multi("Recurrencia", "Recurrencia de compra (TAT)")
    if st.button("🔄 Recargar datos"):
        st.cache_data.clear()
        st.rerun()

datos = tabla.copy()
for col, sel in [("Centro", f_centro), ("Area", f_area), ("Criticidad texto", f_crit),
                 ("Condicion Stock", f_cond), ("Estado gestión", f_gest),
                 ("Estado OC", f_oc), ("Nacionalidad", f_nac), ("Recurrencia", f_rec)]:
    if sel is not None and col in datos.columns:
        datos = datos[datos[col].astype(str).isin(sel)]

if datos.empty:
    st.warning("Ningún material coincide con los filtros seleccionados.")
    st.stop()

# ---------------------------------------------------------------- KPIs
total = len(datos)
sin_stock = int((datos["Condicion Stock"] == "Quiebre Stock").sum()) if "Condicion Stock" in datos else 0
dispo = round(100 * (total - sin_stock) / total, 1) if total else 0
k = st.columns(7)
k[0].metric("Materiales", f"{total:,}".replace(",", "."))
k[1].metric("Disponibilidad", f"{dispo} %")
k[2].metric("Sin stock", sin_stock)
k[3].metric("OC atrasadas", int((datos["Estado OC"] == "Atrasada").sum()))
k[4].metric("OC en curso", int((datos["Estado OC"] == "En curso").sum()))
k[5].metric("Solped bloqueadas", int((datos["Estado gestión"] == "Solped bloqueada").sum()))
k[6].metric("En validación", int((datos["Estado gestión"] == "Validación").sum()))

t1, t2, t3, t4, t5 = st.tabs(
    ["📌  Gestión (solped / OC)", "📊  Resumen", "📦  Demanda vs Stock",
     "📋  Todos los materiales", "➕  Actualizar datos"]
)

# ================================================================ 1) GESTIÓN
with t1:
    st.markdown("#### Gestión de solped y OC")
    st.caption(
        "Aquí se ve **qué está pasando con cada material**: si está en validación, "
        "si su solped está bloqueada, cuántos días lleva la solped o la OC. "
        "Cada bloque es una tabla propia, así **no desaparece ningún material** al "
        "agregar los días de gestión."
    )

    # --- Prioridad 1: validación y solped bloqueada ---
    st.markdown("##### 🚨 Prioridad: en validación y solped bloqueada")
    prio = datos[datos["Estado gestión"].isin(["Validación", "Solped bloqueada"])]
    if prio.empty:
        st.success("No hay materiales en validación ni con solped bloqueada.")
    else:
        cols_p = [c for c in ["Material", "Texto breve de material", "Centro", "Area",
                              "Criticidad texto", "Condicion Stock", "Stock",
                              "Estado gestión", "Observación", "Solped", "Días en solped",
                              "Usuario"] if c in prio.columns]
        st.dataframe(prio[cols_p].sort_values(["Estado gestión", "Material"]),
                     use_container_width=True, hide_index=True)
        st.caption(f"{len(prio)} materiales requieren gestión inmediata.")

    st.markdown("---")

    # --- Solped con días de gestión ---
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("##### 📄 Solped en curso (con días de gestión)")
        sol = datos[datos["Estado gestión"] == "Con Solped"]
        cols_s = [c for c in ["Material", "Texto breve de material", "Solped",
                              "Días en solped", "Rango días solped", "Condicion Stock",
                              "Criticidad texto"] if c in sol.columns]
        if sol.empty:
            st.info("No hay materiales con solped en curso.")
        else:
            st.dataframe(sol[cols_s].sort_values("Días en solped", ascending=False),
                         use_container_width=True, hide_index=True)
            if "Rango días solped" in sol.columns:
                orden = ["0-10 días", "11-20 días", "21-30 días", "31+ días"]
                conteo = {o: int((sol["Rango días solped"] == o).sum()) for o in orden}
                st.plotly_chart(barras(conteo, {}, "Antigüedad de las solped", horizontal=True),
                                use_container_width=True)
    with c2:
        st.markdown("##### 🚚 OC en tránsito (con días y atraso)")
        oc = datos[datos["Estado OC"].isin(["Atrasada", "En curso"])]
        cols_o = [c for c in ["Material", "Texto breve de material", "OC en Transito",
                              "Nacionalidad", "Estado OC", "Días de OC", "Días atraso OC",
                              "Días hasta llegada", "Proveedor"] if c in oc.columns]
        if oc.empty:
            st.info("No hay OC en tránsito.")
        else:
            st.dataframe(oc[cols_o].sort_values("Días atraso OC", ascending=False),
                         use_container_width=True, hide_index=True)
            if "Rango atraso OC" in oc.columns:
                orden = ["1-15 días", "16-30 días", "31-45 días", "46-60 días",
                         "61-75 días", ">75 días"]
                conteo = {o: int((oc["Rango atraso OC"] == o).sum()) for o in orden}
                conteo = {k_: v for k_, v in conteo.items() if v}
                if conteo:
                    st.plotly_chart(barras(conteo, {}, "Atraso de las OC", horizontal=True),
                                    use_container_width=True)

    st.markdown("---")
    st.markdown("##### 🌎 OC nacionales vs internacionales")
    cn1, cn2 = st.columns([1, 2])
    with cn1:
        conteo_n = {k_: int((datos["Nacionalidad"] == k_).sum())
                    for k_ in ["Nacional", "Internacional", "Otro"]}
        conteo_n = {k_: v for k_, v in conteo_n.items() if v}
        st.plotly_chart(barras(conteo_n, COLOR_NAC, "OC por nacionalidad"),
                        use_container_width=True)
    with cn2:
        nac = datos[datos["Nacionalidad"].isin(["Nacional", "Internacional", "Otro"])]
        cols_n = [c for c in ["Material", "Texto breve de material", "OC en Transito",
                              "Nacionalidad", "Estado OC", "Días atraso OC",
                              "TAT Promedio", "Recurrencia", "Proveedor"] if c in nac.columns]
        if not nac.empty:
            st.dataframe(nac[cols_n].sort_values("Nacionalidad"),
                         use_container_width=True, hide_index=True)
        else:
            st.info("No hay OC para clasificar.")

# ================================================================ 2) RESUMEN
with t2:
    g1, g2 = st.columns(2)
    with g1:
        if "Condicion Stock" in datos:
            orden = ["Sobre Stock", "Stock OK", "Bajo Stock", "Quiebre Stock"]
            conteo = {o: int((datos["Condicion Stock"] == o).sum()) for o in orden}
            conteo = {k_: v for k_, v in conteo.items() if v}
            st.plotly_chart(barras(conteo, COLOR_COND, "Materiales por condición de stock"),
                            use_container_width=True)
    with g2:
        st.plotly_chart(barras(dict(datos["Estado gestión"].value_counts()), COLOR_GEST,
                               "Materiales por estado de gestión"), use_container_width=True)

    g3, g4 = st.columns(2)
    with g3:
        if "Criticidad texto" in datos:
            st.plotly_chart(barras(dict(datos["Criticidad texto"].value_counts()), {},
                                   "Materiales por criticidad"), use_container_width=True)
    with g4:
        if "Rango TAT" in datos:
            orden = ["Sin TAT", "1-30 días", "31-60 días", "61-90 días", "91-120 días", ">120 días"]
            conteo = {o: int((datos["Rango TAT"] == o).sum()) for o in orden}
            conteo = {k_: v for k_, v in conteo.items() if v}
            st.plotly_chart(barras(conteo, {}, "Materiales por rango de TAT"),
                            use_container_width=True)

    # Recurrencia de compra y TAT poco predecible
    g5, g6 = st.columns(2)
    with g5:
        if "Recurrencia" in datos.columns and datos["Recurrencia"].notna().any():
            orden_r = ["Muy recurrente", "Recurrente", "Ocasional", "Baja frecuencia"]
            conteo = {o: int((datos["Recurrencia"] == o).sum()) for o in orden_r}
            conteo = {k_: v for k_, v in conteo.items() if v}
            st.plotly_chart(barras(conteo, {}, "Materiales por recurrencia de compra"),
                            use_container_width=True)
    with g6:
        if "TAT CV%" in datos.columns and datos["TAT CV%"].notna().any():
            st.markdown("##### TAT poco predecible (variabilidad > 100 %)")
            inest = datos[datos["TAT CV%"] > 100]
            cols_i = [c for c in ["Material", "Texto breve de material", "TAT Promedio",
                                  "TAT Min", "TAT Max", "TAT CV%", "Recurrencia"]
                      if c in inest.columns]
            if inest.empty:
                st.success("Ningún material tiene un TAT tan variable.")
            else:
                st.caption(f"{len(inest)} materiales cuyo tiempo de entrega es poco confiable.")
                st.dataframe(inest[cols_i].sort_values("TAT CV%", ascending=False).head(30),
                             use_container_width=True, hide_index=True)

    # Disponibilidad por área
    if "Area" in datos.columns and "Condicion Stock" in datos.columns:
        st.markdown("##### Disponibilidad por área")
        res = (datos.assign(_ok=(datos["Condicion Stock"] != "Quiebre Stock").astype(int))
               .groupby("Area").agg(Materiales=("Material", "count"),
                                    Disponibilidad=("_ok", "mean")).reset_index())
        res["Disponibilidad"] = (100 * res["Disponibilidad"]).round(1)
        st.dataframe(res.sort_values("Disponibilidad"), use_container_width=True, hide_index=True)

# ================================================================ 3) DEMANDA VS STOCK
with t3:
    st.markdown("#### Demanda vs stock: ¿alcanza para la próxima demanda?")
    st.caption(
        "Cruza el stock actual con el **pronóstico de demanda** (panel 1) y el "
        "**tiempo que falta hasta la próxima demanda**. "
        "*Cumple* = el stock cubre la demanda y deja el stock de seguridad; "
        "*Alerta* = alcanza pero se come el stock de seguridad; "
        "*Urgente* = queda justo; *No cumple* = el stock no alcanza."
    )
    if "Cumple_Demanda" not in datos.columns or datos["Cumple_Demanda"].eq("Sin pronóstico").all():
        st.warning(
            "Todavía no hay pronóstico cruzado. Esto ocurre si faltan los datos de "
            "**MB51/MB5B** (panel 1) o si los materiales del MRP no coinciden con "
            "los del histórico de demanda."
        )
    else:
        cA, cB = st.columns([1, 2])
        with cA:
            orden = ["No cumple", "Urgente", "Alerta", "Cumple", "Sin pronóstico"]
            conteo = {o: int((datos["Cumple_Demanda"] == o).sum()) for o in orden}
            conteo = {k_: v for k_, v in conteo.items() if v}
            st.plotly_chart(barras(conteo, COLOR_CUMPLE, "Cobertura de la demanda"),
                            use_container_width=True)
        with cB:
            criticos = datos[datos["Cumple_Demanda"].isin(["No cumple", "Urgente", "Alerta"])]
            st.markdown("##### Materiales que requieren atención")
            cols_d = [c for c in ["Material", "Texto breve de material", "Stock",
                                  "Stock Seguridad", "Pronostico_Consolidado",
                                  "Tiempo_Prox_Demanda", "Cumple_Demanda", "Tipo_demanda",
                                  "Estado gestión", "TAT Promedio"] if c in criticos.columns]
            if criticos.empty:
                st.success("Todos los materiales con pronóstico cubren su próxima demanda.")
            else:
                st.dataframe(criticos[cols_d].sort_values("Cumple_Demanda"),
                             use_container_width=True, hide_index=True)

    st.markdown("---")
    st.markdown("##### Detalle de demanda por material")
    cols_dm = [c for c in ["Material", "Texto breve de material", "Centro", "Tipo_demanda",
                           "Pronostico_Consolidado", "Tiempo_Prox_Demanda", "Stock",
                           "Stock Seguridad", "Cumple_Demanda"] if c in datos.columns]
    st.dataframe(datos[cols_dm].sort_values("Material"), use_container_width=True, hide_index=True)

# ================================================================ 4) TODOS LOS MATERIALES
with t4:
    st.markdown("#### Todos los materiales (tabla completa)")
    st.caption(
        f"**{len(datos)} materiales** — esta tabla contiene TODAS las columnas y "
        "**ningún material se pierde**, tengan o no solped, OC, TAT o pronóstico."
    )

    # Buscador
    datos["etiqueta"] = (datos["Material"].astype(str) + "  —  "
                         + datos.get("Texto breve de material",
                                     pd.Series("", index=datos.index)).fillna(""))
    etiqueta = st.selectbox("Buscar un material", ["(ver todos)"] + sorted(datos["etiqueta"]))
    if etiqueta != "(ver todos)":
        mat = etiqueta.split("  —  ")[0].strip()
        info = datos[datos["Material"] == mat].iloc[0]
        st.markdown(f"### {info.get('Texto breve de material', mat)}")
        cA, cB, cC, cD = st.columns(4)
        with cA:
            st.markdown("**Estado de gestión**")
            st.markdown(_chip(info["Estado gestión"]), unsafe_allow_html=True)
            st.caption(f"Centro {info.get('Centro','')} · {info.get('Criticidad texto','')}")
        with cB:
            st.metric("Condición de stock", str(info.get("Condicion Stock", "—")))
            st.caption(f"Stock: {info.get('Stock','—')} · Seg.: {info.get('Stock Seguridad','—')}")
        with cC:
            ds = info.get("Días en solped")
            st.metric("Solped", "—" if pd.isna(info.get("Solped")) else str(info.get("Solped")))
            st.caption("" if pd.isna(ds) else f"{int(ds)} días en gestión")
        with cD:
            st.metric("OC", "—" if pd.isna(info.get("OC en Transito")) else str(info.get("OC en Transito")))
            da = info.get("Días atraso OC", 0)
            st.caption(f"{info.get('Estado OC','')} · {info.get('Nacionalidad','')}"
                       + (f" · {int(da)} d. atraso" if da and da > 0 else ""))
        cE, cF, cG, cH = st.columns(4)
        cE.metric("TAT promedio", "—" if pd.isna(info.get("TAT Promedio")) else f"{info['TAT Promedio']:.0f} d")
        cF.metric("Pronóstico", "—" if pd.isna(info.get("Pronostico_Consolidado")) else f"{info['Pronostico_Consolidado']:.0f}")
        cG.metric("Cobertura", str(info.get("Cumple_Demanda", "—")))
        cH.metric("Valor stock", "—" if pd.isna(info.get("Valor stock")) else f"{info['Valor stock']:,.0f}".replace(",", "."))

        # Detalle del TAT (tiempo de abastecimiento)
        if pd.notna(info.get("TAT Promedio")):
            st.markdown("**Tiempo de abastecimiento (TAT)**")
            t1_, t2_, t3_, t4_, t5_ = st.columns(5)
            t1_.metric("Mínimo", f"{info['TAT Min']:.0f} d" if pd.notna(info.get("TAT Min")) else "—")
            t2_.metric("Máximo", f"{info['TAT Max']:.0f} d" if pd.notna(info.get("TAT Max")) else "—")
            t3_.metric("Variabilidad", f"{info['TAT CV%']:.0f} %" if pd.notna(info.get("TAT CV%")) else "—")
            t4_.metric("Recurrencia", str(info.get("Recurrencia", "—")))
            t5_.metric("Últ. solicitud", f"{info['Días desde última solicitud']:.0f} d atrás"
                       if pd.notna(info.get("Días desde última solicitud")) else "—")
            st.caption(
                f"Basado en {int(info['TAT Registros'])} compras históricas. "
                "Una variabilidad alta (>100%) significa que el tiempo de entrega "
                "es poco predecible."
                if pd.notna(info.get("TAT Registros")) else ""
            )
        else:
            st.info("Este material no tiene historial de compras, por lo que no tiene TAT calculado.")
        with st.expander("Ver todos los datos de este material"):
            st.dataframe(info.drop(labels=["etiqueta"], errors="ignore").rename("Valor").to_frame(),
                         use_container_width=True)
        st.markdown("---")

    vista = datos.drop(columns=["etiqueta"], errors="ignore").sort_values("Material")
    st.dataframe(vista, use_container_width=True, hide_index=True)
    st.download_button("⬇️  Descargar tabla completa (CSV)",
                       data=vista.to_csv(index=False).encode("utf-8-sig"),
                       file_name="abastecimiento_completo.csv", mime="text/csv")

# ================================================================ 5) ACTUALIZAR
with t5:
    st.markdown("#### Actualizar los datos de este panel")
    st.caption("Cada archivo **reemplaza** al anterior (son fotos que se actualizan).")
    with st.expander("📥 Qué descargar de SAP para cada uno"):
        st.markdown(
            """
| Archivo | Transacción / origen | Cada cuánto |
|---|---|---|
| **MRP semanal** | Planificacion_Simpl (hoja `data`) | Semanal |
| **MM60** | MM60 — maestro de materiales | Mensual o al cambiar |
| **ME5A** | ME5A — solicitudes de pedido (solped) | Semanal |
| **ME2M** | ME2M — órdenes de compra | Semanal |
| **TAT** | Vista Ejecutiva de Materiales (hoja `Dias_TAT`) | Mensual |
            """
        )
    modo_github = github_store.disponible()
    if modo_github:
        st.success("🟢 Se guardarán en GitHub (permanente). La app se actualiza sola en ~1 min.")
    else:
        st.info("🟡 Guardado local (temporal en la nube). Configura los *secrets* para GitHub.")

    autorizado = True
    if github_store.password_configurada():
        clave = st.text_input("🔒 Contraseña", type="password", key="ab_pwd")
        autorizado = github_store.password_ok(clave)
        if not autorizado:
            st.caption("Ingresa la contraseña para habilitar la carga.")

    fuentes = [
        ("MRP semanal (Planificacion_Simpl)", "MRP", config.CARPETA_MRP),
        ("MM60 (maestro de materiales)", "MM60", config.CARPETA_MM60),
        ("ME5A (solicitudes / solped)", "ME5A", config.CARPETA_ME5A),
        ("ME2M (órdenes de compra)", "ME2M", config.CARPETA_ME2M),
        ("TAT (MERGE, tiempos de abastecimiento)", "TAT", config.CARPETA_TAT),
    ]
    if autorizado:
        for etiqueta_f, sub, carpeta in fuentes:
            archivo = st.file_uploader(etiqueta_f, type=["xlsx", "xls"], key=f"up_{sub}")
            if archivo is not None and st.button(f"Guardar {sub}", key=f"btn_{sub}"):
                try:
                    if modo_github:
                        github_store.guardar(sub, archivo.name, archivo.getvalue(), reemplazar=True)
                        st.success(f"{sub} guardado en GitHub: {archivo.name}. La app se actualizará sola.")
                    else:
                        data_loading.guardar_local(carpeta, archivo, reemplazar=True)
                        st.cache_data.clear()
                        st.success(f"{sub} guardado localmente: {archivo.name}.")
                        st.rerun()
                except Exception as e:
                    st.error(f"No se pudo guardar {sub}: {e}")

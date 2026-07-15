"""
Panel 2 - Abastecimiento: estado de los materiales (solped, OC, disponibilidad).

Integra varias descargas de SAP en una sola tabla por Material + Centro:

  Base : MRP semanal (Planificacion_Simpl, hoja 'data')  -> estado y stock
  MM60 : maestro de materiales   (une por Material + Centro)  -> precio, ABC
  ME5A : solicitudes de pedido    (une por Solped + Material)  -> antigüedad solped
  ME2M : órdenes de compra (OC)   (une por OC + Material)      -> atraso / en curso
  TAT  : tiempos de abastecimiento (une por Material, todos los centros)

Las llaves de unión son las indicadas por el negocio:
  - Material + Centro para el maestro y el MRP.
  - Solped + Material para las solicitudes.
  - OC + Material para las órdenes de compra.
  - Material (sin centro) para el TAT, porque se estudia sobre todas las compras.
"""

from __future__ import annotations

from dataclasses import dataclass

import numpy as np
import pandas as pd

from . import config
from .data_loading import (
    _buscar_columna,
    _listar_excels,
    _norm_codigo,
    _parse_fecha,
    _parse_numero,
    _renombrar,
)


# --------------------------------------------------------------------------
# Utilidades de lectura
# --------------------------------------------------------------------------
def _leer_primera_hoja(ruta, hoja=None, header=0):
    archivos = _listar_excels(ruta)
    if not archivos:
        raise FileNotFoundError(f"No se encontraron archivos en: {ruta}")
    partes = []
    for a in archivos:
        df = pd.read_excel(a, sheet_name=hoja if hoja is not None else 0, header=header)
        partes.append(df)
    return pd.concat(partes, ignore_index=True)


def _fila_encabezado(ruta, hoja, palabra="Material", max_filas=40):
    """Encuentra la fila (0-index) que contiene 'palabra' en la hoja indicada."""
    archivos = _listar_excels(ruta)
    if not archivos:
        raise FileNotFoundError(f"No se encontraron archivos en: {ruta}")
    crudo = pd.read_excel(archivos[0], sheet_name=hoja, header=None, nrows=max_filas)
    for i in range(len(crudo)):
        fila = [str(x).strip().lower() for x in crudo.iloc[i].tolist()]
        if palabra.lower() in fila:
            return i
    return 0


# --------------------------------------------------------------------------
# Cargadores por fuente
# --------------------------------------------------------------------------
def cargar_mrp(ruta=None) -> pd.DataFrame:
    """MRP semanal: hoja 'data', encabezados donde aparece 'Material'."""
    ruta = ruta or config.CARPETA_MRP
    archivos = _listar_excels(ruta)
    if not archivos:
        raise FileNotFoundError(f"No se encontró el MRP semanal en: {ruta}")
    hoja = "data"
    fila = _fila_encabezado(ruta, hoja)
    partes = [pd.read_excel(a, sheet_name=hoja, header=fila) for a in archivos]
    df = pd.concat(partes, ignore_index=True)

    mapeo = {
        "Material": ("Material",),
        "Texto breve de material": ("Texto breve de material", "Texto breve"),
        "Centro": ("Centro",),
        "Area": ("Area", "Área"),
        "Criticidad": ("Criticidad",),
        "Stock Seguridad": ("Stock Seguridad",),
        "Cantidad de Compra": ("Cantidad de Compra",),
        "Stock": ("Stock",),
        "UMB": ("UMB", "UN"),
        "Condicion Stock": ("Condicion Stock", "Condición Stock"),
        "Solped": ("Solped",),
        "Cantidad Solped": ("Cantidad Solped", "Cantidad"),
        "OC en Transito": ("OC en Transito", "OC en Tránsito", "OC en T"),
        "Cantidad en Transito": ("Cantidad en Transito", "Cantidad Transito"),
        "Proveedor": ("Proveedor",),
        "Fecha de entrega": ("Fecha de entrega",),
        "Usuario": ("Usuario",),
        "Observación": ("Observación", "Observacion"),
    }
    df = _renombrar(df, mapeo)
    df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]
    df = df.dropna(subset=["Material"])
    df["Material"] = _norm_codigo(df["Material"])
    df["Centro"] = _norm_codigo(df.get("Centro"))
    for c in ("Solped", "OC en Transito"):
        if c in df.columns:
            df[c] = _norm_codigo(df[c])
    for c in ("Stock", "Stock Seguridad", "Cantidad de Compra",
              "Cantidad Solped", "Cantidad en Transito"):
        if c in df.columns:
            df[c] = _parse_numero(df[c])
    if "Fecha de entrega" in df.columns:
        df["Fecha de entrega"] = _parse_fecha(df["Fecha de entrega"])
    if "Criticidad" in df.columns:
        df["Criticidad"] = df["Criticidad"].fillna("").astype(str).str.strip()
    # Si hay varias semanas cargadas, se queda con el último estado por material+centro
    df = df.drop_duplicates(subset=["Material", "Centro"], keep="last")
    return df.reset_index(drop=True)


def cargar_mm60(ruta=None) -> pd.DataFrame:
    ruta = ruta or config.CARPETA_MM60
    df = _leer_primera_hoja(ruta)
    mapeo = {
        "Material": ("Material",),
        "Indicador ABC": ("Indicador ABC", "Indicador de ABC"),
        "Precio": ("Precio",),
        "Grupo de compras": ("Grupo de compras",),
        "Centro": ("Centro",),
        "Tipo de material": ("Tipo de material",),
    }
    df = _renombrar(df, mapeo)
    df["Material"] = _norm_codigo(df["Material"])
    df["Centro"] = _norm_codigo(df.get("Centro"))
    if "Precio" in df.columns:
        df["Precio"] = _parse_numero(df["Precio"])
    cols = [c for c in ["Material", "Centro", "Indicador ABC", "Precio",
                        "Grupo de compras", "Tipo de material"] if c in df.columns]
    return df[cols].drop_duplicates(subset=["Material", "Centro"]).reset_index(drop=True)


def cargar_me5a(ruta=None) -> pd.DataFrame:
    """Solicitudes de pedido (solped). Llave: Solicitud de pedido + Material."""
    ruta = ruta or config.CARPETA_ME5A
    df = _leer_primera_hoja(ruta)
    mapeo = {
        "Material": ("Material",),
        "Solped": ("Solicitud de pedido",),
        "Fecha solicitud": ("Fecha de solicitud",),
        "Status solped": ("Status tratamiento", "Stat.trat.sol.ped."),
        "Liberacion solped": ("Indicador liberación", "Indicador liberacion"),
    }
    df = _renombrar(df, mapeo)
    df["Material"] = _norm_codigo(df["Material"])
    df["Solped"] = _norm_codigo(df.get("Solped"))
    if "Fecha solicitud" in df.columns:
        df["Fecha solicitud"] = _parse_fecha(df["Fecha solicitud"])
    cols = [c for c in ["Material", "Solped", "Fecha solicitud",
                        "Status solped", "Liberacion solped"] if c in df.columns]
    df = df[cols].dropna(subset=["Solped"])
    return df.drop_duplicates(subset=["Solped", "Material"]).reset_index(drop=True)


def cargar_me2m(ruta=None) -> pd.DataFrame:
    """Órdenes de compra (OC) en tránsito. Llave: Documento compras (OC) + Material."""
    ruta = ruta or config.CARPETA_ME2M
    df = _leer_primera_hoja(ruta)
    mapeo = {
        "Material": ("Material",),
        "OC": ("Documento compras",),
        "Fecha OC": ("Fecha documento",),
        "Cantidad OC": ("Cantidad de pedido",),
        "Por entregar": ("Por entregar (cantidad)",),
        "Fecha entrega OC": ("Fecha entrega estad.", "Fecha entrega estadística"),
        "Proveedor OC": ("Proveedor/Centro suministrador", "Proveedor"),
        "Valor OC": ("Valor neto de orden",),
    }
    df = _renombrar(df, mapeo)
    df["Material"] = _norm_codigo(df["Material"])
    df["OC"] = _norm_codigo(df.get("OC"))
    for c in ("Fecha OC", "Fecha entrega OC"):
        if c in df.columns:
            df[c] = _parse_fecha(df[c])
    for c in ("Cantidad OC", "Por entregar", "Valor OC"):
        if c in df.columns:
            df[c] = _parse_numero(df[c])
    cols = [c for c in ["Material", "OC", "Fecha OC", "Cantidad OC", "Por entregar",
                        "Fecha entrega OC", "Proveedor OC", "Valor OC"] if c in df.columns]
    df = df[cols].dropna(subset=["OC"])
    return df.drop_duplicates(subset=["OC", "Material"]).reset_index(drop=True)


def cargar_tat(ruta=None) -> pd.DataFrame:
    """
    Estudio TAT (tiempo de abastecimiento) por material.

    Fuente preferida: hoja **'Dias_TAT'** de la Vista Ejecutiva de Materiales
    (archivo '..VISTA_EJECUTIVA_MATERIALES_….xlsx'), que trae una fila por
    material con media, mínimo, máximo, desviación, coeficiente de variación,
    recurrencia y días desde la última solicitud.

    Si no existe esa hoja, cae de vuelta a la hoja 'Analisis por mateial' del
    archivo MERGE antiguo (compatibilidad hacia atrás).

    Llave: Material (sin centro: el TAT se estudia sobre todas las compras).
    """
    ruta = ruta or config.CARPETA_TAT
    archivos = _listar_excels(ruta)
    if not archivos:
        raise FileNotFoundError(f"No se encontró el TAT en: {ruta}")

    xls = pd.ExcelFile(archivos[0])
    hojas = {h.lower(): h for h in xls.sheet_names}

    # --- Formato nuevo: hoja Dias_TAT (vista ejecutiva) ---
    hoja_dias = next((real for low, real in hojas.items() if "dias_tat" in low.replace(" ", "")), None)
    if hoja_dias is not None:
        df = pd.read_excel(archivos[0], sheet_name=hoja_dias)
        mapeo = {
            "Material": ("Material",),
            "TAT Promedio": ("Media TAT", "TAT Promedio"),
            "TAT Min": ("Min TAT", "TAT Min"),
            "TAT Max": ("Máximo TAT", "Maximo TAT"),
            "TAT Std": ("Desviación estándar TAT", "Desviacion estandar TAT", "TAT Std"),
            "TAT CV%": ("Coeficiente variación % TAT", "Coeficiente variacion % TAT"),
            "TAT Registros": ("Registros",),
            "Recurrencia": ("Recurrencia",),
            "Grupo compra principal": ("Grupo compra principal",),
            "Días desde última solicitud": ("Días desde última solicitud",
                                            "Dias desde ultima solicitud"),
        }
        df = _renombrar(df, mapeo)
        df["Material"] = _norm_codigo(df["Material"])
        for c in ("TAT Promedio", "TAT Min", "TAT Max", "TAT Std", "TAT CV%",
                  "TAT Registros", "Días desde última solicitud"):
            if c in df.columns:
                df[c] = _parse_numero(df[c])
        cols = [c for c in ["Material", "TAT Promedio", "TAT Min", "TAT Max", "TAT Std",
                            "TAT CV%", "TAT Registros", "Recurrencia",
                            "Grupo compra principal", "Días desde última solicitud"]
                if c in df.columns]
        df = df[cols].dropna(subset=["Material"])
        return df.drop_duplicates(subset=["Material"], keep="first").reset_index(drop=True)

    # --- Formato antiguo: MERGE, hoja 'Analisis por mateial' ---
    hoja = next((h for h in xls.sheet_names if "analisis" in h.lower()), xls.sheet_names[0])
    df = pd.read_excel(archivos[0], sheet_name=hoja)
    mapeo = {
        "Material": ("Material",),
        "TAT Promedio": ("TAT Promedio",),
        "TAT Min": ("TAT Min",),
        "TAT Std": ("TAT Std",),
    }
    df = _renombrar(df, mapeo)
    df["Material"] = _norm_codigo(df["Material"])
    for c in ("TAT Promedio", "TAT Min", "TAT Std"):
        if c in df.columns:
            df[c] = _parse_numero(df[c])
    cols = [c for c in ["Material", "TAT Promedio", "TAT Min", "TAT Std"] if c in df.columns]
    return df[cols].drop_duplicates(subset=["Material"]).reset_index(drop=True)


# --------------------------------------------------------------------------
# Derivaciones de estado
# --------------------------------------------------------------------------
def _tiene_valor(serie: pd.Series) -> pd.Series:
    """Máscara booleana: True donde hay un código real (no nulo/vacío/'none'/'nan')."""
    s = serie.astype(str).str.strip().str.lower()
    return serie.notna() & ~s.isin(["", "nan", "none"])


def estado_archivos() -> list[dict]:
    """
    Revisa qué Excel están cargados y cuáles faltan, para avisar en el panel.
    Devuelve una lista de dicts: nombre, carpeta, cargado, archivo, para_que.
    """
    fuentes = [
        ("MRP semanal", config.CARPETA_MRP, True,
         "Base del panel: materiales, stock, condición, solped y OC."),
        ("MM60", config.CARPETA_MM60, False,
         "Precio, indicador ABC y grupo de compras."),
        ("ME5A", config.CARPETA_ME5A, False,
         "Fecha de la solped -> días de gestión de la solicitud."),
        ("ME2M", config.CARPETA_ME2M, False,
         "Fecha y entrega de la OC -> días de OC y atrasos."),
        ("TAT (Vista Ejecutiva)", config.CARPETA_TAT, False,
         "Tiempo de abastecimiento por material (hoja Dias_TAT)."),
    ]
    estado = []
    for nombre, carpeta, obligatorio, para_que in fuentes:
        archivos = _listar_excels(carpeta)
        estado.append({
            "nombre": nombre,
            "carpeta": str(carpeta).replace("\\", "/").split("data/")[-1],
            "obligatorio": obligatorio,
            "cargado": len(archivos) > 0,
            "archivos": [a.name for a in archivos],
            "para_que": para_que,
        })
    return estado


def _nacionalidad_oc(oc) -> str:
    """OC 45 -> Nacional, 47 -> Internacional, 35 -> Nacional (Ariba), otro -> Otro."""
    s = str(oc or "").strip()
    if s in ("", "nan", "None"):
        return "Sin OC"
    pref = s[:2]
    if pref in ("45", "35"):
        return "Nacional"
    if pref == "47":
        return "Internacional"
    return "Otro"


def _criticidad_texto(c) -> str:
    c = str(c or "").strip().upper()
    if c == "A":
        return "Alta"
    if c == "C":
        return "Baja"
    if c == "B":
        return "Media"
    return "Sin criticidad"


def _rango_dias_solped(d):
    if pd.isna(d):
        return None
    d = int(d)
    if d <= 10:
        return "0-10 días"
    if d <= 20:
        return "11-20 días"
    if d <= 30:
        return "21-30 días"
    return "31+ días"


def _rango_atraso(d):
    if pd.isna(d) or d <= 0:
        return None
    d = int(d)
    if d <= 15:
        return "1-15 días"
    if d <= 30:
        return "16-30 días"
    if d <= 45:
        return "31-45 días"
    if d <= 60:
        return "46-60 días"
    if d <= 75:
        return "61-75 días"
    return ">75 días"


def _rango_tat(t):
    if pd.isna(t):
        return "Sin TAT"
    if t <= 0:
        return "Sin TAT"
    if t <= 30:
        return "1-30 días"
    if t <= 60:
        return "31-60 días"
    if t <= 90:
        return "61-90 días"
    if t <= 120:
        return "91-120 días"
    return ">120 días"


def _cumple_demanda(stock, demanda, stock_seg):
    """¿El stock actual cubre la próxima demanda (+ stock de seguridad)?"""
    if pd.isna(demanda):
        return "Sin pronóstico"
    if pd.isna(stock):
        return "Sin stock dato"
    seg = 0 if pd.isna(stock_seg) else stock_seg
    if stock < demanda:
        return "No cumple"
    if stock == demanda:
        return "Urgente"
    if (stock - demanda) < seg:
        return "Alerta"
    return "Cumple"


def _estado_gestion(row) -> str:
    """Estado de gestión del material, según solped/OC y la observación."""
    obs = str(row.get("Observación") or "").lower()
    if "bloque" in obs:
        return "Solped bloqueada"
    if "validac" in obs:
        return "Validación"
    if str(row.get("OC en Transito") or "").strip() not in ("", "None", "nan"):
        return "Con OC"
    if str(row.get("Solped") or "").strip() not in ("", "None", "nan"):
        return "Con Solped"
    return "Sin gestión"


# --------------------------------------------------------------------------
# Integración
# --------------------------------------------------------------------------
@dataclass
class ResultadoAbastecimiento:
    tabla: pd.DataFrame
    kpis: dict


def construir_abastecimiento(
    mrp=None, mm60=None, me5a=None, me2m=None, tat=None,
    hoy: pd.Timestamp | None = None,
) -> ResultadoAbastecimiento:
    hoy = (hoy or pd.Timestamp.today()).normalize()

    base = cargar_mrp(mrp)

    # --- MM60 por Material + Centro ---
    try:
        m60 = cargar_mm60(mm60)
        base = base.merge(m60, on=["Material", "Centro"], how="left")
    except FileNotFoundError:
        pass

    # --- ME5A por Solped + Material ---
    try:
        s = cargar_me5a(me5a)
        base = base.merge(s, on=["Solped", "Material"], how="left")
    except FileNotFoundError:
        pass

    # --- ME2M por OC + Material ---
    try:
        oc = cargar_me2m(me2m).rename(columns={"OC": "OC en Transito"})
        base = base.merge(oc, on=["OC en Transito", "Material"], how="left")
    except FileNotFoundError:
        pass

    # --- TAT por Material (todos los centros) ---
    try:
        t = cargar_tat(tat)
        base = base.merge(t, on="Material", how="left")
    except FileNotFoundError:
        pass

    # --- Derivaciones ---
    base["Material_Centro"] = base["Material"].astype(str) + "_" + base["Centro"].astype(str)
    base["Estado gestión"] = base.apply(_estado_gestion, axis=1)
    base["Criticidad texto"] = base["Criticidad"].apply(_criticidad_texto) if "Criticidad" in base.columns else "Sin criticidad"
    base["Nacionalidad"] = base["OC en Transito"].apply(_nacionalidad_oc)

    # Disponibilidad conservadora: disponible solo si Stock OK o Sobre Stock
    if "Condicion Stock" in base.columns:
        base["Disponible conservador"] = base["Condicion Stock"].isin(["Stock OK", "Sobre Stock"]).astype(int)

    # Antigüedades (días de gestión) — NO se pierden materiales: quedan en blanco si no aplica
    if "Fecha solicitud" in base.columns:
        base["Días en solped"] = (hoy - base["Fecha solicitud"]).dt.days
    if "Fecha OC" in base.columns:
        base["Días de OC"] = (hoy - base["Fecha OC"]).dt.days

    # Estado de la OC: atrasada / en curso
    fecha_entrega = base["Fecha entrega OC"] if "Fecha entrega OC" in base.columns else base.get("Fecha de entrega")
    tiene_oc = _tiene_valor(base["OC en Transito"])
    if fecha_entrega is not None:
        atrasada = tiene_oc & fecha_entrega.notna() & (fecha_entrega < hoy)
        base["Estado OC"] = np.where(~tiene_oc, "Sin OC",
                             np.where(atrasada, "Atrasada", "En curso"))
        base["Días atraso OC"] = np.where(atrasada, (hoy - fecha_entrega).dt.days, 0)
        base["Días hasta llegada"] = np.where(tiene_oc & fecha_entrega.notna(),
                                              (fecha_entrega - hoy).dt.days, np.nan)
    else:
        base["Estado OC"] = np.where(tiene_oc, "En curso", "Sin OC")
        base["Días atraso OC"] = 0
        base["Días hasta llegada"] = np.nan

    # Rangos (para agrupar en gráficos)
    if "Días en solped" in base.columns:
        base["Rango días solped"] = base["Días en solped"].apply(_rango_dias_solped)
    base["Rango atraso OC"] = base["Días atraso OC"].apply(_rango_atraso)
    if "TAT Promedio" in base.columns:
        base["Rango TAT"] = base["TAT Promedio"].apply(_rango_tat)

    # Valorizaciones
    if "Precio" in base.columns and "Stock" in base.columns:
        base["Valor stock"] = base["Precio"] * base["Stock"]
    if "Precio" in base.columns and "Cantidad en Transito" in base.columns:
        base["Costo en tránsito"] = base["Precio"] * base["Cantidad en Transito"]

    # --- Conexión con la DEMANDA (panel 1): pronóstico, tiempo y Cumple_Demanda ---
    base = _unir_demanda(base)

    # --- KPIs ---
    total = len(base)
    quiebre = int((base["Condicion Stock"] == "Quiebre Stock").sum()) if "Condicion Stock" in base.columns else 0
    kpis = {
        "materiales": total,
        "sin_stock": quiebre,
        "disponibilidad": round(100 * (total - quiebre) / total, 2) if total else 0,
        "oc_atrasadas": int((base["Estado OC"] == "Atrasada").sum()),
        "oc_en_curso": int((base["Estado OC"] == "En curso").sum()),
        "con_solped": int((base["Estado gestión"] == "Con Solped").sum()),
        "con_oc": int((base["Estado gestión"] == "Con OC").sum()),
        "solped_bloqueada": int((base["Estado gestión"] == "Solped bloqueada").sum()),
        "validacion": int((base["Estado gestión"] == "Validación").sum()),
        "nacional": int((base["Nacionalidad"] == "Nacional").sum()),
        "internacional": int((base["Nacionalidad"] == "Internacional").sum()),
    }
    if "Disponible conservador" in base.columns:
        kpis["disponibilidad_conservadora"] = round(100 * base["Disponible conservador"].mean(), 2)
    if "Cumple_Demanda" in base.columns:
        kpis["no_cumple_demanda"] = int(base["Cumple_Demanda"].isin(["No cumple", "Urgente"]).sum())
    # Disponibilidad de críticos (A)
    if "Criticidad" in base.columns:
        crit = base[base["Criticidad"] == "A"]
        if len(crit):
            q = int((crit["Condicion Stock"] == "Quiebre Stock").sum())
            kpis["disponibilidad_A"] = round(100 * (len(crit) - q) / len(crit), 2)

    return ResultadoAbastecimiento(tabla=base.reset_index(drop=True), kpis=kpis)


def _unir_demanda(base: pd.DataFrame) -> pd.DataFrame:
    """
    Une la información de demanda (panel 1) por Material y calcula:
      - Pronostico_Consolidado (según tipo de demanda)
      - Tiempo_Prox_Demanda (meses hasta la próxima demanda)
      - Cumple_Demanda (¿el stock cubre la próxima demanda + seguridad?)
    Si no hay datos de demanda (MB51/MB5B), agrega las columnas en blanco.
    """
    try:
        from .pipeline import construir as construir_demanda
        dem = construir_demanda().resultado.copy()
    except Exception:
        base["Tipo_demanda"] = pd.NA
        base["Pronostico_Consolidado"] = pd.NA
        base["Tiempo_Prox_Demanda"] = pd.NA
        base["Cumple_Demanda"] = "Sin pronóstico"
        return base

    # Pronóstico consolidado según el tipo de demanda
    def _consolidado(r):
        tipo = r.get("Tipo_demanda")
        if tipo in ("Constante", "Errática"):
            return r.get("Pronostico_redondeado")
        if tipo in ("Intermitente", "Irregular"):
            return r.get("PR_Pronostico_redondeado")
        return pd.NA
    dem["Pronostico_Consolidado"] = dem.apply(_consolidado, axis=1)
    # Tiempo hasta la próxima demanda en meses (desde días)
    if "Dias_hasta_demanda" in dem.columns:
        dem["Tiempo_Prox_Demanda"] = pd.to_numeric(dem["Dias_hasta_demanda"], errors="coerce") / 30.0

    cols = ["Material", "Tipo_demanda", "Pronostico_Consolidado", "Tiempo_Prox_Demanda"]
    cols = [c for c in cols if c in dem.columns]
    dem_u = dem[cols].drop_duplicates(subset=["Material"], keep="first")
    base = base.merge(dem_u, on="Material", how="left")

    # Cumple_Demanda: comparar stock actual vs pronóstico (+ stock de seguridad)
    if "Stock" in base.columns and "Pronostico_Consolidado" in base.columns:
        seg = base["Stock Seguridad"] if "Stock Seguridad" in base.columns else pd.Series(0, index=base.index)
        base["Cumple_Demanda"] = [
            _cumple_demanda(s, d, sg)
            for s, d, sg in zip(base["Stock"], base["Pronostico_Consolidado"], seg)
        ]
    return base

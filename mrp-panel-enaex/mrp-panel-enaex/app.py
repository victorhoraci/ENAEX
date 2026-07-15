"""
========================================================================
PANEL MRP · ENAEX S.A.  —  Aplicación completa en un solo archivo
========================================================================

Contiene TODO: carga de datos, cálculos de demanda y pronóstico,
integración de abastecimiento (MRP + MM60 + ME5A + ME2M + TAT) y las
páginas de visualización.

Ejecutar:  streamlit run app.py

Datos: se leen de la carpeta 'data/' que debe estar junto a este archivo:
    data/MB51   data/MB5B   data/MRP   data/MM60   data/ME5A   data/ME2M   data/TAT
También se pueden subir desde la página "Cargar archivos".
"""

from __future__ import annotations

import base64
import io
import math
import re
import types
from dataclasses import dataclass
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

# ==========================================================================
# CONFIGURACIÓN DE LA PÁGINA  (debe ir antes que cualquier otro st.*)
# ==========================================================================
st.set_page_config(
    page_title="MRP · Enaex",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
)

RAIZ_PROYECTO = Path(__file__).resolve().parent

# ========================================================================
#  CONFIGURACIÓN DEL MODELO
# ========================================================================
# --------------------------------------------------------------------------
# RUTAS DE DATOS
# --------------------------------------------------------------------------
# Carpeta raíz del proyecto (…/mrp-panel-demanda)

# Carpetas donde se dejan los Excel descargados de SAP HANA.
# MB51  -> movimientos de material (un archivo, o varios, con histórico).
# MB5B  -> foto de stock por mes (un archivo por mes descargado).
CARPETA_MB51 = RAIZ_PROYECTO / "data" / "MB51"
CARPETA_MB5B = RAIZ_PROYECTO / "data" / "MB5B"

# --- Panel 2: Abastecimiento (estado de materiales) ---
# MRP  -> planificación semanal (Planificacion_Simpl, hoja 'data')
# MM60 -> maestro de materiales (precio, ABC, grupo de compras)
# ME5A -> solicitudes de pedido (solped)
# ME2M -> órdenes de compra (OC) en tránsito
# TAT  -> estudio de tiempos de abastecimiento (MERGE, hoja resumen por material)
CARPETA_MRP = RAIZ_PROYECTO / "data" / "MRP"
CARPETA_MM60 = RAIZ_PROYECTO / "data" / "MM60"
CARPETA_ME5A = RAIZ_PROYECTO / "data" / "ME5A"
CARPETA_ME2M = RAIZ_PROYECTO / "data" / "ME2M"
CARPETA_TAT = RAIZ_PROYECTO / "data" / "TAT"


# --------------------------------------------------------------------------
# PARÁMETROS DEL MODELO DE DEMANDA
# --------------------------------------------------------------------------
# Fecha desde la que se construye la serie mensual completa.
# En el Excel original era 01-01-2023.
FECHA_INICIO = "2023-01-01"

# Factor de suavizamiento exponencial usado por SES, SBA y COMBINADO.
ALFA = 0.3

# Clases de movimiento de MB51 que representan la operación real:
#   101 -> entrada de mercancía (ingreso)
#   201 / 261 -> salidas por consumo (egreso = demanda real)
CLASES_MOVIMIENTO = ["101", "201", "261"]

# En MB51 la cantidad de entrada viene POSITIVA para ingresos (101)
# y NEGATIVA para consumos (201, 261). Con eso separamos Entrada vs Demanda.


# --------------------------------------------------------------------------
# CLASIFICACIÓN DE DEMANDA (metodología ADI / CV²  -  Syntetos & Boylan)
# --------------------------------------------------------------------------
# ADI  = Average Demand Interval  (cada cuánto, en promedio, hay demanda)
# CV²  = Coeficiente de variación al cuadrado del tamaño de la demanda
CORTE_ADI = 1.32   # frontera intermitencia
CORTE_CV2 = 0.49   # frontera variabilidad

# Mapa base: tipo de demanda -> método de pronóstico.
# Nota: para Intermitente / Irregular el método final (SBA o PR) depende de
# cuántas demandas históricas tenga el material (ver MIN_DEMANDAS_PR).
METODO_POR_TIPO = {
    "Constante": "SES",        # antes "Suave"
    "Errática": "COMBINADO",
    "Intermitente": "SBA",     # o "PR" si tiene >= MIN_DEMANDAS_PR demandas
    "Irregular": "SBA",        # o "PR" si tiene >= MIN_DEMANDAS_PR demandas
    "Sin demanda": "Sin cálculo",
}

# Para demanda Intermitente / Irregular:
#   - con MENOS de MIN_DEMANDAS_PR demandas históricas -> método SBA
#     (tiempo hasta la próxima demanda = "Indeterminado")
#   - con MIN_DEMANDAS_PR o más -> método PR (Proceso de Renovación)
#     (tiempo hasta la próxima demanda estimado en días)
MIN_DEMANDAS_PR = 4

# Para expresar en DÍAS los intervalos que el modelo calcula en meses.
DIAS_POR_MES = 30


# --------------------------------------------------------------------------
# PROCESO DE RENOVACIÓN (Renewal Process) - demanda intermitente
# --------------------------------------------------------------------------
MIN_EVENTOS = 3      # mínimo de eventos de demanda para generar pronóstico
Z_95 = 1.96          # valor Z para intervalos de confianza al 95 %
HORIZONTE = 12       # períodos (meses) hacia adelante para demanda acumulada


# --------------------------------------------------------------------------
# APARIENCIA DEL GRÁFICO
# --------------------------------------------------------------------------
COLOR_ENTRADA = "#2E86DE"   # azul  -> ingresos de material (barras)
COLOR_DEMANDA = "#E74C3C"   # rojo  -> egresos / demanda   (barras)
COLOR_STOCK = "#27AE60"     # verde -> stock de cierre      (línea)

# Objeto `config` para que el código de los módulos siga funcionando igual.
config = types.SimpleNamespace(
    RAIZ_PROYECTO=RAIZ_PROYECTO,
    CARPETA_MB51=CARPETA_MB51, CARPETA_MB5B=CARPETA_MB5B,
    CARPETA_MRP=CARPETA_MRP, CARPETA_MM60=CARPETA_MM60,
    CARPETA_ME5A=CARPETA_ME5A, CARPETA_ME2M=CARPETA_ME2M,
    CARPETA_TAT=CARPETA_TAT,
    FECHA_INICIO=FECHA_INICIO, ALFA=ALFA, CLASES_MOVIMIENTO=CLASES_MOVIMIENTO,
    CORTE_ADI=CORTE_ADI, CORTE_CV2=CORTE_CV2, METODO_POR_TIPO=METODO_POR_TIPO,
    MIN_DEMANDAS_PR=MIN_DEMANDAS_PR, DIAS_POR_MES=DIAS_POR_MES,
    MIN_EVENTOS=MIN_EVENTOS, Z_95=Z_95, HORIZONTE=HORIZONTE,
    COLOR_ENTRADA=COLOR_ENTRADA, COLOR_DEMANDA=COLOR_DEMANDA, COLOR_STOCK=COLOR_STOCK,
)



# ========================================================================
#  CARGA DE DATOS (Excel de SAP)
# ========================================================================
# --------------------------------------------------------------------------
# Utilidades internas
# --------------------------------------------------------------------------
def _listar_excels(ruta: str | Path) -> list[Path]:
    """Devuelve la lista de archivos Excel de una ruta (archivo o carpeta)."""
    ruta = Path(ruta)
    if ruta.is_file():
        return [ruta]
    if ruta.is_dir():
        return sorted(
            p for p in ruta.iterdir()
            if p.suffix.lower() in (".xlsx", ".xls") and not p.name.startswith("~$")
        )
    return []


def _buscar_columna(df: pd.DataFrame, *alias: str) -> str | None:
    """Busca la primera columna que coincida con alguno de los alias dados."""
    normal = {str(c).strip().lower(): c for c in df.columns}
    for a in alias:
        clave = a.strip().lower()
        if clave in normal:
            return normal[clave]
    return None


def _renombrar(df: pd.DataFrame, mapeo: dict[str, tuple[str, ...]]) -> pd.DataFrame:
    """Renombra columnas usando {nombre_estandar: (alias1, alias2, ...)}."""
    renombres = {}
    for estandar, alias in mapeo.items():
        real = _buscar_columna(df, *alias)
        if real is not None:
            renombres[real] = estandar
    return df.rename(columns=renombres)


def _norm_codigo(serie: pd.Series) -> pd.Series:
    """
    Normaliza un código (Material, Centro) a texto.
    100001.0 -> '100001'; '  3000 ' -> '3000'; deja alfanumericos tal cual.
    """
    def f(x):
        if pd.isna(x):
            return None
        if isinstance(x, float) and x.is_integer():
            return str(int(x))
        if isinstance(x, int):
            return str(x)
        return str(x).strip()
    return serie.map(f)


def _parse_fecha(serie: pd.Series) -> pd.Series:
    """
    Parsea fechas de forma robusta.
    - Si ya son datetime, las deja igual.
    - Si son texto (export SAP en dd.mm.yyyy), usa dayfirst=True.
    """
    if pd.api.types.is_datetime64_any_dtype(serie):
        return serie
    return pd.to_datetime(serie, errors="coerce", dayfirst=True)


def _parse_numero(serie: pd.Series) -> pd.Series:
    """Convierte a numero aceptando coma decimal (formato es-CL)."""
    if pd.api.types.is_numeric_dtype(serie):
        return serie
    limpia = (
        serie.astype(str)
        .str.replace(".", "", regex=False)
        .str.replace(",", ".", regex=False)
    )
    return pd.to_numeric(limpia, errors="coerce")


# --------------------------------------------------------------------------
# MB51 - Movimientos de material
# --------------------------------------------------------------------------
def cargar_mb51(ruta: str | Path | None = None, archivos=None) -> pd.DataFrame:
    """
    Lee el/los Excel de MB51 y devuelve los movimientos ya filtrados por las
    clases de movimiento relevantes (101, 201, 261).

    Puede leer desde:
      - una carpeta/archivo en disco (parámetro `ruta`), o
      - una lista de archivos subidos en el navegador (parámetro `archivos`,
        objetos tipo file de st.file_uploader).

    Columnas de salida:
        Material (str), Centro (str), Clase de movimiento (str),
        Fecha contabiliz. (datetime), Ctd.en UM entrada (float)
    """
    if archivos:
        fuentes = list(archivos)
    else:
        ruta = ruta or config.CARPETA_MB51
        fuentes = _listar_excels(ruta)
        if not fuentes:
            raise FileNotFoundError(
                f"No se encontraron archivos MB51 en: {ruta}. "
                "Deja el/los Excel de MB51 en esa carpeta o súbelos en el panel."
            )

    mapeo = {
        "Material": ("Material",),
        "Centro": ("Centro",),
        "Clase de movimiento": ("Clase de movimiento", "Clase movimiento", "Clase mov."),
        "Fecha contabiliz.": ("Fecha contabiliz.", "Fecha contabilizacion",
                              "Fecha de contabilizacion", "Fecha contab."),
        "Ctd.en UM entrada": ("Ctd.en UM entrada", "Ctd. en UM entrada",
                             "Cantidad en UM entrada", "Ctd en UM entrada"),
    }

    partes = [_renombrar(pd.read_excel(a, sheet_name=0), mapeo) for a in fuentes]
    df = pd.concat(partes, ignore_index=True)

    faltan = [c for c in mapeo if c not in df.columns]
    if faltan:
        raise KeyError(f"MB51: faltan columnas {faltan}. Columnas leidas: {list(df.columns)}")

    df["Material"] = _norm_codigo(df["Material"])
    df["Centro"] = _norm_codigo(df["Centro"])
    df["Clase de movimiento"] = _norm_codigo(df["Clase de movimiento"])
    df["Fecha contabiliz."] = _parse_fecha(df["Fecha contabiliz."])
    df["Ctd.en UM entrada"] = _parse_numero(df["Ctd.en UM entrada"])

    df = df[df["Clase de movimiento"].isin(config.CLASES_MOVIMIENTO)].copy()
    df = df.dropna(subset=["Material", "Fecha contabiliz.", "Ctd.en UM entrada"])
    return df.reset_index(drop=True)


# --------------------------------------------------------------------------
# MB5B - Stock mensual
# --------------------------------------------------------------------------
def cargar_mb5b(ruta: str | Path | None = None, archivos=None) -> pd.DataFrame:
    """
    Lee todos los Excel de MB5B (uno por mes) y los concatena.

    Puede leer desde una carpeta en disco (`ruta`) o desde archivos subidos en
    el navegador (`archivos`, objetos de st.file_uploader).

    Columnas de salida:
        Material (str), Descripcion del material (str),
        De fecha (datetime), A fecha (datetime),
        Stock inicial, Total ctd.entrada mcia., Total cantidades salida,
        Stock de cierre (float)
    """
    if archivos:
        fuentes = list(archivos)
    else:
        ruta = ruta or config.CARPETA_MB5B
        fuentes = _listar_excels(ruta)
        if not fuentes:
            raise FileNotFoundError(
                f"No se encontraron archivos MB5B en: {ruta}. "
                "Deja los Excel mensuales de MB5B en esa carpeta o súbelos en el panel."
            )

    mapeo = {
        "Material": ("Material",),
        "Descripción del material": ("Descripción del material", "Descripcion del material",
                                     "Texto breve material", "Texto material"),
        "De fecha": ("De fecha", "Desde fecha", "Fecha desde"),
        "A fecha": ("A fecha", "Hasta fecha", "Fecha hasta"),
        "Stock inicial": ("Stock inicial",),
        "Total ctd.entrada mcía.": ("Total ctd.entrada mcía.", "Total ctd. entrada mcía.",
                                    "Total entrada mercancia", "Total ctd.entrada mercancía"),
        "Total cantidades salida": ("Total cantidades salida", "Total cantidad salida",
                                    "Total salidas"),
        "Stock de cierre": ("Stock de cierre", "Stock cierre", "Stock final"),
    }

    partes = []
    for archivo in fuentes:
        df = _renombrar(pd.read_excel(archivo, sheet_name=0), mapeo)
        df["Source.Name"] = getattr(archivo, "name", str(archivo))
        partes.append(df)
    df = pd.concat(partes, ignore_index=True)

    for c in ("Material", "De fecha", "Stock de cierre"):
        if c not in df.columns:
            raise KeyError(f"MB5B: falta la columna '{c}'. Columnas leidas: {list(df.columns)}")

    df["Material"] = _norm_codigo(df["Material"])
    for col in ("De fecha", "A fecha"):
        if col in df.columns:
            df[col] = _parse_fecha(df[col])
    for col in ("Stock inicial", "Total ctd.entrada mcía.",
                "Total cantidades salida", "Stock de cierre"):
        if col in df.columns:
            df[col] = _parse_numero(df[col])
    if "Descripción del material" not in df.columns:
        df["Descripción del material"] = None

    df = df.dropna(subset=["Material", "De fecha"])
    return df.reset_index(drop=True)


# --------------------------------------------------------------------------
# Guardar archivos nuevos (para la sección "Agregar datos" del panel)
# --------------------------------------------------------------------------
def reemplazar_mb51(archivo, carpeta: str | Path | None = None) -> str:
    """
    Guarda un nuevo Excel de MB51 REEMPLAZANDO el/los anteriores (MB51 se
    descarga cada semana y siempre pisa al previo). Devuelve el nombre guardado.
    """
    carpeta = Path(carpeta or config.CARPETA_MB51)
    carpeta.mkdir(parents=True, exist_ok=True)
    # Borrar Excel anteriores (se conserva cualquier .gitkeep)
    for viejo in carpeta.iterdir():
        if viejo.suffix.lower() in (".xlsx", ".xls"):
            viejo.unlink()
    destino = carpeta / getattr(archivo, "name", "MB51.xlsx")
    with open(destino, "wb") as f:
        f.write(archivo.getbuffer() if hasattr(archivo, "getbuffer") else archivo.read())
    return destino.name


def agregar_mb5b(archivo, carpeta: str | Path | None = None) -> str:
    """
    Guarda un nuevo Excel de MB5B SIN borrar los anteriores (cada mes se agrega
    uno más). Si el nombre ya existe, lo sobreescribe. Devuelve el nombre.
    """
    carpeta = Path(carpeta or config.CARPETA_MB5B)
    carpeta.mkdir(parents=True, exist_ok=True)
    destino = carpeta / getattr(archivo, "name", "MB5B.xlsx")
    with open(destino, "wb") as f:
        f.write(archivo.getbuffer() if hasattr(archivo, "getbuffer") else archivo.read())
    return destino.name


def guardar_local(carpeta, archivo, reemplazar: bool = True) -> str:
    """Guardado genérico en una carpeta local (reemplaza o agrega)."""
    carpeta = Path(carpeta)
    carpeta.mkdir(parents=True, exist_ok=True)
    if reemplazar:
        for viejo in carpeta.iterdir():
            if viejo.suffix.lower() in (".xlsx", ".xls"):
                viejo.unlink()
    destino = carpeta / getattr(archivo, "name", "archivo.xlsx")
    with open(destino, "wb") as f:
        f.write(archivo.getvalue() if hasattr(archivo, "getvalue") else archivo.read())
    return destino.name


def archivos_detectados() -> dict:
    """
    Diagnóstico: muestra la ruta real y los archivos Excel que la app ve en
    este momento en las carpetas MB51 y MB5B (lectura en vivo, sin caché).
    Sirve para verificar si está leyendo los datos reales o aún los de ejemplo.
    """
    def _ls(p):
        p = Path(p)
        if not p.exists():
            return str(p), ["(la carpeta no existe)"]
        nombres = [f.name for f in _listar_excels(p)]
        return str(p), (nombres if nombres else ["(vacía)"])

    ruta51, arch51 = _ls(config.CARPETA_MB51)
    ruta5b, arch5b = _ls(config.CARPETA_MB5B)
    return {
        "mb51_ruta": ruta51, "mb51_archivos": arch51,
        "mb5b_ruta": ruta5b, "mb5b_archivos": arch5b,
    }


def estado_mb5b(carpeta: str | Path | None = None) -> dict:
    """
    Revisa qué meses de MB5B hay cargados y si falta el del mes recién cerrado.
    Devuelve un dict con: meses (lista de Timestamps), ultimo, falta (bool),
    mes_faltante (Timestamp | None) y n_archivos.
    """
    carpeta = Path(carpeta or config.CARPETA_MB5B)
    archivos = _listar_excels(carpeta)
    meses: list[pd.Timestamp] = []
    if archivos:
        try:
            df = cargar_mb5b(carpeta)
            meses = sorted(df["De fecha"].dt.to_period("M").dt.to_timestamp().unique())
        except Exception:
            meses = []

    # Mes recién cerrado (el que deberían haber cargado al inicio de este mes)
    mes_cerrado = (pd.Timestamp.today().to_period("M") - 1).to_timestamp()
    falta = bool(meses) and (mes_cerrado not in set(meses))
    if not meses:
        falta = True

    return {
        "meses": meses,
        "ultimo": meses[-1] if meses else None,
        "falta": falta,
        "mes_faltante": mes_cerrado if falta else None,
        "n_archivos": len(archivos),
    }


# ========================================================================
#  TRANSFORMACIONES (serie mensual)
# ========================================================================
def demanda_desagregada(mb51: pd.DataFrame) -> pd.DataFrame:
    """
    A partir de los movimientos MB51, separa cada línea en:
      - Demanda: consumo (cantidad negativa -> se toma su valor absoluto)
      - Entrada: ingreso (cantidad positiva)

    Equivale a %Demanda Realista desagregada.
    """
    df = mb51.copy()
    cantidad = df["Ctd.en UM entrada"]
    df["Demanda"] = cantidad.where(cantidad < 0, 0).abs()
    df["Entrada"] = cantidad.where(cantidad > 0, 0)
    return df


def demanda_real_mes(desagregada: pd.DataFrame, mb5b: pd.DataFrame) -> pd.DataFrame:
    """
    Agrupa la demanda/entrada por Material + Centro + mes y le pega el
    Stock de cierre correspondiente desde MB5B.

    Equivale a %Demanda Real Mes.
    """
    df = desagregada.copy()
    df["Año"] = df["Fecha contabiliz."].dt.year
    df["Mes"] = df["Fecha contabiliz."].dt.month

    agrupado = (
        df.groupby(["Material", "Centro", "Año", "Mes"], as_index=False)
        .agg(**{
            "Demanda Mensual": ("Demanda", "sum"),
            "Entrada Mensual": ("Entrada", "sum"),
        })
    )
    # Primer día del mes -> clave temporal
    agrupado["FechaMes"] = pd.to_datetime(
        dict(year=agrupado["Año"], month=agrupado["Mes"], day=1)
    )

    # Unir stock de cierre desde MB5B por (Material, FechaMes = De fecha)
    stock = _stock_mensual(mb5b)
    agrupado = agrupado.merge(stock, on=["Material", "FechaMes"], how="left")

    return agrupado.sort_values(["Material", "Centro", "FechaMes"]).reset_index(drop=True)


def stock_mensual(mb5b: pd.DataFrame) -> pd.DataFrame:
    """
    Prepara el stock de cierre por (Material, FechaMes) desde MB5B.
    'De fecha' es el primer día del mes de la foto, así que lo normalizamos
    al primer día del mes para que empalme con FechaMes.

    Se calcula a nivel de Material (MB5B es por material, sin centro) y luego
    se adjunta a la serie por (Material, FechaMes) para TODOS los meses.
    """
    stock = mb5b[["Material", "De fecha", "Stock de cierre"]].copy()
    stock["FechaMes"] = stock["De fecha"].dt.to_period("M").dt.to_timestamp()
    stock = (
        stock.groupby(["Material", "FechaMes"], as_index=False)["Stock de cierre"]
        .last()
    )
    return stock


# Alias interno para retrocompatibilidad
_stock_mensual = stock_mensual


def serie_completa(
    real_mes: pd.DataFrame,
    stock: pd.DataFrame | None = None,
    fecha_inicio: str | None = None,
    fecha_fin: str | pd.Timestamp | None = None,
) -> pd.DataFrame:
    """
    Genera una fila por cada combinación Material × Centro × Mes, desde
    FECHA_INICIO hasta el mes actual, rellenando con 0 los meses sin
    movimiento. Adjunta el Stock de cierre (de MB5B) en TODOS los meses
    disponibles, para que la línea del gráfico quede completa.

    Equivale a %SerieCompleta (más el arrastre del stock para el gráfico).

    Parámetros
    ----------
    real_mes : DataFrame de demanda_real_mes (demanda/entrada por mes).
    stock    : DataFrame de stock_mensual(mb5b). Opcional; si es None, el stock
               se toma del que traiga real_mes (solo meses con movimiento).
    """
    fecha_inicio = pd.Timestamp(fecha_inicio or config.FECHA_INICIO)
    if fecha_fin is None:
        fecha_fin = pd.Timestamp.today().to_period("M").to_timestamp()
    else:
        fecha_fin = pd.Timestamp(fecha_fin).to_period("M").to_timestamp()

    meses = pd.date_range(fecha_inicio, fecha_fin, freq="MS")

    # Combinaciones únicas de Material + Centro
    combos = real_mes[["Material", "Centro"]].drop_duplicates()

    # Producto cartesiano combos × meses
    combos = combos.assign(_key=1)
    tabla_meses = pd.DataFrame({"FechaMes": meses, "_key": 1})
    rejilla = combos.merge(tabla_meses, on="_key").drop(columns="_key")

    # Unir demanda / entrada reales
    serie = rejilla.merge(
        real_mes[["Material", "Centro", "FechaMes", "Demanda Mensual", "Entrada Mensual"]],
        on=["Material", "Centro", "FechaMes"],
        how="left",
    )
    serie["Demanda Mensual"] = serie["Demanda Mensual"].fillna(0.0)
    serie["Entrada Mensual"] = serie["Entrada Mensual"].fillna(0.0)

    # Adjuntar stock de cierre por (Material, FechaMes) en todos los meses
    if stock is not None:
        serie = serie.merge(stock, on=["Material", "FechaMes"], how="left")
    elif "Stock de cierre" in real_mes.columns:
        serie = serie.merge(
            real_mes[["Material", "Centro", "FechaMes", "Stock de cierre"]],
            on=["Material", "Centro", "FechaMes"], how="left",
        )
    else:
        serie["Stock de cierre"] = pd.NA

    serie["Año"] = serie["FechaMes"].dt.year
    serie["Mes"] = serie["FechaMes"].dt.month

    serie = serie[[
        "Material", "Centro", "Año", "Mes", "FechaMes",
        "Demanda Mensual", "Entrada Mensual", "Stock de cierre",
    ]]
    return serie.sort_values(["Material", "Centro", "FechaMes"]).reset_index(drop=True)


# ========================================================================
#  CLASIFICACIÓN DE DEMANDA (ADI / CV²)
# ========================================================================
def _clasificar_fila(adi: float | None, cv2: float | None) -> str:
    """Regla de clasificación ADI/CV² (idéntica al Excel)."""
    if adi is None or cv2 is None or pd.isna(adi) or pd.isna(cv2):
        return "Sin demanda"
    if adi < config.CORTE_ADI and cv2 < config.CORTE_CV2:
        return "Constante"
    if adi < config.CORTE_ADI and cv2 >= config.CORTE_CV2:
        return "Errática"
    if adi >= config.CORTE_ADI and cv2 < config.CORTE_CV2:
        return "Intermitente"
    return "Irregular"


def clasificar_demanda(serie: pd.DataFrame) -> pd.DataFrame:
    """
    Recibe la serie completa y devuelve una fila por Material + Centro con su
    clasificación y método recomendado.
    """
    filas = []
    for (material, centro), grupo in serie.groupby(["Material", "Centro"], sort=True):
        demanda = grupo["Demanda Mensual"]
        positivas = demanda[demanda > 0]

        n_meses = len(grupo)
        meses_con_demanda = int((demanda > 0).sum())

        if len(positivas) == 0:
            promedio = np.nan
        else:
            promedio = positivas.mean()

        # Desviación estándar MUESTRAL (ddof=1). Con <=1 dato -> 0 (como en Excel).
        if len(positivas) <= 1:
            desv = 0.0
        else:
            desv = positivas.std(ddof=1)

        adi = (n_meses / meses_con_demanda) if meses_con_demanda > 0 else np.nan
        if pd.isna(promedio) or promedio == 0:
            cv2 = np.nan
        else:
            cv2 = (desv / promedio) ** 2

        tipo = _clasificar_fila(adi, cv2)
        metodo = config.METODO_POR_TIPO.get(tipo, "Sin cálculo")

        # Para Intermitente / Irregular el método depende de cuántas demandas
        # históricas haya: pocas -> SBA; suficientes -> PR (proceso de renovación).
        if tipo in ("Intermitente", "Irregular"):
            if meses_con_demanda >= config.MIN_DEMANDAS_PR:
                metodo = "PR"
            else:
                metodo = "SBA"

        filas.append({
            "Material": material,
            "Centro": centro,
            "N_meses": n_meses,
            "Meses_con_demanda": meses_con_demanda,
            "Promedio_demandas_positivas": promedio,
            "DesvEst_demandas_positivas": desv,
            "CV2": cv2,
            "ADI": adi,
            "Tipo_demanda": tipo,
            "Metodo": metodo,
        })

    resultado = pd.DataFrame(filas)
    return resultado.sort_values(["Material", "Centro"]).reset_index(drop=True)


# ========================================================================
#  PRONÓSTICOS (SES · SBA · COMBINADO · PR)
# ========================================================================
ALFA = config.ALFA


# --------------------------------------------------------------------------
# Base de pronóstico: serie + tipo/método
# --------------------------------------------------------------------------
def base_pronostico(serie: pd.DataFrame, clasificacion: pd.DataFrame) -> pd.DataFrame:
    """Une la serie completa con su clasificación (Tipo_demanda, Metodo)."""
    base = serie.merge(
        clasificacion[["Material", "Centro", "Tipo_demanda", "Metodo"]],
        on=["Material", "Centro"],
        how="left",
    )
    return base.sort_values(["Material", "Centro", "FechaMes"]).reset_index(drop=True)


# --------------------------------------------------------------------------
# Utilidades de suavizamiento
# --------------------------------------------------------------------------
def _ses_serie(demandas: list[float], alfa: float = ALFA) -> list[float]:
    """
    Suavizamiento Exponencial Simple a un paso.
      f[0] = d[0]
      f[t] = alfa * d[t-1] + (1-alfa) * f[t-1]
    """
    n = len(demandas)
    if n == 0:
        return []
    if n == 1:
        return [demandas[0]]
    f = [demandas[0]]
    for t in range(1, n):
        f.append(alfa * demandas[t - 1] + (1 - alfa) * f[t - 1])
    return f


# --------------------------------------------------------------------------
# SES  (demanda Suave)
# --------------------------------------------------------------------------
def pronostico_ses(base: pd.DataFrame) -> pd.DataFrame:
    """Pronóstico SES por material para los materiales con Metodo == 'SES'."""
    datos = base[base["Metodo"] == "SES"].copy()
    partes = []
    for material, grupo in datos.groupby("Material", sort=True):
        g = grupo.sort_values("FechaMes").reset_index(drop=True)
        g["Pronostico_SES"] = _ses_serie(g["Demanda Mensual"].tolist())
        partes.append(g)
    if not partes:
        return datos.assign(Pronostico_SES=pd.Series(dtype=float))
    return pd.concat(partes, ignore_index=True)


# --------------------------------------------------------------------------
# SBA  (demanda Intermitente / Irregular)  -  Croston con corrección SBA
# --------------------------------------------------------------------------
def _sba_serie(demandas: list[float], alfa: float = ALFA) -> list[float | None]:
    """
    Croston-SBA. Devuelve el pronóstico por período.
      z = tamaño suavizado, p = intervalo suavizado
      pronóstico = (1 - alfa/2) * (z / p)
    """
    contador = 1
    z = None
    p = None
    salida: list[float | None] = []
    for d in demandas:
        intervalo = contador
        if d > 0:
            z = d if z is None else alfa * d + (1 - alfa) * z
            p = intervalo if p is None else alfa * intervalo + (1 - alfa) * p
        if z is None or p is None or p == 0:
            pron = None
        else:
            pron = (1 - alfa / 2) * (z / p)
        salida.append(pron)
        contador = 1 if d > 0 else contador + 1
    return salida


def pronostico_sba(base: pd.DataFrame) -> pd.DataFrame:
    """Pronóstico SBA por material para los materiales con Metodo == 'SBA'."""
    datos = base[base["Metodo"] == "SBA"].copy()
    partes = []
    for material, grupo in datos.groupby("Material", sort=True):
        g = grupo.sort_values("FechaMes").reset_index(drop=True)
        g["Pronostico_SBA"] = _sba_serie(g["Demanda Mensual"].tolist())
        partes.append(g)
    if not partes:
        return datos.assign(Pronostico_SBA=pd.Series(dtype=float))
    return pd.concat(partes, ignore_index=True)


# --------------------------------------------------------------------------
# COMBINADO  (demanda Errática)  =  promedio de SES + PM3 + PM6
# --------------------------------------------------------------------------
def _media_movil_previa(demandas: list[float], ventana: int) -> list[float | None]:
    """Media móvil que usa SOLO los 'ventana' períodos anteriores a cada i."""
    salida: list[float | None] = []
    for i in range(len(demandas)):
        if i < ventana:
            salida.append(None)
        else:
            salida.append(sum(demandas[i - ventana:i]) / ventana)
    return salida


def pronostico_combinado(base: pd.DataFrame) -> pd.DataFrame:
    """Pronóstico COMBINADO por material (Metodo == 'COMBINADO')."""
    datos = base[base["Metodo"] == "COMBINADO"].copy()
    partes = []
    for material, grupo in datos.groupby("Material", sort=True):
        g = grupo.sort_values("FechaMes").reset_index(drop=True)
        demandas = g["Demanda Mensual"].tolist()
        ses = _ses_serie(demandas)
        pm3 = _media_movil_previa(demandas, 3)
        pm6 = _media_movil_previa(demandas, 6)
        comb = [
            None if (pm3[i] is None or pm6[i] is None) else (ses[i] + pm3[i] + pm6[i]) / 3
            for i in range(len(demandas))
        ]
        g["Pronostico_SES"] = ses
        g["PM3"] = pm3
        g["PM6"] = pm6
        g["Pronostico_COMBINADO"] = comb
        partes.append(g)
    if not partes:
        return datos.assign(Pronostico_COMBINADO=pd.Series(dtype=float))
    return pd.concat(partes, ignore_index=True)


# --------------------------------------------------------------------------
# PROCESO DE RENOVACIÓN  (Renewal Process)  -  para intermitentes (SBA)
# --------------------------------------------------------------------------
def _proceso_renovacion(demandas: list[float]) -> list[dict]:
    """
    Recorre la serie de demandas y acumula, período a período, las
    estadísticas de intervalos y tamaños. Devuelve una lista de dicts
    (una por período) con todas las métricas del proceso de renovación.
    """
    MIN = config.MIN_EVENTOS
    Z = config.Z_95
    H = config.HORIZONTE

    contador = 1
    n_eventos = 0
    n_intervalos = 0
    sum_int = 0.0
    sum_sq_int = 0.0
    sum_tam = 0.0
    sum_sq_tam = 0.0
    salida = []

    for d in demandas:
        ocurrencia = 1 if d > 0 else 0
        intervalo_actual = contador

        n_ev = n_eventos + 1 if d > 0 else n_eventos
        # el intervalo solo se registra a partir del 2º evento
        n_int = n_intervalos + 1 if (d > 0 and n_eventos >= 1) else n_intervalos

        s_int = sum_int + intervalo_actual if (d > 0 and n_eventos >= 1) else sum_int
        s_sq_int = (sum_sq_int + intervalo_actual ** 2) if (d > 0 and n_eventos >= 1) else sum_sq_int
        s_tam = sum_tam + d if d > 0 else sum_tam
        s_sq_tam = sum_sq_tam + d * d if d > 0 else sum_sq_tam

        # --- estadísticas de intervalos ---
        media_int = s_int / n_int if n_int >= 1 else None
        var_int = ((s_sq_int - s_int * s_int / n_int) / (n_int - 1)) if n_int >= 2 else None
        var_int_safe = _var_safe(var_int)
        de_int = math.sqrt(var_int_safe) if (var_int_safe is not None and var_int_safe > 0) else (0.0 if var_int_safe is not None else None)
        cv_int = (de_int / media_int) if (de_int is not None and media_int not in (None, 0)) else None

        # --- estadísticas de tamaños ---
        media_tam = s_tam / n_ev if n_ev >= 1 else None
        var_tam = ((s_sq_tam - s_tam * s_tam / n_ev) / (n_ev - 1)) if n_ev >= 2 else None
        var_tam_safe = _var_safe(var_tam)
        de_tam = math.sqrt(var_tam_safe) if (var_tam_safe is not None and var_tam_safe > 0) else (0.0 if var_tam_safe is not None else None)
        cv_tam = (de_tam / media_tam) if (de_tam is not None and media_tam not in (None, 0)) else None

        hay_int = n_int >= MIN and media_int is not None and media_int > 0
        hay_tam = n_ev >= MIN and media_tam is not None

        dist_int = _dist_intervalo(cv_int)
        dist_tam = _dist_tamano(cv_tam)

        tasa = (media_tam / media_int) if (hay_int and hay_tam) else None

        if hay_int:
            if d > 0:
                periodos_hasta = round(media_int, 1)
            else:
                periodos_hasta = round(max(media_int - contador, 0), 1)
        else:
            periodos_hasta = None

        tam_esperado = round(media_tam, 2) if hay_tam else None

        ic_inf_int = round(max(media_int - Z * de_int, 1), 1) if (hay_int and de_int is not None) else None
        ic_sup_int = round(media_int + Z * de_int, 1) if (hay_int and de_int is not None) else None
        ic_inf_tam = round(max(media_tam - Z * de_tam, 0), 2) if (hay_tam and de_tam is not None) else None
        ic_sup_tam = round(media_tam + Z * de_tam, 2) if (hay_tam and de_tam is not None) else None

        dem_acum = round(tasa * H, 2) if tasa is not None else None

        salida.append({
            "Ocurrencia": ocurrencia,
            "Intervalo_Obs": intervalo_actual if (d > 0 and n_eventos >= 1) else None,
            "N_Eventos": n_ev,
            "N_Intervalos": n_int,
            "Media_Intervalo": round(media_int, 2) if media_int is not None else None,
            "DE_Intervalo": round(de_int, 2) if de_int is not None else None,
            "CV_Intervalo": round(cv_int, 4) if cv_int is not None else None,
            "Dist_Intervalo": dist_int,
            "Media_Tamano": round(media_tam, 2) if media_tam is not None else None,
            "DE_Tamano": round(de_tam, 2) if de_tam is not None else None,
            "CV_Tamano": round(cv_tam, 4) if cv_tam is not None else None,
            "Dist_Tamano": dist_tam,
            "Periodos_Hasta_Prox": periodos_hasta,
            "Tamano_Esperado": tam_esperado,
            "Pronostico_PR": round(tasa, 4) if tasa is not None else None,
            "IC95_Inf_Intervalo": ic_inf_int,
            "IC95_Sup_Intervalo": ic_sup_int,
            "IC95_Inf_Tamano": ic_inf_tam,
            "IC95_Sup_Tamano": ic_sup_tam,
            "Demanda_Acum_12P": dem_acum,
        })

        # actualizar estado para el siguiente período
        n_eventos = n_ev
        n_intervalos = n_int
        sum_int, sum_sq_int = s_int, s_sq_int
        sum_tam, sum_sq_tam = s_tam, s_sq_tam
        contador = 1 if d > 0 else contador + 1

    return salida


def _var_safe(var):
    """La varianza no puede ser negativa (errores de redondeo)."""
    if var is None:
        return None
    return var if var > 0 else 0.0


def _dist_intervalo(cv):
    if cv is None:
        return None
    if cv < 0.1:
        return "Deterministica"
    if cv < 0.85:
        return "Gamma"
    if cv <= 1.15:
        return "Exponencial"
    return "Log-Normal"


def _dist_tamano(cv):
    if cv is None:
        return None
    if cv <= 0.5:
        return "Poisson"
    if cv <= 1:
        return "Binomial_Negativa"
    return "Log-Normal"


def pronostico_pr(base: pd.DataFrame) -> pd.DataFrame:
    """Proceso de renovación completo (una fila por período) para método PR."""
    datos = base[base["Metodo"] == "PR"].copy()
    partes = []
    for material, grupo in datos.groupby("Material", sort=True):
        g = grupo.sort_values("FechaMes").reset_index(drop=True)
        metricas = pd.DataFrame(_proceso_renovacion(g["Demanda Mensual"].tolist()))
        partes.append(pd.concat([g.reset_index(drop=True), metricas], axis=1))
    if not partes:
        return datos
    return pd.concat(partes, ignore_index=True)


def pronostico_pr_final(base: pd.DataFrame) -> pd.DataFrame:
    """Solo la ÚLTIMA fila del proceso de renovación por material (Metodo == 'PR')."""
    completo = pronostico_pr(base)
    if completo.empty:
        return completo
    ultimo = (
        completo.sort_values(["Material", "Centro", "FechaMes"])
        .groupby(["Material", "Centro"], as_index=False)
        .tail(1)
    )
    return ultimo.reset_index(drop=True)


# --------------------------------------------------------------------------
# CONSOLIDACIÓN  ->  ResultadoFinal (una fila por Material + Centro)
# --------------------------------------------------------------------------
def _dias_hasta_proximo_mes(hoy: pd.Timestamp | None = None) -> int:
    """Días de calendario que faltan desde hoy hasta el 1° del próximo mes."""
    hoy = (hoy or pd.Timestamp.today()).normalize()
    prox = (hoy.to_period("M") + 1).to_timestamp()
    return int((prox - hoy).days)


def resultado_final(
    ses: pd.DataFrame,
    combinado: pd.DataFrame,
    sba: pd.DataFrame,
    pr_final: pd.DataFrame,
    clasificacion: pd.DataFrame,
) -> pd.DataFrame:
    """
    Consolida una fila por Material + Centro con el último pronóstico y el
    'tiempo hasta la próxima demanda' según el método:

      - Constante (SES) / Errática (COMBINADO): pronóstico del próximo mes;
        tiempo = días que faltan hasta el próximo mes.
      - Intermitente / Irregular con pocas demandas (SBA): tiempo = "Indeterminado".
      - Intermitente / Irregular con >= MIN_DEMANDAS_PR (PR): pronóstico = tamaño
        esperado; tiempo = días estimados hasta la próxima demanda.
      - Sin demanda: pronóstico 0.
    """
    def _sel(df, col):
        vacio = pd.DataFrame(columns=["Material", "Centro", "FechaMes",
                                      "Tipo_demanda", "Metodo", "Pronostico"])
        if df.empty or col not in df.columns:
            return vacio
        out = df[["Material", "Centro", "FechaMes", "Tipo_demanda", "Metodo", col]].copy()
        return out.rename(columns={col: "Pronostico"})

    # --- Métodos mensuales: SES, COMBINADO, SBA -> último mes por material ---
    unidos = pd.concat([
        _sel(ses, "Pronostico_SES"),
        _sel(combinado, "Pronostico_COMBINADO"),
        _sel(sba, "Pronostico_SBA"),
    ], ignore_index=True)

    if not unidos.empty:
        mensual = (
            unidos.sort_values(["Material", "Centro", "FechaMes"])
            .groupby(["Material", "Centro"], as_index=False)
            .tail(1).reset_index(drop=True)
        )
    else:
        mensual = unidos

    # --- Método PR: el pronóstico es el tamaño esperado de la próxima demanda ---
    if not pr_final.empty:
        pr_base = pr_final[["Material", "Centro", "FechaMes",
                            "Tipo_demanda", "Metodo", "Media_Tamano"]].copy()
        pr_base = pr_base.rename(columns={"Media_Tamano": "Pronostico"})
    else:
        pr_base = pd.DataFrame(columns=["Material", "Centro", "FechaMes",
                                        "Tipo_demanda", "Metodo", "Pronostico"])

    # --- Materiales sin demanda -> pronóstico 0 ---
    sin_dem = clasificacion[clasificacion["Tipo_demanda"] == "Sin demanda"].copy()
    if not sin_dem.empty:
        sin_dem["FechaMes"] = pd.NaT
        sin_dem["Pronostico"] = 0.0
        sin_dem = sin_dem[["Material", "Centro", "FechaMes",
                           "Tipo_demanda", "Metodo", "Pronostico"]]

    final = pd.concat([mensual, pr_base, sin_dem], ignore_index=True)

    # --- Pegar detalle del proceso de renovación (solo materiales PR) ---
    cols_pr = {
        "Media_Intervalo": "PR_Media_Intervalo",
        "Media_Tamano": "PR_Tamano_Esperado",
        "Periodos_Hasta_Prox": "PR_Periodos_Hasta_Prox",
        "IC95_Inf_Intervalo": "PR_IC95_Inf_Intervalo",
        "IC95_Sup_Intervalo": "PR_IC95_Sup_Intervalo",
        "IC95_Inf_Tamano": "PR_IC95_Inf_Tamano",
        "IC95_Sup_Tamano": "PR_IC95_Sup_Tamano",
    }
    if not pr_final.empty:
        pr = pr_final[["Material", "Centro", *cols_pr.keys()]].rename(columns=cols_pr)
        final = final.merge(pr, on=["Material", "Centro"], how="left")
    else:
        for nuevo in cols_pr.values():
            final[nuevo] = np.nan

    # --- Columnas calculadas ---
    final["MesPronosticado"] = final["FechaMes"].apply(
        lambda f: (f + pd.DateOffset(months=1)) if pd.notna(f) else pd.NaT
    )
    final["Pronostico_redondeado"] = final["Pronostico"].apply(
        lambda x: int(math.ceil(x)) if pd.notna(x) else pd.NA
    )
    final["PR_Pronostico_redondeado"] = final["PR_Tamano_Esperado"].apply(
        lambda x: int(math.ceil(x)) if pd.notna(x) else pd.NA
    )

    # --- Tiempo hasta la próxima demanda (unificado, según el método) ---
    dias_prox_mes = _dias_hasta_proximo_mes()

    def _tiempo(row):
        """Devuelve (texto, dias) del tiempo hasta la próxima demanda."""
        metodo = row["Metodo"]
        if metodo in ("SES", "COMBINADO"):
            # Constante y Errática: la próxima demanda es el próximo mes.
            return f"{dias_prox_mes} días", dias_prox_mes
        if metodo == "SBA":
            # Pocas demandas históricas: no se puede estimar cuándo.
            return "Indeterminado", pd.NA
        if metodo == "PR":
            periodos = row.get("PR_Periodos_Hasta_Prox")
            if pd.isna(periodos):
                return "Indeterminado", pd.NA
            dias = int(round(periodos * config.DIAS_POR_MES))
            return f"{dias} días", dias
        return "Sin demanda", pd.NA

    tiempos = final.apply(_tiempo, axis=1, result_type="expand")
    final["Tiempo_hasta_demanda"] = tiempos[0]
    final["Dias_hasta_demanda"] = tiempos[1]

    return final.sort_values(["Material", "Centro"]).reset_index(drop=True)

# ==========================================================================
#  PIPELINE DE DEMANDA  (orquesta el cálculo del panel 1)
# ==========================================================================
@dataclass
class ResultadoMRP:
    serie: pd.DataFrame
    clasificacion: pd.DataFrame
    resultado: pd.DataFrame
    materiales: pd.DataFrame
    tabla_final: pd.DataFrame


def _catalogo_materiales(mb5b: pd.DataFrame) -> pd.DataFrame:
    """Código + descripción (la más reciente) de cada material, desde MB5B."""
    return mb5b.sort_values("De fecha").groupby("Material", as_index=False).agg(
        **{"Descripción del material": ("Descripción del material", "last")}
    )


def _ordenar_tabla_final(tabla: pd.DataFrame) -> pd.DataFrame:
    orden = [
        "Material", "Descripción del material", "Centro", "Tipo_demanda", "Metodo",
        "Pronostico", "Pronostico_redondeado", "Tiempo_hasta_demanda",
        "Dias_hasta_demanda", "MesPronosticado", "PR_Periodos_Hasta_Prox",
        "PR_Media_Intervalo", "PR_Tamano_Esperado", "PR_Pronostico_redondeado",
        "PR_IC95_Inf_Intervalo", "PR_IC95_Sup_Intervalo",
        "PR_IC95_Inf_Tamano", "PR_IC95_Sup_Tamano",
    ]
    existentes = [c for c in orden if c in tabla.columns]
    resto = [c for c in tabla.columns if c not in existentes]
    return tabla[existentes + resto]


def construir(mb51=None, mb5b=None, mb51_archivos=None, mb5b_archivos=None,
              fecha_fin=None) -> ResultadoMRP:
    """Ejecuta el pipeline de demanda completo y devuelve todas las tablas."""
    df_mb51 = cargar_mb51(mb51, archivos=mb51_archivos)
    df_mb5b = cargar_mb5b(mb5b, archivos=mb5b_archivos)

    desag = demanda_desagregada(df_mb51)
    real_mes = demanda_real_mes(desag, df_mb5b)
    stock = stock_mensual(df_mb5b)
    serie = serie_completa(real_mes, stock=stock, fecha_fin=fecha_fin)

    clasif = clasificar_demanda(serie)

    base = base_pronostico(serie, clasif)
    ses = pronostico_ses(base)
    comb = pronostico_combinado(base)
    sba = pronostico_sba(base)
    pr_f = pronostico_pr_final(base)

    res = resultado_final(ses, comb, sba, pr_f, clasif)
    catalogo = _catalogo_materiales(df_mb5b)
    tabla = _ordenar_tabla_final(res.merge(catalogo, on="Material", how="left"))

    return ResultadoMRP(serie=serie, clasificacion=clasif, resultado=res,
                        materiales=catalogo, tabla_final=tabla)



# ========================================================================
#  ABASTECIMIENTO (MRP + MM60 + ME5A + ME2M + TAT)
# ========================================================================
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
        dem = construir().resultado.copy()
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


# ========================================================================
#  GUARDADO EN GITHUB (opcional, con contraseña)
# ========================================================================
API = "https://api.github.com"


# --------------------------------------------------------------------------
# Acceso seguro a los secrets
# --------------------------------------------------------------------------
def _secret(clave: str, defecto=None):
    """Lee un secret de Streamlit sin reventar si no está configurado."""
    try:
        if clave in st.secrets:
            return st.secrets[clave]
    except Exception:
        pass
    return defecto


def _config():
    """Devuelve la configuración de GitHub, o None si no está completa."""
    token = _secret("GITHUB_TOKEN")
    repo = _secret("GITHUB_REPO")
    if not token or not repo:
        return None
    return {
        "token": token,
        "repo": repo,
        "branch": _secret("GITHUB_BRANCH", "main"),
        "prefix": str(_secret("GITHUB_DATA_PREFIX", "data")).strip("/"),
    }


def gh_disponible() -> bool:
    """True si la app puede guardar en GitHub (hay token y repo configurados)."""
    return _config() is not None


# --------------------------------------------------------------------------
# Contraseña
# --------------------------------------------------------------------------
def gh_password_configurada() -> bool:
    return _secret("APP_PASSWORD") is not None


def gh_password_ok(clave: str) -> bool:
    """Compara la clave ingresada con la configurada. Si no hay clave configurada, permite."""
    real = _secret("APP_PASSWORD")
    if real is None:
        return True
    return bool(clave) and clave == real


# --------------------------------------------------------------------------
# Llamadas a la API de GitHub
# --------------------------------------------------------------------------
def _headers(token: str) -> dict:
    return {
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
    }


def _listar_dir(cfg: dict, carpeta: str) -> list[dict]:
    """Lista los archivos de una carpeta del repo (name, path, sha)."""
    url = f"{API}/repos/{cfg['repo']}/contents/{carpeta}"
    r = requests.get(url, headers=_headers(cfg["token"]),
                     params={"ref": cfg["branch"]}, timeout=30)
    if r.status_code == 200 and isinstance(r.json(), list):
        return r.json()
    return []


def _sha_de(cfg: dict, ruta: str) -> str | None:
    """Devuelve el sha del archivo si ya existe (necesario para actualizarlo)."""
    url = f"{API}/repos/{cfg['repo']}/contents/{ruta}"
    r = requests.get(url, headers=_headers(cfg["token"]),
                     params={"ref": cfg["branch"]}, timeout=30)
    if r.status_code == 200:
        return r.json().get("sha")
    return None


def _put(cfg: dict, ruta: str, contenido: bytes, mensaje: str):
    """Crea o actualiza un archivo en el repo."""
    url = f"{API}/repos/{cfg['repo']}/contents/{ruta}"
    data = {
        "message": mensaje,
        "content": base64.b64encode(contenido).decode("ascii"),
        "branch": cfg["branch"],
    }
    sha = _sha_de(cfg, ruta)
    if sha:
        data["sha"] = sha
    r = requests.put(url, headers=_headers(cfg["token"]), json=data, timeout=120)
    r.raise_for_status()


def _delete(cfg: dict, ruta: str, sha: str, mensaje: str):
    """Borra un archivo del repo."""
    url = f"{API}/repos/{cfg['repo']}/contents/{ruta}"
    data = {"message": mensaje, "sha": sha, "branch": cfg["branch"]}
    r = requests.delete(url, headers=_headers(cfg["token"]), json=data, timeout=120)
    r.raise_for_status()


# --------------------------------------------------------------------------
# Operaciones de alto nivel (usadas por el panel)
# --------------------------------------------------------------------------
def gh_guardar(subcarpeta: str, nombre: str, contenido: bytes, reemplazar: bool = True) -> str:
    """
    Guardado genérico en cualquier subcarpeta de datos del repo.
    reemplazar=True borra los Excel previos de esa subcarpeta (snapshots que se
    reemplazan, como MRP/MM60/ME5A/ME2M/TAT). reemplazar=False solo agrega.
    """
    cfg = _config()
    if cfg is None:
        raise RuntimeError("GitHub no está configurado (falta GITHUB_TOKEN o GITHUB_REPO).")
    carpeta = f"{cfg['prefix']}/{subcarpeta}"
    if reemplazar:
        for item in _listar_dir(cfg, carpeta):
            if item["name"].lower().endswith((".xlsx", ".xls")):
                _delete(cfg, item["path"], item["sha"], f"Reemplazar {subcarpeta}: borrar {item['name']}")
    _put(cfg, f"{carpeta}/{nombre}", contenido, f"Actualizar {subcarpeta}: {nombre}")
    return nombre


def gh_guardar_mb51(nombre: str, contenido: bytes) -> str:
    """
    MB51 REEMPLAZA: borra los Excel que haya en la carpeta MB51 del repo y
    sube el nuevo. Devuelve el nombre guardado.
    """
    cfg = _config()
    if cfg is None:
        raise RuntimeError("GitHub no está configurado (falta GITHUB_TOKEN o GITHUB_REPO).")
    carpeta = f"{cfg['prefix']}/MB51"
    for item in _listar_dir(cfg, carpeta):
        if item["name"].lower().endswith((".xlsx", ".xls")):
            _delete(cfg, item["path"], item["sha"], f"Reemplazar MB51: borrar {item['name']}")
    _put(cfg, f"{carpeta}/{nombre}", contenido, f"Actualizar MB51: {nombre}")
    return nombre


def gh_agregar_mb5b(nombre: str, contenido: bytes) -> str:
    """
    MB5B SE AGREGA: sube el archivo del mes sin borrar los anteriores (si el
    nombre ya existe, lo actualiza). Devuelve el nombre guardado.
    """
    cfg = _config()
    if cfg is None:
        raise RuntimeError("GitHub no está configurado (falta GITHUB_TOKEN o GITHUB_REPO).")
    carpeta = f"{cfg['prefix']}/MB5B"
    _put(cfg, f"{carpeta}/{nombre}", contenido, f"Agregar MB5B: {nombre}")
    return nombre


# ========================================================================
#  INTERFAZ (páginas y navegación)
# ========================================================================


# ==========================================================================
#  ESTILO VISUAL
# ==========================================================================
st.markdown(
    """
    <style>
      :root { --tinta:#1B2A3A; --acero:#14618A; --niebla:#6B7A8D; --linea:#E2E6EB; }
      .stApp { background:#F5F6F8; }
      .hdr { color:#fff; padding:20px 26px; border-radius:14px; margin-bottom:16px; }
      .hdr-azul   { background:linear-gradient(100deg,#1B2A3A 0%,#14618A 100%); }
      .hdr-ambar  { background:linear-gradient(100deg,#1B2A3A 0%,#B9770E 100%); }
      .hdr h1 { font-size:1.5rem; margin:0; font-weight:700; }
      .hdr p  { margin:4px 0 0; opacity:.85; font-size:.9rem; }
      .metric-card {
        background:#fff; border:1px solid var(--linea); border-radius:12px;
        padding:14px 16px; height:100%; border-left:4px solid var(--acero);
      }
      .metric-card .lbl { font-size:.72rem; text-transform:uppercase;
        letter-spacing:.6px; color:var(--niebla); margin-bottom:4px; }
      .metric-card .val { font-size:1.3rem; font-weight:700; color:var(--tinta); }
      .metric-card .sub { font-size:.78rem; color:var(--niebla); }
      .mono { font-family:"SFMono-Regular","Consolas",monospace; }
      .chip { display:inline-block; padding:3px 12px; border-radius:999px;
              font-size:.82rem; font-weight:600; }
      .chip-constante {background:#E3F2FD;color:#1565C0;}
      .chip-erratica {background:#FFF3E0;color:#E65100;}
      .chip-intermitente {background:#F3E5F5;color:#6A1B9A;}
      .chip-irregular {background:#FBE9E7;color:#C62828;}
      .chip-sin {background:#ECEFF1;color:#546E7A;}
      .e-oc{background:#E3F2FD;color:#1565C0;} .e-solped{background:#E8F5E9;color:#2E7D32;}
      .e-bloq{background:#FBE9E7;color:#C62828;} .e-valid{background:#FFF3E0;color:#E65100;}
      .e-sin{background:#ECEFF1;color:#546E7A;}
      h2,h3 { color:var(--tinta); }
    </style>
    """,
    unsafe_allow_html=True,
)

COLOR_COND = {"Sobre Stock": "#2E86DE", "Stock OK": "#27AE60",
              "Bajo Stock": "#F39C12", "Quiebre Stock": "#E74C3C"}
COLOR_GEST = {"Con OC": "#1565C0", "Con Solped": "#2E7D32", "Solped bloqueada": "#C62828",
              "Validación": "#E65100", "Sin gestión": "#90A4AE"}
COLOR_CUMPLE = {"Cumple": "#27AE60", "Alerta": "#F39C12", "Urgente": "#E67E22",
                "No cumple": "#E74C3C", "Sin pronóstico": "#B0BEC5",
                "Sin stock dato": "#B0BEC5"}
COLOR_NAC = {"Nacional": "#2E86DE", "Internacional": "#8E44AD",
             "Otro": "#95A5A6", "Sin OC": "#CFD8DC"}


# ==========================================================================
#  CARGA CACHEADA
# ==========================================================================
@st.cache_data(show_spinner="Calculando demanda y pronósticos…")
def cargar_demanda():
    r = construir()
    return r.serie, r.clasificacion, r.resultado, r.tabla_final


@st.cache_data(show_spinner="Integrando MRP + MM60 + ME5A + ME2M + TAT…")
def cargar_abastecimiento():
    r = construir_abastecimiento()
    return r.tabla, r.kpis


def _card(lbl, val, sub=""):
    return (f'<div class="metric-card"><div class="lbl">{lbl}</div>'
            f'<div class="val">{val}</div><div class="sub">{sub}</div></div>')


def _chip_demanda(tipo):
    clase = {"Constante": "chip-constante", "Errática": "chip-erratica",
             "Intermitente": "chip-intermitente", "Irregular": "chip-irregular",
             "Sin demanda": "chip-sin"}.get(tipo, "chip-sin")
    return f'<span class="chip {clase}">{tipo}</span>'


def _chip_gestion(estado):
    clase = {"Con OC": "e-oc", "Con Solped": "e-solped", "Solped bloqueada": "e-bloq",
             "Validación": "e-valid", "Sin gestión": "e-sin"}.get(estado, "e-sin")
    return f'<span class="chip {clase}">{estado}</span>'


def barras(conteo, colores, titulo, horizontal=False):
    conteo = {k: v for k, v in conteo.items() if pd.notna(k)}
    if horizontal:
        fig = go.Figure(go.Bar(y=list(conteo.keys()), x=list(conteo.values()),
                               orientation="h",
                               marker_color=[colores.get(k, "#607D8B") for k in conteo],
                               text=list(conteo.values()), textposition="outside"))
    else:
        fig = go.Figure(go.Bar(x=list(conteo.keys()), y=list(conteo.values()),
                               marker_color=[colores.get(k, "#607D8B") for k in conteo],
                               text=list(conteo.values()), textposition="outside"))
    fig.update_layout(title=titulo, height=300, margin=dict(l=10, r=10, t=40, b=10),
                      plot_bgcolor="#fff", paper_bgcolor="#fff",
                      xaxis=dict(showgrid=False),
                      yaxis=dict(showgrid=True, gridcolor="#EEF1F4"))
    return fig


def tabla_estado_archivos(expandido=False):
    """Muestra qué Excel están cargados y cuáles faltan."""
    est = estado_archivos_todos()
    faltan = [e for e in est if not e["cargado"]]
    titulo = ("📁 Archivos cargados — todo listo" if not faltan
              else f"📁 Archivos: faltan {len(faltan)} por subir")
    with st.expander(titulo, expanded=expandido):
        st.caption("Estos son los Excel que alimentan los paneles. "
                   "Se suben en la página **Cargar archivos**.")
        filas = [{
            "Excel": e["nombre"] + (" (obligatorio)" if e["obligatorio"] else ""),
            "Estado": "✅ Cargado" if e["cargado"] else "❌ Falta",
            "Carpeta": f"data/{e['carpeta']}",
            "Archivo": ", ".join(e["archivos"]) if e["archivos"] else "—",
            "Para qué sirve": e["para_que"],
        } for e in est]
        st.dataframe(pd.DataFrame(filas), use_container_width=True, hide_index=True)
        if faltan:
            st.warning("Sin los archivos que faltan, los paneles igual funcionan, pero "
                       "esas columnas quedarán en blanco.")


def estado_archivos_todos():
    """Estado de TODOS los Excel (demanda + abastecimiento)."""
    fuentes = [
        ("MB51 (movimientos)", CARPETA_MB51, True,
         "Panel Demanda: ingresos y consumos por material."),
        ("MB5B (stock mensual)", CARPETA_MB5B, True,
         "Panel Demanda: stock de cierre de cada mes."),
        ("MRP semanal", CARPETA_MRP, True,
         "Panel MRP E002: materiales, stock, condición, solped y OC."),
        ("MM60", CARPETA_MM60, False, "Precio, indicador ABC y grupo de compras."),
        ("ME5A", CARPETA_ME5A, False, "Fecha de la solped → días de gestión."),
        ("ME2M", CARPETA_ME2M, False, "Fecha y entrega de la OC → días y atrasos."),
        ("TAT (Vista Ejecutiva)", CARPETA_TAT, False,
         "Tiempo de abastecimiento por material (hoja Dias_TAT)."),
    ]
    est = []
    for nombre, carpeta, obligatorio, para_que in fuentes:
        archivos = _listar_excels(carpeta)
        est.append({"nombre": nombre, "carpeta": Path(carpeta).name,
                    "obligatorio": obligatorio, "cargado": len(archivos) > 0,
                    "archivos": [a.name for a in archivos], "para_que": para_que})
    return est


# ==========================================================================
#  PÁGINA · INICIO
# ==========================================================================
def pagina_inicio():
    st.markdown(
        '<div class="hdr hdr-azul"><h1>📦 MRP · Enaex S.A.</h1>'
        '<p>Planificación de materiales — Demanda, pronóstico y abastecimiento</p></div>',
        unsafe_allow_html=True,
    )
    st.markdown("""
Bienvenido al panel de **planificación de materiales (MRP)** de Enaex.
Usa el menú de la izquierda para moverte entre las visualizaciones:

- **📈 Demanda y Pronóstico** — histórico de cada material (ingresos, egresos y stock),
  su clasificación de demanda y el pronóstico de la próxima demanda.
- **🚚 MRP E002** — estado de los materiales: solped, OC, días de gestión,
  compras nacionales o internacionales, disponibilidad y cobertura de la demanda.
- **📥 Cargar archivos** — subir las descargas de SAP para actualizar los cálculos.
- **📖 Cómo usar** — guía completa y explicación de los métodos.
    """)
    st.markdown("---")
    tabla_estado_archivos(expandido=True)

    # Resumen rápido si hay datos
    try:
        tabla_ab, kpis = cargar_abastecimiento()
        st.markdown("#### Resumen de abastecimiento")
        c = st.columns(5)
        c[0].metric("Materiales", f"{kpis['materiales']:,}".replace(",", "."))
        c[1].metric("Disponibilidad", f"{kpis['disponibilidad']} %")
        c[2].metric("Sin stock", kpis["sin_stock"])
        c[3].metric("OC atrasadas", kpis["oc_atrasadas"])
        c[4].metric("Solped bloqueadas", kpis["solped_bloqueada"])
    except Exception:
        st.info("Sube los archivos en **Cargar archivos** para ver el resumen.")


# ==========================================================================
#  PÁGINA · DEMANDA Y PRONÓSTICO
# ==========================================================================
def grafico_material(serie_mat, info):
    fig = go.Figure()
    fig.add_bar(x=serie_mat["FechaMes"], y=serie_mat["Entrada Mensual"],
                name="Ingreso de material", marker_color=COLOR_ENTRADA,
                hovertemplate="%{x|%b %Y}<br>Ingreso: %{y:.0f}<extra></extra>")
    fig.add_bar(x=serie_mat["FechaMes"], y=serie_mat["Demanda Mensual"],
                name="Egreso por uso (demanda)", marker_color=COLOR_DEMANDA,
                hovertemplate="%{x|%b %Y}<br>Demanda: %{y:.0f}<extra></extra>")
    stock = serie_mat.dropna(subset=["Stock de cierre"])
    if not stock.empty:
        fig.add_trace(go.Scatter(
            x=stock["FechaMes"], y=stock["Stock de cierre"], name="Stock de cierre",
            mode="lines+markers", line=dict(color=COLOR_STOCK, width=2.5),
            marker=dict(size=5), yaxis="y2",
            hovertemplate="%{x|%b %Y}<br>Stock: %{y:.0f}<extra></extra>"))
    if info is not None and pd.notna(info.get("MesPronosticado")):
        indet = str(info.get("Tiempo_hasta_demanda", "")) == "Indeterminado"
        if indet:
            fig.add_trace(go.Scatter(
                x=[info["MesPronosticado"]], y=[info["Pronostico"]],
                name="Pronóstico (fecha indeterminada) *", mode="markers",
                marker=dict(color="#8E44AD", size=16, symbol="asterisk",
                            line=dict(color="#8E44AD", width=2)),
                hovertemplate="Pronóstico: %{y:.1f}<br>(momento indeterminado)<extra></extra>"))
        else:
            fig.add_trace(go.Scatter(
                x=[info["MesPronosticado"]], y=[info["Pronostico"]],
                name="Pronóstico próxima demanda ★", mode="markers",
                marker=dict(color="#F39C12", size=14, symbol="star",
                            line=dict(color="#B9770E", width=1)),
                hovertemplate="%{x|%b %Y}<br>Pronóstico: %{y:.1f}<extra></extra>"))
    fig.update_layout(
        barmode="group", height=460, margin=dict(l=10, r=10, t=30, b=10),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0),
        plot_bgcolor="#fff", paper_bgcolor="#fff", hovermode="x unified",
        xaxis=dict(title="", showgrid=False),
        yaxis=dict(title="Cantidad (ingresos / egresos)", showgrid=True, gridcolor="#EEF1F4"),
        yaxis2=dict(title="Stock de cierre", overlaying="y", side="right", showgrid=False))
    return fig


def pagina_demanda():
    st.markdown(
        '<div class="hdr hdr-azul"><h1>📈 Histórico y Pronóstico de Demanda</h1>'
        '<p>Enaex S.A. — Datos de SAP HANA (MB51 + MB5B)</p></div>',
        unsafe_allow_html=True,
    )
    try:
        serie, clasif, resultado, tabla = cargar_demanda()
    except FileNotFoundError:
        st.error("Faltan los archivos **MB51** y/o **MB5B**, que alimentan este panel.")
        tabla_estado_archivos(expandido=True)
        st.info("Súbelos en la página **Cargar archivos** y vuelve aquí.")
        return
    except Exception as e:
        st.error(f"No se pudieron calcular los pronósticos: {e}")
        return

    with st.sidebar:
        st.markdown("### Filtros")
        centros = sorted(tabla["Centro"].dropna().unique().tolist())
        centro_sel = st.multiselect("Centro", centros, default=centros)
        tipos = ["Constante", "Errática", "Intermitente", "Irregular", "Sin demanda"]
        tipos_pres = [t for t in tipos if t in tabla["Tipo_demanda"].unique()]
        tipo_sel = st.multiselect("Clasificación de demanda", tipos_pres, default=tipos_pres)
        st.caption(f"Datos hasta: **{serie['FechaMes'].max():%b %Y}**")
        st.caption(f"Materiales: **{tabla['Material'].nunique()}**")
        if st.button("🔄 Recargar datos"):
            st.cache_data.clear()
            st.rerun()
        with st.expander("❓ ¿Qué significan los tipos de demanda?"):
            img = RAIZ_PROYECTO / "assets" / "Tipos_de_demanda.png"
            if img.exists():
                st.image(str(img), use_container_width=True)
            else:
                st.caption(
                    "**Constante:** demanda frecuente y estable → SES.\n\n"
                    "**Errática:** frecuente pero muy variable → COMBINADO.\n\n"
                    "**Intermitente:** aparece en pocos meses, tamaño estable.\n\n"
                    "**Irregular:** esporádica y muy variable.\n\n"
                    "Con menos de 4 demandas → SBA (tiempo indeterminado); "
                    "con 4 o más → PR (días hasta la próxima demanda)."
                )

    tabla_f = tabla[tabla["Centro"].isin(centro_sel)
                    & tabla["Tipo_demanda"].isin(tipo_sel)].copy()
    if tabla_f.empty:
        st.warning("Ningún material coincide con los filtros.")
        return

    tabla_f["etiqueta"] = (tabla_f["Material"].astype(str) + "  —  "
                           + tabla_f["Descripción del material"].fillna("(sin descripción)"))
    st.markdown("#### Selecciona un material")
    st.caption("Escribe el código o parte de la descripción para filtrar.")
    etiqueta = st.selectbox("Material", sorted(tabla_f["etiqueta"]),
                            label_visibility="collapsed")
    material = etiqueta.split("  —  ")[0].strip()
    info = tabla_f[tabla_f["Material"] == material].iloc[0]

    st.markdown(f"### {info['Descripción del material'] or material}")
    c1, c2, c3, c4, c5 = st.columns(5)
    with c1:
        st.markdown(_card("Material", f'<span class="mono">{info["Material"]}</span>',
                          f'Centro {info["Centro"]}'), unsafe_allow_html=True)
    with c2:
        st.markdown(_card("Clasificación", _chip_demanda(info["Tipo_demanda"]),
                          f'Método: {info["Metodo"]}'), unsafe_allow_html=True)
    with c3:
        p = info.get("Pronostico_redondeado")
        st.markdown(_card("Pronóstico", "—" if pd.isna(p) else f"{int(p)}",
                          "próxima demanda"), unsafe_allow_html=True)
    with c4:
        st.markdown(_card("Tiempo hasta demanda",
                          str(info.get("Tiempo_hasta_demanda", "—")), "estimado"),
                    unsafe_allow_html=True)
    with c5:
        mi = info.get("PR_Media_Intervalo")
        st.markdown(_card("Intervalo de demanda", "—" if pd.isna(mi) else f"{mi:g} meses",
                          "promedio entre demandas"), unsafe_allow_html=True)

    st.markdown("")
    serie_mat = serie[(serie["Material"] == material)
                      & (serie["Centro"] == info["Centro"])].sort_values("FechaMes")
    st.plotly_chart(grafico_material(serie_mat, info), use_container_width=True)

    with st.expander("Ver todos los datos de este material"):
        st.dataframe(info.drop(labels=["etiqueta"], errors="ignore")
                     .rename("Valor").to_frame(), use_container_width=True)

    st.markdown("---")
    st.markdown("#### Tabla de materiales (según filtros)")
    cols = ["Material", "Descripción del material", "Centro", "Tipo_demanda", "Metodo",
            "Pronostico_redondeado", "Tiempo_hasta_demanda", "MesPronosticado",
            "PR_Periodos_Hasta_Prox", "PR_Media_Intervalo", "PR_Tamano_Esperado",
            "PR_IC95_Inf_Tamano", "PR_IC95_Sup_Tamano"]
    cols = [c for c in cols if c in tabla_f.columns]
    vista = tabla_f[cols].sort_values("Material")
    st.dataframe(vista, use_container_width=True, hide_index=True,
                 column_config={
                     "Descripción del material": "Descripción",
                     "Tipo_demanda": "Tipo de demanda", "Metodo": "Método",
                     "Pronostico_redondeado": st.column_config.NumberColumn("Pronóstico"),
                     "Tiempo_hasta_demanda": "Tiempo hasta demanda",
                     "MesPronosticado": st.column_config.DateColumn("Mes pronosticado",
                                                                    format="MMM YYYY"),
                 })
    st.download_button("⬇️  Descargar tabla (CSV)",
                       data=vista.to_csv(index=False).encode("utf-8-sig"),
                       file_name="mrp_demanda.csv", mime="text/csv")


# ==========================================================================
#  PÁGINA · MRP E002 (abastecimiento)
# ==========================================================================
def pagina_mrp_e002():
    st.markdown(
        '<div class="hdr hdr-ambar"><h1>🚚 MRP E002 · Estado y control de materiales</h1>'
        '<p>Solped, OC, días de gestión, nacionalidad, disponibilidad, TAT y demanda</p></div>',
        unsafe_allow_html=True,
    )
    try:
        tabla, kpis = cargar_abastecimiento()
    except FileNotFoundError:
        st.error("Falta el **MRP semanal**, que es la base de este panel.")
        tabla_estado_archivos(expandido=True)
        st.info("Súbelo en la página **Cargar archivos** y vuelve aquí.")
        return
    except Exception as e:
        st.error(f"No se pudieron integrar los datos: {e}")
        tabla_estado_archivos(expandido=True)
        return

    tabla_estado_archivos()

    with st.sidebar:
        st.markdown("### Filtros")

        def _multi(col, etiqueta):
            if col not in tabla.columns:
                return None
            ops = sorted(str(x) for x in tabla[col].dropna().unique() if str(x).strip())
            return st.multiselect(etiqueta, ops, default=ops)

        f_centro = _multi("Centro", "Centro")
        f_area = _multi("Area", "Área")
        f_crit = _multi("Criticidad texto", "Criticidad")
        f_cond = _multi("Condicion Stock", "Condición de stock")
        f_gest = _multi("Estado gestión", "Estado de gestión")
        f_oc = _multi("Estado OC", "Estado de la OC")
        f_nac = _multi("Nacionalidad", "Nacionalidad de la OC")
        f_rec = _multi("Recurrencia", "Recurrencia de compra (TAT)")
        if st.button("🔄 Recargar datos", key="rec_ab"):
            st.cache_data.clear()
            st.rerun()

    datos = tabla.copy()
    for col, sel in [("Centro", f_centro), ("Area", f_area), ("Criticidad texto", f_crit),
                     ("Condicion Stock", f_cond), ("Estado gestión", f_gest),
                     ("Estado OC", f_oc), ("Nacionalidad", f_nac), ("Recurrencia", f_rec)]:
        if sel is not None and col in datos.columns:
            datos = datos[datos[col].astype(str).isin(sel)]

    if datos.empty:
        st.warning("Ningún material coincide con los filtros.")
        return

    total = len(datos)
    sin_stock = int((datos["Condicion Stock"] == "Quiebre Stock").sum())
    dispo = round(100 * (total - sin_stock) / total, 1) if total else 0
    k = st.columns(7)
    k[0].metric("Materiales", f"{total:,}".replace(",", "."))
    k[1].metric("Disponibilidad", f"{dispo} %")
    k[2].metric("Sin stock", sin_stock)
    k[3].metric("OC atrasadas", int((datos["Estado OC"] == "Atrasada").sum()))
    k[4].metric("OC en curso", int((datos["Estado OC"] == "En curso").sum()))
    k[5].metric("Solped bloqueadas", int((datos["Estado gestión"] == "Solped bloqueada").sum()))
    k[6].metric("En validación", int((datos["Estado gestión"] == "Validación").sum()))

    t1, t2, t3, t4 = st.tabs(["📌  Gestión (solped / OC)", "📊  Resumen",
                              "📦  Demanda vs Stock", "📋  Todos los materiales"])

    # ---------------- Gestión ----------------
    with t1:
        st.markdown("#### Gestión de solped y OC")
        st.caption("Cada bloque es una tabla propia: así **no desaparece ningún material** "
                   "al mostrar los días de gestión.")

        st.markdown("##### 🚨 Prioridad: en validación y solped bloqueada")
        prio = datos[datos["Estado gestión"].isin(["Validación", "Solped bloqueada"])]
        if prio.empty:
            st.success("No hay materiales en validación ni con solped bloqueada.")
        else:
            cols_p = [c for c in ["Material", "Texto breve de material", "Centro", "Area",
                                  "Criticidad texto", "Condicion Stock", "Stock",
                                  "Estado gestión", "Observación", "Solped",
                                  "Días en solped", "Usuario"] if c in prio.columns]
            st.dataframe(prio[cols_p].sort_values(["Estado gestión", "Material"]),
                         use_container_width=True, hide_index=True)
            st.caption(f"{len(prio)} materiales requieren gestión inmediata.")

        st.markdown("---")
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
                orden = ["0-10 días", "11-20 días", "21-30 días", "31+ días"]
                conteo = {o: int((sol["Rango días solped"] == o).sum()) for o in orden}
                st.plotly_chart(barras(conteo, {}, "Antigüedad de las solped", True),
                                use_container_width=True)
        with c2:
            st.markdown("##### 🚚 OC en tránsito (con días y atraso)")
            oc = datos[datos["Estado OC"].isin(["Atrasada", "En curso"])]
            cols_o = [c for c in ["Material", "Texto breve de material", "OC en Transito",
                                  "Nacionalidad", "Estado OC", "Días de OC",
                                  "Días atraso OC", "Días hasta llegada", "Proveedor"]
                      if c in oc.columns]
            if oc.empty:
                st.info("No hay OC en tránsito.")
            else:
                st.dataframe(oc[cols_o].sort_values("Días atraso OC", ascending=False),
                             use_container_width=True, hide_index=True)
                orden = ["1-15 días", "16-30 días", "31-45 días", "46-60 días",
                         "61-75 días", ">75 días"]
                conteo = {o: int((oc["Rango atraso OC"] == o).sum()) for o in orden}
                conteo = {a: b for a, b in conteo.items() if b}
                if conteo:
                    st.plotly_chart(barras(conteo, {}, "Atraso de las OC", True),
                                    use_container_width=True)

        st.markdown("---")
        st.markdown("##### 🌎 OC nacionales vs internacionales")
        cn1, cn2 = st.columns([1, 2])
        with cn1:
            conteo_n = {x: int((datos["Nacionalidad"] == x).sum())
                        for x in ["Nacional", "Internacional", "Otro"]}
            conteo_n = {a: b for a, b in conteo_n.items() if b}
            st.plotly_chart(barras(conteo_n, COLOR_NAC, "OC por nacionalidad"),
                            use_container_width=True)
        with cn2:
            nac = datos[datos["Nacionalidad"].isin(["Nacional", "Internacional", "Otro"])]
            cols_n = [c for c in ["Material", "Texto breve de material", "OC en Transito",
                                  "Nacionalidad", "Estado OC", "Días atraso OC",
                                  "TAT Promedio", "Recurrencia", "Proveedor"]
                      if c in nac.columns]
            if nac.empty:
                st.info("No hay OC para clasificar.")
            else:
                st.dataframe(nac[cols_n].sort_values("Nacionalidad"),
                             use_container_width=True, hide_index=True)

    # ---------------- Resumen ----------------
    with t2:
        g1, g2 = st.columns(2)
        with g1:
            orden = ["Sobre Stock", "Stock OK", "Bajo Stock", "Quiebre Stock"]
            conteo = {o: int((datos["Condicion Stock"] == o).sum()) for o in orden}
            conteo = {a: b for a, b in conteo.items() if b}
            st.plotly_chart(barras(conteo, COLOR_COND, "Materiales por condición de stock"),
                            use_container_width=True)
        with g2:
            st.plotly_chart(barras(dict(datos["Estado gestión"].value_counts()), COLOR_GEST,
                                   "Materiales por estado de gestión"),
                            use_container_width=True)
        g3, g4 = st.columns(2)
        with g3:
            if "Criticidad texto" in datos:
                st.plotly_chart(barras(dict(datos["Criticidad texto"].value_counts()), {},
                                       "Materiales por criticidad"), use_container_width=True)
        with g4:
            if "Rango TAT" in datos:
                orden = ["Sin TAT", "1-30 días", "31-60 días", "61-90 días",
                         "91-120 días", ">120 días"]
                conteo = {o: int((datos["Rango TAT"] == o).sum()) for o in orden}
                conteo = {a: b for a, b in conteo.items() if b}
                st.plotly_chart(barras(conteo, {}, "Materiales por rango de TAT"),
                                use_container_width=True)
        g5, g6 = st.columns(2)
        with g5:
            if "Recurrencia" in datos.columns and datos["Recurrencia"].notna().any():
                orden_r = ["Muy recurrente", "Recurrente", "Ocasional", "Baja frecuencia"]
                conteo = {o: int((datos["Recurrencia"] == o).sum()) for o in orden_r}
                conteo = {a: b for a, b in conteo.items() if b}
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
                    st.caption(f"{len(inest)} materiales con tiempo de entrega poco confiable.")
                    st.dataframe(inest[cols_i].sort_values("TAT CV%", ascending=False).head(30),
                                 use_container_width=True, hide_index=True)

        if "Area" in datos.columns:
            st.markdown("##### Disponibilidad por área")
            res = (datos.assign(_ok=(datos["Condicion Stock"] != "Quiebre Stock").astype(int))
                   .groupby("Area").agg(Materiales=("Material", "count"),
                                        Disponibilidad=("_ok", "mean")).reset_index())
            res["Disponibilidad"] = (100 * res["Disponibilidad"]).round(1)
            st.dataframe(res.sort_values("Disponibilidad"),
                         use_container_width=True, hide_index=True)

    # ---------------- Demanda vs Stock ----------------
    with t3:
        st.markdown("#### Demanda vs stock: ¿alcanza para la próxima demanda?")
        st.caption("*Cumple* = el stock cubre la demanda y deja el stock de seguridad · "
                   "*Alerta* = alcanza pero se come el stock de seguridad · "
                   "*Urgente* = queda justo · *No cumple* = no alcanza.")
        if "Cumple_Demanda" not in datos.columns or datos["Cumple_Demanda"].eq("Sin pronóstico").all():
            st.warning("Todavía no hay pronóstico cruzado. Revisa que **MB51 y MB5B** estén "
                       "cargados con datos reales y que sus materiales coincidan con el MRP.")
        else:
            cA, cB = st.columns([1, 2])
            with cA:
                orden = ["No cumple", "Urgente", "Alerta", "Cumple", "Sin pronóstico"]
                conteo = {o: int((datos["Cumple_Demanda"] == o).sum()) for o in orden}
                conteo = {a: b for a, b in conteo.items() if b}
                st.plotly_chart(barras(conteo, COLOR_CUMPLE, "Cobertura de la demanda"),
                                use_container_width=True)
            with cB:
                st.markdown("##### Materiales que requieren atención")
                crit = datos[datos["Cumple_Demanda"].isin(["No cumple", "Urgente", "Alerta"])]
                cols_d = [c for c in ["Material", "Texto breve de material", "Stock",
                                      "Stock Seguridad", "Pronostico_Consolidado",
                                      "Tiempo_Prox_Demanda", "Cumple_Demanda",
                                      "Tipo_demanda", "Estado gestión", "TAT Promedio"]
                          if c in crit.columns]
                if crit.empty:
                    st.success("Todos los materiales con pronóstico cubren su próxima demanda.")
                else:
                    st.dataframe(crit[cols_d].sort_values("Cumple_Demanda"),
                                 use_container_width=True, hide_index=True)
        st.markdown("---")
        st.markdown("##### Detalle de demanda por material")
        cols_dm = [c for c in ["Material", "Texto breve de material", "Centro",
                               "Tipo_demanda", "Pronostico_Consolidado",
                               "Tiempo_Prox_Demanda", "Stock", "Stock Seguridad",
                               "Cumple_Demanda"] if c in datos.columns]
        st.dataframe(datos[cols_dm].sort_values("Material"),
                     use_container_width=True, hide_index=True)

    # ---------------- Todos los materiales ----------------
    with t4:
        st.markdown("#### Todos los materiales")
        st.caption(f"**{len(datos)} materiales** — ningún material se pierde, "
                   "tengan o no solped, OC, TAT o pronóstico.")
        datos["etiqueta"] = (datos["Material"].astype(str) + "  —  "
                             + datos.get("Texto breve de material",
                                         pd.Series("", index=datos.index)).fillna(""))
        etiqueta = st.selectbox("Buscar un material",
                                ["(ver todos)"] + sorted(datos["etiqueta"]))
        if etiqueta != "(ver todos)":
            mat = etiqueta.split("  —  ")[0].strip()
            info = datos[datos["Material"] == mat].iloc[0]
            st.markdown(f"### {info.get('Texto breve de material', mat)}")
            cA, cB, cC, cD = st.columns(4)
            with cA:
                st.markdown("**Estado de gestión**")
                st.markdown(_chip_gestion(info["Estado gestión"]), unsafe_allow_html=True)
                st.caption(f"Centro {info.get('Centro','')} · {info.get('Criticidad texto','')}")
            with cB:
                st.metric("Condición de stock", str(info.get("Condicion Stock", "—")))
                st.caption(f"Stock: {info.get('Stock','—')} · Seg.: {info.get('Stock Seguridad','—')}")
            with cC:
                ds = info.get("Días en solped")
                st.metric("Solped", "—" if pd.isna(info.get("Solped")) else str(info.get("Solped")))
                st.caption("" if pd.isna(ds) else f"{int(ds)} días en gestión")
            with cD:
                st.metric("OC", "—" if pd.isna(info.get("OC en Transito"))
                          else str(info.get("OC en Transito")))
                da = info.get("Días atraso OC", 0)
                st.caption(f"{info.get('Estado OC','')} · {info.get('Nacionalidad','')}"
                           + (f" · {int(da)} d. atraso" if da and da > 0 else ""))
            cE, cF, cG, cH = st.columns(4)
            cE.metric("TAT promedio", "—" if pd.isna(info.get("TAT Promedio"))
                      else f"{info['TAT Promedio']:.0f} d")
            cF.metric("Pronóstico", "—" if pd.isna(info.get("Pronostico_Consolidado"))
                      else f"{info['Pronostico_Consolidado']:.0f}")
            cG.metric("Cobertura", str(info.get("Cumple_Demanda", "—")))
            cH.metric("Valor stock", "—" if pd.isna(info.get("Valor stock"))
                      else f"{info['Valor stock']:,.0f}".replace(",", "."))
            if pd.notna(info.get("TAT Promedio")):
                st.markdown("**Tiempo de abastecimiento (TAT)**")
                t_ = st.columns(5)
                t_[0].metric("Mínimo", f"{info['TAT Min']:.0f} d" if pd.notna(info.get("TAT Min")) else "—")
                t_[1].metric("Máximo", f"{info['TAT Max']:.0f} d" if pd.notna(info.get("TAT Max")) else "—")
                t_[2].metric("Variabilidad", f"{info['TAT CV%']:.0f} %" if pd.notna(info.get("TAT CV%")) else "—")
                t_[3].metric("Recurrencia", str(info.get("Recurrencia", "—")))
                t_[4].metric("Últ. solicitud",
                             f"{info['Días desde última solicitud']:.0f} d atrás"
                             if pd.notna(info.get("Días desde última solicitud")) else "—")
            with st.expander("Ver todos los datos de este material"):
                st.dataframe(info.drop(labels=["etiqueta"], errors="ignore")
                             .rename("Valor").to_frame(), use_container_width=True)
            st.markdown("---")

        vista = datos.drop(columns=["etiqueta"], errors="ignore").sort_values("Material")
        st.dataframe(vista, use_container_width=True, hide_index=True)
        st.download_button("⬇️  Descargar tabla completa (CSV)",
                           data=vista.to_csv(index=False).encode("utf-8-sig"),
                           file_name="mrp_e002_materiales.csv", mime="text/csv")


# ==========================================================================
#  PÁGINA · CARGAR ARCHIVOS
# ==========================================================================
def pagina_cargar():
    st.markdown(
        '<div class="hdr hdr-azul"><h1>📥 Cargar archivos</h1>'
        '<p>Sube aquí las descargas de SAP para actualizar los cálculos</p></div>',
        unsafe_allow_html=True,
    )
    tabla_estado_archivos(expandido=True)

    with st.expander("📖 Qué descargar de SAP y con qué layout", expanded=False):
        st.markdown("""
| Excel | Transacción / origen | Layout / hoja | Cada cuánto | Al subirlo |
|---|---|---|---|---|
| **MB51** | MB51 | Layout **`/CALCDEMANDA`** ("MOV. PARA PRONOSTICO DE DEMANDA") | Semanal | **Reemplaza** |
| **MB5B** | MB5B | Columnas: Material · Descripción del material · De fecha · A fecha · Stock inicial · Total ctd.entrada mcía. · Total cantidades salida · Stock de cierre · Unidad medida base · Stock especial | Mensual | **Se agrega** |
| **MRP semanal** | Planificacion_Simpl | Hoja `data` | Semanal | **Reemplaza** |
| **MM60** | MM60 | — | Mensual | **Reemplaza** |
| **ME5A** | ME5A | — | Semanal | **Reemplaza** |
| **ME2M** | ME2M | — | Semanal | **Reemplaza** |
| **TAT** | Vista Ejecutiva de Materiales | Hoja `Dias_TAT` | Mensual | **Reemplaza** |
        """)

    st.warning("📅 **Recordatorio:** al inicio de **cada mes** agrega el nuevo **MB5B** "
               "del mes recién cerrado. El **MB51**, **MRP**, **ME5A** y **ME2M** "
               "conviene actualizarlos cada semana.")

    modo_github = gh_disponible()
    if modo_github:
        st.success("🟢 Los archivos se guardarán **en GitHub** (permanente). "
                   "La app se actualizará sola en ~1 minuto.")
    else:
        st.info("🟡 Guardado **local**: en la nube estos archivos son temporales "
                "(se pierden al reiniciar). Para que queden permanentes, configura los "
                "*secrets* de GitHub (ver página **Cómo usar**) o súbelos al repositorio.")

    autorizado = True
    if gh_password_configurada():
        clave = st.text_input("🔒 Contraseña para cargar archivos", type="password")
        autorizado = gh_password_ok(clave)
        if not autorizado:
            st.caption("Ingresa la contraseña para habilitar la carga.")
    else:
        st.caption("⚠️ No hay contraseña configurada (`APP_PASSWORD` en los secrets).")

    if not autorizado:
        return

    st.markdown("---")
    fuentes = [
        ("MB51 — movimientos (Demanda)", "MB51", CARPETA_MB51, True),
        ("MB5B — stock del mes (Demanda)", "MB5B", CARPETA_MB5B, False),
        ("MRP semanal — Planificacion_Simpl", "MRP", CARPETA_MRP, True),
        ("MM60 — maestro de materiales", "MM60", CARPETA_MM60, True),
        ("ME5A — solicitudes (solped)", "ME5A", CARPETA_ME5A, True),
        ("ME2M — órdenes de compra", "ME2M", CARPETA_ME2M, True),
        ("TAT — Vista Ejecutiva (hoja Dias_TAT)", "TAT", CARPETA_TAT, True),
    ]
    for etiqueta, sub, carpeta, reemplaza in fuentes:
        modo = "reemplaza el anterior" if reemplaza else "**se agrega** a los anteriores"
        st.markdown(f"##### {etiqueta}")
        st.caption(f"Al subirlo, {modo}.")
        archivo = st.file_uploader(etiqueta, type=["xlsx", "xls"], key=f"up_{sub}",
                                   label_visibility="collapsed")
        if archivo is not None and st.button(f"Guardar {sub}", key=f"btn_{sub}"):
            try:
                if modo_github:
                    gh_guardar(sub, archivo.name, archivo.getvalue(), reemplazar=reemplaza)
                    st.success(f"{sub} guardado en GitHub: {archivo.name}. "
                               "La app se actualizará sola en ~1 minuto.")
                else:
                    guardar_local(carpeta, archivo, reemplazar=reemplaza)
                    st.cache_data.clear()
                    st.success(f"{sub} guardado: {archivo.name}.")
                    st.rerun()
            except Exception as e:
                st.error(f"No se pudo guardar {sub}: {e}")
        st.markdown("")


# ==========================================================================
#  PÁGINA · CÓMO USAR
# ==========================================================================
def pagina_ayuda():
    st.markdown(
        '<div class="hdr hdr-azul"><h1>📖 Cómo usar y actualizar</h1>'
        '<p>Guía completa del panel MRP</p></div>',
        unsafe_allow_html=True,
    )
    st.markdown("""
### Las visualizaciones

**📈 Demanda y Pronóstico** — por cada material:
- **Barras azules**: ingresos de material (clase 101).
- **Barras rojas**: egresos por uso = demanda real (clases 201 y 261).
- **Línea verde**: stock de cierre de cada mes (MB5B).
- **★ estrella**: pronóstico de la próxima demanda.
- **✳ asterisco**: pronóstico con momento *indeterminado*.

**🚚 MRP E002** — estado de cada material: solped, OC, días de gestión,
validación, solped bloqueada, nacional/internacional, TAT y cobertura de demanda.

---

### Cómo se elige el método y el "tiempo hasta demanda"

| Tipo de demanda | Método | Tiempo hasta la próxima demanda |
|---|---|---|
| **Constante** | **SES** | Días que faltan hasta el próximo mes |
| **Errática** | **COMBINADO** (SES + medias móviles 3 y 6) | Días que faltan hasta el próximo mes |
| **Intermitente / Irregular** con **menos de 4** demandas | **SBA** | **Indeterminado** |
| **Intermitente / Irregular** con **4 o más** demandas | **PR** (Proceso de Renovación) | **Días estimados** |
| **Sin demanda** | — | — |

La clasificación usa **ADI / CV²** (Syntetos & Boylan): ADI mide cada cuánto hay
demanda; CV² cuánto varía su tamaño. Cortes: ADI = 1,32 · CV² = 0,49.

---

### Cómo se unen las tablas del MRP E002

| Fuente | Se une por |
|---|---|
| MRP semanal | tabla **base** (ningún material se pierde) |
| MM60 | Material + Centro |
| ME5A | Solped + Material |
| ME2M | OC + Material |
| TAT | Material (todos los centros) |
| Demanda | Material |

Todas las uniones son *left join*: un material sin solped, sin OC, sin TAT o sin
pronóstico **igual aparece**, con esos campos en blanco.

---

### Cómo actualizar los datos

Ve a **📥 Cargar archivos**. Ahí está el detalle de cada Excel, su layout de SAP
y cada cuánto se actualiza. El panel recalcula solo.

---

### Guardar en GitHub desde la app (opcional)

Para que los archivos que subes queden **permanentes**, configura los *secrets*
en Streamlit (menú de la app → **Settings → Secrets**):

```toml
APP_PASSWORD       = "una-clave-secreta"
GITHUB_TOKEN       = "github_pat_xxxxxxxx"
GITHUB_REPO        = "usuario/repositorio"
GITHUB_BRANCH      = "main"
GITHUB_DATA_PREFIX = "ruta/a/la/carpeta/data"
```

El token se crea en GitHub → Settings → Developer settings → Personal access
tokens → Fine-grained tokens, con permiso **Contents: Read and write** sobre el
repositorio.
    """)


# ==========================================================================
#  NAVEGACIÓN
# ==========================================================================
def main():
    paginas = {
        "Inicio": [
            st.Page(pagina_inicio, title="Inicio", icon="🏠", default=True),
        ],
        "01 Visualizaciones": [
            st.Page(pagina_demanda, title="Demanda y Pronóstico", icon="📈"),
            st.Page(pagina_mrp_e002, title="MRP E002", icon="🚚"),
        ],
        "02 Datos": [
            st.Page(pagina_cargar, title="Cargar archivos", icon="📥"),
            st.Page(pagina_ayuda, title="Cómo usar", icon="📖"),
        ],
    }
    st.navigation(paginas).run()


if __name__ == "__main__":
    main()


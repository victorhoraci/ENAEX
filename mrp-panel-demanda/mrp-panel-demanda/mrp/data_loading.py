"""
Carga de los Excel descargados de SAP HANA.

Dos fuentes:
  - MB51  -> movimientos de material (histórico de ingresos y consumos).
  - MB5B  -> foto de stock por mes (uno o varios archivos, uno por mes).

Estas funciones son la "traducción" de las dos primeras consultas de
Power Query (%Data y %Demandas). Aceptan tanto un archivo suelto como una
carpeta con varios archivos, y toleran pequeñas variaciones en los nombres
de las columnas que suele producir la exportación de SAP.
"""

from __future__ import annotations

from pathlib import Path

import pandas as pd

from . import config


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

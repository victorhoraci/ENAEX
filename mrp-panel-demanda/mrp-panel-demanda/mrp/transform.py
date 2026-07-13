"""
Transformaciones de datos: de movimientos crudos a serie mensual completa.

Replica estas consultas de Power Query:
  - %Demanda Realista desagregada  -> separar Entrada vs Demanda
  - %Demanda Real Mes              -> agregar por mes y unir stock (MB5B)
  - %SerieCompleta                 -> rejilla mensual completa (rellena con 0)

La "serie completa" es la base de TODO lo demás: clasificación, pronósticos
y el gráfico del panel.
"""

from __future__ import annotations

import pandas as pd

from . import config


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

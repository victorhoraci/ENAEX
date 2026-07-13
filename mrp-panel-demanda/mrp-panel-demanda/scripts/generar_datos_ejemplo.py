"""
Genera archivos MB51 y MB5B de EJEMPLO (datos sintéticos) para poder probar
el panel sin datos reales de SAP.

Crea materiales con distintos patrones de demanda (suave, errática,
intermitente y sin demanda) para ejercitar los cuatro métodos de pronóstico.

Uso:
    python scripts/generar_datos_ejemplo.py

Deja los Excel en data/MB51 y data/MB5B.
"""

from __future__ import annotations

import sys
from pathlib import Path

import numpy as np
import pandas as pd

RAIZ = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RAIZ))

CENTRO = "3000"
FECHA_INICIO = pd.Timestamp("2023-01-01")
FECHA_FIN = pd.Timestamp.today().to_period("M").to_timestamp()
MESES = pd.date_range(FECHA_INICIO, FECHA_FIN, freq="MS")

rng = np.random.default_rng(42)

# (código, descripción, patrón)
MATERIALES = [
    (100001, "VALVULA BOLA 2 PULG ACERO INOX", "suave"),
    (100002, "SELLO MECANICO BOMBA CENTRIFUGA", "erratica"),
    (100003, "RODAMIENTO SKF 6205 2RS", "intermitente"),
    (100004, "FILTRO HIDRAULICO 10 MICRONES", "suave"),
    (100005, "CORREA TRANSPORTADORA 800MM EP400", "intermitente"),
    (100006, "MANOMETRO GLICERINA 0-10 BAR", "erratica"),
    (100007, "REPUESTO OBSOLETO SIN USO", "sin_demanda"),
    (100008, "ACOPLE FLEXIBLE TIPO L095", "intermitente"),
    (100009, "REPUESTO CRITICO USO MUY RARO", "intermitente_pocas"),
]


def _demanda_mensual(patron: str, n: int) -> np.ndarray:
    """Genera una serie de consumos mensuales según el patrón."""
    if patron == "suave":
        return np.maximum(0, rng.normal(40, 6, n)).round()
    if patron == "erratica":
        # gamma con forma ~1.2 -> demanda casi todos los meses pero muy variable
        # (CV alto -> CV² >= 0.49 -> se clasifica como Errática -> método COMBINADO)
        return np.maximum(1, rng.gamma(1.2, 25, n)).round()
    if patron == "intermitente":
        d = np.zeros(n)
        for i in range(n):
            if rng.random() < 0.35:  # ~1 de cada 3 meses
                d[i] = rng.integers(5, 40)
        return d
    if patron == "intermitente_pocas":
        # Solo 2 demandas en toda la serie -> método SBA, tiempo "Indeterminado"
        d = np.zeros(n)
        posiciones = rng.choice(n, size=2, replace=False)
        for pos in posiciones:
            d[pos] = rng.integers(5, 30)
        return d
    if patron == "sin_demanda":
        return np.zeros(n)
    return np.zeros(n)


def _entradas_mensuales(demanda: np.ndarray) -> np.ndarray:
    """Compras/ingresos: reponen de vez en cuando en lotes."""
    n = len(demanda)
    entradas = np.zeros(n)
    for i in range(n):
        if rng.random() < 0.4:
            entradas[i] = max(demanda[max(0, i - 1):i + 1].sum(), rng.integers(20, 80))
    return entradas.round()


def generar():
    filas_mb51 = []
    filas_mb5b = []

    for material, desc, patron in MATERIALES:
        n = len(MESES)
        demanda = _demanda_mensual(patron, n)
        entrada = _entradas_mensuales(demanda)

        stock = 100.0
        for i, mes in enumerate(MESES):
            stock_ini = stock
            stock = stock_ini + entrada[i] - demanda[i]
            if stock < 0:  # nunca negativo
                stock = 0.0

            # --- MB5B: una fila por material y mes ---
            filas_mb5b.append({
                "Material": material,
                "Descripción del material": desc,
                "De fecha": mes,
                "A fecha": (mes + pd.offsets.MonthEnd(0)),
                "Stock inicial": stock_ini,
                "Total ctd.entrada mcía.": entrada[i],
                "Total cantidades salida": demanda[i],
                "Stock de cierre": stock,
                "Unidad medida base": "UN",
            })

            # --- MB51: movimientos individuales del mes ---
            if entrada[i] > 0:
                filas_mb51.append({
                    "Material": material,
                    "Centro": CENTRO,
                    "Almacén": 1000,
                    "Clase de movimiento": "101",
                    "Fecha contabiliz.": mes + pd.Timedelta(days=int(rng.integers(0, 25))),
                    "Ctd.en UM entrada": float(entrada[i]),
                    "Un.medida de entrada": "UN",
                })
            if demanda[i] > 0:
                clase = rng.choice(["201", "261"])
                filas_mb51.append({
                    "Material": material,
                    "Centro": CENTRO,
                    "Almacén": 1000,
                    "Clase de movimiento": clase,
                    "Fecha contabiliz.": mes + pd.Timedelta(days=int(rng.integers(0, 25))),
                    "Ctd.en UM entrada": -float(demanda[i]),  # negativo = consumo
                    "Un.medida de entrada": "UN",
                })

    df_mb51 = pd.DataFrame(filas_mb51)
    df_mb5b = pd.DataFrame(filas_mb5b)

    (RAIZ / "data" / "MB51").mkdir(parents=True, exist_ok=True)
    (RAIZ / "data" / "MB5B").mkdir(parents=True, exist_ok=True)

    ruta_mb51 = RAIZ / "data" / "MB51" / "MB51_ejemplo.xlsx"
    ruta_mb5b = RAIZ / "data" / "MB5B" / "MB5B_ejemplo.xlsx"
    df_mb51.to_excel(ruta_mb51, index=False)
    df_mb5b.to_excel(ruta_mb5b, index=False)

    print(f"MB51 generado: {ruta_mb51}  ({len(df_mb51)} movimientos)")
    print(f"MB5B generado: {ruta_mb5b}  ({len(df_mb5b)} filas de stock)")


if __name__ == "__main__":
    generar()

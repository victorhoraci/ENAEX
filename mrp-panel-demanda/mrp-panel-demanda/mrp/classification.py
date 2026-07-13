"""
Clasificación de demanda por material (metodología ADI / CV²).

Replica la consulta %ClasificacionDemanda.

Para cada Material + Centro calcula:
  - N_meses                     : cantidad de meses en la serie
  - Meses_con_demanda           : meses con demanda > 0
  - Promedio_demandas_positivas : media de las demandas positivas
  - DesvEst_demandas_positivas  : desviación estándar MUESTRAL (n-1) de las positivas
  - ADI = N_meses / Meses_con_demanda
  - CV2 = (DesvEst / Promedio)²
  - Tipo_demanda : Suave / Errática / Intermitente / Irregular / Sin demanda
  - Metodo       : método de pronóstico recomendado
"""

from __future__ import annotations

import numpy as np
import pandas as pd

from . import config


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

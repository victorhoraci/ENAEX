"""
Paquete MRP - Cálculo de histórico y pronóstico de demanda para Enaex S.A.

Traducción a Python del modelo desarrollado originalmente en Power Query
(Excel). Expone un pipeline sencillo:

    from mrp.pipeline import construir
    r = construir()
    r.serie          # serie mensual (demanda, entrada, stock)
    r.clasificacion  # tipo de demanda por material
    r.resultado      # último pronóstico por material
    r.tabla_final    # tabla lista para el panel
"""

from . import (
    classification,
    config,
    data_loading,
    forecasting,
    pipeline,
    transform,
)

__all__ = [
    "config",
    "data_loading",
    "transform",
    "classification",
    "forecasting",
    "pipeline",
]

"""
Orquestador del pipeline completo.

Toma las carpetas de MB51 y MB5B y devuelve, en un solo objeto, todas las
tablas que el panel necesita:

  serie        -> serie mensual completa (Demanda, Entrada, Stock) por material
  clasificacion-> tipo de demanda y método por material
  resultado    -> ResultadoFinal (último pronóstico + proceso de renovación)
  materiales   -> catálogo (código + descripción) para el buscador
  tabla_final  -> tabla lista para mostrar/exportar en el panel

Uso típico:
    from mrp.pipeline import construir
    r = construir()                      # usa las carpetas por defecto
    r = construir(mb51="...", mb5b="...") # o rutas específicas
"""

from __future__ import annotations

from dataclasses import dataclass

import pandas as pd

from . import classification, data_loading, forecasting, transform


@dataclass
class ResultadoMRP:
    serie: pd.DataFrame
    clasificacion: pd.DataFrame
    resultado: pd.DataFrame
    materiales: pd.DataFrame
    tabla_final: pd.DataFrame


def _catalogo_materiales(mb5b: pd.DataFrame) -> pd.DataFrame:
    """Código + descripción (la más reciente) de cada material, desde MB5B."""
    cat = mb5b.sort_values("De fecha").groupby("Material", as_index=False).agg(
        **{"Descripción del material": ("Descripción del material", "last")}
    )
    return cat


def construir(
    mb51: str | None = None,
    mb5b: str | None = None,
    mb51_archivos=None,
    mb5b_archivos=None,
    fecha_fin: str | None = None,
) -> ResultadoMRP:
    """
    Ejecuta el pipeline completo y devuelve todas las tablas.

    Los datos pueden venir de carpetas en disco (mb51 / mb5b) o de archivos
    subidos en el navegador (mb51_archivos / mb5b_archivos).
    """
    # 1) Cargar fuentes
    df_mb51 = data_loading.cargar_mb51(mb51, archivos=mb51_archivos)
    df_mb5b = data_loading.cargar_mb5b(mb5b, archivos=mb5b_archivos)

    # 2) Transformar -> serie mensual completa (stock de MB5B en todos los meses)
    desag = transform.demanda_desagregada(df_mb51)
    real_mes = transform.demanda_real_mes(desag, df_mb5b)
    stock = transform.stock_mensual(df_mb5b)
    serie = transform.serie_completa(real_mes, stock=stock, fecha_fin=fecha_fin)

    # 3) Clasificar demanda
    clasif = classification.clasificar_demanda(serie)

    # 4) Pronósticos por método
    base = forecasting.base_pronostico(serie, clasif)
    ses = forecasting.pronostico_ses(base)
    comb = forecasting.pronostico_combinado(base)
    sba = forecasting.pronostico_sba(base)
    pr_final = forecasting.pronostico_pr_final(base)

    # 5) Consolidar
    resultado = forecasting.resultado_final(ses, comb, sba, pr_final, clasif)

    # 6) Catálogo de materiales (para el buscador)
    catalogo = _catalogo_materiales(df_mb5b)

    # 7) Tabla final lista para mostrar
    tabla = resultado.merge(catalogo, on="Material", how="left")
    tabla = _ordenar_tabla_final(tabla)

    return ResultadoMRP(
        serie=serie,
        clasificacion=clasif,
        resultado=resultado,
        materiales=catalogo,
        tabla_final=tabla,
    )


def _ordenar_tabla_final(tabla: pd.DataFrame) -> pd.DataFrame:
    """Deja las columnas más útiles primero, con nombres legibles."""
    orden = [
        "Material",
        "Descripción del material",
        "Centro",
        "Tipo_demanda",
        "Metodo",
        "Pronostico",
        "Pronostico_redondeado",
        "Tiempo_hasta_demanda",
        "Dias_hasta_demanda",
        "MesPronosticado",
        "PR_Periodos_Hasta_Prox",
        "PR_Media_Intervalo",
        "PR_Tamano_Esperado",
        "PR_Pronostico_redondeado",
        "PR_IC95_Inf_Intervalo",
        "PR_IC95_Sup_Intervalo",
        "PR_IC95_Inf_Tamano",
        "PR_IC95_Sup_Tamano",
    ]
    existentes = [c for c in orden if c in tabla.columns]
    resto = [c for c in tabla.columns if c not in existentes]
    return tabla[existentes + resto]

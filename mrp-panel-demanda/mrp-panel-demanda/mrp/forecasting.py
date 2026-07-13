"""
Métodos de pronóstico de demanda y consolidación final.

Replica las consultas:
  - %BasePronostico
  - %Pronostico_SES          (demanda Suave)
  - %Pronostico_SBA          (demanda Intermitente / Irregular)
  - %Pronostico_COMBINADO    (demanda Errática = SES + PM3 + PM6)
  - %Pronostico_PR / _final  (Proceso de Renovación para intermitentes)
  - %ResultadoFinal          (consolidación: último pronóstico por material)

Cada material usa el método que le asignó la clasificación ADI/CV².
"""

from __future__ import annotations

import math

import numpy as np
import pandas as pd

from . import config

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
    """Proceso de renovación completo (una fila por período) para SBA."""
    datos = base[base["Metodo"] == "SBA"].copy()
    partes = []
    for material, grupo in datos.groupby("Material", sort=True):
        g = grupo.sort_values("FechaMes").reset_index(drop=True)
        metricas = pd.DataFrame(_proceso_renovacion(g["Demanda Mensual"].tolist()))
        partes.append(pd.concat([g.reset_index(drop=True), metricas], axis=1))
    if not partes:
        return datos
    return pd.concat(partes, ignore_index=True)


def pronostico_pr_final(base: pd.DataFrame) -> pd.DataFrame:
    """Solo la ÚLTIMA fila del proceso de renovación por material (Metodo == 'SBA')."""
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
def resultado_final(
    ses: pd.DataFrame,
    combinado: pd.DataFrame,
    sba: pd.DataFrame,
    pr_final: pd.DataFrame,
    clasificacion: pd.DataFrame,
) -> pd.DataFrame:
    """
    Toma el ÚLTIMO pronóstico de cada material (según su método), agrega los
    materiales 'Sin demanda' con pronóstico 0, y le pega la información del
    proceso de renovación (tiempo hasta próxima demanda, intervalo, etc.).
    """
    def _sel(df, col):
        if df.empty or col not in df.columns:
            return pd.DataFrame(columns=["Material", "Centro", "FechaMes",
                                         "Tipo_demanda", "Metodo", "Pronostico"])
        out = df[["Material", "Centro", "FechaMes", "Tipo_demanda", "Metodo", col]].copy()
        return out.rename(columns={col: "Pronostico"})

    unidos = pd.concat([
        _sel(ses, "Pronostico_SES"),
        _sel(combinado, "Pronostico_COMBINADO"),
        _sel(sba, "Pronostico_SBA"),
    ], ignore_index=True)

    # Último mes por Material + Centro
    if not unidos.empty:
        ultimo = (
            unidos.sort_values(["Material", "Centro", "FechaMes"])
            .groupby(["Material", "Centro"], as_index=False)
            .tail(1)
            .reset_index(drop=True)
        )
    else:
        ultimo = unidos

    # Materiales sin demanda -> FechaMes nulo, Pronostico 0
    sin_demanda = clasificacion[clasificacion["Tipo_demanda"] == "Sin demanda"].copy()
    if not sin_demanda.empty:
        sin_demanda["FechaMes"] = pd.NaT
        sin_demanda["Pronostico"] = 0.0
        sin_demanda = sin_demanda[["Material", "Centro", "FechaMes",
                                   "Tipo_demanda", "Metodo", "Pronostico"]]
        final = pd.concat([ultimo, sin_demanda], ignore_index=True)
    else:
        final = ultimo

    # Pegar datos del proceso de renovación (por Material + Centro)
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

    # Columnas calculadas
    final["MesPronosticado"] = final["FechaMes"].apply(
        lambda f: (f + pd.DateOffset(months=1)) if pd.notna(f) else pd.NaT
    )
    final["Pronostico_redondeado"] = final["Pronostico"].apply(
        lambda x: int(math.ceil(x)) if pd.notna(x) else pd.NA
    )
    final["PR_Pronostico_redondeado"] = final["PR_Tamano_Esperado"].apply(
        lambda x: int(math.ceil(x)) if pd.notna(x) else pd.NA
    )

    return final.sort_values(["Material", "Centro"]).reset_index(drop=True)

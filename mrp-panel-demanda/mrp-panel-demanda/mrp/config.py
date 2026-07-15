"""
Configuración central del panel MRP - Histórico y Pronóstico de Demanda.

Todos los parámetros del modelo viven aquí para que sean fáciles de ajustar
sin tener que tocar la lógica de cálculo. Si mañana cambia el alfa del
suavizamiento o los cortes de clasificación, solo se edita este archivo.
"""

from __future__ import annotations

from pathlib import Path

# --------------------------------------------------------------------------
# RUTAS DE DATOS
# --------------------------------------------------------------------------
# Carpeta raíz del proyecto (…/mrp-panel-demanda)
RAIZ_PROYECTO = Path(__file__).resolve().parent.parent

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

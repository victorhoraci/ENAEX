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

# Mapa: tipo de demanda -> método de pronóstico recomendado.
METODO_POR_TIPO = {
    "Suave": "SES",
    "Errática": "COMBINADO",
    "Intermitente": "SBA",
    "Irregular": "SBA",
    "Sin demanda": "Sin cálculo",
}


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

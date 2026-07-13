# 📦 Panel MRP · Histórico y Pronóstico de Demanda — Enaex S.A.

Panel de visualización en **Streamlit** para el estudio de planificación de
materiales (MRP). Reemplaza el reporte de Power BI por una aplicación en
Python, versionable en GitHub y fácil de mantener.

Para cada material muestra, en un solo gráfico, sus **ingresos** y **egresos**
mensuales (barras) y el **stock de cierre** (línea), junto a su **clasificación
de demanda** y el **pronóstico** del próximo mes. Toda la lógica de cálculo que
antes vivía en el editor de Power Query de Excel está reescrita en Python.

---

## ✨ Qué hace

- **Un gráfico por material** con tres variables a la vez: ingreso (barra azul),
  egreso/demanda (barra roja) y stock de cierre (línea verde), más el pronóstico
  del próximo mes (estrella naranja).
- **Buscador escribible**: escribe el código o parte de la descripción y la
  lista se filtra sola.
- **Tarjetas** con código, descripción, centro, clasificación, método,
  pronóstico, tiempo hasta la próxima demanda e intervalo de demanda.
- **Tabla completa** de todos los materiales, filtrable y descargable en CSV.
- **Clasificación ADI / CV²** y **cuatro métodos de pronóstico** (SES, SBA,
  COMBINADO y Proceso de Renovación), elegidos automáticamente por material.

---

## 🗂️ Estructura del proyecto

```
mrp-panel-demanda/
├── app.py                     # Panel de Streamlit (la interfaz)
├── requirements.txt           # Dependencias
├── README.md                  # Este archivo
├── .streamlit/config.toml     # Tema visual
│
├── mrp/                        # Lógica de cálculo (traducción del Power Query)
│   ├── config.py               # Parámetros: α, cortes ADI/CV², rutas, colores
│   ├── data_loading.py         # Lee los Excel de MB51 y MB5B
│   ├── transform.py            # Demanda desagregada, mensual y serie completa
│   ├── classification.py       # Clasificación ADI / CV²
│   ├── forecasting.py          # SES, SBA, COMBINADO, Proceso de Renovación
│   └── pipeline.py             # Une todo y entrega las tablas al panel
│
├── scripts/
│   └── generar_datos_ejemplo.py  # Crea datos sintéticos para probar sin SAP
│
└── data/                       # Los Excel de SAP van aquí (no se suben al repo)
    ├── MB51/                    # Movimientos de material
    └── MB5B/                    # Stock mensual (un archivo por mes)
```

---

## 🚀 Puesta en marcha

Requiere **Python 3.10 o superior**.

```bash
# 1) Clonar el repositorio
git clone <URL-de-tu-repositorio>
cd mrp-panel-demanda

# 2) Crear un entorno virtual (recomendado)
python -m venv .venv
# Windows:
.venv\Scripts\activate
# macOS / Linux:
source .venv/bin/activate

# 3) Instalar dependencias
pip install -r requirements.txt

# 4a) (Opcional) Generar datos de ejemplo para probar sin SAP
python scripts/generar_datos_ejemplo.py

# 4b) O dejar los Excel reales en data/MB51 y data/MB5B (ver abajo)

# 5) Levantar el panel
streamlit run app.py
```

Se abrirá en el navegador (por defecto `http://localhost:8501`).

---

## 🔄 Cómo actualizar los datos

El panel se alimenta de dos descargas de **SAP HANA**:

| Transacción | Qué es | Cada cuánto | Carpeta | Al actualizar |
|---|---|---|---|---|
| **MB51** | Movimientos (ingresos 101, consumos 201/261) | **Semanal** | `data/MB51/` | **Reemplaza** al anterior |
| **MB5B** | Stock del mes por material | **Mensual** | `data/MB5B/` | **Se agrega** uno nuevo |

Tienes dos formas de actualizar:

- **Desde el panel** (recomendado): pestaña **➕ Agregar datos**. Subes el MB51
  (reemplaza) o el MB5B (se agrega) y el panel recalcula solo. *Nota:* en
  Streamlit Cloud estos archivos son temporales; para que queden permanentes,
  súbelos al repo o usa el panel en tu PC.
- **Desde las carpetas**: deja los Excel directamente en `data/MB51/` y
  `data/MB5B/` (por ejemplo subiéndolos a GitHub) y refresca el panel.

El panel reconoce las columnas por su nombre y tolera variantes comunes de los
encabezados de SAP. Si falta una columna clave, avisa cuál no encontró.

---

## 🧮 Cómo se calcula el pronóstico

1. **Demanda mensual** — se suman los consumos (clases 201/261) por material y mes.
   Los ingresos (clase 101) se suman aparte como entradas.
2. **Serie completa** — se arma una fila por material y mes desde enero-2023 hasta
   el mes actual, rellenando con 0 los meses sin movimiento.
3. **Clasificación ADI / CV²** (Syntetos & Boylan):
   - **ADI** = meses totales ÷ meses con demanda (cada cuánto hay demanda).
   - **CV²** = (desviación ÷ promedio)² del tamaño de la demanda.
   - Resultado: *Suave*, *Errática*, *Intermitente* o *Irregular*
     (o *Sin demanda* si nunca hubo consumo). Cortes: ADI = 1.32, CV² = 0.49.
4. **Método y "tiempo hasta la próxima demanda" según el tipo:**
   | Tipo | Método | Tiempo hasta demanda |
   |---|---|---|
   | **Constante** (antes "Suave") | **SES** | Días hasta el próximo mes |
   | **Errática** | **COMBINADO** (SES + media móvil 3 + media móvil 6) | Días hasta el próximo mes |
   | **Intermitente / Irregular** con **< 4** demandas | **SBA** (Croston-SBA) | **Indeterminado** |
   | **Intermitente / Irregular** con **≥ 4** demandas | **PR** (Proceso de Renovación) | **Días** estimados hasta la próxima demanda |
   | **Sin demanda** | — | — |
5. **Proceso de Renovación** (materiales PR): estima cuándo ocurrirá la próxima
   demanda (en meses, que el panel convierte a días), su tamaño esperado y los
   intervalos de confianza al 95 %.
6. **Resultado final** — se toma el último pronóstico de cada material.

El umbral entre SBA y PR (4 demandas) y la conversión de meses a días se ajustan
en `mrp/config.py` (`MIN_DEMANDAS_PR`, `DIAS_POR_MES`).

Todos los parámetros (α = 0.3, cortes de clasificación, horizonte = 12) están en
[`mrp/config.py`](mrp/config.py) y se ajustan sin tocar la lógica.

---

## 🔌 Integración futura

Las tablas que produce el pipeline (`serie`, `clasificacion`, `resultado`,
`tabla_final`) son DataFrames de pandas, listos para cruzarse con otras fuentes
(costos, criticidad, lead time, etc.) y así analizar relaciones entre variables.
Se accede a ellas con:

```python
from mrp.pipeline import construir
r = construir()
r.serie          # serie mensual (demanda, entrada, stock)
r.clasificacion  # tipo de demanda por material
r.resultado      # último pronóstico + proceso de renovación
r.tabla_final    # tabla lista para mostrar/exportar
```

---

## 📤 Subir a GitHub

```bash
git init
git add .
git commit -m "Panel MRP: histórico y pronóstico de demanda"
git branch -M main
git remote add origin <URL-de-tu-repositorio>
git push -u origin main
```

> Los Excel de `data/` **no se suben** (están en `.gitignore`) porque son
> información interna. El repositorio conserva solo las carpetas vacías.

Para publicarlo online sin servidor propio, puedes usar
[Streamlit Community Cloud](https://streamlit.io/cloud): conecta el repositorio
de GitHub y apunta a `app.py`.

---

## 🧪 Notas técnicas

- La clasificación y los pronósticos se calculan por **Material + Centro**, para
  que un mismo material en distintos centros no se mezcle.
- La desviación estándar usa el estimador **muestral (n-1)**, igual que
  `List.StandardDeviation` de Power Query.
- La demanda intermitente/irregular usa **PR** solo si tiene **4 o más**
  demandas históricas (`MIN_DEMANDAS_PR`); con menos, usa **SBA** y su tiempo
  hasta la próxima demanda queda como **"Indeterminado"**.

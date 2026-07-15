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

**Cómo descargarlos desde SAP HANA (layouts):**

- **MB51:** transacción **MB51** con el layout **`/CALCDEMANDA`** —
  *"MOV. PARA PRONOSTICO DE DEMANDA"*. Exporta a Excel.
- **MB5B:** transacción **MB5B** con un layout que deje estas columnas, en este
  orden: Material · Descripción del material · De fecha · A fecha · Stock inicial ·
  Total ctd.entrada mcía. · Total cantidades salida · Stock de cierre ·
  Unidad medida base · Stock especial.

El panel reconoce las columnas por su nombre y tolera variantes comunes de los
encabezados de SAP. Si falta una columna clave, avisa cuál no encontró.

---

## 🔐 Guardar en GitHub desde la app (con contraseña)

Por defecto, los archivos que subes en la pestaña **➕ Agregar datos** se guardan
solo temporalmente en la nube. Puedes activar que se **guarden directamente en el
repositorio de GitHub** (permanentes) y que **pidan una contraseña**. Se
configura una sola vez con los *secrets* de Streamlit.

**1) Crea un token de GitHub (permite a la app escribir en el repo):**
1. En GitHub: foto de perfil → **Settings** → abajo **Developer settings**.
2. **Personal access tokens → Fine-grained tokens → Generate new token**.
3. Nombre y expiración a gusto. En **Repository access** elige *Only select
   repositories* → tu repo `ENAEX`.
4. En **Permissions → Repository permissions → Contents**, selecciona
   **Read and write**.
5. Genera el token y **cópialo** (empieza con `github_pat_...`). Solo se muestra una vez.

**2) Configura los *secrets* en Streamlit Cloud:**
1. Abre tu app en share.streamlit.io → menú **⋮** o **Settings → Secrets**.
2. Pega esto (reemplazando con tus valores) y guarda:

   ```toml
   APP_PASSWORD       = "una-clave-secreta"
   GITHUB_TOKEN       = "github_pat_xxxxxxxx"
   GITHUB_REPO        = "victorhoraci/ENAEX"
   GITHUB_BRANCH      = "main"
   GITHUB_DATA_PREFIX = "mrp-panel-demanda/mrp-panel-demanda/data"
   ```

Al guardar, la app se reinicia. Desde entonces, en **➕ Agregar datos** te pedirá
la contraseña y, al subir un Excel, lo dejará guardado en el repo (MB51 reemplaza,
MB5B se agrega). La app vuelve a leer los datos sola en ~1 minuto.

> El archivo `secrets.toml` **nunca** se sube al repo (está en `.gitignore`).
> Para probar localmente, copia `.streamlit/secrets.toml.example` como
> `.streamlit/secrets.toml` y complétalo.

---

## 🚚 Panel 2 · Abastecimiento (seguimiento y control)

Segundo panel (menú lateral → **Abastecimiento**) para hacer seguimiento y tomar
decisiones sobre los materiales: qué está en validación o con solped bloqueada,
cuántos días llevan las solped y las OC, qué OC están atrasadas, si son
nacionales o internacionales, y si el stock alcanza para la próxima demanda.

### Qué archivos subir (y cada cuánto)

| Archivo | De dónde sale | Cada cuánto | Carpeta |
|---|---|---|---|
| **MRP semanal** | `Planificacion_Simpl…xlsx`, hoja `data` (encabezados en la fila 16) | Semanal | `data/MRP` |
| **MM60** | Transacción MM60 (maestro: precio, ABC, grupo compra) | Mensual | `data/MM60` |
| **ME5A** | Transacción ME5A (solicitudes / solped) | Semanal | `data/ME5A` |
| **ME2M** | Transacción ME2M (órdenes de compra) | Semanal | `data/ME2M` |
| **TAT** | `…VISTA_EJECUTIVA_MATERIALES_….xlsx`, hoja **`Dias_TAT`** | Mensual | `data/TAT` |

Todos se pueden subir desde la pestaña **➕ Actualizar datos** del panel (con
contraseña y guardado en GitHub). Cada uno **reemplaza** al anterior.

> El TAT se lee de la hoja **`Dias_TAT`** de la *Vista Ejecutiva de Materiales*
> (una fila por material: media, mínimo, máximo, desviación, variabilidad,
> recurrencia y días desde la última solicitud). Si se sube el archivo MERGE
> antiguo, el panel también lo entiende (compatibilidad hacia atrás).

### Cómo se unen las tablas (llaves)

| Fuente | Se une por |
|---|---|
| MRP semanal | tabla **base** (nadie se pierde) |
| MM60 | Material + Centro |
| ME5A | Solped + Material |
| ME2M | OC + Material |
| TAT | Material (todos los centros, sin centro) |
| Demanda (panel 1) | Material |

### 🔑 Ningún material se pierde

Todas las uniones son *left join* sobre el MRP: un material sin solped, sin OC,
sin TAT o sin pronóstico **igual aparece**, con esos campos en blanco. Además,
la pestaña **Gestión** separa en tablas propias los materiales en *validación* y
con *solped bloqueada*, para que no desaparezcan al mostrar los días de gestión
(el problema que ocurría en Power BI). La pestaña **Todos los materiales** siempre
muestra el total.

### Medidas replicadas del DAX

| Medida (Power BI) | En el panel |
|---|---|
| `Observación` (Con OC / Con Solped / Solped bloqueada / Validación) | **Estado gestión** |
| `Estado OC` (Atrasada / En curso) | **Estado OC** |
| `Días desde Solped`, `1Días Duración OC` | **Días en solped**, **Días de OC** |
| `Dias hasta llegada material` | **Días hasta llegada** |
| `Rango Días Solicitud`, `Rango Días Atraso Material` | **Rango días solped**, **Rango atraso OC** |
| `Nacionalidad OC` (45→Nacional, 47→Internacional) | **Nacionalidad** |
| `Disponibilidad Conservadora` | **Disponible conservador** |
| `Criticidad Texto` (A→Alta, C→Baja, resto→Media) | **Criticidad texto** |
| `Rango TAT` | **Rango TAT** |
| `Pronostico_Consolidado`, `Tiempo_Prox_Demanda`, `Cumple_Demanda` | mismas columnas |
| `Valor Stock Total`, `1Precio en Tránsito` | **Valor stock**, **Costo en tránsito** |

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

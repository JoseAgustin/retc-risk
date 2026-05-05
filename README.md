# RETC Nuevo León – Análisis de Riesgo por Emisiones al Aire

Herramienta de procesamiento de datos del **Registro de Emisiones y Transferencia de Contaminantes (RETC)** para estimar el riesgo cancerígeno y no cancerígeno por inhalación de contaminantes emitidos al aire en el estado de **Nuevo León**, México, con énfasis en la **Zona Metropolitana de Monterrey (ZMM)**.

---

## Tabla de contenidos

- [Descripción](#descripción)
- [Archivos de entrada](#archivos-de-entrada)
- [Archivos de salida](#archivos-de-salida)
- [Instalación](#instalación)
- [Uso](#uso)
- [Metodología](#metodología)
- [Parámetros](#parámetros)
- [Estructura del proyecto](#estructura-del-proyecto)
- [Limitaciones y advertencias](#limitaciones-y-advertencias)
- [Preguntas frecuentes](#preguntas-frecuentes)
- [Licencia](#licencia)

---

## Descripción

El script `process_retc_nl.py` automatiza el procesamiento de los archivos anuales del RETC (disponibles en [gob.mx](https://www.gob.mx/semarnat/acciones-y-programas/retc)) para:

1. Filtrar las emisiones al aire de empresas ubicadas en Nuevo León.
2. Identificar los municipios de la Zona Metropolitana de Monterrey.
3. Calcular concentraciones estimadas de contaminantes en aire ambiente.
4. Estimar puntajes de riesgo cancerígeno y no cancerígeno usando valores toxicológicos de referencia (IUR y REL).
5. Exportar los resultados en archivos CSV listos para análisis o visualización.

---

## Archivos de entrada

| Archivo | Descripción |
|---|---|
| `retc YYYY.xlsx` | Archivo anual del RETC. Debe contener las hojas **"Datos Generales"** y **"Emisiones y Transferencias"** con encabezados a partir de la fila 10. Puede haber uno o varios archivos por año. |
| `contable.xlsx` | Tabla de referencia toxicológica con valores **IUR** (Inhalation Unit Risk) y **REL** (Reference Exposure Level) por número CAS. |

### Estructura esperada de `retc YYYY.xlsx`

**Hoja: Datos Generales** (encabezados en fila 10)

| Columna | Descripción |
|---|---|
| NRA | Clave de identificación de la empresa |
| NOMBRE | Razón social |
| ESTADO / ENTIDAD | Estado donde opera la empresa |
| MUNICIPIO | Municipio |
| LATITUD | Coordenada geográfica (grados decimales) |
| LONGITUD | Coordenada geográfica (grados decimales) |
| UTMX / columna J | Coordenada UTM X en metros (Zona 14N) |
| UTMY / columna K | Coordenada UTM Y en metros (Zona 14N) |

**Hoja: Emisiones y Transferencias** (encabezados en fila 10)

| Columna | Descripción |
|---|---|
| NRA | Clave de la empresa (llave foránea) |
| CAS | Número CAS de la sustancia |
| SUSTANCIA | Nombre de la sustancia |
| AIRE | Emisión anual al aire en kg/año |

### Estructura esperada de `contable.xlsx`

El script detecta automáticamente las columnas buscando palabras clave en los encabezados:

| Palabra clave buscada | Columna que representa |
|---|---|
| `CHEMICAL ABSTRACT` o `CAS` | Número CAS |
| `INHALATION UNIT RISK` | IUR – (µg/m³)⁻¹ |
| `CHRONIC INHALATION` | REL – µg/m³ |
| `SUBSTANCE` | Nombre de la sustancia |

---

## Archivos de salida

Para cada año procesado se crea una subcarpeta `output_{YYYY}/` dentro del directorio de salida, con los siguientes archivos:

| Archivo | Descripción |
|---|---|
| `{year}_companies.csv` | Empresas de Nuevo León con coordenadas geográficas y UTM (km). |
| `{year}_toxics.csv` | Catálogo de sustancias tóxicas emitidas (CAS + nombre). |
| `{year}_company_emissions_risk.csv` | Tabla pivote con emisiones por sustancia (kg/año), concentraciones estimadas (µg/m³) y puntajes de riesgo **cancerígeno** (Métodos 1 y 2) por empresa. |
| `{year}_company_noncancer_score.csv` | Puntajes de riesgo **no cancerígeno** por empresa y sustancia. |

### Columnas clave en `{year}_company_emissions_risk.csv`

| Columna | Descripción |
|---|---|
| `NRA`, `NOMBRE`, `MUNICIPIO` | Identificadores de la empresa |
| `UTMX_KM`, `UTMY_KM` | Coordenadas UTM en kilómetros |
| `CONC_ugm3_{CAS}` | Concentración estimada en µg/m³ (Método 1) |
| `RISK_{CAS}` | Riesgo cancerígeno unitario (Método 1) |
| `CONC_SCORE_{CAS}` | Concentración estimada (Método 2) |
| `CANCER_SCORE_{CAS}` | Puntaje de riesgo cancerígeno (Método 2) |
| `RISK_TOTAL` | Suma del riesgo cancerígeno de todas las sustancias (Método 1) |
| `CANCER_SCORE_TOTAL` | Suma del puntaje cancerígeno de todas las sustancias (Método 2) |

### Columnas clave en `{year}_company_noncancer_score.csv`

| Columna | Descripción |
|---|---|
| `NONCANCER_SCORE_{CAS}` | Puntaje no cancerígeno por sustancia (`NaN` si la sustancia no tiene REL) |
| `NONCANCER_SCORE_TOTAL` | Suma de puntajes no cancerígenos (ignora `NaN`) |

---

## Instalación

### Requisitos

- Python ≥ 3.8
- pip

### Dependencias

```bash
pip install pandas openpyxl
```

> `openpyxl` es el motor que usa pandas para leer archivos `.xlsx`.

### Clonar el repositorio

```bash
git clone https://github.com/tu-usuario/retc-nuevo-leon.git
cd retc-nuevo-leon
```

---

## Uso

### Uso básico

```bash
python3 process_retc_nl.py \
    --input-dir /ruta/a/los/datos \
    --output-dir /ruta/de/salida
```

### Con factores de dilución personalizados

```bash
python3 process_retc_nl.py \
    --input-dir /datos/retc \
    --output-dir /datos/retc/resultados \
    --dilution 3.10 \
    --dilution2 0.30
```

### Ejemplo completo

```
/datos/retc/
├── retc 2018.xlsx
├── retc 2021.xlsx
├── retc 2022.xlsx
├── retc 2023.xlsx
└── contable.xlsx

python3 process_retc_nl.py -i /datos/retc -o /datos/retc/resultados
```

**Resultado esperado en consola:**

```
Cargando valores IUR y REL desde contable.xlsx...
  Columnas encontradas en contable.xlsx:
    - CAS  : Chemical Abstract Number
    - IUR  : Inhalation Unit Risk
    - REL  : Chronic Inhalation REL
    - Nombre: Substance Name
  Total de sustancias cargadas : 187
  Con valores IUR              : 112
  Con valores REL              : 134
  Factor de dilución DF1       : 2.47
  Factor de dilución DF2       : 0.25

Archivos RETC encontrados: 4

Procesando retc 2018.xlsx...
  - Registros en Nuevo León    : 4320
  - Empresas únicas            : 198
  - Sustancias tóxicas         : 73
  - Sustancias con REL válido  : 61
  ⚠  12 sustancias sin valor REL (puntaje no-cáncer = NaN):
      - 7664-41-7 (Ammonia)
      ...
  ✓ Resultados guardados en: /datos/retc/resultados/output_2018
```

---

## Metodología

### Base metodológica

Este script implementa el procedimiento de evaluación de riesgo por inhalación definido por la **California Air Resources Board (CARB)** en el programa **AB 2588 – Air Toxics Hot Spots**:

> [https://ww2.arb.ca.gov/our-work/programs/ab-2588-air-toxics-hot-spots/hot-spots-risk-assessment](https://ww2.arb.ca.gov/our-work/programs/ab-2588-air-toxics-hot-spots/hot-spots-risk-assessment)

El programa AB 2588 establece un marco para identificar fuentes de emisión de contaminantes tóxicos al aire, cuantificar el riesgo para la salud de la población cercana y priorizar acciones de reducción. Las ecuaciones de concentración y los factores de dilución utilizados en este script siguen las guías técnicas de dicho programa.

### Tabla de valores toxicológicos (`contable.xlsx`)

Los valores de IUR y REL provienen de la tabla de salud de referencia publicada por CARB:

> **Health Values Table** – California Air Resources Board  
> [https://ww2.arb.ca.gov/sites/default/files/classic/toxics/healthval/contable09252025.pdf](https://ww2.arb.ca.gov/sites/default/files/classic/toxics/healthval/contable09252025.pdf)

Para preparar el archivo `contable.xlsx` a partir de ese PDF:

1. Descarga el PDF desde el enlace anterior.
2. Conviértelo a Excel (`.xlsx`) usando la herramienta de tu preferencia (Adobe Acrobat, LibreOffice, etc.).
3. **Elimina los renglones de encabezado repetidos** que aparecen al inicio de cada página subsiguiente a la primera (son artefactos de la conversión del PDF paginado).
4. Guarda el archivo resultante como `contable.xlsx` en el mismo directorio que los archivos RETC.

> **Nota:** La tabla de CARB se actualiza periódicamente. Verifica la fecha del archivo PDF para asegurarte de usar la versión más reciente de los valores toxicológicos.

### Filtrado geográfico

Se filtran empresas cuyo campo `ESTADO` (o `ENTIDAD`) corresponda a **Nuevo León**, tolerando variantes tipográficas (con/sin tilde, con/sin espacios). Adicionalmente, se marca con `ES_ZM_MONTERREY = True` si el municipio pertenece a la Zona Metropolitana de Monterrey.

### Cálculo de concentración y riesgo

#### Método 1 – Riesgo estándar

Apropiado para chimeneas de ~9 m de altura a 100 m de distancia.

```
Concentración [µg/m³] = Emisión [kg/año] × 2.20462 [lb/kg] / 365 [días/año] × DF1

Riesgo cancerígeno = Concentración × IUR
```

#### Método 2 – Puntaje simplificado

```
Concentración [µg/m³] = Emisión [kg/año] × 2.20462 × 7700 × DF2

Puntaje cancerígeno = Concentración × IUR
```

#### Puntaje no cancerígeno

```
Puntaje = Emisión [lb/año] × DF2 / REL
```

> Si una sustancia no tiene REL en `contable.xlsx`, su puntaje no cancerígeno se reporta como `NaN` y **no se suma** al total.

### Glosario

| Término | Definición |
|---|---|
| **IUR** | *Inhalation Unit Risk*: incremento de riesgo de cáncer por unidad de concentración inhalada durante toda la vida, en (µg/m³)⁻¹. |
| **REL** | *Reference Exposure Level*: concentración crónica en aire que no se espera cause efectos adversos no cancerígenos, en µg/m³. |
| **DF1** | Factor de dilución estándar: 2.47 µg/m³ por lb/día (modelo de dispersión simplificado para fuentes puntuales). |
| **DF2** | Factor de dilución para puntaje: 0.25 (µg/m³)/(lb/año). |
| **ZMM** | Zona Metropolitana de Monterrey: 18 municipios incluidos en el análisis. |

---

## Parámetros

| Argumento | Corto | Tipo | Por defecto | Descripción |
|---|---|---|---|---|
| `--input-dir` | `-i` | str | *(requerido)* | Directorio con los archivos RETC y `contable.xlsx`. |
| `--output-dir` | `-o` | str | *(requerido)* | Directorio donde se guardarán los CSVs. |
| `--dilution` | `-d` | float | `2.47` | Factor de dilución DF1 (µg/m³ por lb/día). |
| `--dilution2` | `-d2` | float | `0.25` | Factor de dilución DF2 ((µg/m³)/(lb/año)). |

---

## Estructura del proyecto

```
retc-nuevo-leon/
├── process_retc_nl.py   # Script principal
├── README.md            # Este archivo
└── data/                # (No incluida en el repo) – coloca aquí tus datos
    ├── contable.xlsx
    ├── retc 2018.xlsx
    └── ...
```

---

## Limitaciones y advertencias

- **Precisión del modelo de dispersión**: Los factores de dilución DF1 y DF2 son aproximaciones simplificadas. Para evaluaciones de riesgo regulatorias se recomienda utilizar modelos de dispersión atmosférica certificados (AERMOD, CALPUFF).

- **Sustancias sin REL**: Cuando una sustancia no tiene valor REL en `contable.xlsx`, su puntaje no cancerígeno se reporta como `NaN`. El total acumulado `NONCANCER_SCORE_TOTAL` **subestima** el riesgo real para esas empresas.

- **Emisiones reportadas**: Los datos del RETC representan emisiones autodeclaradas por las empresas obligadas a reportar. No incluyen fuentes menores no obligadas ni emisiones fugitivas no capturadas.

- **Coordenadas UTM**: El script asume **Zona UTM 14N** (EPSG:32614), que corresponde a la mayor parte de Nuevo León. Verifica que tus archivos RETC usen el mismo datum.

- **Variaciones de formato**: El script maneja variaciones comunes entre versiones anuales del RETC (nombres de columna, tildes, espacios). Si un año produce errores, revisa que las hojas tengan los nombres y estructura esperados.

---

## Preguntas frecuentes

**¿Dónde descargo los archivos RETC?**
En el portal de datos abiertos de SEMARNAT: [http://sinat.semarnat.gob.mx](http://sinat.semarnat.gob.mx/retc/retc/index.php)) → sección RETC.

**¿Qué pasa si el archivo RETC tiene un formato diferente?**
Verifica que las hojas se llamen exactamente `Datos Generales` y `Emisiones y Transferencias`, y que los encabezados inicien en la fila 10. Si el formato difiere, ajusta el parámetro `header` en `_load_retc_sheets()`.

**¿Por qué algunas sustancias aparecen con NONCANCER_SCORE = NaN?**
Porque no tienen un valor REL definido en `contable.xlsx`. Esto puede deberse a que la sustancia es exclusivamente cancerígena (sin umbral de efecto no cancerígeno conocido) o a que el valor no está incluido en la versión de la tabla que usas.

**¿Puedo usar otros estados además de Nuevo León?**
Sí. Modifica la lista `NUEVO_LEON_PATTERNS` y la constante `MONTERREY_MUNICIPALITIES` en el script para adaptar el filtrado a otro estado o zona metropolitana.

**¿Dónde descargo la tabla contable.pdf?**
En el portal de California Air Resources Board  [ww2.arb.ca.gov/](https://ww2.arb.ca.gov/es/resources/documents/consolidated-table-oehha-carb-approved-risk-assessment-health-values)) → en la liga __Consolidated Table of OEHHA/ARB Approved Risk Assessment Health Values__.

---

## Licencia

Este proyecto se distribuye bajo la licencia **MIT**. Consulta el archivo `LICENSE` para más detalles.

---

*Desarrollado para análisis ambiental de emisiones industriales en la Zona Metropolitana de Monterrey, Nuevo León, México.*

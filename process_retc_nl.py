#!/usr/bin/env python3
"""
process_retc_nl.py
==================
Procesamiento de archivos RETC (Registro de Emisiones y Transferencia de Contaminantes)
para el estado de Nuevo León / Zona Metropolitana de Monterrey.

Descripción
-----------
Este script lee archivos Excel del RETC (retc YYYY.xlsx) y un archivo de referencia
toxicológica (contable.xlsx) que contiene valores IUR y REL, filtra las emisiones al aire
correspondientes al estado de Nuevo León y genera cuatro archivos CSV por año con datos
de empresas, sustancias tóxicas, puntuaciones de riesgo cancerígeno y no cancerígeno.

Archivos de salida (por año)
-----------------------------
  - {year}_companies.csv              : Empresas registradas en Nuevo León con coordenadas.
  - {year}_toxics.csv                 : Catálogo de sustancias tóxicas emitidas.
  - {year}_company_emissions_risk.csv : Emisiones por empresa con concentraciones y
                                        puntajes de riesgo cancerígeno (Métodos 1 y 2).
  - {year}_company_noncancer_score.csv: Puntajes de riesgo no cancerígeno por empresa.

Uso (GNU/Linux)
---------------
  1. Coloca los archivos retc YYYY.xlsx y contable.xlsx en la misma carpeta.
  2. Ejecuta:
       python3 process_retc_nl.py --input-dir /ruta/datos --output-dir /ruta/salida
  3. Ajusta MUNICIPALITIES o los factores de dilución si es necesario.

Metodología de cálculo de riesgo
----------------------------------
  Método 1 – Riesgo estándar (chimeneas de ~9 m a 100 m de distancia):
    Concentración (µg/m³) = Emisión (kg/año) × 2.20462 [lb/kg] / 365 [días/año] × DF1
    Riesgo cancerígeno    = Concentración × IUR

  Método 2 – Puntaje de riesgo cancerígeno (enfoque simplificado):
    Concentración (µg/m³) = Emisión (kg/año) × 2.20462 × 7700 × DF2
    Puntaje cancerígeno   = Concentración × IUR

  Puntaje no cancerígeno:
    Puntaje = Emisión (lb/año) × DF2 / REL

Parámetros por defecto
-----------------------
  DF1 (dilution_factor)  = 2.47  (µg/m³ por lb/día)
  DF2 (dilution_factor2) = 0.25  ((µg/m³) / (lb/año))

Notas técnicas
--------------
  - El archivo RETC debe contener hojas "Datos Generales" y "Emisiones y Transferencias"
    con encabezados a partir de la fila 10 (header=9 en pandas).
  - Las emisiones al aire se leen de la columna 'AIRE' (kg/año).
  - Las coordenadas UTM (Zona 14N) se convierten de metros a kilómetros en la salida.
  - IUR  : Inhalation Unit Risk  (riesgo por µg/m³ de exposición vitalicia).
  - REL  : Reference Exposure Level (µg/m³, umbral de efectos no cancerígenos).
  - Si una sustancia no tiene REL en contable.xlsx, su puntaje no cancerígeno = NaN.

Dependencias
------------
  Python >= 3.8
  pandas >= 1.3
  openpyxl >= 3.0  (para leer archivos .xlsx)

Autor
-----
  Proyecto de análisis ambiental – Nuevo León, México.
"""

import argparse
import os
import re
from pathlib import Path
from typing import Dict, Any, Optional, Tuple

import pandas as pd


# ---------------------------------------------------------------------------
# Constantes globales
# ---------------------------------------------------------------------------

#: Factor de conversión de kilogramos a libras
KG_TO_LB: float = 2.20462262185

#: Municipios que forman la Zona Metropolitana de Monterrey (ZMM).
#: Se incluyen variantes con/sin tildes para robustez.
MONTERREY_MUNICIPALITIES = {
    'MONTERREY',
    'SAN NICOLAS DE LOS GARZA', 'SAN NICOLÁS DE LOS GARZA',
    'GUADALUPE',
    'APODACA',
    'GENERAL ESCOBEDO', 'ESCOBEDO', 'ESCÓBEDO',
    'SANTA CATARINA',
    'GARCIA', 'GARCÍA',
    'JUAREZ', 'JUÁREZ',
    'SAN PEDRO GARZA GARCIA', 'SAN PEDRO GARZA GARCÍA',
    'PESQUERIA', 'PESQUERÍA',
    'EL CARMEN',
    'CADEREYTA JIMENEZ', 'CADEREYTA JIMÉNEZ',
    'CIÉNEGA DE FLORES', 'CIENEGA DE FLORES',
}

#: Patrones para identificar registros de Nuevo León (sin tildes, sin espacios ni puntos)
NUEVO_LEON_PATTERNS = [
    'NUEVOLEON',
    'NL',
    'NUEVOLEÓN',
    'NUEVOLÉN',
]


# ---------------------------------------------------------------------------
# Carga de datos toxicológicos
# ---------------------------------------------------------------------------

def load_iur_rel(contable_path: Path) -> Dict[str, Dict[str, Any]]:
    """
    Carga los valores IUR (Inhalation Unit Risk) y REL (Reference Exposure Level)
    desde el archivo contable.xlsx.

    El archivo debe tener una hoja con columnas que contengan (sin importar mayúsculas):
      - CAS  : número CAS de la sustancia (puede llamarse "Chemical Abstract..." o "CAS")
      - IUR  : "Inhalation Unit Risk" (riesgo unitario por inhalación, en (µg/m³)⁻¹)
      - REL  : "Chronic Inhalation Reference Exposure Level" (µg/m³)
      - Nombre: columna que contenga "SUBSTANCE"

    Parameters
    ----------
    contable_path : Path
        Ruta al archivo Excel contable.xlsx.

    Returns
    -------
    dict
        Diccionario indexado por número CAS. Cada valor es un dict con claves:
          - 'iur'  : float o None
          - 'rel'  : float o None
          - 'name' : str o None
    """
    data_by_cas: Dict[str, Dict[str, Any]] = {}

    dfc = pd.read_excel(contable_path, sheet_name=0, header=0)
    # Normalizar nombres de columna a mayúsculas para búsqueda robusta
    cols = [str(c).upper() for c in dfc.columns]

    # Identificar columnas por palabras clave
    cas_col: Optional[str] = None
    iur_col: Optional[str] = None
    rel_col: Optional[str] = None
    name_col: Optional[str] = None

    for idx, col_name in enumerate(cols):
        if 'CHEMICAL' in col_name and 'ABSTRACT' in col_name:
            cas_col = dfc.columns[idx]
        if 'INHALATION' in col_name and 'UNIT' in col_name and 'RISK' in col_name:
            iur_col = dfc.columns[idx]
        if 'CHRONIC' in col_name and 'INHALATION' in col_name:
            rel_col = dfc.columns[idx]
        if 'SUBSTANCE' in col_name:
            name_col = dfc.columns[idx]

    # Fallback: buscar columna CAS por nombre corto
    if cas_col is None:
        for idx, col_name in enumerate(cols):
            if 'CAS' in col_name:
                cas_col = dfc.columns[idx]

    # Informar columnas detectadas (útil para diagnóstico)
    print("  Columnas encontradas en contable.xlsx:")
    print(f"    - CAS  : {cas_col}")
    print(f"    - IUR  : {iur_col}")
    print(f"    - REL  : {rel_col}")
    print(f"    - Nombre: {name_col}")

    # Construir diccionario sustancia→{iur, rel, name}
    for row_idx, row in dfc.iterrows():
        # Número CAS
        cas = (
            str(row[cas_col]).strip()
            if cas_col is not None and pd.notna(row[cas_col])
            else None
        )
        if not cas:
            continue

        # IUR – puede ser NaN o no numérico en algunas filas
        iur: Optional[float] = None
        if iur_col is not None and pd.notna(row[iur_col]):
            try:
                iur = float(row[iur_col])
            except (ValueError, TypeError):
                pass

        # REL – ídem
        rel: Optional[float] = None
        if rel_col is not None and pd.notna(row[rel_col]):
            try:
                rel = float(row[rel_col])
            except (ValueError, TypeError):
                pass

        # Nombre de la sustancia
        name: Optional[str] = (
            str(row[name_col]).strip()
            if name_col is not None and pd.notna(row[name_col])
            else None
        )

        data_by_cas[cas] = {'iur': iur, 'rel': rel, 'name': name}

        # Muestra de diagnóstico para las primeras filas con REL
        if row_idx < 10 and rel is not None:
            print(f"    Muestra: CAS={cas}, Nombre={name}, REL={rel}")

    return data_by_cas


# ---------------------------------------------------------------------------
# Procesamiento de un archivo RETC individual
# ---------------------------------------------------------------------------

def _find_utm_columns(df: pd.DataFrame) -> Tuple[Optional[str], Optional[str]]:
    """
    Detecta las columnas UTM X e UTM Y en un DataFrame.

    Primero intenta encontrarlas por nombre (buscando 'UTM' + 'X'/'Y').
    Si no las encuentra, usa las columnas en la posición 9 y 10 (J y K en Excel).

    Parameters
    ----------
    df : pd.DataFrame
        DataFrame de la hoja "Datos Generales".

    Returns
    -------
    tuple
        (nombre_columna_utmx, nombre_columna_utmy) – cualquiera puede ser None.
    """
    utmx: Optional[str] = None
    utmy: Optional[str] = None

    for col in df.columns:
        col_upper = str(col).upper()
        if 'UTM' in col_upper and 'X' in col_upper:
            utmx = col
        elif 'UTM' in col_upper and 'Y' in col_upper:
            utmy = col

    # Fallback por posición (columnas J=9, K=10 en los archivos RETC estándar)
    if utmx is None and len(df.columns) > 9:
        utmx = df.columns[9]
    if utmy is None and len(df.columns) > 10:
        utmy = df.columns[10]

    return utmx, utmy


def _load_retc_sheets(
    retc_path: Path,
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    Lee las hojas "Datos Generales" y "Emisiones y Transferencias" de un archivo RETC.

    Los encabezados en los archivos RETC comienzan en la fila 10 (índice 9 en base-0).

    Parameters
    ----------
    retc_path : Path
        Ruta al archivo .xlsx del RETC.

    Returns
    -------
    tuple
        (datos_generales_df, emisiones_df)
    """
    datos_generales = pd.read_excel(retc_path, sheet_name='Datos Generales', header=9)
    emisiones = pd.read_excel(retc_path, sheet_name='Emisiones y Transferencias', header=9)
    return datos_generales, emisiones


def _build_emissions_table(emisiones: pd.DataFrame) -> pd.DataFrame:
    """
    Extrae y limpia la tabla de emisiones al aire.

    Se seleccionan las columnas NRA, CAS, SUSTANCIA y AIRE, se convierten los valores
    a numérico y se eliminan filas sin CAS válido.

    Parameters
    ----------
    emisiones : pd.DataFrame
        Hoja "Emisiones y Transferencias" del archivo RETC.

    Returns
    -------
    pd.DataFrame
        Tabla con columnas: NRA, CAS, SUSTANCIA, EMISION_KG_ANIO.
    """
    df = emisiones[['NRA', 'CAS', 'SUSTANCIA', 'AIRE']].copy()
    df.rename(columns={'AIRE': 'EMISION_KG_ANIO'}, inplace=True)

    df['CAS'] = df['CAS'].astype(str).str.strip()
    # Eliminar filas sin CAS (nan, none, cadena vacía)
    df = df[~df['CAS'].str.lower().isin(['nan', 'none', ''])]
    df['EMISION_KG_ANIO'] = pd.to_numeric(df['EMISION_KG_ANIO'], errors='coerce').fillna(0)

    return df


def _build_company_table(
    datos_generales: pd.DataFrame,
    utmx_col: Optional[str],
    utmy_col: Optional[str],
) -> pd.DataFrame:
    """
    Construye la tabla de empresas con columnas estandarizadas y coordenadas UTM en km.

    Detecta variaciones en los nombres de columna entre años (p.ej. ESTADO vs ENTIDAD)
    y añade las columnas UTMX_KM y UTMY_KM convirtiendo de metros a kilómetros.

    Parameters
    ----------
    datos_generales : pd.DataFrame
        Hoja "Datos Generales" del archivo RETC.
    utmx_col : str or None
        Nombre de la columna UTM X detectada.
    utmy_col : str or None
        Nombre de la columna UTM Y detectada.

    Returns
    -------
    pd.DataFrame
        Tabla con columnas: NRA, NOMBRE, ESTADO, MUNICIPIO, LATITUD, LONGITUD,
        UTMX (opcional), UTMY (opcional), UTMX_KM, UTMY_KM.
    """
    # Detectar variantes de nombre de columna según año del archivo
    col_estado = 'ESTADO' if 'ESTADO' in datos_generales.columns else 'ENTIDAD'
    col_lat = next(
        (c for c in datos_generales.columns if 'LATIT' in str(c).upper()), None
    )
    col_lon = next(
        (c for c in datos_generales.columns if 'LONGIT' in str(c).upper()), None
    )

    # Seleccionar columnas base
    base_cols = ['NRA', 'NOMBRE', col_estado, 'MUNICIPIO', col_lat, col_lon]
    if utmx_col is not None:
        base_cols.append(utmx_col)
    if utmy_col is not None:
        base_cols.append(utmy_col)

    dg = datos_generales[base_cols].copy()

    # Renombrar a nombres estandarizados
    rename_map = {
        col_estado: 'ESTADO',
        col_lat: 'LATITUD',
        col_lon: 'LONGITUD',
    }
    if utmx_col is not None:
        rename_map[utmx_col] = 'UTMX'
    if utmy_col is not None:
        rename_map[utmy_col] = 'UTMY'
    dg.rename(columns=rename_map, inplace=True)

    # Convertir UTM de metros a kilómetros
    dg['UTMX_KM'] = (
        pd.to_numeric(dg['UTMX'], errors='coerce') / 1000.0
        if 'UTMX' in dg.columns else None
    )
    dg['UTMY_KM'] = (
        pd.to_numeric(dg['UTMY'], errors='coerce') / 1000.0
        if 'UTMY' in dg.columns else None
    )

    return dg


def _filter_nuevo_leon(merged: pd.DataFrame) -> pd.DataFrame:
    """
    Filtra el DataFrame para conservar únicamente registros de Nuevo León y agrega
    la columna ES_ZM_MONTERREY indicando si el municipio pertenece a la ZMM.

    La normalización elimina puntos, espacios y tildes para comparación robusta.

    Parameters
    ----------
    merged : pd.DataFrame
        DataFrame combinado de emisiones + datos generales.

    Returns
    -------
    pd.DataFrame
        Subconjunto de registros de Nuevo León con columna ES_ZM_MONTERREY.
    """
    # Normalizar nombre del estado: sin puntos, sin espacios
    merged['ESTADO_NORM'] = (
        merged['ESTADO']
        .astype(str)
        .str.upper()
        .str.strip()
        .str.replace('.', '', regex=False)
        .str.replace(' ', '', regex=False)
    )

    # Filtrar por patrones exactos o expresión regular (acepta tildes en la Ó)
    mask = (
        merged['ESTADO_NORM'].isin(NUEVO_LEON_PATTERNS)
        | merged['ESTADO_NORM'].str.contains(
            r'NUEVO.*LE[OÓ]N', na=False, regex=True
        )
    )
    nl = merged[mask].copy()

    # Indicador de pertenencia a la Zona Metropolitana de Monterrey
    nl.loc[:, 'MUNICIPIO_NORM'] = nl['MUNICIPIO'].astype(str).str.upper().str.strip()
    nl.loc[:, 'ES_ZM_MONTERREY'] = nl['MUNICIPIO_NORM'].isin(MONTERREY_MUNICIPALITIES)

    return nl


def _calculate_risk_scores(
    pivot: pd.DataFrame,
    toxics: pd.DataFrame,
    contable_data: Dict[str, Dict[str, Any]],
    dilution_factor: float,
    dilution_factor2: float,
) -> Tuple[pd.DataFrame, pd.DataFrame, list]:
    """
    Calcula concentraciones estimadas y puntajes de riesgo cancerígeno y no cancerígeno
    para cada empresa y sustancia tóxica.

    Método 1 – Riesgo estándar (CONC_ugm3, RISK):
        Conc = Emisión_kg/año × KG_TO_LB / 365 × DF1          [µg/m³]
        Riesgo_cancerígeno = Conc × IUR

    Método 2 – Puntaje simplificado (CONC_SCORE, CANCER_SCORE):
        Conc = Emisión_kg/año × KG_TO_LB × 7700 × DF2         [µg/m³]
        Cancer_score = Conc × IUR

    Puntaje no cancerígeno (NONCANCER_SCORE):
        Score = Emisión_lb/año × DF2 / REL
        (NaN si REL no disponible o REL = 0)

    Parameters
    ----------
    pivot : pd.DataFrame
        Tabla pivote con emisiones por empresa (filas) y CAS (columnas).
    toxics : pd.DataFrame
        Catálogo de sustancias con columnas CAS y SUSTANCIA.
    contable_data : dict
        Diccionario {CAS: {iur, rel, name}} cargado desde contable.xlsx.
    dilution_factor : float
        Factor de dilución DF1 (µg/m³ por lb/día).
    dilution_factor2 : float
        Factor de dilución DF2 ((µg/m³)/(lb/año)).

    Returns
    -------
    tuple
        (pivot_con_riesgo, pivot_no_cancer, sustancias_sin_rel)
          - pivot_con_riesgo  : pivot original + columnas de concentración y riesgo cancerígeno.
          - pivot_no_cancer   : DataFrame con columnas de puntaje no cancerígeno.
          - sustancias_sin_rel: lista de CAS (con nombre) que carecen de REL válido.
    """
    id_cols = [
        'NRA', 'NOMBRE', 'ESTADO', 'MUNICIPIO',
        'LATITUD', 'LONGITUD', 'UTMX_KM', 'UTMY_KM',
    ]

    # DataFrame base para puntajes no cancerígenos (solo columnas identificadoras)
    pivot_noncancer = pivot[id_cols].copy()

    cancer_cols_dict: Dict[str, Any] = {}
    noncancer_cols_dict: Dict[str, Any] = {}
    substances_without_rel: list = []

    for cas in toxics['CAS'].unique():
        entry = contable_data.get(cas, {})
        iur: float = entry.get('iur') or 0.0
        rel: Optional[float] = entry.get('rel')

        # ---- Riesgo cancerígeno (Métodos 1 y 2) ----
        if cas in pivot.columns:
            emission_kg = pivot[cas].astype(float)

            # Método 1: concentración diaria promedio
            conc_m1 = emission_kg * KG_TO_LB / 365.0 * dilution_factor
            cancer_cols_dict[f'CONC_ugm3_{cas}'] = conc_m1
            cancer_cols_dict[f'RISK_{cas}'] = conc_m1 * iur

            # Método 2: puntaje simplificado
            conc_m2 = emission_kg * KG_TO_LB * 7700.0 * dilution_factor2
            cancer_cols_dict[f'CONC_SCORE_{cas}'] = conc_m2
            cancer_cols_dict[f'CANCER_SCORE_{cas}'] = conc_m2 * iur

            # ---- Riesgo no cancerígeno ----
            emission_lb_yr = emission_kg * KG_TO_LB

            if rel is not None and rel > 0:
                noncancer_cols_dict[f'NONCANCER_SCORE_{cas}'] = (
                    emission_lb_yr * dilution_factor2 / rel
                )
            else:
                # Sustancia sin REL: puntaje indefinido → NaN
                noncancer_cols_dict[f'NONCANCER_SCORE_{cas}'] = pd.Series(
                    [float('nan')] * len(pivot), index=pivot.index
                )
                # Registrar para reporte diagnóstico
                subst_name = entry.get(
                    'name',
                    toxics.loc[toxics['CAS'] == cas, 'SUSTANCIA'].iloc[0]
                    if cas in toxics['CAS'].values else 'Desconocida',
                )
                substances_without_rel.append(f"{cas} ({subst_name})")
        else:
            # Sustancia registrada pero sin emisión reportada por ninguna empresa
            cancer_cols_dict[f'CONC_ugm3_{cas}'] = 0.0
            cancer_cols_dict[f'RISK_{cas}'] = 0.0
            cancer_cols_dict[f'CONC_SCORE_{cas}'] = 0.0
            cancer_cols_dict[f'CANCER_SCORE_{cas}'] = 0.0
            noncancer_cols_dict[f'NONCANCER_SCORE_{cas}'] = 0.0

    # Agregar columnas calculadas a los DataFrames
    for col_name, col_data in cancer_cols_dict.items():
        pivot[col_name] = col_data

    for col_name, col_data in noncancer_cols_dict.items():
        pivot_noncancer[col_name] = col_data

    # Totales de riesgo cancerígeno (Método 1 y 2)
    pivot['RISK_TOTAL'] = pivot[
        [c for c in pivot.columns if c.startswith('RISK_')]
    ].sum(axis=1)

    pivot['CANCER_SCORE_TOTAL'] = pivot[
        [c for c in pivot.columns if c.startswith('CANCER_SCORE_')]
    ].sum(axis=1)

    # Total de puntaje no cancerígeno (ignorando NaN para no penalizar falta de REL)
    pivot_noncancer['NONCANCER_SCORE_TOTAL'] = pivot_noncancer[
        [c for c in pivot_noncancer.columns if c.startswith('NONCANCER_SCORE_')]
    ].sum(axis=1, skipna=True)

    return pivot, pivot_noncancer, substances_without_rel


def process_file(
    retc_path: Path,
    contable_data: Dict[str, Dict[str, Any]],
    output_dir: str,
    dilution_factor: float = 2.47,
    dilution_factor2: float = 0.25,
) -> Path:
    """
    Procesa un archivo RETC anual y genera los cuatro CSV de salida.

    Flujo de procesamiento:
      1. Leer hojas "Datos Generales" y "Emisiones y Transferencias".
      2. Construir tabla de empresas (con UTM en km) y tabla de emisiones al aire.
      3. Combinar (merge) por NRA y filtrar únicamente Nuevo León.
      4. Identificar municipios de la Zona Metropolitana de Monterrey.
      5. Calcular concentraciones y puntajes de riesgo cancerígeno/no cancerígeno.
      6. Guardar los CSV en output_dir/output_{year}/.

    Parameters
    ----------
    retc_path : Path
        Ruta al archivo retc YYYY.xlsx.
    contable_data : dict
        Diccionario toxicológico {CAS: {iur, rel, name}} de contable.xlsx.
    output_dir : str
        Directorio base para guardar los resultados.
    dilution_factor : float, optional
        Factor de dilución DF1 (µg/m³ por lb/día). Por defecto 2.47.
    dilution_factor2 : float, optional
        Factor de dilución DF2 ((µg/m³)/(lb/año)). Por defecto 0.25.

    Returns
    -------
    Path
        Ruta a la carpeta de salida generada (output_dir/output_{year}/).
    """
    # 1. Leer hojas del archivo RETC
    datos_generales, emisiones = _load_retc_sheets(retc_path)

    # 2. Construir tabla de empresas con coordenadas normalizadas
    utmx_col, utmy_col = _find_utm_columns(datos_generales)
    company_table = _build_company_table(datos_generales, utmx_col, utmy_col)

    # 3. Construir tabla de emisiones al aire
    emissions_table = _build_emissions_table(emisiones)

    # 4. Combinar emisiones con datos de empresa y filtrar Nuevo León
    merged = emissions_table.merge(company_table, on='NRA', how='left')
    merged_nl = _filter_nuevo_leon(merged)

    # 5. Generar catálogos de salida
    toxics = (
        merged_nl[['CAS', 'SUSTANCIA']]
        .drop_duplicates()
        .reset_index(drop=True)
    )
    companies = (
        company_table
        .merge(merged_nl[['NRA']].drop_duplicates(), on='NRA', how='inner')
        .drop_duplicates()
        .reset_index(drop=True)
    )

    # 6. Tabla pivote: filas = empresa, columnas = CAS, valores = emisión kg/año
    pivot_index = [
        'NRA', 'NOMBRE', 'ESTADO', 'MUNICIPIO',
        'LATITUD', 'LONGITUD', 'UTMX_KM', 'UTMY_KM',
    ]
    pivot = merged_nl.pivot_table(
        index=pivot_index,
        columns='CAS',
        values='EMISION_KG_ANIO',
        aggfunc='sum',
        fill_value=0,
    ).reset_index()

    # 7. Calcular riesgo cancerígeno y no cancerígeno
    pivot, pivot_noncancer, substances_without_rel = _calculate_risk_scores(
        pivot, toxics, contable_data, dilution_factor, dilution_factor2
    )

    # 8. Determinar el año del archivo y crear carpeta de salida
    year_match = re.search(r'(\d{4})', retc_path.name)
    year = year_match.group(1) if year_match else retc_path.stem
    out_folder = Path(output_dir) / f'output_{year}'
    out_folder.mkdir(parents=True, exist_ok=True)

    # 9. Guardar CSVs
    companies.to_csv(out_folder / f'{year}_companies.csv', index=False, encoding='utf-8')
    toxics.to_csv(out_folder / f'{year}_toxics.csv', index=False, encoding='utf-8')
    pivot.to_csv(
        out_folder / f'{year}_company_emissions_risk.csv', index=False, encoding='utf-8'
    )
    pivot_noncancer.to_csv(
        out_folder / f'{year}_company_noncancer_score.csv', index=False, encoding='utf-8'
    )

    # 10. Reporte de resultados en consola
    valid_rel_count = sum(
        1
        for cas in toxics['CAS']
        if contable_data.get(cas, {}).get('rel') is not None
        and contable_data.get(cas, {}).get('rel') > 0
    )
    print(f"  - Registros en Nuevo León    : {len(merged_nl)}")
    print(f"  - Empresas únicas            : {len(companies)}")
    print(f"  - Sustancias tóxicas         : {len(toxics)}")
    print(f"  - Sustancias con REL válido  : {valid_rel_count}")

    if substances_without_rel:
        print(
            f"  ⚠  {len(substances_without_rel)} sustancias sin valor REL "
            f"(puntaje no-cáncer = NaN):"
        )
        for subst in substances_without_rel[:5]:
            print(f"      - {subst}")
        if len(substances_without_rel) > 5:
            print(f"      ... y {len(substances_without_rel) - 5} más")

    return out_folder


# ---------------------------------------------------------------------------
# Punto de entrada principal
# ---------------------------------------------------------------------------

def main() -> None:
    """
    Punto de entrada del script. Parsea argumentos, carga contable.xlsx y
    procesa cada archivo RETC encontrado en el directorio de entrada.
    """
    parser = argparse.ArgumentParser(
        description=(
            'Procesa archivos RETC (.xlsx) para calcular riesgo por exposición a '
            'emisiones al aire en Nuevo León, México.'
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            'Ejemplo de uso:\n'
            '  python3 process_retc_nl.py \\\n'
            '      --input-dir /datos/retc \\\n'
            '      --output-dir /datos/retc/resultados \\\n'
            '      --dilution 2.47 \\\n'
            '      --dilution2 0.25'
        ),
    )
    parser.add_argument(
        '--input-dir', '-i', required=True,
        help='Carpeta con los archivos retc YYYY.xlsx y contable.xlsx.',
    )
    parser.add_argument(
        '--output-dir', '-o', required=True,
        help='Carpeta donde se guardarán los CSVs de salida.',
    )
    parser.add_argument(
        '--dilution', '-d', type=float, default=2.47,
        help='Factor de dilución DF1 en µg/m³ por lb/día (default: 2.47).',
    )
    parser.add_argument(
        '--dilution2', '-d2', type=float, default=0.25,
        help='Factor de dilución DF2 en (µg/m³)/(lb/año) (default: 0.25).',
    )
    args = parser.parse_args()

    input_dir = Path(args.input_dir)
    contable_path = input_dir / 'contable.xlsx'

    if not contable_path.exists():
        raise FileNotFoundError(
            f"No se encontró 'contable.xlsx' en: {input_dir}"
        )

    # Cargar datos toxicológicos
    print("Cargando valores IUR y REL desde contable.xlsx...")
    contable_data = load_iur_rel(contable_path)

    substances_with_iur = sum(
        1 for v in contable_data.values()
        if v.get('iur') is not None and v.get('iur') > 0
    )
    substances_with_rel = sum(
        1 for v in contable_data.values()
        if v.get('rel') is not None and v.get('rel') > 0
    )
    print(f"  Total de sustancias cargadas : {len(contable_data)}")
    print(f"  Con valores IUR              : {substances_with_iur}")
    print(f"  Con valores REL              : {substances_with_rel}")
    print(f"  Factor de dilución DF1       : {args.dilution}")
    print(f"  Factor de dilución DF2       : {args.dilution2}")

    # Buscar archivos RETC en el directorio de entrada
    retc_files = [
        f for f in input_dir.iterdir()
        if f.name.lower().startswith('retc') and f.suffix.lower() in ('.xlsx', '.xls')
    ]

    if not retc_files:
        print('No se encontraron archivos RETC en el directorio de entrada.')
        return

    retc_files.sort()
    print(f"\nArchivos RETC encontrados: {len(retc_files)}\n")

    # Procesar cada archivo
    for rf in retc_files:
        print(f"Procesando {rf.name}...")
        try:
            out_path = process_file(
                rf, contable_data, args.output_dir, args.dilution, args.dilution2
            )
            print(f"  ✓ Resultados guardados en: {out_path}\n")
        except Exception as exc:
            print(f"  ✗ Error procesando {rf.name}: {exc}\n")
            import traceback
            traceback.print_exc()


if __name__ == '__main__':
    main()

# data_manager.py
# Gestor de datos de alto rendimiento para el Dashboard APS - Hospital Malvinas
# Convierte archivos Excel pesados a formato Parquet y consolida las atenciones mensuales.

import re
import unicodedata
from pathlib import Path
import pandas as pd
import numpy as np
import streamlit as st

def _norm_doc(v):
    """Normaliza el número de documento extrayendo únicamente los dígitos."""
    if v is None or (isinstance(v, float) and (pd.isna(v) or np.isnan(v))):
        return ""
    if isinstance(v, (int, np.integer)):
        return str(int(v))
    if isinstance(v, float):
        if np.isfinite(v):
            if float(v).is_integer():
                return str(int(v))
            return re.sub(r"\D", "", str(v).split(".")[0])
        return ""
    s = str(v).strip()
    try:
        if re.match(r'^\s*\d+(\.\d+)?e\+\d+\s*$', s, re.I):
            f = float(s); return str(int(round(f)))
    except Exception:
        pass
    if re.match(r'^\d+\.\d+$', s):
        left, right = s.split(".", 1)
        if set(right) == {"0"}:
            return left
    return re.sub(r"\D", "", s)

def get_paths():
    """Retorna las rutas del proyecto en disco G:\\aps-malvinas"""
    base_dir = Path("G:/aps-malvinas")
    if not base_dir.exists():
        # Fallback por si acaso al directorio local actual
        base_dir = Path.cwd()
    return {
        "base_dir": base_dir,
        "char_xlsx": base_dir / "HMHOOV.xlsx",
        "char_parquet": base_dir / "caracterizacion_consolidado.parquet",
        "atenciones_dir": base_dir / "atenciones",
        "atenciones_parquet": base_dir / "atenciones_consolidado.parquet"
    }

def consolidate_caracterizacion(force=False) -> str:
    """
    Lee HMHOOV.xlsx y lo convierte a caracterizacion_consolidado.parquet
    para cargas instantáneas en el dashboard.
    """
    paths = get_paths()
    xlsx_path = paths["char_xlsx"]
    parquet_path = paths["char_parquet"]

    if not xlsx_path.exists():
        return "No se encontró el archivo base HMHOOV.xlsx"

    # Verificar si el parquet ya existe y es más reciente que el xlsx
    if parquet_path.exists() and not force:
        xlsx_mtime = xlsx_path.stat().st_mtime
        parquet_mtime = parquet_path.stat().st_mtime
        if parquet_mtime > xlsx_mtime:
            return "El archivo Parquet de caracterización ya está actualizado."

    try:
        # Cargar excel de caracterización (lectura única y optimizada)
        df = pd.read_excel(xlsx_path)
        
        # Guardar en formato parquet comprimido (muy rápido de leer en el futuro)
        df.to_parquet(parquet_path, compression="snappy", index=False)
        return "¡Consolidación de caracterización exitosa! Creado caracterizacion_consolidado.parquet"
    except Exception as e:
        return f"Error consolidando caracterización: {e}"

def find_header_row(filepath) -> int:
    """
    Escanea las primeras 10 filas de un archivo Excel para detectar dinámicamente 
    el índice de fila que contiene las cabeceras del reporte (ej.: CodiInst o Documento).
    Esto maneja de forma transparente archivos con y sin cabecera institucional.
    """
    try:
        df_temp = pd.read_excel(filepath, nrows=10, header=None)
        for idx, row in df_temp.iterrows():
            row_vals = [str(x).lower().strip() for x in row.values if pd.notna(x)]
            if 'codiinst' in row_vals or 'documento' in row_vals or 'paciente' in row_vals:
                return idx
    except Exception as e:
        print(f"Error detectando cabecera en {Path(filepath).name}: {e}")
    return 5 # Fallback por defecto (fila 6 en excel, 0-indexed index 5)

def consolidate_atenciones(force=False) -> str:
    """
    Busca todos los archivos Excel mensuales en G:\\aps-malvinas\\atenciones,
    los unifica, normaliza el documento del paciente y los guarda en atenciones_consolidado.parquet.
    """
    paths = get_paths()
    atenciones_dir = paths["atenciones_dir"]
    parquet_path = paths["atenciones_parquet"]

    if not atenciones_dir.exists() or not atenciones_dir.is_dir():
        return "No existe la carpeta 'atenciones'."

    xlsx_files = list(atenciones_dir.glob("*.xlsx"))
    xlsx_files = [f for f in xlsx_files if not f.name.startswith("~$")] # Ignorar temporales de Excel

    if not xlsx_files:
        return "No se encontraron archivos Excel en la carpeta 'atenciones'."

    # Verificar si es necesario regenerar
    if parquet_path.exists() and not force:
        latest_xlsx_mtime = max(f.stat().st_mtime for f in xlsx_files)
        parquet_mtime = parquet_path.stat().st_mtime
        if parquet_mtime > latest_xlsx_mtime:
            return "El archivo Parquet de atenciones ya está actualizado."

    dfs = []
    for filepath in xlsx_files:
        try:
            # Detectar fila de cabeceras de forma dinámica
            header_idx = find_header_row(filepath)
            df_m = pd.read_excel(filepath, header=header_idx)
            
            # Limpiar filas completamente vacías
            df_m = df_m.dropna(how='all')
            
            # Identificar columnas relevantes
            doc_col = None
            for col in df_m.columns:
                if str(col).lower().strip() in ['documento', 'documento_paciente', 'num_doc', 'cc', 'identificacion', 'pacientedocumento']:
                    doc_col = col
                    break
            
            if doc_col is None:
                # Fallback por si la columna no se llama Documento exactamente
                # Encontramos la primera que contenga "Documento"
                candidates = [c for c in df_m.columns if 'documento' in str(c).lower()]
                if candidates:
                    doc_col = candidates[0]
            
            if doc_col is not None:
                # Normalizar el documento
                df_m["_doc_norm"] = df_m[doc_col].apply(_norm_doc)
                # Mantener solo registros con documentos válidos
                df_m = df_m[df_m["_doc_norm"] != ""]
            else:
                # Si no encontramos columna de documento, saltamos
                continue

            # Agregar nombre de archivo como columna de origen para referencia del mes
            df_m["Origen_Archivo"] = filepath.stem
            dfs.append(df_m)
            
        except Exception as e:
            print(f"Error procesando archivo {filepath.name}: {e}")
            continue

    if not dfs:
        return "No se pudo procesar ningún archivo de atenciones válido."

    try:
        # Concatenar todos los meses
        df_total = pd.concat(dfs, ignore_index=True)
        
        # Guardar en parquet
        df_total.to_parquet(parquet_path, compression="snappy", index=False)
        return f"¡Unificación exitosa! Consolidados {len(xlsx_files)} meses en atenciones_consolidado.parquet"
    except Exception as e:
        return f"Error consolidando atenciones: {e}"

@st.cache_data(show_spinner=False)
def load_caracterizacion_parquet() -> pd.DataFrame:
    """Carga rápida del parquet de caracterización."""
    paths = get_paths()
    parquet_path = paths["char_parquet"]
    if parquet_path.exists():
        return pd.read_parquet(parquet_path)
    # Fallback al xlsx original si no se ha consolidado
    xlsx_path = paths["char_xlsx"]
    if xlsx_path.exists():
        return pd.read_excel(xlsx_path)
    raise FileNotFoundError("No se encontró ninguna base de datos de caracterización (HMHOOV.xlsx o Parquet).")

@st.cache_data(show_spinner=False)
def load_atenciones_parquet() -> pd.DataFrame:
    """Carga rápida del parquet de atenciones cruzadas."""
    paths = get_paths()
    parquet_path = paths["atenciones_parquet"]
    if parquet_path.exists():
        return pd.read_parquet(parquet_path)
    return pd.DataFrame()

def query_atenciones_by_doc(doc_number: str) -> pd.DataFrame:
    """
    Busca todas las atenciones médicas del hospital asociadas a una cédula normalizada.
    Retorna un DataFrame limpio con el historial del paciente.
    """
    df_aten = load_atenciones_parquet()
    if df_aten.empty or "_doc_norm" not in df_aten.columns:
        return pd.DataFrame()
    
    # Filtrar por documento normalizado
    norm_query = _norm_doc(doc_number)
    matches = df_aten[df_aten["_doc_norm"] == norm_query].copy()
    
    if matches.empty:
        return pd.DataFrame()
    
    # Seleccionar y renombrar columnas interesantes para el usuario
    # Cabeceras típicas encontradas: 'Sede', 'Profesional', 'especialidad', 'estadoCita', 'Fecha', 'Hora', 'Paciente'
    cols_to_show = {}
    for col in matches.columns:
        col_lower = str(col).lower().strip()
        if col_lower == 'fecha':
            cols_to_show[col] = 'Fecha Cita'
        elif col_lower == 'hora':
            cols_to_show[col] = 'Hora'
        elif col_lower == 'profesional':
            cols_to_show[col] = 'Profesional/Médico'
        elif col_lower == 'especialidad':
            cols_to_show[col] = 'Especialidad/Servicio'
        elif col_lower == 'estadocita':
            cols_to_show[col] = 'Estado de la Cita'
        elif col_lower == 'sede':
            cols_to_show[col] = 'Sede'
        elif col_lower == 'origen_archivo':
            cols_to_show[col] = 'Mes Reporte'
            
    if not cols_to_show:
        # Si no encontramos las columnas típicas, retornamos todas
        return matches
        
    res = matches[list(cols_to_show.keys())].rename(columns=cols_to_show)
    
    # Ordenar por fecha de la cita (más reciente primero)
    if 'Fecha Cita' in res.columns:
        try:
            res['Fecha Cita'] = pd.to_datetime(res['Fecha Cita'], errors='coerce')
            res = res.sort_values('Fecha Cita', ascending=False)
            res['Fecha Cita'] = res['Fecha Cita'].dt.strftime('%Y-%m-%d')
        except Exception:
            pass
            
    return res.reset_index(drop=True)

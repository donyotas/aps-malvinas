 # aps2_styled.py
# Dashboard APS 2025 — tarjetas + 3 paneles + planes de cuidado
# Auto-carga: HMHOO.xlsx o el XLSX más reciente en ./ o ./data
# BG = creationDateFormulario (fecha) | BH = creatorFormulario (responsable) | DK = EAPB (código → nombre)
# Mejoras: filtro por fecha, desagregación por sexo, ámbito Urbano/Rural por EBS, grupos etarios solicitados

import io
import re
import base64
import hashlib
import unicodedata
from pathlib import Path
import pandas as pd
import numpy as np
import plotly.express as px
import streamlit as st
import data_manager

st.set_page_config(page_title="APS HMHOO 2025 — DASHBOARD", layout="wide")

# ===== Ocultar menú de Streamlit (hamburguesa, Deploy y footer) =====
st.markdown("""
<style>
#MainMenu {visibility:hidden;}
footer {visibility:hidden;}
header {visibility: visible;}
.stDeployButton {display:none;}
</style>
""", unsafe_allow_html=True)

# ===============================
# Utilidades de carga con cache
# ===============================
@st.cache_data(show_spinner=False)
def load_excel_from_bytes(file_bytes: bytes) -> pd.DataFrame:
    return pd.read_excel(io.BytesIO(file_bytes))

@st.cache_data(show_spinner=False)
def load_excel_from_path(path: str, version_token: str) -> pd.DataFrame:
    return pd.read_excel(path)

def file_md5(path: str) -> str:
    with open(path, "rb") as f:
        return hashlib.md5(f.read()).hexdigest()

def pick_default_excel():
    """
    1) ./HMHOO.xlsx
    2) ./data/HMHOO.xlsx
    3) XLSX más reciente en ./ o ./data (patrón *.xlsx)
       (ignorando temporales de Excel que empiezan por ~$)
    """
    cwd = Path.cwd()
    priority = [cwd / "HMHOO.xlsx", cwd / "data" / "HMHOO.xlsx"]
    for p in priority:
        if p.exists() and not p.name.startswith("~$"):
            return p
    candidates = []
    for folder in [cwd, cwd / "data"]:
        if folder.exists():
            candidates.extend([f for f in folder.glob("*.xlsx") if not f.name.startswith("~$")])
    if candidates:
        return max(candidates, key=lambda p: p.stat().st_mtime)
    return None

# ===============================
# Helpers básicos
# ===============================
def _find_col(df_like, candidates):
    # Normaliza: quita tildes, espacios y guiones bajos; a minúsculas
    def _norm_key(x):
        s = str(x).strip().lower()
        s = "".join(ch for ch in unicodedata.normalize("NFD", s) if unicodedata.category(ch) != "Mn")
        return s.replace(" ", "").replace("_", "")
    cols = {_norm_key(c): c for c in getattr(df_like, "columns", [])}
    for c in candidates:
        key = _norm_key(c)
        if key in cols:
            return cols[key]
    return None

def _norm_doc(v):
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

def _format_doc_for_display(v) -> str:
    s = _norm_doc(v)
    return s if s else (str(v) if pd.notna(v) else "—")

def _norm_text(s: str) -> str:
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    s = str(s)
    s = "".join(ch for ch in unicodedata.normalize("NFD", s) if unicodedata.category(ch) != "Mn")
    s = s.upper()
    s = " ".join(s.split())
    return s

def _norm_code(s: str) -> str:
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    return re.sub(r"[^A-Za-z0-9]", "", str(s)).upper()

def _full_name(row: pd.Series) -> str:
    df_row = row.to_frame().T
    n1_col = _find_col(df_row, ["primernombre","primer_nombre","nombre","nombres"])
    n2_col = _find_col(df_row, ["segundonombre","segundo_nombre"])
    a1_col = _find_col(df_row, ["primerapellido","primer_apellido","apellido","apellidos"])
    a2_col = _find_col(df_row, ["segundoapellido","segundo_apellido"])
    parts = [
        str(row.get(n1_col, "")).strip() if n1_col else "",
        str(row.get(n2_col, "")).strip() if n2_col else "",
        str(row.get(a1_col, "")).strip() if a1_col else "",
        str(row.get(a2_col, "")).strip() if a2_col else "",
    ]
    parts = [p for p in parts if p and p.lower() != "nan"]
    return " ".join(parts)

def _pick_one_location(row: pd.Series) -> str:
    df_row = row.to_frame().T
    loc_col = _find_col(df_row, ["barrio","vereda","barriovereda","barrio_vereda",
                                 "corregimiento","comuna","sector","resguardo"])
    val = row.get(loc_col, None) if loc_col else None
    return str(val) if pd.notna(val) else "—"

def _pick_address(row: pd.Series) -> str:
    df_row = row.to_frame().T
    addr_col = _find_col(df_row, ["direccion","dirección","direccion_residencia","direccionresidencia",
                                  "dir_residencia","dir","direccionhogar","direcciónhogar"])
    val = row.get(addr_col, None) if addr_col else None
    return str(val) if pd.notna(val) else "—"

def _pick_home_location(row: pd.Series) -> str:
    df_row = row.to_frame().T
    uh_col = _find_col(df_row, ["ubicacionHogar","ubicacion_hogar","ubicaciónhogar","ubicación_hogar"])
    val = row.get(uh_col, None) if uh_col else None
    return str(val) if pd.notna(val) else "—"

# === NUEVOS HELPERS: territorio/microterritorio por CÓDIGO (H/J) y rol familia ===
def _pick_territory_code(row: pd.Series) -> str:
    df_row = row.to_frame().T
    col = _find_col(df_row, ["codTerritorio","cod_territorio","id_territorio","territorioid"])
    if not col:
        # Fallback por posición (H = idx 7)
        try:
            return str(row.iloc[7]) if pd.notna(row.iloc[7]) else "—"
        except Exception:
            return "—"
    val = row.get(col, None)
    return str(val) if pd.notna(val) else "—"

def _pick_microterritory_code(row: pd.Series) -> str:
    df_row = row.to_frame().T
    col = _find_col(df_row, ["codMicroterritorio","cod_microterritorio","id_microterritorio","microterritorioid"])
    if not col:
        # Fallback por posición (J = idx 9)
        try:
            return str(row.iloc[9]) if pd.notna(row.iloc[9]) else "—"
        except Exception:
            return "—"
    val = row.get(col, None)
    return str(val) if pd.notna(val) else "—"

ROLE_MAP = {
    1: "Jefe de familia",
    2: "Cónyuge",
    3: "Hijo(a)",
    4: "Hermano(a)",
    5: "Padre o madre",
    6: "Otro",
}

def _pick_family_role(row: pd.Series) -> str:
    df_row = row.to_frame().T
    col = _find_col(df_row, ["rolEnLaFamilia","rol_familia","rolfamilia","rol en la familia","parentesco","rol"])
    if not col:
        return "—"
    v = pd.to_numeric(row.get(col, np.nan), errors="coerce")
    if pd.isna(v):
        return "—"
    return ROLE_MAP.get(int(v), "—")

def _pick_char_date(row: pd.Series) -> str:
    df_row = row.to_frame().T
    date_col = _find_col(df_row, ["creationDateFormulario"]) or _find_col(df_row, [
        "fechaCaracterizacion","fechaCaracterización","fecha_caracterizacion",
        "fecha_atencion","fecha_atención","fechaRegistro","fecha_registro","fecha"
    ])
    val = row.get(date_col, None) if date_col else None
    if pd.isna(val):
        return "—"
    try:
        dt = pd.to_datetime(val, errors="coerce", dayfirst=True)
        return dt.strftime("%Y-%m-%d") if pd.notna(dt) else str(val)
    except Exception:
        return str(val)

def make_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8-sig")

def make_xlsx_bytes(df: pd.DataFrame) -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Resultado")
        ws = writer.sheets["Resultado"]
        for i, col in enumerate(df.columns):
            try:
                max_len = int(df[col].astype(str).str.len().max())
            except Exception:
                max_len = 12
            width = max(12, min(42, max_len + 2))
            ws.set_column(i, i, width)
    buffer.seek(0)
    return buffer.getvalue()

def _printable_html(df: pd.DataFrame, title="Consulta"):
    table_html = df.to_html(index=False, border=0)
    return f"""
<!doctype html>
<html>
<head>
<meta charset="utf-8" />
<title>{title}</title>
<style>
  body{{font-family: Arial, sans-serif; margin:24px; color:#111}}
  h1{{font-size:18px; margin-bottom:10px}}
  .meta{{color:#555; font-size:12px; margin-bottom:12px}}
  table{{border-collapse: collapse; width:100%}}
  th,td{{border:1px solid #e5e7eb; padding:8px; text-align:left; font-size:13px}}
  th{{background:#f8fafc}}
</style>
</head>
<body>
<h1>{title}</h1>
<div class="meta">Generado por APS HMHOO 2025</div>
{table_html}
<script>window.print()</script>
</body>
</html>
"""

@st.cache_data(show_spinner=False)
def _build_fullname_series(df_for_names: pd.DataFrame, _version_token: str) -> pd.Series:
    cols = df_for_names.columns.str.lower()
    has_split = any(c in cols for c in [
        "primernombre","primer_nombre","nombres","nombre",
        "primerapellido","primer_apellido","apellidos","apellido"
    ])
    if has_split:
        def pick(df, cands):
            col = _find_col(df, cands)
            return df[col].astype(str) if col else ""
        n1 = pick(df_for_names, ["primernombre","primer_nombre","nombre","nombres"])
        n2 = pick(df_for_names, ["segundonombre","segundo_nombre"])
        a1 = pick(df_for_names, ["primerapellido","primer_apellido","apellido","apellidos"])
        a2 = pick(df_for_names, ["segundoapellido","segundo_apellido"])
        full = (n1.fillna("") + " " + n2.fillna("") + " " + a1.fillna("") + " " + a2.fillna("")).str.strip()
    else:
        full = df_for_names.apply(_full_name, axis=1)
    return full.map(_norm_text)

# ====== EAPB mapping ======
def _parse_eapb_from_pdf(file_bytes: bytes) -> dict:
    try:
        import pdfplumber
    except Exception:
        st.warning("Para leer EAPB desde PDF instala: pip install pdfplumber")
        return {}
    mapping = {}
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            for line in [l.strip() for l in text.split("\n") if l.strip()]:
                m = re.search(r"\b([A-Za-z0-9\-]+)\s*[-–—:]\s*(.+)$", line)
                if not m:
                    m = re.search(r"\b(CODIGO|CÓDIGO)\s*[:\-]?\s*([A-Za-z0-9\-]+).{0,6}(NOMBRE|RAZON\s*SOCIAL)\s*[:\-]?\s*(.+)$", line, flags=re.I)
                    if m:
                        code, name = m.group(2), m.group(4)
                    else:
                        m2 = re.search(r"^\s*([A-Za-z0-9\-]{2,})\s{2,}(.+)$", line)
                        if not m2:
                            continue
                        code, name = m2.group(1), m2.group(2)
                else:
                    code, name = m.group(1), m.group(2)
                code_n = _norm_code(code)
                name_n = str(name).strip()
                if code_n and name_n:
                    mapping[code_n] = name_n
    return mapping

def _load_eapb_mapping(uploaded_file) -> dict:
    if uploaded_file is None:
        return {}
    name = uploaded_file.name.lower()
    raw = uploaded_file.read()
    bio = io.BytesIO(raw)
    try:
        if name.endswith(".pdf"):
            return _parse_eapb_from_pdf(raw)
        elif name.endswith(".csv"):
            bio.seek(0); dfmap = pd.read_csv(bio)
        elif name.endswith(".xlsx") or name.endswith(".xls"):
            bio.seek(0); dfmap = pd.read_excel(bio)
        else:
            st.warning("Formato no soportado. Usa PDF/CSV/XLS/XLSX.")
            return {}
    except Exception as e:
        st.error(f"No pude leer el archivo de EAPB. Error: {e}")
        return {}
    code_col = _find_col(dfmap, ["codigo","código","cod_eapb","code","id","eapb_codigo","cod"])
    name_col = _find_col(dfmap, ["nombre","eapb","razon_social","razón_social","name","entidad","eps","aseguradora"])
    if not code_col or not name_col:
        if dfmap.shape[1] >= 2:
            code_col, name_col = dfmap.columns[0], dfmap.columns[1]
        else:
            st.warning("No encontré columnas (código/nombre) en el archivo EAPB.")
            return {}
    mapping = {}
    for _, r in dfmap[[code_col, name_col]].dropna().iterrows():
        c = _norm_code(r[code_col]); n = str(r[name_col]).strip()
        if c and n: mapping[c] = n
    return mapping

def _map_eapb_value(raw_value, mapping: dict) -> str:
    if not mapping:
        return str(raw_value) if pd.notna(raw_value) else "—"
    code_n = _norm_code(raw_value)
    if code_n in mapping:
        return mapping[code_n]
    if isinstance(raw_value, str) and len(raw_value) > 4 and " " in raw_value:
        return raw_value
    return str(raw_value) if pd.notna(raw_value) else "—"

# ===============================
# Planes de cuidado (reglas) — con alias seguros
# ===============================
@st.cache_data(show_spinner=False)
def _load_care_rules(uploaded_xlsx) -> pd.DataFrame:
    if uploaded_xlsx is None:
        return pd.DataFrame()
    try:
        df_rules = pd.read_excel(uploaded_xlsx)
    except Exception as e:
        st.error(f"No pude leer el Excel de planes de cuidado: {e}")
        return pd.DataFrame()
    df_rules.columns = df_rules.columns.str.strip().str.lower()
    if 'condicion' in df_rules.columns and 'condicion_python' not in df_rules.columns:
        df_rules['condicion_python'] = df_rules['condicion']
    if 'recomendacion' not in df_rules.columns:
        df_rules['recomendacion'] = ''
    if 'categoria' not in df_rules.columns:
        df_rules['categoria'] = 'General'
    if 'prioridad' not in df_rules.columns:
        df_rules['prioridad'] = 3
    if 'seguimiento_dias' not in df_rules.columns:
        df_rules['seguimiento_dias'] = np.nan
    return df_rules

SAFE_NAMES = {'np': np, 'pd': pd}

def _safe_alias(name: str) -> str:
    s = "".join(ch for ch in unicodedata.normalize("NFD", str(name)) if unicodedata.category(ch) != "Mn")
    s = re.sub(r"\W+", "_", s.strip().lower())
    s = re.sub(r"^(\d)", r"_\1", s)
    return s

def _row_with_aliases(row: pd.Series) -> dict:
    env = {}
    for k, v in row.items():
        env[_safe_alias(k)] = v
    return env

def _eval_condition(cond: str, row: pd.Series) -> bool:
    if not cond or not isinstance(cond, str):
        return False
    env = _row_with_aliases(row)
    env.update(SAFE_NAMES)
    try:
        return bool(eval(cond, {"__builtins__": {}}, env))
    except Exception:
        return False

BUILTIN_RULES = [
    {'categoria':'Gestante','condicion_python':"gestante_flag == 1",
     'recomendacion':"Control prenatal según ruta materno perinatal; educación en signos de alarma; remitir a EAPB/EBS si no está en control; verificar vacunación y suplementación.",
     'prioridad':1,'seguimiento_dias':7},
    {'categoria':'Primera infancia','condicion_python':"pd.to_numeric(edad, errors='coerce') < 5",
     'recomendacion':"Control de crecimiento y desarrollo; valoración nutricional; revisión de esquema de vacunación; educación a cuidadores.",
     'prioridad':2,'seguimiento_dias':30},
    {'categoria':'Adulto mayor','condicion_python':"pd.to_numeric(edad, errors='coerce') >= 60",
     'recomendacion':"Valoración integral (funcional, cognitiva, social); prevención de caídas; revisión de polifarmacia; vacunación.",
     'prioridad':2,'seguimiento_dias':30},
]

def build_plan_for_person(row: pd.Series, rules_df: pd.DataFrame) -> pd.DataFrame:
    out = []
    for _, r in rules_df.iterrows():
        if _eval_condition(r.get('condicion_python', ''), row):
            out.append({'Categoría': r.get('categoria','General'),
                        'Recomendación': r.get('recomendacion',''),
                        'Prioridad': r.get('prioridad',3),
                        'Seguimiento (días)': r.get('seguimiento_dias', np.nan)})
    for r in BUILTIN_RULES:
        if _eval_condition(r.get('condicion_python', ''), row):
            out.append({'Categoría': r['categoria'],
                        'Recomendación': r['recomendacion'],
                        'Prioridad': r['prioridad'],
                        'Seguimiento (días)': r['seguimiento_dias']})
    if not out:
        return pd.DataFrame()
    plan = pd.DataFrame(out).sort_values(['Prioridad']).reset_index(drop=True)
    return plan

# Diccionario de nivel educativo
NIVEL_MAP = {
    1:'Sin escolaridad', 2:'Preescolar', 3:'Primaria completa',
    4:'Primaria incompleta', 5:'Secundaria completa', 6:'Secundaria incompleta',
    7:'Media técnica', 8:'Media académica', 9:'Técnica profesional',
    10:'Tecnológica', 11:'Universitaria', 12:'Universitaria', 13:'Posgrado'
}

# ===============================
# Estilos (CSS)
# ===============================
STYLES = """
<style>
.reportview-container, .main, .block-container { background: #f7fafc; }
.block-container { padding-top: 1.2rem; max-width: 96%; }
.headerbar{display:flex;justify-content:space-between;align-items:center;
background:#ffffff;border:1px solid #e5e7eb;border-radius:14px;padding:16px 16px;
box-shadow:0 1px 3px rgba(0,0,0,.06);margin-bottom:12px}
.header-left{display:flex;gap:16px;align-items:center}
.header-titles b{font-size:18px;color:#111827}
.header-sub{font-size:12px;color:#6b7280;line-height:1.3}
.card-row{display:flex;gap:14px;flex-wrap:wrap;margin-bottom:10px}
.card{flex:1 1 220px;display:flex;align-items:center;gap:12px;padding:16px 18px;
border:1px solid #e5e7eb;border-radius:16px;background:#ffffff;
box-shadow:0 1px 3px rgba(0,0,0,.06)}
.icon{font-size:26px;line-height:1}
.label{font-size:12px;color:#6b7280}
.value{font-size:22px;font-weight:800}
.sub{font-size:11px;color:#9ca3af}
.badge{display:inline-block;padding:2px 8px;border-radius:999px;background:#f3f4f6;color:#111827;font-size:11px}
.section-title{font-size:15px;font-weight:700;margin:8px 0 0 2px;color:#111827}
.panel{background:#ffffff;border:1px solid #e5e7eb;border-radius:16px;
padding:6px 10px 2px 10px;box-shadow:0 1px 3px rgba(0,0,0,.06)}
.quick .stButton > button { width:100%; padding:12px 0; border-radius:10px; font-weight:700; }
.logo-img{height:54px;object-fit:contain}
</style>
"""
st.markdown(STYLES, unsafe_allow_html=True)

# ===============================
# Datos base — Carga optimizada con data_manager y Parquet
# ===============================
st.sidebar.title("Administración de Datos")

# Botón para regenerar / optimizar datos
if st.sidebar.button("🔁 Optimizar y Consolidar Datos (Excel → Parquet)"):
    st.cache_data.clear()
    with st.spinner("Procesando archivos Excel y convirtiendo a Parquet..."):
        msg_char = data_manager.consolidate_caracterizacion(force=True)
        msg_aten = data_manager.consolidate_atenciones(force=True)
    st.sidebar.success(f"{msg_char}\n{msg_aten}")
    st.rerun()

st.sidebar.markdown("---")
st.sidebar.title("Filtros")

df = pd.DataFrame()
source_name = "—"
BASE_VERSION = "0"

try:
    # Cargar caracterización consolidada
    df = data_manager.load_caracterizacion_parquet()
    source_name = "caracterizacion_consolidado.parquet"
    st.sidebar.caption("⚡ Datos de caracterización cargados instantáneamente (Parquet)")
except Exception as e:
    st.sidebar.warning("No se encontró la base de datos optimizada en Parquet.")
    source_path = pick_default_excel()
    if source_path and source_path.exists():
        BASE_VERSION = file_md5(str(source_path))
        df = load_excel_from_path(str(source_path), BASE_VERSION)
        source_name = source_path.name
        st.sidebar.caption(f"Cargado desde Excel: {source_name}")
    else:
        st.error("No se encontró ninguna base de datos de caracterización (HMHOOV.xlsx o Parquet).")
        st.stop()

# Mostrar estado de atenciones médicas
if Path("G:/aps-malvinas/atenciones_consolidado.parquet").exists():
    st.sidebar.caption("🏥 Base de atenciones médicas cargada y cruzada")
else:
    st.sidebar.caption("⚠️ Sin base de atenciones médicas. Colócalas en 'atenciones' y presiona Optimizar.")

# ===============================
# Encabezado con identidad visual
# ===============================
try:
    BASE_DIR = Path(__file__).resolve().parent
except Exception:
    BASE_DIR = Path.cwd()

def find_logo_path():
    patterns = [
        "LOGO.png","LOGO .png","logo.png","logo .png",
        "LOGO.jpg","logo.jpg","LOGO.jpeg","logo.jpeg",
        "LOGO.webp","logo.webp",
        "assets/LOGO.png","assets/LOGO .png","assets/logo.png","assets/logo .png",
        "assets/LOGO.jpg","assets/logo.jpg","assets/LOGO.jpeg","assets/logo.jpeg",
        "assets/LOGO.webp","assets/logo.webp",
        "*logo*.png","*logo*.jpg","*logo*.jpeg","*logo*.webp",
        "assets/*logo*.png","assets/*logo*.jpg","assets/*logo*.jpeg","assets/*logo*.webp",
    ]
    for p in patterns[:10]:
        path = BASE_DIR / p
        if path.exists():
            return path
    for pat in patterns[10:]:
        matches = list((BASE_DIR).glob(pat))
        if matches:
            return matches[0]
    return None

def img_to_data_uri(p: Path) -> str:
    ext = p.suffix.lower().lstrip(".") or "png"
    mime = f"image/{'jpeg' if ext in ['jpg','jpeg'] else ('webp' if ext=='webp' else 'png')}"
    b64 = base64.b64encode(p.read_bytes()).decode()
    return f"data:{mime};base64,{b64}"

logo_file = find_logo_path()
logo_tag = '<div style="width:56px;height:56px;background:#f3f4f6;border-radius:10px"></div>'
if logo_file:
    data_uri = img_to_data_uri(logo_file)
    logo_tag = f'<img src="{data_uri}" alt="Logo" style="height:56px;object-fit:contain" />'

st.markdown(f"""
<div style="
  display:flex;justify-content:space-between;align-items:center;
  background:#fff;border:1px solid #e5e7eb;border-radius:14px;
  padding:14px 16px;margin-bottom:12px;box-shadow:0 1px 3px rgba(0,0,0,.06)
">
  <div style="display:flex;align-items:center;gap:12px">{logo_tag}</div>
  <div style="display:flex;flex-direction:column;align-items:flex-end;text-align:right;gap:2px">
    <div style="font-weight:700;font-size:14px;color:#111827">Atención Primaria en Salud</div>
    <div style="font-size:13px;color:#374151">Ing. Carlos Fernando Franco Monje</div>
  </div>
</div>
""", unsafe_allow_html=True)

st.caption(f"Caracterizaciones — Fuente: {source_name}")

# ===============================
# Normalizaciones de la base
# ===============================
# Edad
if 'edad' in df.columns:
    df['edad'] = pd.to_numeric(df['edad'], errors='coerce')

# Gestantes
if 'esGestante' in df.columns:
    df['gestante_flag'] = (pd.to_numeric(df['esGestante'], errors='coerce') == 1).astype(int)
elif 'gestantes' in df.columns:
    df['gestante_flag'] = (pd.to_numeric(df['gestantes'], errors='coerce') == 1).astype(int)
else:
    df['gestante_flag'] = 0

# (Opcional) Lactancia: detecta si existe alguna columna equivalente
col_lact = _find_col(df, ["lactante","lactancia","esLactante","lactando","lactante_flag"])
if col_lact:
    df['lactante_flag'] = (pd.to_numeric(df[col_lact], errors='coerce') == 1).astype(int)
else:
    df['lactante_flag'] = 0  # si no hay dato, quedará 0 y se mostrará nota

# Nivel educativo
if 'nivelEducativo' in df.columns:
    df['nivelEducativo'] = df['nivelEducativo'].map(NIVEL_MAP).fillna(df['nivelEducativo'])

# Pertenencia étnica
if 'pertenenciaEtnica' in df.columns:
    ETNIA_MAP = {1:'Indígena',2:'ROM (Gitano)',3:'Raizal',4:'Palenquero',5:'Negro/Afrocolombiano',6:'Otra',7:'Ninguna'}
    df['pertenenciaEtnica'] = pd.to_numeric(df['pertenenciaEtnica'], errors='coerce').map(ETNIA_MAP).fillna('No reportado')

# Columnas claves
col_ebs = _find_col(df, ['nroIdentificacionEBS','ebs','id_territorio','idEBS'])
col_familia = _find_col(df, ['parentglobFamilia','idFamilia','familia','grupoFamiliar'])
col_id = 'globalid' if 'globalid' in df.columns else df.columns[0]  # Código de ficha
col_sexo = _find_col(df, ['sexo','género','genero','sexo_biologico','sex'])

# ===== Vectorización de fecha base para filtros y ordenamientos
date_col = _find_col(df, ["creationDateFormulario"]) or _find_col(df, [
    "fechaCaracterizacion","fechaCaracterización","fecha_caracterizacion",
    "fecha_atencion","fecha_atención","fechaRegistro","fecha_registro","fecha"
])
if date_col:
    df['_dt'] = pd.to_datetime(df[date_col], errors='coerce', dayfirst=True)
else:
    df['_dt'] = pd.NaT

# ===== Ámbito Urbano/Rural por EBS (000–016 Urbano, 017–027 Rural)
def _ambito_from_ebs(x) -> str:
    if pd.isna(x):
        return "No definido"
    s = str(x)
    m = re.search(r'(\d+)', s)
    if not m:
        return "No definido"
    n = int(m.group(1))
    if 0 <= n <= 16:
        return "Urbano"
    if 17 <= n <= 27:
        return "Rural"
    return "No definido"

if col_ebs:
    df['Ámbito'] = df[col_ebs].apply(_ambito_from_ebs)
else:
    df['Ámbito'] = "No definido"

# ===== Sexo normalizado (M/F/No reportado)
def _sexo_norm(v):
    s = str(v).strip().lower()
    if s in ["m","masculino","h","hombre","male","masc"]:
        return "Masculino"
    if s in ["f","femenino","mujer","female","feme","fem"]:
        return "Femenino"
    return "No reportado"

if col_sexo:
    df['Sexo'] = df[col_sexo].apply(_sexo_norm)
else:
    df['Sexo'] = "No reportado"

# ===== Nombre completo cacheado
try:
    df["_nombre_completo_norm"] = _build_fullname_series(df, BASE_VERSION)
except Exception:
    df["_nombre_completo_norm"] = ""

# ===== Grupos etarios solicitados
def _grupo_etario(edad):
    e = pd.to_numeric(edad, errors='coerce')
    if pd.isna(e):
        return "Sin dato"
    e = int(e)
    if e < 0:
        return "Sin dato"
    if 0 <= e <= 5:
        return "Primera infancia (0–5)"
    if 6 <= e <= 11:
        return "Infancia (6–11)"
    if 12 <= e <= 17:
        return "Adolescencia (12–17)"
    if 18 <= e <= 28:
        return "Juventud (18–28)"
    if 29 <= e <= 59:
        return "Adultez (29–59)"
    if e >= 60:
        return "Persona mayor (60+)"
    return "Sin dato"

df['Grupo etario'] = df['edad'].apply(_grupo_etario) if 'edad' in df.columns else "Sin dato"

# ===== Conjunto Gestantes y Lactantes
df['Gestantes y lactantes'] = np.where((df['gestante_flag'] == 1) | (df['lactante_flag'] == 1), 1, 0)

# ===============================
# Filtros + Cargas auxiliares (EAPB y Planes de Cuidado)
# ===============================
# Filtro EBS
if col_ebs:
    ebss = df[col_ebs].dropna().astype(str).unique()
    selected_ebs = st.sidebar.multiselect("Seleccione EBS", ebss, default=ebss)
else:
    st.sidebar.info("No se encontró la columna de EBS ('nroIdentificacionEBS').")
    selected_ebs = None

# Filtro Ámbito
ambitos = df['Ámbito'].unique().tolist()
sel_amb = st.sidebar.multiselect("Ámbito (Urbano/Rural)", ambitos, default=ambitos)

# Filtro Sexo
sexos = df['Sexo'].unique().tolist()
sel_sexo = st.sidebar.multiselect("Sexo", sexos, default=sexos)

# Filtro por fecha (rango)
st.sidebar.markdown("---")
st.sidebar.markdown("### 🗓️ Filtro por fecha de caracterización")
if df['_dt'].notna().any():
    min_date = pd.to_datetime(df['_dt']).min().date()
    max_date = pd.to_datetime(df['_dt']).max().date()
    fecha_rango = st.sidebar.date_input("Rango (desde – hasta)", value=(min_date, max_date))
    if isinstance(fecha_rango, tuple) and len(fecha_rango) == 2:
        f_ini, f_fin = pd.to_datetime(fecha_rango[0]), pd.to_datetime(fecha_rango[1])
    else:
        f_ini, f_fin = pd.to_datetime(min_date), pd.to_datetime(max_date)
else:
    st.sidebar.info("No hay fechas válidas para filtrar.")
    f_ini, f_fin = None, None

st.sidebar.markdown("---")
st.sidebar.markdown("### 🏥 Tabla EAPB (PDF/CSV/Excel)")
eapb_file = st.sidebar.file_uploader("Cargar mapeo EAPB (código → nombre)", type=["pdf","csv","xlsx","xls"])
EAPB_MAP = _load_eapb_mapping(eapb_file) if eapb_file else {}

st.sidebar.markdown("---")
st.sidebar.markdown("### 📋 Planes de cuidado (.xlsx)")
care_file = st.sidebar.file_uploader("Cargar reglas de cuidado", type=["xlsx"], key="care_rules")
CARE_RULES = _load_care_rules(care_file) if care_file else pd.DataFrame()

# Diagnóstico rápido (opcional)
try:
    _terr_detect = _find_col(df, ["codTerritorio","id_territorio"])
    _micro_detect = _find_col(df, ["codMicroterritorio","id_microterritorio"])
    _rol_detect = _find_col(df, ["rolEnLaFamilia","rol_familia","rolfamilia","parentesco","rol"])
    st.sidebar.caption(f"🧭 Territorio(cód): {str(_terr_detect)} | Microterritorio(cód): {str(_micro_detect)} | Rol: {str(_rol_detect)}")
except Exception:
    pass

# Formularios (única instancia)
st.sidebar.markdown("---")
with st.sidebar.form("form_doc"):
    st.markdown("### 🔎 Buscar por cédula")
    doc_query = st.text_input("Número de cédula (solo dígitos):", key="doc_query")
    submitted_doc = st.form_submit_button("Buscar cédula")

with st.sidebar.form("form_name"):
    st.markdown("### 🧍‍♀️🧍‍♂️ Buscar por nombre")
    name_query = st.text_input("Nombre (parcial, sin tildes):", key="name_query")
    submitted_name = st.form_submit_button("Buscar por nombre")

# ===== Aplicar filtros seleccionados a df base =====
df_f = df.copy()

if selected_ebs is not None and len(selected_ebs) > 0 and col_ebs:
    df_f = df_f[df_f[col_ebs].astype(str).isin(selected_ebs)]

if sel_amb:
    df_f = df_f[df_f['Ámbito'].isin(sel_amb)]

if sel_sexo:
    df_f = df_f[df_f['Sexo'].isin(sel_sexo)]

if f_ini is not None and f_fin is not None and df_f['_dt'].notna().any():
    df_f = df_f[(df_f['_dt'] >= f_ini) & (df_f['_dt'] <= f_fin)]

# ===============================
# KPI (tarjetas)
# ===============================
total_personas = len(df_f)
familias = df_f[col_familia].nunique() if col_familia else np.nan
gestantes = int(df_f['gestante_flag'].sum()) if 'gestante_flag' in df_f.columns else np.nan
adultos_mayores = int((df_f['edad'] >= 60).sum()) if 'edad' in df_f.columns else np.nan
menores_5 = int((df_f['edad'] < 5).sum()) if 'edad' in df_f.columns else np.nan

fmt = lambda x: f"{int(x):,}".replace(",", ".") if pd.notna(x) else "—"
cards = [
    {"icon":"👥","label":"Personas caracterizadas","value":fmt(total_personas),"color":"#2563eb","helper":"Total filtrado"},
    {"icon":"🏠","label":"Familias caracterizadas","value":fmt(familias),"color":"#059669","helper":"Únicos en parentglobFamilia"},
    {"icon":"🤰","label":"Gestantes","value":fmt(gestantes),"color":"#dc2626","helper":"Marcadas como 1"},
    {"icon":"🧓","label":"Adultos mayores (60+)","value":fmt(adultos_mayores),"color":"#7c3aed","helper":"Edad ≥ 60"},
    {"icon":"👶","label":"Menores de 5 años","value":fmt(menores_5),"color":"#f59e0b","helper":"Edad < 5"},
]
card_tpl = ("<div class='card'>"
            "<div class='icon'>{icon}</div>"
            "<div><div class='label'>{label}</div>"
            "<div class='value' style='color:{color}'>{value}</div>"
            "<div class='sub'>{helper}</div></div></div>")
st.markdown("<div class='card-row'>" + "".join([card_tpl.format(**c) for c in cards]) + "</div>", unsafe_allow_html=True)

# ===== Helper: listado estándar
def _series_or_fill(df_sub, colname, fill="—"):
    return df_sub[colname] if (colname in df_sub.columns) else pd.Series([fill]*len(df_sub))

def _make_people_listing(df_sub: pd.DataFrame) -> pd.DataFrame:
    if df_sub is None or df_sub.empty:
        return pd.DataFrame()

    doc_col  = _find_col(df_sub, [
        "nroDocumento","numDocumento","numeroDocumento","documento",
        "cedula","cédula","cc","doc","nro_doc","identificacion","identificación"
    ])
    edad_col = _find_col(df_sub, ["edad","edad_en_anios","edad_en_años"])
    date_col_local = _find_col(df_sub, ["creationDateFormulario"]) or _find_col(df_sub, [
        "fechaCaracterizacion","fechaCaracterización","fecha_caracterizacion",
        "fecha_atencion","fecha_atención","fechaRegistro","fecha_registro","fecha"
    ])
    resp_col = _find_col(df_sub, ["creatorFormulario"])
    fam_col  = _find_col(df_sub, ["parentglobFamilia","idFamilia","familia","grupoFamiliar"])
    eapb_col = _find_col(df_sub, ["EAPB","eapb"])
    ebs_col  = _find_col(df_sub, ["nroIdentificacionEBS","ebs","id_territorio","idEBS"])

    def _safe_int(x):
        try:
            v = pd.to_numeric(x, errors="coerce")
            return int(v) if pd.notna(v) else "—"
        except Exception:
            return "—"

    eapb_series = _series_or_fill(df_sub, eapb_col) if eapb_col else pd.Series(["—"] * len(df_sub))
    eapb_series = eapb_series.apply(lambda x: _map_eapb_value(x, EAPB_MAP))

    out = pd.DataFrame({
        "EBS": _series_or_fill(df_sub, ebs_col).astype(str) if ebs_col else pd.Series(["—"]*len(df_sub)),
        "Ámbito": df_sub.get("Ámbito", pd.Series(["—"]*len(df_sub))),
        "Sexo": df_sub.get("Sexo", pd.Series(["No reportado"]*len(df_sub))),
        "Grupo etario": df_sub.get("Grupo etario", pd.Series(["Sin dato"]*len(df_sub))),
        "Nombre completo": df_sub["_nombre_completo_norm"] if "_nombre_completo_norm" in df_sub.columns else df_sub.apply(_full_name, axis=1),
        "No. identificación": _series_or_fill(df_sub, doc_col).apply(_format_doc_for_display) if doc_col else pd.Series(["—"]*len(df_sub)),
        "Edad": _series_or_fill(df_sub, edad_col).map(_safe_int) if edad_col else pd.Series(["—"]*len(df_sub)),
        "Barrio/Vereda": df_sub.apply(_pick_one_location, axis=1),
        "Dirección": df_sub.apply(_pick_address, axis=1),
        "Ubicación del hogar": df_sub.apply(_pick_home_location, axis=1),

        # NUEVOS
        "Territorio (cód.)": df_sub.apply(_pick_territory_code, axis=1),
        "Microterritorio (cód.)": df_sub.apply(_pick_microterritory_code, axis=1),
        "Rol en la familia": df_sub.apply(_pick_family_role, axis=1),
        "Código de ficha": df_sub[col_id].astype(str) if col_id else pd.Series(["—"]*len(df_sub)),
        "Num. Identificación Hogar": df_sub.get("numIdentificacionHogar", pd.Series(["—"]*len(df_sub))),
        "Num. Identificación Familia": df_sub.get("numIdentificacionFamilia", pd.Series(["—"]*len(df_sub))),

        "Fecha de caracterización": df_sub.apply(_pick_char_date, axis=1) if date_col_local else pd.Series(["—"]*len(df_sub)),
        "Responsable (creatorFormulario)": _series_or_fill(df_sub, resp_col).astype(str) if resp_col else pd.Series(["—"]*len(df_sub)),
        "Familia": _series_or_fill(df_sub, fam_col).astype(str) if fam_col else pd.Series(["—"]*len(df_sub)),
        "EAPB": eapb_series.astype(str),
    })

    try:
        out["_dt_sort"] = pd.to_datetime(out["Fecha de caracterización"], errors="coerce", dayfirst=True)
        out = out.sort_values(["_dt_sort","Nombre completo"], ascending=[False, True]).drop(columns=["_dt_sort"])
    except Exception:
        out = out.sort_values(["Nombre completo"], ascending=[True])

    return out

# ===============================
# Tres paneles (gráficos)
# ===============================
c1, c2, c3 = st.columns((1.2, 1, 1))

with c1:
    st.markdown("<div class='section-title'>Personas caracterizadas por EBS</div>", unsafe_allow_html=True)
    with st.container():
        st.markdown("<div class='panel'>", unsafe_allow_html=True)
        if col_ebs:
            personas_ebs = (
                df_f.groupby(col_ebs, dropna=True)
                  .size().reset_index(name='Personas')
                  .sort_values('Personas', ascending=False)
            )
            personas_ebs[col_ebs] = personas_ebs[col_ebs].astype(str)
            fig_ebs = px.bar(
                personas_ebs, x=col_ebs, y='Personas', text='Personas',
                color=col_ebs, template='simple_white',
            )
            fig_ebs.update_traces(textposition='outside')
            fig_ebs.update_layout(
                showlegend=False, height=430,
                margin=dict(l=10,r=10,t=40,b=10),
                xaxis_title="EBS", yaxis_title="Personas",
                uniformtext_minsize=10, uniformtext_mode='hide'
            )
            st.plotly_chart(fig_ebs, use_container_width=True)
        else:
            st.info("Agrega 'nroIdentificacionEBS' para ver este gráfico.")
        st.markdown("</div>", unsafe_allow_html=True)

with c2:
    st.markdown("<div class='section-title'>Distribución por Sexo y Grupo etario</div>", unsafe_allow_html=True)
    st.markdown("<div class='panel'>", unsafe_allow_html=True)
    if 'Sexo' in df_f.columns and 'Grupo etario' in df_f.columns:
        sex_grp = df_f.groupby(['Grupo etario','Sexo']).size().reset_index(name='Personas')
        sex_grp = sex_grp.sort_values(['Grupo etario','Personas'])
        fig_sex = px.bar(
            sex_grp, x='Personas', y='Grupo etario',
            color='Sexo', orientation='h', barmode='group', text='Personas',
            template='simple_white'
        )
        fig_sex.update_traces(textposition='outside', cliponaxis=False)
        fig_sex.update_layout(showlegend=True, height=320,
                              margin=dict(l=10,r=10,t=40,b=10),
                              xaxis_title="Personas", yaxis_title="")
        st.plotly_chart(fig_sex, use_container_width=True)
    else:
        st.info("No se encontró 'Sexo' o 'Grupo etario'.")
    st.markdown("</div>", unsafe_allow_html=True)

with c3:
    st.markdown("<div class='section-title'>Distribución por Ámbito y Grupo etario</div>", unsafe_allow_html=True)
    st.markdown("<div class='panel'>", unsafe_allow_html=True)
    if 'Ámbito' in df_f.columns and 'Grupo etario' in df_f.columns:
        amb_grp = df_f.groupby(['Grupo etario','Ámbito']).size().reset_index(name='Personas')
        amb_grp = amb_grp.sort_values(['Grupo etario','Personas'])
        fig_amb = px.bar(
            amb_grp, x='Personas', y='Grupo etario',
            color='Ámbito', orientation='h', barmode='group', text='Personas',
            template='simple_white'
        )
        fig_amb.update_traces(textposition='outside', cliponaxis=False)
        fig_amb.update_layout(showlegend=True, height=430,
                              margin=dict(l=10,r=10,t=40,b=10),
                              xaxis_title="Personas", yaxis_title="")
        st.plotly_chart(fig_amb, use_container_width=True)
    else:
        st.info("No se encontró 'Ámbito' o 'Grupo etario'.")
    st.markdown("</div>", unsafe_allow_html=True)

# ===============================
# NUEVA SECCIÓN: Listados rápidos (incluye grupos etarios + Gestantes/Lactantes)
# ===============================
st.markdown("<div class='section-title'>Listados rápidos</div>", unsafe_allow_html=True)
st.markdown("<div class='panel quick'>", unsafe_allow_html=True)

b1, b2, b3 = st.columns(3)
b4, b5, b6 = st.columns(3)
b7, _, _ = st.columns(3)

show = st.session_state.get("_quick_show", None)
with b1:
    if st.button("🤰 Gestantes / Lactantes"):
        show = "gest_lact"
with b2:
    if st.button("👶 Primera infancia (0–5)"):
        show = "g_0_5"
with b3:
    if st.button("🧒 Infancia (6–11)"):
        show = "g_6_11"
with b4:
    if st.button("🧑 Adolescencia (12–17)"):
        show = "g_12_17"
with b5:
    if st.button("🧑‍🎓 Juventud (18–28)"):
        show = "g_18_28"
with b6:
    if st.button("🧍 Adultez (29–59)"):
        show = "g_29_59"
with b7:
    if st.button("🧓 Persona mayor (60+)"):
        show = "g_60_plus"
st.session_state["_quick_show"] = show




if show is not None:
    # Subconjunto según botón
    if show == "gest_lact":
        df_sub = df_f[(df_f['gestante_flag'] == 1) | (df_f['lactante_flag'] == 1)].copy()
        titulo = "Madres Gestantes y Lactantes"
        if col_lact is None:
            st.info("Nota: No se encontró columna de lactancia; se muestran solo Gestantes.")
    elif show == "g_0_5":
        df_sub = df_f[df_f['Grupo etario'] == "Primera infancia (0–5)"].copy()
        titulo = "Primera infancia (0–5)"
    elif show == "g_6_11":
        df_sub = df_f[df_f['Grupo etario'] == "Infancia (6–11)"].copy()
        titulo = "Infancia (6–11)"
    elif show == "g_12_17":
        df_sub = df_f[df_f['Grupo etario'] == "Adolescencia (12–17)"].copy()
        titulo = "Adolescencia (12–17)"
    elif show == "g_18_28":
        df_sub = df_f[df_f['Grupo etario'] == "Juventud (18–28)"].copy()
        titulo = "Juventud (18–28)"
    elif show == "g_29_59":
        df_sub = df_f[df_f['Grupo etario'] == "Adultez (29–59)"].copy()
        titulo = "Adultez (29–59)"
    else:
        df_sub = df_f[df_f['Grupo etario'] == "Persona mayor (60+)"].copy()
        titulo = "Persona mayor (60+)"

    if df_sub.empty:
        st.info("No hay registros para mostrar con el filtro seleccionado.")
    else:
        listado_quick = _make_people_listing(df_sub)
        st.markdown(f"**{titulo}** (total: {len(listado_quick)})")
        st.dataframe(listado_quick, use_container_width=True, height=460)

        qc1, qc2, qc3 = st.columns([1,1,1])
        with qc1:
            st.download_button("⬇️ CSV", data=make_csv_bytes(listado_quick),
                               file_name=f"{titulo.lower().replace(' ', '_')}.csv", mime="text/csv")
        with qc2:
            st.download_button("⬇️ Excel", data=make_xlsx_bytes(listado_quick),
                               file_name=f"{titulo.lower().replace(' ', '_')}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with qc3:
            html_q = _printable_html(listado_quick, title=titulo)
            b64q = base64.b64encode(html_q.encode("utf-8")).decode()
            st.markdown(f"<a href='data:text/html;base64,{b64q}' target='_blank'>🖨️ Imprimir / Guardar como PDF</a>", unsafe_allow_html=True)
st.markdown("</div>", unsafe_allow_html=True)

# ===============================
# Buscador por cédula → ficha + PLAN DE CUIDADO + FAMILIA
# ===============================
st.markdown("<div class='section-title'>Buscador por número de cédula</div>", unsafe_allow_html=True)
st.markdown("<div class='panel'>", unsafe_allow_html=True)

doc_col = _find_col(df, [
    "nroDocumento","numDocumento","numeroDocumento","documento",
    "cedula","cédula","cc","doc","nro_doc","identificacion","identificación"
])

if not doc_col:
    st.info("No encontré la columna del número de documento (ej.: 'nroDocumento' o 'cedula').")
else:
    if submitted_doc:
        query = _norm_doc(doc_query)
        if not query:
            st.warning("Escribe un número de cédula válido (solo dígitos).")
        else:
            tmp = df_f.copy()  # usa el df filtrado por fecha/sexo/ámbito/EBS
            tmp["_doc_norm"] = tmp[doc_col].apply(_norm_doc)
            matches = tmp[tmp["_doc_norm"] == query].copy()

            if matches.empty:
                st.error("No se encontraron personas con ese número de cédula.")
            else:
                date_col_m = _find_col(matches, [
                    "creationDateFormulario",
                    "fechaCaracterizacion","fechaCaracterización","fecha_caracterizacion",
                    "fecha_atencion","fecha_atención","fechaRegistro","fecha_registro","fecha"
                ])
                if date_col_m:
                    matches["_parsed_date"] = pd.to_datetime(matches[date_col_m], errors="coerce", dayfirst=True)
                    matches = matches.sort_values("_parsed_date", ascending=False, na_position="last")
                row = matches.iloc[0]

                # Ficha
                nombre_completo = _full_name(row)
                no_ident = _format_doc_for_display(row.get(doc_col, "—"))

                edad_col2 = _find_col(matches, ["edad","edad_en_anios","edad_en_años"])
                edad_val = row.get(edad_col2, "—") if edad_col2 else "—"
                try:
                    edad_num = pd.to_numeric(edad_val, errors="coerce")
                    edad = int(edad_num) if pd.notna(edad_num) else "—"
                except Exception:
                    edad = "—"

                barrio_vereda = _pick_one_location(row)
                direccion = _pick_address(row)
                ubic_hogar = _pick_home_location(row)
                territorio_cod = _pick_territory_code(row)
                microterr_cod  = _pick_microterritory_code(row)
                rol_familia    = _pick_family_role(row)
                fecha_carac = _pick_char_date(row)

                ebs_col_m = _find_col(df, ["nroIdentificacionEBS","ebs","id_territorio","idEBS"])
                fam_col_g = _find_col(df, ["parentglobFamilia","idFamilia","familia","grupoFamiliar"])
                resp_col = _find_col(df, ["creatorFormulario"])
                eapb_col = _find_col(df, ["EAPB","eapb"])

                ebs_val  = str(row.get(ebs_col_m, "—"))  if ebs_col_m else "—"
                fam_val  = str(row.get(fam_col_g, "—"))  if fam_col_g else "—"
                resp_val = str(row.get(resp_col, "—")) if resp_col else "—"
                eapb_val = _map_eapb_value(row.get(eapb_col, "—"), EAPB_MAP) if eapb_col else "—"
                cod_ficha = str(row.get(col_id, "—")) if col_id else "—"
                sexo_val = row.get('Sexo', 'No reportado')
                amb_val = row.get('Ámbito', 'No definido')
                grupo_val = row.get('Grupo etario', 'Sin dato')

                res_df = pd.DataFrame([{
                    "Nombre completo": nombre_completo if nombre_completo else "—",
                    "No. identificación": no_ident,
                    "Sexo": sexo_val,
                    "Ámbito": amb_val,
                    "Grupo etario": grupo_val,
                    "Edad": edad,
                    "Barrio/Vereda": barrio_vereda,
                    "Dirección": direccion,
                    "Ubicación del hogar": ubic_hogar,
                    "Territorio (cód.)": territorio_cod,
                    "Microterritorio (cód.)": microterr_cod,
                    "Rol en la familia": rol_familia,
                    "Código de ficha": cod_ficha,
                    "Num. Identificación Hogar": row.get("numIdentificacionHogar", "—"),
                    "Num. Identificación Familia": row.get("numIdentificacionFamilia", "—"),
                    "Fecha de caracterización": fecha_carac,
                    "EBS": ebs_val,
                    "Familia": fam_val,
                    "Responsable (creatorFormulario)": resp_val,
                    "EAPB": eapb_val,
                }])

                st.success("Ficha de la persona consultada:")
                st.dataframe(res_df, use_container_width=True, height=300)

                # ====== HISTORIAL DE ATENCIONES MÉDICAS (CRUCE) ======
                st.markdown("### 🏥 Historial de Atención Prestada (Hospital Malvinas)")
                df_atenciones_cruzadas = data_manager.query_atenciones_by_doc(query)
                if df_atenciones_cruzadas.empty:
                    st.info("No se registraron atenciones médicas para este usuario en las bases de datos cruzadas.")
                else:
                    st.dataframe(df_atenciones_cruzadas, use_container_width=True)
                    ac1, ac2 = st.columns(2)
                    with ac1:
                        st.download_button("⬇️ CSV (Atenciones)", data=make_csv_bytes(df_atenciones_cruzadas),
                                           file_name=f"atenciones_{query}.csv", mime="text/csv")
                    with ac2:
                        st.download_button("⬇️ Excel (Atenciones)", data=make_xlsx_bytes(df_atenciones_cruzadas),
                                           file_name=f"atenciones_{query}.xlsx",
                                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

                # ====== PLAN DE CUIDADO ======
                plan_df = build_plan_for_person(row, CARE_RULES)
                st.markdown("**Sugerencias de Plan de Cuidado**")
                if plan_df.empty:
                    st.info("Sin hallazgos para plan de cuidado con las reglas actuales.")
                else:
                    st.dataframe(plan_df, use_container_width=True, height=220)
                    pc1, pc2, pc3 = st.columns([1,1,1])
                    with pc1:
                        st.download_button("⬇️ CSV (plan)", data=make_csv_bytes(plan_df),
                                           file_name=f"plan_cuidado_{query}.csv", mime="text/csv")
                    with pc2:
                        st.download_button("⬇️ Excel (plan)", data=make_xlsx_bytes(plan_df),
                                           file_name=f"plan_cuidado_{query}.xlsx",
                                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                    with pc3:
                        html_plan = _printable_html(plan_df, title=f"Plan de cuidado — {nombre_completo}")
                        b64p = base64.b64encode(html_plan.encode("utf-8")).decode()
                        st.markdown(f"<a href='data:text/html;base64,{b64p}' target='_blank'>🖨️ Imprimir / Guardar como PDF</a>", unsafe_allow_html=True)

                # ====== Descargas ficha ======
                cdl1, cdl2, cdl3 = st.columns([1,1,1])
                with cdl1:
                    st.download_button("⬇️ CSV (persona)", data=make_csv_bytes(res_df),
                        file_name=f"consulta_{query}.csv", mime="text/csv")
                with cdl2:
                    st.download_button("⬇️ Excel (persona)", data=make_xlsx_bytes(res_df),
                        file_name=f"consulta_{query}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                with cdl3:
                    html_str = _printable_html(res_df, title=f"Consulta cédula {query}")
                    b64 = base64.b64encode(html_str.encode("utf-8")).decode()
                    st.markdown(f"<a href='data:text/html;base64,{b64}' target='_blank'>🖨️ Imprimir / Guardar como PDF</a>", unsafe_allow_html=True)

                # ====== Resumen + Listado FAMILIA ======
                if fam_col_g and pd.notna(row.get(fam_col_g, np.nan)):
                    fam_id = str(row.get(fam_col_g))
                    df_fam = df_f[df_f[fam_col_g].astype(str) == fam_id].copy()

                    edades = pd.to_numeric(df_fam.get('edad', np.nan), errors='coerce')
                    fam_total     = len(df_fam)
                    fam_mayores   = int((edades >= 60).sum()) if 'edad' in df_fam.columns else 0
                    fam_menores5  = int((edades < 5).sum())   if 'edad' in df_fam.columns else 0
                    fam_gestantes = int(pd.to_numeric(df_fam.get('gestante_flag', 0), errors='coerce').fillna(0).sum()) if 'gestante_flag' in df_fam.columns else 0
                    fam_prom_edad = f"{float(edades.mean()):.1f}" if 'edad' in df_fam.columns and pd.notna(edades.mean()) else "—"

                    st.markdown(f"**Resumen de la familia {fam_id}**")
                    fam_cards = [
                        {"icon":"👨‍👩‍👧‍👦","label":"Integrantes","value":fmt(fam_total),"color":"#2563eb","helper":"Personas en la familia"},
                        {"icon":"🧓","label":"Adultos mayores (60+)","value":fmt(fam_mayores),"color":"#7c3aed","helper":"Edad ≥ 60"},
                        {"icon":"👶","label":"Menores de 5 años","value":fmt(fam_menores5),"color":"#f59e0b","helper":"Edad < 5"},
                        {"icon":"🤰","label":"Gestantes","value":fmt(fam_gestantes),"color":"#dc2626","helper":"Marcadas como 1"},
                        {"icon":"📈","label":"Promedio de edad","value":fam_prom_edad,"color":"#059669","helper":"Años"},
                    ]
                    st.markdown("<div class='card-row'>" + "".join([card_tpl.format(**c) for c in fam_cards]) + "</div>", unsafe_allow_html=True)

                    fam_listado = _make_people_listing(df_fam)
                    doc_col_f  = _find_col(df_fam, [
                        "nroDocumento","numDocumento","numeroDocumento","documento",
                        "cedula","cédula","cc","doc","nro_doc","identificacion","identificación"
                    ])
                    if doc_col_f:
                        fam_listado.insert(0, "Consultado",
                            (df_fam[doc_col_f].apply(_norm_doc) == query).map(lambda x: "✅" if x else "").values)
                    else:
                        fam_listado.insert(0, "Consultado", [""] * len(fam_listado))

                    st.markdown(f"**Integrantes de la familia {fam_id}** (total: {len(fam_listado)})")
                    st.dataframe(fam_listado, use_container_width=True, height=460)

                    fc1, fc2, fc3 = st.columns([1,1,1])
                    with fc1:
                        st.download_button("⬇️ CSV (familia)", data=make_csv_bytes(fam_listado),
                                           file_name=f"familia_{fam_id}.csv", mime="text/csv")
                    with fc2:
                        st.download_button("⬇️ Excel (familia)", data=make_xlsx_bytes(fam_listado),
                                           file_name=f"familia_{fam_id}.xlsx",
                                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                    with fc3:
                        html_fam = _printable_html(fam_listado, title=f"Familia {fam_id} — consulta cédula {query}")
                        b64f = base64.b64encode(html_fam.encode("utf-8")).decode()
                        st.markdown(f"<a href='data:text/html;base64,{b64f}' target='_blank'>🖨️ Imprimir / Guardar como PDF</a>", unsafe_allow_html=True)
                else:
                    st.info("No se pudo identificar la familia (columna 'parentglobFamilia' u homólogas ausente o vacía).")

st.markdown("</div>", unsafe_allow_html=True)

# ===============================
# Buscador por nombre → ficha + PLAN DE CUIDADO + FAMILIA (con filtros aplicados)
# ===============================
st.markdown("<div class='section-title'>Buscador por nombre</div>", unsafe_allow_html=True)
st.markdown("<div class='panel'>", unsafe_allow_html=True)

if submitted_name:
    names_norm = _build_fullname_series(df_f, BASE_VERSION)
    q = _norm_text(name_query)
    if not q:
        st.warning("Escribe un nombre o parte del nombre para buscar.")
    else:
        idx = names_norm.str.contains(q, na=False)
        found = df_f[idx].copy()
        if found.empty:
            st.error("No se encontraron personas que coincidan con ese nombre con los filtros actuales.")
        else:
            date_col_n = _find_col(found, [
                "creationDateFormulario",
                "fechaCaracterizacion","fechaCaracterización","fecha_caracterizacion",
                "fecha_atencion","fecha_atención","fechaRegistro","fecha_registro","fecha"
            ])
            if date_col_n:
                found["_parsed_date"] = pd.to_datetime(found[date_col_n], errors="coerce", dayfirst=True)
                found = found.sort_values("_parsed_date", ascending=False, na_position="last")

            doc_col_n  = _find_col(found, [
                "nroDocumento","numDocumento","numeroDocumento","documento",
                "cedula","cédula","cc","doc","nro_doc","identificacion","identificación"
            ])
            ebs_col_n  = _find_col(found, ["nroIdentificacionEBS","ebs","id_territorio","idEBS"])
            fam_col_gn = _find_col(found, ["parentglobFamilia","idFamilia","familia","grupoFamiliar"])
            resp_col_n = _find_col(found, ["creatorFormulario"])
            eapb_col_n = _find_col(found, ["EAPB","eapb"])

            def _rn(row):
                nombre = _full_name(row) or "—"
                doc = _format_doc_for_display(row.get(doc_col_n, "—")) if doc_col_n else "—"
                fecha = _pick_char_date(row)
                return f"{nombre} — {doc} — {fecha}"

            options = found.apply(_rn, axis=1).tolist()
            choice = options[0] if len(options) == 1 else st.selectbox(
                "Se encontraron varias coincidencias. Selecciona una:",
                options, index=0
            )
            row_sel = found.iloc[ options.index(choice) ] if len(found) > 1 else found.iloc[0]

            # Ficha individual seleccionada
            nombre_completo = _full_name(row_sel)
            no_ident = _format_doc_for_display(row_sel.get(doc_col_n, "—")) if doc_col_n else "—"

            edad_col_n2 = _find_col(found, ["edad","edad_en_anios","edad_en_años"])
            edad_val = row_sel.get(edad_col_n2, "—") if edad_col_n2 else "—"
            try:
                edad_num = pd.to_numeric(edad_val, errors="coerce")
                edad = int(edad_num) if pd.notna(edad_num) else "—"
            except Exception:
                edad = "—"

            barrio_vereda = _pick_one_location(row_sel)
            direccion = _pick_address(row_sel)
            ubic_hogar = _pick_home_location(row_sel)
            territorio_cod = _pick_territory_code(row_sel)
            microterr_cod  = _pick_microterritory_code(row_sel)
            rol_familia    = _pick_family_role(row_sel)
            fecha_carac = _pick_char_date(row_sel)
            ebs_val  = str(row_sel.get(ebs_col_n, "—")) if ebs_col_n else "—"
            fam_val  = str(row_sel.get(fam_col_gn, "—"))  if fam_col_gn else "—"
            resp_val = str(row_sel.get(resp_col_n, "—")) if resp_col_n else "—"
            eapb_val = _map_eapb_value(row_sel.get(eapb_col_n, "—"), EAPB_MAP) if eapb_col_n else "—"
            cod_ficha = str(row_sel.get(col_id, "—")) if col_id else "—"
            sexo_val = row_sel.get('Sexo', 'No reportado')
            amb_val = row_sel.get('Ámbito', 'No definido')
            grupo_val = row_sel.get('Grupo etario', 'Sin dato')

            res_name_df = pd.DataFrame([{
                "Nombre completo": nombre_completo if nombre_completo else "—",
                "No. identificación": no_ident,
                "Sexo": sexo_val,
                "Ámbito": amb_val,
                "Grupo etario": grupo_val,
                "Edad": edad,
                "Barrio/Vereda": barrio_vereda,
                "Dirección": direccion,
                "Ubicación del hogar": ubic_hogar,
                "Territorio (cód.)": territorio_cod,
                "Microterritorio (cód.)": microterr_cod,
                "Rol en la familia": rol_familia,
                "Código de ficha": cod_ficha,
                "Num. Identificación Hogar": row_sel.get("numIdentificacionHogar", "—"),
                "Num. Identificación Familia": row_sel.get("numIdentificacionFamilia", "—"),
                "Fecha de caracterización": fecha_carac,
                "EBS": ebs_val,
                "Familia": fam_val,
                "Responsable (creatorFormulario)": resp_val,
                "EAPB": eapb_val,
            }])

            st.success("Ficha de la persona seleccionada:")
            st.dataframe(res_name_df, use_container_width=True, height=300)

            # ====== HISTORIAL DE ATENCIONES MÉDICAS (CRUCE) ======
            st.markdown("### 🏥 Historial de Atención Prestada (Hospital Malvinas)")
            df_atenciones_cruzadas = data_manager.query_atenciones_by_doc(no_ident)
            if df_atenciones_cruzadas.empty:
                st.info("No se registraron atenciones médicas para este usuario en las bases de datos cruzadas.")
            else:
                st.dataframe(df_atenciones_cruzadas, use_container_width=True)
                ac1, ac2 = st.columns(2)
                with ac1:
                    st.download_button("⬇️ CSV (Atenciones)", data=make_csv_bytes(df_atenciones_cruzadas),
                                       file_name=f"atenciones_{no_ident}.csv", mime="text/csv")
                with ac2:
                    st.download_button("⬇️ Excel (Atenciones)", data=make_xlsx_bytes(df_atenciones_cruzadas),
                                       file_name=f"atenciones_{no_ident}.xlsx",
                                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

            # ====== PLAN DE CUIDADO ======
            plan_df = build_plan_for_person(row_sel, CARE_RULES)
            st.markdown("**Sugerencias de Plan de Cuidado**")
            if plan_df.empty:
                st.info("Sin hallazgos para plan de cuidado con las reglas actuales.")
            else:
                st.dataframe(plan_df, use_container_width=True, height=220)
                pc1b, pc2b, pc3b = st.columns([1,1,1])
                with pc1b:
                    st.download_button("⬇️ CSV (plan)", data=make_csv_bytes(plan_df),
                                       file_name=f"plan_cuidado_nombre.csv", mime="text/csv")
                with pc2b:
                    st.download_button("⬇️ Excel (plan)", data=make_xlsx_bytes(plan_df),
                                       file_name=f"plan_cuidado_nombre.xlsx",
                                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                with pc3b:
                    html_plan = _printable_html(plan_df, title=f"Plan de cuidado — {nombre_completo}")
                    b64p = base64.b64encode(html_plan.encode("utf-8")).decode()
                    st.markdown(f"<a href='data:text/html;base64,{b64p}' target='_blank'>🖨️ Imprimir / Guardar como PDF</a>", unsafe_allow_html=True)

            # ====== Resumen + Listado FAMILIA ======
            if fam_col_gn and pd.notna(row_sel.get(fam_col_gn, np.nan)):
                fam_id = str(row_sel.get(fam_col_gn))
                df_fam = df_f[df_f[fam_col_gn].astype(str) == fam_id].copy()

                edades = pd.to_numeric(df_fam.get('edad', np.nan), errors='coerce')
                fam_total     = len(df_fam)
                fam_mayores   = int((edades >= 60).sum()) if 'edad' in df_fam.columns else 0
                fam_menores5  = int((edades < 5).sum())   if 'edad' in df_fam.columns else 0
                fam_gestantes = int(pd.to_numeric(df_fam.get('gestante_flag', 0), errors='coerce').fillna(0).sum()) if 'gestante_flag' in df_fam.columns else 0
                fam_prom_edad = f"{float(edades.mean()):.1f}" if 'edad' in df_fam.columns and pd.notna(edades.mean()) else "—"

                st.markdown(f"**Resumen de la familia {fam_id}**")
                fam_cards = [
                    {"icon":"👨‍👩‍👧‍👦","label":"Integrantes","value":fmt(fam_total),"color":"#2563eb","helper":"Personas en la familia"},
                    {"icon":"🧓","label":"Adultos mayores (60+)","value":fmt(fam_mayores),"color":"#7c3aed","helper":"Edad ≥ 60"},
                    {"icon":"👶","label":"Menores de 5 años","value":fmt(fam_menores5),"color":"#f59e0b","helper":"Edad < 5"},
                    {"icon":"🤰","label":"Gestantes","value":fmt(fam_gestantes),"color":"#dc2626","helper":"Marcadas como 1"},
                    {"icon":"📈","label":"Promedio de edad","value":fam_prom_edad,"color":"#059669","helper":"Años"},
                ]
                st.markdown("<div class='card-row'>" + "".join([card_tpl.format(**c) for c in fam_cards]) + "</div>", unsafe_allow_html=True)

                fam_listado = _make_people_listing(df_fam)
                doc_col_f2  = _find_col(df_fam, [
                    "nroDocumento","numDocumento","numeroDocumento","documento",
                    "cedula","cédula","cc","doc","nro_doc","identificacion","identificación"
                ])
                if doc_col_f2 and doc_col_n:
                    sel_doc_norm = _norm_doc(row_sel.get(doc_col_n, ""))
                    fam_listado.insert(0, "Consultado",
                        (df_fam[doc_col_f2].apply(_norm_doc) == sel_doc_norm).map(lambda x: "✅" if x else "").values)
                else:
                    fam_listado.insert(0, "Consultado", [""] * len(fam_listado))

                st.markdown(f"**Integrantes de la familia {fam_id}** (total: {len(fam_listado)})")
                st.dataframe(fam_listado, use_container_width=True, height=460)

                fn1, fn2, fn3 = st.columns([1,1,1])
                with fn1:
                    st.download_button("⬇️ CSV (familia)", data=make_csv_bytes(fam_listado),
                                       file_name=f"familia_{fam_id}.csv", mime="text/csv")
                with fn2:
                    st.download_button("⬇️ Excel (familia)", data=make_xlsx_bytes(fam_listado),
                                       file_name=f"familia_{fam_id}.xlsx",
                                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                with fn3:
                    html_fam = _printable_html(fam_listado, title=f"Familia {fam_id} — consulta por nombre")
                    b64f = base64.b64encode(html_fam.encode("utf-8")).decode()
                    st.markdown(f"<a href='data:text/html;base64,{b64f}' target='_blank'>🖨️ Imprimir / Guardar como PDF</a>", unsafe_allow_html=True)
            else:
                st.info("No se pudo identificar la familia (columna 'parentglobFamilia' u homólogas ausente o vacía).")

st.markdown("</div>", unsafe_allow_html=True)

# ===============================
# Listado general por EBS (con Ámbito y Sexo)
# ===============================
st.markdown("<div class='section-title'>Listado de personas caracterizadas por EBS</div>", unsafe_allow_html=True)
st.markdown("<div class='panel'>", unsafe_allow_html=True)

if col_ebs:
    df_list = df_f.copy()
    listado = _make_people_listing(df_list)
    st.dataframe(listado, use_container_width=True, height=500)

    cdl1, cdl2, cdl3 = st.columns([1,1,1])
    with cdl1:
        st.download_button("⬇️ CSV (listado EBS)", data=make_csv_bytes(listado),
            file_name="listado_por_EBS.csv", mime="text/csv")
    with cdl2:
        st.download_button("⬇️ Excel (listado EBS)", data=make_xlsx_bytes(listado),
            file_name="listado_por_EBS.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with cdl3:
        html_str = _printable_html(listado, title="Listado de personas por EBS")
        b64 = base64.b64encode(html_str.encode("utf-8")).decode()
        st.markdown(f"<a href='data:text/html;base64,{b64}' target='_blank'>🖨️ Imprimir / Guardar como PDF</a>", unsafe_allow_html=True)
else:
    st.info("Agrega 'nroIdentificacionEBS' para ver el listado por EBS.")
st.markdown("</div>", unsafe_allow_html=True)

# ===============================
# Footer
# ===============================
st.markdown(
    """
    <div class='badge'>
        Versión estilizada • Listo para usar con Streamlit
    </div>
    """,
    unsafe_allow_html=True
)

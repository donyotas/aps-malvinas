
# aps2_styled.py (versión optimizada para carga rápida)
# - Prefiere Parquet/CSV optimizado y hace fallback a XLSX
# - Lee solo columnas necesarias (usecols) y tipa columnas (dtypes)
# - Precálculos vectorizados (evita apply por fila)
# - KPIs + tarjetas (incluye Jefes de Hogar H/M)
# - Gráficos: Personas por EBS, Nivel educativo, Pertenencia étnica
# - Listados rápidos y listado general (controlado por checkbox)

import io
import re
import base64
import hashlib
from pathlib import Path
import pandas as pd
import numpy as np
import plotly.express as px
import streamlit as st

st.set_page_config(page_title="APS HMHOO 2025 — Optimizado", layout="wide")

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
# Estilos (CSS)
# ===============================
STYLES = """
<style>
.reportview-container, .main, .block-container { background: #f7fafc; }
.block-container { padding-top: 1.0rem; max-width: 96%; }
.card-row{display:flex;gap:14px;flex-wrap:wrap;margin-bottom:10px}
.card{flex:1 1 220px;display:flex;align-items:center;gap:12px;padding:16px 18px;
border:1px solid #e5e7eb;border-radius:16px;background:#ffffff;
box-shadow:0 1px 3px rgba(0,0,0,.06)}
.icon{font-size:26px;line-height:1}
.label{font-size:12px;color:#6b7280}
.value{font-size:22px;font-weight:800}
.sub{font-size:11px;color:#9ca3af}
.section-title{font-size:15px;font-weight:700;margin:8px 0 0 2px;color:#111827}
.panel{background:#ffffff;border:1px solid #e5e7eb;border-radius:16px;
padding:6px 10px 2px 10px;box-shadow:0 1px 3px rgba(0,0,0,.06)}
.logo-img{height:54px;object-fit:contain}
.badge{display:inline-block;padding:4px 10px;border-radius:999px;background:#f3f4f6;color:#111827;font-size:11px}
</style>
"""
st.markdown(STYLES, unsafe_allow_html=True)

# ===============================
# Datos base — preferir PARQUET/CSV optimizado con fallback a XLSX
# ===============================
st.sidebar.title("Fuente de datos")

DATA_DIR = Path(__file__).resolve().parent / "data"

# Lee solo lo que usa la app; ajusta si tus nombres cambian
USECOLS = [
    "nroDocumento","edad","nivelEducativo","pertenenciaEtnica",
    "nroIdentificacionEBS","parentglobFamilia","creationDateFormulario",
    "creatorFormulario","EAPB","globalid",
    "numIdentificacionHogar","numIdentificacionFamilia",
    "rolEnLaFamilia","sexo",
    # ubicación (ajusta según tu esquema)
    "barrio","vereda","direccion","ubicacionHogar",
    # territorio / microterritorio (ajusta si corresponde)
    "codTerritorio","id_territorio","codMicroterritorio","id_microterritorio",
    # nombres desagregados
    "primernombre","segundonombre","primerapellido","segundoapellido",
]

DTYPES = {
    "nroDocumento": "string",
    "nroIdentificacionEBS": "string",
    "parentglobFamilia": "string",
    "creatorFormulario": "string",
    "EAPB": "string",
    "globalid": "string",
    "numIdentificacionHogar": "string",
    "numIdentificacionFamilia": "string",
    "rolEnLaFamilia": "Int8",
    "sexo": "string",
    "nivelEducativo": "Int8",
    "pertenenciaEtnica": "Int8",
}

PREF_FILES = [
    DATA_DIR / "HMHOO_optimizado.parquet",
    DATA_DIR / "HMHOO.parquet",
    DATA_DIR / "HMHOO_optimizado.csv",
    # cualquier CSV en /data (más reciente)
    DATA_DIR / "HMHOO.xlsx",
]

def _file_md5(path: Path) -> str:
    with open(path, "rb") as f:
        return hashlib.md5(f.read()).hexdigest()

@st.cache_data(show_spinner=False)
def _load_parquet(path: Path, token: str) -> pd.DataFrame:
    return pd.read_parquet(path)

@st.cache_data(show_spinner=False)
def _load_csv(path: Path, token: str) -> pd.DataFrame:
    try:
        return pd.read_csv(path, usecols=lambda c: (c in USECOLS), dtype=DTYPES)
    except Exception:
        return pd.read_csv(path)

@st.cache_data(show_spinner=False)
def _load_excel(path: Path, token: str) -> pd.DataFrame:
    try:
        return pd.read_excel(path, usecols=[c for c in USECOLS if c])
    except Exception:
        return pd.read_excel(path)

def _pick_source() -> Path | None:
    if not DATA_DIR.exists():
        return None
    for p in PREF_FILES[:3]:
        if p.exists():
            return p
    csvs = [p for p in DATA_DIR.glob("*.csv")]
    if csvs:
        return max(csvs, key=lambda p: p.stat().st_mtime)
    x = DATA_DIR / "HMHOO.xlsx"
    if x.exists() and not x.name.startswith("~$"):
        return x
    xlsxs = [p for p in DATA_DIR.glob("*.xlsx") if not p.name.startswith("~$")]
    if xlsxs:
        return max(xlsxs, key=lambda p: p.stat().st_mtime)
    return None

if st.sidebar.button("🔁 Recargar datos"):
    st.cache_data.clear()
    st.rerun()

src_path = _pick_source()
if not src_path:
    st.error("No encontré datos en `data/`. Sube Parquet/CSV optimizado o tu XLSX.")
    st.stop()

try:
    ext = src_path.suffix.lower()
    token = str(src_path.stat().st_mtime_ns)
    if ext == ".parquet":
        df = _load_parquet(src_path, token)
    elif ext == ".csv":
        df = _load_csv(src_path, token)
    else:
        token = _file_md5(src_path) if src_path.stat().st_size < 50_000_000 else token
        df = _load_excel(src_path, token)
    source_name = str(src_path)
    st.sidebar.caption(f"Fuente: **{src_path.name}** (/data)")
except Exception as e:
    st.error(f"No pude cargar {src_path.name}. Error: {e}")
    st.stop()

# ===============================
# Encabezado con identidad visual
# ===============================
def find_logo_path():
    for p in ["assets/logo.png","assets/logo.jpg","assets/logo.jpeg","assets/logo.webp",
              "assets/LOGO.png","assets/LOGO.jpg","assets/LOGO.jpeg","assets/LOGO.webp"]:
        if (Path(__file__).parent / p).exists():
            return Path(__file__).parent / p
    return None

logo_path = find_logo_path()
left, right = st.columns([1,3])
with left:
    if logo_path:
        st.image(str(logo_path), caption="Atención Primaria en Salud", use_container_width=True)
    else:
        st.markdown("<div class='badge'>Sin logo (sube assets/logo.png)</div>", unsafe_allow_html=True)
with right:
    st.markdown(f"""
    <div style="display:flex;justify-content:flex-end;gap:8px;align-items:center;">
      <div class='badge'>Fuente: {src_path.name}</div>
      <div class='badge'>Registros: {len(df):,}</div>
    </div>
    """, unsafe_allow_html=True)

# ===============================
# Precálculos (evita apply por fila)
# ===============================
@st.cache_data(show_spinner=False)
def _precompute(df_in: pd.DataFrame) -> pd.DataFrame:
    df2 = df_in.copy()

    # Edad numérica
    if "edad" in df2.columns:
        df2["edad"] = pd.to_numeric(df2["edad"], errors="coerce")

    # Fecha caracterización (primera disponible)
    date_col = None
    for c in ["creationDateFormulario","fechaCaracterizacion","fecha_caracterizacion","fecha_atencion","fechaRegistro","fecha"]:
        if c in df2.columns:
            date_col = c; break
    if date_col:
        df2["__fecha_char"] = pd.to_datetime(df2[date_col], errors="coerce", dayfirst=True)
    else:
        df2["__fecha_char"] = pd.NaT

    # Nombre completo (vectorizado si existen partes)
    if any(c in df2.columns for c in ["primernombre","primerapellido"]):
        n1 = df2.get("primernombre","").astype("string").fillna("")
        n2 = df2.get("segundonombre","").astype("string").fillna("")
        a1 = df2.get("primerapellido","").astype("string").fillna("")
        a2 = df2.get("segundoapellido","").astype("string").fillna("")
        df2["__nombre_completo"] = (
            (n1 + " " + n2 + " " + a1 + " " + a2)
            .str.replace(r"\s+", " ", regex=True)
            .str.strip()
        )
    else:
        df2["__nombre_completo"] = ""

    # Documento normalizado y mostrable
    if "nroDocumento" in df2.columns:
        doc_raw = df2["nroDocumento"].astype("string").fillna("")
        df2["__doc_norm"] = doc_raw.str.replace(r"\D", "", regex=True)
        df2["__doc_disp"] = df2["__doc_norm"].mask(df2["__doc_norm"] == "", doc_raw)
    else:
        df2["__doc_norm"] = ""
        df2["__doc_disp"] = "—"

    # Rol legible
    ROLE_MAP = {1:"Jefe de familia",2:"Cónyuge",3:"Hijo(a)",4:"Hermano(a)",5:"Padre o madre",6:"Otro"}
    if "rolEnLaFamilia" in df2.columns:
        df2["__rol_familia_txt"] = (
            pd.to_numeric(df2["rolEnLaFamilia"], errors="coerce")
              .map(ROLE_MAP).fillna("—")
        )
    else:
        df2["__rol_familia_txt"] = "—"

    # Sexo normalizado para KPIs
    if "sexo" in df2.columns:
        s = df2["sexo"].astype("string").str.upper().str.strip()
        df2["__sexo_norm"] = np.where(s.str.startswith("M"), "Hombre",
                               np.where(s.str.startswith("F"), "Mujer", "No reportado"))
        df2["__sexo_norm"] = df2["__sexo_norm"].astype("category")
    else:
        df2["__sexo_norm"] = "No reportado"

    # Territorio / microterritorio (elige la disponible)
    terr = None
    for c in ["codTerritorio","id_territorio"]:
        if c in df2.columns: terr = c; break
    micro = None
    for c in ["codMicroterritorio","id_microterritorio"]:
        if c in df2.columns: micro = c; break
    df2["__territorio"] = df2.get(terr, "—").astype("string") if terr else "—"
    df2["__microterritorio"] = df2.get(micro, "—").astype("string") if micro else "—"

    # Gestante flag
    if "esGestante" in df2.columns:
        df2["gestante_flag"] = (pd.to_numeric(df2["esGestante"], errors="coerce") == 1).astype(int)
    elif "gestantes" in df2.columns:
        df2["gestante_flag"] = (pd.to_numeric(df2["gestantes"], errors="coerce") == 1).astype(int)
    else:
        df2["gestante_flag"] = 0

    # Categorías para ahorrar RAM
    for cat in ["nivelEducativo","pertenenciaEtnica","creatorFormulario","EAPB","nroIdentificacionEBS","parentglobFamilia"]:
        if cat in df2.columns:
            try:
                df2[cat] = df2[cat].astype("category")
            except Exception:
                pass

    return df2

df = _precompute(df)

# ===============================
# KPIs (tarjetas)
# ===============================
def fmt(x): 
    return f"{int(x):,}".replace(",", ".") if isinstance(x,(int,np.integer)) or (isinstance(x,float) and np.isfinite(x)) else "—"

total_personas = len(df)
familias = df["parentglobFamilia"].nunique() if "parentglobFamilia" in df.columns else np.nan
gestantes = int(df.get("gestante_flag", 0).sum()) if "gestante_flag" in df.columns else 0
adultos_mayores = int((df.get("edad", np.nan) >= 60).sum()) if "edad" in df.columns else 0
menores_5 = int((df.get("edad", np.nan) < 5).sum()) if "edad" in df.columns else 0

# Jefes de hogar por sexo
jefes = df[df.get("rolEnLaFamilia","").astype(str) == "1"]
jefes_hombres = int((jefes.get("__sexo_norm","") == "Hombre").sum()) if not jefes.empty else 0
jefes_mujeres = int((jefes.get("__sexo_norm","") == "Mujer").sum()) if not jefes.empty else 0

cards = [
    {"icon":"👥","label":"Personas caracterizadas","value":fmt(total_personas),"color":"#2563eb","helper":"Total filtrado"},
    {"icon":"🏠","label":"Familias caracterizadas","value":fmt(familias),"color":"#059669","helper":"Únicos en parentglobFamilia"},
    {"icon":"🤰","label":"Gestantes","value":fmt(gestantes),"color":"#dc2626","helper":"Marcadas como 1"},
    {"icon":"🧓","label":"Adultos mayores (60+)","value":fmt(adultos_mayores),"color":"#7c3aed","helper":"Edad ≥ 60"},
    {"icon":"👶","label":"Menores de 5 años","value":fmt(menores_5),"color":"#f59e0b","helper":"Edad < 5"},
    {"icon":"👨","label":"Jefes de hogar (H)","value":fmt(jefes_hombres),"color":"#3b82f6","helper":"rolEnLaFamilia = 1 y sexo M"},
    {"icon":"👩","label":"Jefes de hogar (M)","value":fmt(jefes_mujeres),"color":"#ec4899","helper":"rolEnLaFamilia = 1 y sexo F"},
]
card_tpl = ("<div class='card'>"
            "<div class='icon'>{icon}</div>"
            "<div><div class='label'>{label}</div>"
            "<div class='value' style='color:{color}'>{value}</div>"
            "<div class='sub'>{helper}</div></div></div>")
st.markdown("<div class='card-row'>" + "".join([card_tpl.format(**c) for c in cards]) + "</div>", unsafe_allow_html=True)

# ===============================
# Tres paneles (gráficos)
# ===============================
c1, c2, c3 = st.columns((1.2, 1, 1))

with c1:
    st.markdown("<div class='section-title'>Personas caracterizadas por EBS</div>", unsafe_allow_html=True)
    with st.container():
        st.markdown("<div class='panel'>", unsafe_allow_html=True)
        col_ebs = "nroIdentificacionEBS" if "nroIdentificacionEBS" in df.columns else None
        if col_ebs:
            personas_ebs = (
                df.groupby(col_ebs, dropna=True)
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
    st.markdown("<div class='section-title'>Nivel educativo</div>", unsafe_allow_html=True)
    st.markdown("<div class='panel'>", unsafe_allow_html=True)
    if 'nivelEducativo' in df.columns:
        edu_counts = (
            df['nivelEducativo'].astype(str).replace({'nan':'No reportado'})
              .value_counts(dropna=False).rename_axis('Nivel Educativo')
              .reset_index(name='Personas').sort_values('Personas')
        )
        fig_edu = px.bar(
            edu_counts, x='Personas', y='Nivel Educativo',
            orientation='h', text='Personas', color='Nivel Educativo',
            template='simple_white'
        )
        fig_edu.update_traces(textposition='outside', cliponaxis=False)
        fig_edu.update_layout(showlegend=False, height=320,
                              margin=dict(l=10,r=10,t=40,b=10),
                              xaxis_title="Personas", yaxis_title="")
        st.plotly_chart(fig_edu, use_container_width=True)
    else:
        st.info("No se encontró 'nivelEducativo'.")
    st.markdown("</div>", unsafe_allow_html=True)

with c3:
    st.markdown("<div class='section-title'>Pertenencia étnica</div>", unsafe_allow_html=True)
    st.markdown("<div class='panel'>", unsafe_allow_html=True)
    if 'pertenenciaEtnica' in df.columns:
        etn_counts = (
            df['pertenenciaEtnica'].astype(str)
              .value_counts(dropna=False).rename_axis('Pertenencia étnica')
              .reset_index(name='Personas').sort_values('Personas')
        )
        fig_etn = px.bar(
            etn_counts, x='Personas', y='Pertenencia étnica',
            orientation='h', text='Personas', color='Pertenencia étnica',
            template='simple_white'
        )
        fig_etn.update_traces(textposition='outside', cliponaxis=False)
        fig_etn.update_layout(showlegend=False, height=430,
                              margin=dict(l=10,r=10,t=40,b=10),
                              xaxis_title="Personas", yaxis_title="")
        st.plotly_chart(fig_etn, use_container_width=True)
    else:
        st.info("No se encontró 'pertenenciaEtnica'.")
    st.markdown("</div>", unsafe_allow_html=True)

# ===============================
# Listados rápidos
# ===============================
st.markdown("<div class='section-title'>Listados rápidos</div>", unsafe_allow_html=True)
st.markdown("<div class='panel'>", unsafe_allow_html=True)

col1, col2, col3 = st.columns(3)
with col1:
    show_g = st.button("🤰 Ver gestantes")
with col2:
    show_may = st.button("🧓 Ver adultos mayores (60+)")
with col3:
    show_m5 = st.button("👶 Ver menores de 5 años")

def _make_people_listing_fast(df_sub: pd.DataFrame) -> pd.DataFrame:
    if df_sub is None or df_sub.empty:
        return pd.DataFrame()
    out = pd.DataFrame({
        "EBS": df_sub.get("nroIdentificacionEBS", "—").astype(str),
        "Nombre completo": df_sub.get("__nombre_completo", "—"),
        "No. identificación": df_sub.get("__doc_disp", "—"),
        "Edad": df_sub.get("edad", "—"),
        "Barrio/Vereda": df_sub.get("barrio", df_sub.get("vereda","—")).astype(str),
        "Dirección": df_sub.get("direccion", "—").astype(str),
        "Ubicación del hogar": df_sub.get("ubicacionHogar", "—").astype(str),
        "Territorio (cód.)": df_sub.get("__territorio","—").astype(str),
        "Microterritorio (cód.)": df_sub.get("__microterritorio","—").astype(str),
        "Rol en la familia": df_sub.get("__rol_familia_txt","—"),
        "Código de ficha": df_sub.get("globalid","—").astype(str),
        "Num. Identificación Hogar": df_sub.get("numIdentificacionHogar","—").astype(str),
        "Num. Identificación Familia": df_sub.get("numIdentificacionFamilia","—").astype(str),
        "Fecha de caracterización": df_sub["__fecha_char"].dt.strftime("%Y-%m-%d").fillna("—") if "__fecha_char" in df_sub.columns else "—",
        "Responsable (creatorFormulario)": df_sub.get("creatorFormulario","—").astype(str),
        "Familia": df_sub.get("parentglobFamilia","—").astype(str),
        "EAPB": df_sub.get("EAPB","—").astype(str),
    })
    if "__fecha_char" in df_sub.columns:
        out = out.assign(_sort=df_sub["__fecha_char"]).sort_values(["_sort","Nombre completo"], ascending=[False, True]).drop(columns=["_sort"])
    else:
        out = out.sort_values(["Nombre completo"], ascending=True)
    return out

if show_g or show_may or show_m5:
    if show_g:
        df_sub = df[df.get("gestante_flag", 0) == 1].copy()
        titulo = "Listado de gestantes"
    elif show_may:
        df_sub = df[pd.to_numeric(df.get("edad", np.nan), errors="coerce") >= 60].copy()
        titulo = "Listado de adultos mayores (60+)"
    else:
        df_sub = df[pd.to_numeric(df.get("edad", np.nan), errors="coerce") < 5].copy()
        titulo = "Listado de menores de 5 años"

    if df_sub.empty:
        st.info("No hay registros para mostrar con el filtro seleccionado.")
    else:
        listado_quick = _make_people_listing_fast(df_sub)
        st.markdown(f"**{titulo}** (total: {len(listado_quick)})")
        st.dataframe(listado_quick, use_container_width=True, height=460)

        colA, colB = st.columns(2)
        with colA:
            st.download_button("⬇️ CSV", data=listado_quick.to_csv(index=False).encode("utf-8-sig"),
                               file_name=f"{titulo.lower().replace(' ', '_')}.csv", mime="text/csv")
        with colB:
            # Export a Excel rápido
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                listado_quick.to_excel(writer, index=False, sheet_name="Resultado")
            buffer.seek(0)
            st.download_button("⬇️ Excel", data=buffer.getvalue(),
                               file_name=f"{titulo.lower().replace(' ', '_')}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
st.markdown("</div>", unsafe_allow_html=True)

# ===============================
# Listado general por EBS (checkbox para no renderizar gigante por defecto)
# ===============================
st.markdown("<div class='section-title'>Listado de personas caracterizadas por EBS</div>", unsafe_allow_html=True)
st.markdown("<div class='panel'>", unsafe_allow_html=True)

if "nroIdentificacionEBS" in df.columns:
    df_list = df.copy()
    listado = _make_people_listing_fast(df_list)

    show_table = st.checkbox("Mostrar tabla completa", value=False)
    st.caption(f"Registros: {len(listado):,}".replace(",", "."))

    if show_table:
        st.dataframe(listado, use_container_width=True, height=500)
    else:
        st.info("Activa el check para ver la tabla. Puedes descargar abajo.")

    cdl1, cdl2 = st.columns(2)
    with cdl1:
        st.download_button("⬇️ CSV (listado EBS)", data=listado.to_csv(index=False).encode("utf-8-sig"),
            file_name="listado_por_EBS.csv", mime="text/csv")
    with cdl2:
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
            listado.to_excel(writer, index=False, sheet_name="Resultado")
        buffer.seek(0)
        st.download_button("⬇️ Excel (listado EBS)", data=buffer.getvalue(),
            file_name="listado_por_EBS.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
else:
    st.info("Agrega 'nroIdentificacionEBS' para ver el listado por EBS.")
st.markdown("</div>", unsafe_allow_html=True)

# ===============================
# Footer
# ===============================
st.markdown(
    "<div class='badge'>Versión optimizada • Parquet/CSV preferidos • Precálculos sin apply</div>",
    unsafe_allow_html=True
)

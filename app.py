# -*- coding: utf-8 -*-
"""
Vascular Health Analyzer - EVA / SUPERNOVA
Informe PDF profesional para exportación médica + importación de PDF/TXT.
"""

import datetime
import json
import os
import re
import sys
import traceback
import unicodedata
from io import BytesIO

# ---- Imports con manejo amigable de errores ----
_MISSING = []
try:
    import matplotlib
    matplotlib.use("Agg")  # backend headless: imprescindible para servidores/Streamlit Cloud
    import matplotlib.pyplot as plt
    from matplotlib.gridspec import GridSpec
except ImportError:
    _MISSING.append("matplotlib")
try:
    import numpy as np
except ImportError:
    _MISSING.append("numpy")
try:
    import pandas as pd
except ImportError:
    _MISSING.append("pandas")
try:
    import openpyxl  # noqa: F401 - requerido para exportación Excel
except ImportError:
    _MISSING.append("openpyxl")
try:
    import streamlit as st
except ImportError:
    print("ERROR: streamlit no esta instalado. Ejecute:\n"
          "  pip install streamlit matplotlib numpy pandas fpdf2 openpyxl pdfplumber Pillow",
          file=sys.stderr)
    sys.exit(1)
try:
    from fpdf import FPDF
except ImportError:
    _MISSING.append("fpdf2")

# Modulo local de gestion de usuarios (users.py debe estar al lado de app.py)
UserStore = None
_USERS = None
_USER_ERR = ""
try:
    from users import UserStore as _US
    UserStore = _US
    _USERS = UserStore()
except Exception as _e:
    _USER_ERR = f"{type(_e).__name__}: {_e}"

if _MISSING:
    st.set_page_config(page_title="Faltan dependencias", layout="centered")
    st.error(
        "Faltan paquetes Python para ejecutar la app: **" + ", ".join(_MISSING) + "**.\n\n"
        "Instalelos en su terminal con:\n\n"
        "```\npip install streamlit matplotlib numpy pandas fpdf2 openpyxl pdfplumber Pillow\n```"
    )
    st.stop()

# Lector de PDF (para importar archivos PDF; TXT no requiere librerías adicionales)
try:
    import pdfplumber  # type: ignore
    _PDF_BACKEND = "pdfplumber"
except ImportError:
    try:
        from pypdf import PdfReader  # type: ignore
        _PDF_BACKEND = "pypdf"
    except ImportError:
        try:
            from PyPDF2 import PdfReader  # type: ignore
            _PDF_BACKEND = "pypdf2"
        except ImportError:
            _PDF_BACKEND = None

# Captura opcional de imagen desde PDF (curva carotídeo-femoral)
try:
    import fitz  # PyMuPDF  # type: ignore
    _FITZ_OK = True
except Exception:
    fitz = None  # type: ignore
    _FITZ_OK = False

try:
    from PIL import Image  # type: ignore
    _PIL_OK = True
except Exception:
    Image = None  # type: ignore
    _PIL_OK = False


# ---------------------------------------------------------------------------
# CONFIGURACIÓN
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="Vascular Health Analyzer - EVA/SUPERNOVA",
    layout="wide",
    page_icon="🫀",
)

COLOR_SUPERNOVA = "#1F6FB2"
COLOR_NORMAL = "#1E8449"
COLOR_EVA = "#C0392B"
COLOR_BG_HEADER = (31, 78, 120)
COLOR_BG_SECTION = (230, 236, 245)
COLOR_BG_ALERT = (253, 237, 236)
COLOR_BG_OK = (232, 245, 233)

REFERENCIAS_BIBLIOGRAFICAS = [
    "Laurent S, Cockcroft J, Van Bortel L, et al. Expert consensus document on arterial stiffness: methodological issues and clinical applications. Eur Heart J. 2006;27(21):2588-2605.",
    "Van Bortel LM, Laurent S, Boutouyrie P, et al. Expert consensus document on the measurement of aortic stiffness in daily practice using carotid-femoral pulse wave velocity. J Hypertens. 2012;30(3):445-448.",
    "Vlachopoulos C, Aznaouridis K, Stefanadis C. Prediction of cardiovascular events and all-cause mortality with arterial stiffness: a systematic review and meta-analysis. J Am Coll Cardiol. 2010;55(13):1318-1327.",
    "Williams B, Mancia G, Spiering W, et al. 2018 ESC/ESH Guidelines for the management of arterial hypertension. Eur Heart J. 2018;39(33):3021-3104.",
    "Mancia G, Kreutz R, Brunström M, et al. 2023 ESH Guidelines for the management of arterial hypertension. J Hypertens. 2023;41(12):1874-2071.",
    "Reference Values for Arterial Stiffness Collaboration. Determinants of pulse wave velocity in healthy people and in the presence of cardiovascular risk factors. Eur Heart J. 2010;31(19):2338-2350.",
]

NOTA_METODOLOGICA_RIESGO = (
    "La estimación vascular de riesgo de esta app no reemplaza scores validados como SCORE2, "
    "SCORE2-OP, Framingham o ASCVD. Integra marcadores vasculares y hemodinámicos, "
    "especialmente VOP carótido-femoral, presión arterial y presión de pulso. "
    "El punto de corte VOP > 10 m/s se consigna como marcador de daño vascular/lesión de órgano blanco "
    "según consensos y guías de hipertensión arterial."
)


# ---------------------------------------------------------------------------
# ARCHIVOS POR USUARIO: FIRMA Y SELLO DIGITAL
# ---------------------------------------------------------------------------
APP_DIR = os.path.dirname(os.path.abspath(__file__))
USER_ASSETS_DIR = os.path.join(APP_DIR, "user_assets")
DATA_DIR = os.path.join(APP_DIR, "data")
PATIENT_DB_JSONL = os.path.join(DATA_DIR, "pacientes_vop_registros.jsonl")


def _usuario_seguro(username: str) -> str:
    txt = (username or "usuario").strip().lower()
    txt = re.sub(r"[^a-z0-9_.-]+", "_", txt)
    return txt[:60] or "usuario"


def _guardar_asset_usuario(username: str, uploaded_file, tipo: str):
    """Guarda firma/sello en carpeta aislada por usuario autenticado/registrado."""
    if uploaded_file is None:
        return None
    tipo = "sello" if tipo == "sello" else "firma"
    user_dir = os.path.join(USER_ASSETS_DIR, _usuario_seguro(username))
    os.makedirs(user_dir, exist_ok=True)
    ext = os.path.splitext(getattr(uploaded_file, "name", ""))[1].lower()
    if ext not in (".png", ".jpg", ".jpeg"):
        ext = ".png"
    path = os.path.join(user_dir, f"{tipo}{ext}")
    uploaded_file.seek(0)
    with open(path, "wb") as f:
        f.write(uploaded_file.read())
    return path


def _asset_usuario(username: str, tipo: str):
    """Devuelve la ruta del asset del usuario, si existe. Nunca usa assets de otro usuario."""
    tipo = "sello" if tipo == "sello" else "firma"
    user_dir = os.path.join(USER_ASSETS_DIR, _usuario_seguro(username))
    for ext in (".png", ".jpg", ".jpeg"):
        path = os.path.join(user_dir, f"{tipo}{ext}")
        if os.path.exists(path):
            return path
    return None


def _perfil_usuario_actual(username: str):
    return {
        "firma_path": _asset_usuario(username, "firma"),
        "sello_path": _asset_usuario(username, "sello"),
    }


# ---------------------------------------------------------------------------
# BASE DE DATOS SIMPLE POR USUARIO + EXPORTACIÓN EXCEL
# ---------------------------------------------------------------------------
def _asegurar_data_dir():
    os.makedirs(DATA_DIR, exist_ok=True)


def _guardar_registro_paciente(registro: dict):
    """Guarda un registro clínico en JSONL persistente.

    Cada línea pertenece a un paciente evaluado y queda asociada al usuario
    autenticado. Un médico solo exporta sus registros; el administrador exporta
    todos los usuarios.
    """
    _asegurar_data_dir()
    reg = dict(registro or {})
    reg.setdefault("timestamp", datetime.datetime.now().isoformat(timespec="seconds"))
    with open(PATIENT_DB_JSONL, "a", encoding="utf-8") as f:
        f.write(json.dumps(reg, ensure_ascii=False, default=str) + "\n")


def _cargar_registros_pacientes():
    if not os.path.exists(PATIENT_DB_JSONL):
        return []
    registros = []
    with open(PATIENT_DB_JSONL, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            try:
                registros.append(json.loads(line))
            except Exception:
                continue
    return registros


def _df_registros(usuario_actual: str, rol: str):
    registros = _cargar_registros_pacientes()
    if rol != "admin":
        u = _usuario_seguro(usuario_actual)
        registros = [r for r in registros if r.get("usuario_id") == u]
    if not registros:
        return pd.DataFrame()
    df = pd.DataFrame(registros)
    columnas_preferidas = [
        "timestamp", "fecha_estudio", "usuario", "usuario_id", "rol",
        "paciente", "documento", "edad", "sexo", "medico_solicitante",
        "vop_cf_ms", "distancia_cf_cm", "tiempo_transito_cf_ms",
        "pas_mmhg", "pad_mmhg", "pam_mmhg", "pp_mmhg",
        "p10_ms", "p25_ms", "p50_ms", "p75_ms", "p90_ms",
        "edad_vascular", "fenotipo_vascular_unico", "estimacion_vascular_riesgo",
        "lob_vop_mayor_10", "estimacion_vascular_elevada_por_vop_mayor_10",
        "fuente_vop", "archivo_importado"
    ]
    columnas = [c for c in columnas_preferidas if c in df.columns] + [c for c in df.columns if c not in columnas_preferidas]
    return df[columnas]


def _excel_bytes_registros(df: pd.DataFrame, nombre_hoja="Pacientes"):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=nombre_hoja[:31])
        ws = writer.book[nombre_hoja[:31]]
        # Congelar encabezado y ajustar ancho de columnas para lectura clínica.
        ws.freeze_panes = "A2"
        for col in ws.columns:
            max_len = 10
            letter = col[0].column_letter
            for cell in col:
                try:
                    max_len = max(max_len, len(str(cell.value)))
                except Exception:
                    pass
            ws.column_dimensions[letter].width = min(max_len + 2, 42)
    return output.getvalue()


# ---------------------------------------------------------------------------
# LÓGICA CLÍNICA
# ---------------------------------------------------------------------------
class VascularEngine:
    @staticmethod
    def calcular_vop(distancia_cm: float, tiempo_ms: float) -> float:
        """Calcula VOP cf desde distancia y tiempo, sin factor corrector 0.8.
        Este cálculo queda como respaldo; si el PDF informa VOP cf medida,
        se usa el valor medido por el equipo como dato primario.
        """
        if tiempo_ms <= 0:
            return 0.0
        d_m = distancia_cm / 100
        t_s = tiempo_ms / 1000
        return round(d_m / t_s, 2)

    @staticmethod
    def obtener_percentiles(edad: int, sexo: str):
        base = 5.5 if sexo == "Femenino" else 6.0
        inc = 0.08 * (edad - 20) if edad > 20 else 0
        p50 = base + inc
        return tuple(round(v, 2) for v in (p50 * 0.80, p50 * 0.90, p50, p50 * 1.15, p50 * 1.30))

    @staticmethod
    def clasificar_fenotipo(vop, p10, p90):
        """Clasificacion segun guias ARTERY 2024:
           - SUPERNOVA: VOP < p10 (segun sexo y edad)
           - HVA (Envejecimiento Vascular Saludable): p10 <= VOP <= p90
           - EVA: VOP > p90
        LOB (lesion de organo blanco) es una marca adicional cuando VOP > 10 m/s,
        independiente del fenotipo (se reporta por separado).
        """
        if vop < p10:
            return "SUPERNOVA (Envejecimiento Vascular Supranormal)", COLOR_SUPERNOVA, "Bajo"
        if vop > p90:
            return "EVA (Envejecimiento Vascular Acelerado)", COLOR_EVA, "Alto"
        return "HVA (Envejecimiento Vascular Saludable)", COLOR_NORMAL, "Bajo"

    @staticmethod
    def edad_vascular(vop, sexo):
        base = 5.5 if sexo == "Femenino" else 6.0
        if vop <= base:
            return 20
        return int(round(max(20, min(20 + (vop - base) / 0.08, 100))))

    @staticmethod
    def presion_pulso(pas, pad):
        return round(pas - pad, 1)

    @staticmethod
    def presion_arterial_media(pas, pad):
        return round(pad + (pas - pad) / 3, 1)

    @staticmethod
    def estimacion_vascular_riesgo_global(vop, pas, pad, edad):
        pp = pas - pad
        hta = pas >= 140 or pad >= 90
        if vop > 13 or (vop > 10 and hta and edad >= 60):
            return "Muy Alto"
        if vop > 10 or pp >= 60 or hta:
            return "Alto"
        if vop > 9 or pp >= 50:
            return "Moderado"
        return "Bajo"

    @staticmethod
    def recomendaciones(fenotipo, riesgo, vop, pas, pad):
        recs = []
        if "EVA" in fenotipo:
            recs += [
                "Control intensivo de presion arterial (objetivo < 130/80 mmHg).",
                "Dieta DASH/mediterranea, reduccion de sodio (<5 g/dia).",
                "Actividad fisica aerobica >=150 min/semana + entrenamiento de fuerza.",
                "Perfil lipidico, glicemia y funcion renal en proximas 4 semanas.",
                "Reevaluacion de VOP en 6-12 meses para monitorizar progresion.",
            ]
        elif "SUPERNOVA" in fenotipo:
            recs += [
                "Mantener habitos cardioprotectores actuales.",
                "Reforzar adherencia a actividad fisica regular y dieta equilibrada.",
                "Reevaluacion bianual de rigidez arterial.",
            ]
        else:
            recs += [
                "Promover habitos cardiosaludables (dieta, ejercicio, sueno).",
                "Control anual de presion arterial y perfil metabolico.",
                "Reevaluacion de VOP cada 2 anos o ante nuevos factores de riesgo.",
            ]
        if pas >= 140 or pad >= 90:
            recs.append("Confirmar diagnostico de HTA con MAPA/AMPA e iniciar tratamiento (ESC/ESH).")
        if riesgo in ("Alto", "Muy Alto"):
            recs.append("Derivacion a Cardiologia/Medicina Vascular para evaluacion integral.")
        return recs


# ---------------------------------------------------------------------------
# IMPORTADOR PDF/TXT
# ---------------------------------------------------------------------------
class GenericReportParser:
    """Extrae datos desde PDF o TXT priorizando VOP carotídeo-femoral.

    Puntos corregidos:
    - No usa VOP radial como referencia.
    - Busca explícitamente VOP cf / carotídeo-femoral / cfPWV.
    - Importa distancia CF, tiempo de tránsito y VOP cf medida desde el archivo original.
    - Si un dato no se detecta, queda en 0/SD para que no aparezcan valores por defecto falsos.
    """

    NEGATIVOS_CF = [
        "radial", "braquial", "brachial", "perifer", "peripheral",
        "aortica", "aórtica", "central", "augmentation", "aix",
        "referencia", "reference", "normal", "umbral", "cutoff", "objetivo",
    ]
    POSITIVOS_CF = [
        "cf", "c-f", "c/f", "carot", "femor", "femoral", "carotido", "carótido",
        "carotideo", "carotídeo", "carotid", "cfpwv", "pwvcf",
    ]

    PATRONES = {
        "nombre": [
            r"(?:Nombre\s+y\s+Apellido|Apellido\s+y\s+Nombre|Paciente|Nombre)\s*[:\-]?\s*([A-Za-zÁÉÍÓÚÑáéíóúñ\.\s]{3,80}?)(?=\s*(?:\n|Edad|DNI|Documento|Sexo|Fecha|$))",
        ],
        "documento": [
            r"(?:DNI|Documento|D\.N\.I\.?|N[°º]\s*Documento|CI|Cedula|Cédula)\s*[:\-]?\s*([\d\.\-]{6,15})",
        ],
        "edad": [
            r"Edad\s*[:\-]?\s*(\d{1,3})\s*(?:a[nñ]os|años|a\.?)?",
        ],
        "sexo": [
            r"Sexo\s*[:\-]?\s*(Masculino|Femenino|M|F|Mas|Fem)\b",
        ],
    }

    @staticmethod
    def _norm_num(v):
        if v is None:
            return None
        try:
            txt = str(v).strip().replace(" ", "")
            txt = txt.replace("m/s", "").replace("ms", "").replace("cm", "")
            if "," in txt and "." in txt:
                txt = txt.replace(".", "").replace(",", ".")
            else:
                txt = txt.replace(",", ".")
            return float(txt)
        except Exception:
            return None

    @staticmethod
    def _ventana(texto, ini, fin, n=120):
        return texto[max(0, ini-n):min(len(texto), fin+n)].lower()

    @classmethod
    def _es_cf(cls, ventana: str) -> bool:
        if any(x in ventana for x in cls.NEGATIVOS_CF):
            # Permitir "carotido-femoral" aunque también diga referencia solo si el contexto no parece tabla de valores normales.
            if any(x in ventana for x in ["radial", "braquial", "brachial", "perifer", "peripheral", "central", "aortica", "aórtica"]):
                return False
            if any(x in ventana for x in ["referencia", "reference", "normal", "umbral", "cutoff", "objetivo"]):
                return False
        return any(x in ventana for x in cls.POSITIVOS_CF)

    @staticmethod
    def _rango(v, lo, hi):
        return v is not None and lo <= v <= hi

    @staticmethod
    def _lineas_reales(txt: str):
        """Devuelve líneas no vacías conservando el orden original del archivo."""
        if not txt:
            return []
        return [re.sub(r"\s+", " ", x).strip() for x in str(txt).splitlines() if x and x.strip()]

    @staticmethod
    def _normalizar_ocr(txt: str) -> str:
        """Normaliza variantes frecuentes de PDF/TXT sin formato y errores de OCR."""
        if txt is None:
            return ""
        txt = str(txt)
        txt = txt.replace("\u00a0", " ").replace("│", "|").replace("¦", "|")
        txt = txt.replace("Carótida", "Carotida").replace("Carótido", "Carotido")
        txt = txt.replace("carótida", "carotida").replace("carótido", "carotido")
        txt = txt.replace("Fem.", "Fem.")
        return txt

    @staticmethod
    def _limpiar_texto(txt: str) -> str:
        txt = txt.replace("\u00a0", " ").replace("\t", " ")
        # Unificar separadores frecuentes de OCR / texto plano sin perder saltos de línea.
        txt = txt.replace("：", ":").replace("＝", "=").replace("–", "-").replace("—", "-")
        txt = re.sub(r"[ ]+", " ", txt)
        txt = re.sub(r"\n{3,}", "\n\n", txt)
        return txt

    @staticmethod
    def _nombre_desde_archivo(file_obj):
        """Respaldo cuando el TXT/PDF no trae etiqueta clara de paciente.
        Ej.: 'DAIANA ELIZABETH ALMEIDA 1.txt' -> 'DAIANA ELIZABETH ALMEIDA'.
        """
        nombre_archivo = getattr(file_obj, "name", "") or ""
        base = re.sub(r"\.[A-Za-z0-9]{1,6}$", "", nombre_archivo).strip()
        base = re.sub(r"[_-]+", " ", base)
        base = re.sub(r"\b(?:PDF|TXT|INFORME|REPORTE|VOP|EVA|SUPERNOVA)\b", " ", base, flags=re.IGNORECASE)
        base = re.sub(r"\b(?:copia|copy|final|corregido|corr|nuevo|curva|cf)\b", " ", base, flags=re.IGNORECASE)
        base = re.sub(r"\b\d+\b", " ", base)
        base = re.sub(r"\s+", " ", base).strip(" -_:;,.()[]{}")
        if len(base) >= 5 and re.search(r"[A-Za-zÁÉÍÓÚÑáéíóúñ]", base):
            return base.upper() if base.isupper() else base.title()
        return None

    @classmethod
    def _extraer_nombre_paciente(cls, texto, file_obj=None):
        """Extractor tolerante para nombre/apellido del paciente.
        Acepta etiquetas en una línea, en líneas consecutivas o separadas por '|'.
        """
        t = cls._limpiar_texto(texto or "")
        # 1) Etiquetas explícitas en la misma línea.
        patrones = [
            r"(?:Apellido\s*y\s*Nombre|Nombre\s*y\s*Apellido|Paciente|Nombre\s*Completo|ApyNom|Ape\.?\s*y\s*Nom\.?)\s*(?:[:=\-]|\|)?\s*([^\n\r|]{3,90})",
            r"(?:Patient|Name)\s*(?:[:=\-]|\|)?\s*([^\n\r|]{3,90})",
        ]
        cortes = r"\b(?:Edad|DNI|Documento|Sexo|Fecha|Obra\s*Social|M[eé]dico|Diagn[oó]stico|PAS|PAD|VOP|Vel\.|TCF|Dist\.)\b"
        for pat in patrones:
            for m in re.finditer(pat, t, flags=re.IGNORECASE):
                cand = m.group(1).strip(" :-|,.;")
                cand = re.split(cortes, cand, maxsplit=1, flags=re.IGNORECASE)[0].strip(" :-|,.;")
                cand = re.sub(r"\s+", " ", cand)
                # Evitar capturar el nombre de una variable o institución.
                if 3 <= len(cand) <= 90 and re.search(r"[A-Za-zÁÉÍÓÚÑáéíóúñ]", cand):
                    if not re.search(r"^(edad|sexo|dni|documento|fecha|vel|vop|tcf|dist)", cand, flags=re.IGNORECASE):
                        return cand
        # 2) Etiqueta en una línea y nombre en la siguiente.
        lineas = [x.strip(" :-|,.;") for x in t.splitlines() if x.strip(" :-|,.;")]
        for i, lin in enumerate(lineas[:-1]):
            if re.fullmatch(r"(?:Paciente|Nombre|Apellido\s*y\s*Nombre|Nombre\s*y\s*Apellido|Nombre\s*Completo)", lin, flags=re.IGNORECASE):
                cand = re.sub(r"\s+", " ", lineas[i+1]).strip(" :-|,.;")
                if 3 <= len(cand) <= 90 and re.search(r"[A-Za-zÁÉÍÓÚÑáéíóúñ]", cand):
                    return cand
        # 3) Respaldo desde nombre de archivo.
        if file_obj is not None:
            return cls._nombre_desde_archivo(file_obj)
        return None

    @staticmethod
    def extraer_texto(file_obj) -> str:
        file_obj.seek(0)
        nombre = getattr(file_obj, "name", "") or ""
        es_txt = nombre.lower().endswith(".txt")
        if es_txt:
            raw = file_obj.read()
            if isinstance(raw, bytes):
                for enc in ("utf-8", "latin-1", "cp1252"):
                    try:
                        return GenericReportParser._limpiar_texto(raw.decode(enc, errors="ignore"))
                    except Exception:
                        continue
                return GenericReportParser._limpiar_texto(raw.decode("utf-8", errors="ignore"))
            return GenericReportParser._limpiar_texto(str(raw))

        if _PDF_BACKEND is None:
            raise RuntimeError("Instale 'pdfplumber' o 'pypdf' para leer PDFs. Los archivos TXT no requieren librerías adicionales.")
        file_obj.seek(0)
        txt = ""
        if _PDF_BACKEND == "pdfplumber":
            with pdfplumber.open(file_obj) as pdf:
                for p in pdf.pages:
                    txt += (p.extract_text(x_tolerance=2, y_tolerance=3) or "") + "\n"
                    # Muchas veces los valores están en tablas y extract_text los separa mal.
                    try:
                        for table in p.extract_tables() or []:
                            for row in table or []:
                                vals = [str(c).strip() for c in row if c not in (None, "")]
                                if vals:
                                    txt += " | ".join(vals) + "\n"
                    except Exception:
                        pass
        else:
            r = PdfReader(file_obj)
            for p in r.pages:
                txt += (p.extract_text() or "") + "\n"
        return GenericReportParser._limpiar_texto(txt)

    @classmethod
    def _extraer_pa(cls, texto):
        patrones_par = [
            r"(?:TA|PA|Presi[oó]n\s+Arterial|Brachial\s+BP|BP)\s*[:\-=]?\s*(\d{2,3})\s*[/\\]\s*(\d{2,3})\s*(?:mmHg)?",
            r"(?:Sist[oó]lica|PAS|SBP)[^\d]{0,20}(\d{2,3}).{0,60}(?:Diast[oó]lica|PAD|DBP)[^\d]{0,20}(\d{2,3})",
            r"(\d{2,3})\s*[/\\]\s*(\d{2,3})\s*mmHg",
        ]
        for pat in patrones_par:
            for m in re.finditer(pat, texto, flags=re.IGNORECASE | re.DOTALL):
                pas, pad = int(m.group(1)), int(m.group(2))
                ven = cls._ventana(texto, m.start(), m.end(), 100)
                if any(x in ven for x in ["referencia", "normal", "objetivo", "umbral", "percentil"]):
                    continue
                if 70 <= pas <= 260 and 30 <= pad <= 160 and pas > pad:
                    return pas, pad
        # Etiquetas sueltas
        pas = None; pad = None
        for pat in [r"(?:PAS|TAS|SBP|Sist[oó]lica)\s*[:\-=]?\s*(\d{2,3})"]:
            m = re.search(pat, texto, flags=re.IGNORECASE)
            if m: pas = int(m.group(1))
        for pat in [r"(?:PAD|TAD|DBP|Diast[oó]lica)\s*[:\-=]?\s*(\d{2,3})"]:
            m = re.search(pat, texto, flags=re.IGNORECASE)
            if m: pad = int(m.group(1))
        if pas and pad and 70 <= pas <= 260 and 30 <= pad <= 160 and pas > pad:
            return pas, pad
        return None, None

    @classmethod
    def _buscar_numero_cf(cls, texto, tipo):
        """tipo: distancia, tiempo o vop. Devuelve el mejor valor CF, nunca radial.

        Reglas explícitas del equipo:
        - Dist. Car. Fem. = distancia carótido-femoral REAL (cm).
        - TCF = tiempo de tránsito carótido-femoral (ms/mseg).
        - Vel. Carótida-Femoral = VOP carótido-femoral medida (m/s).
        Estas etiquetas se priorizan por encima de cualquier búsqueda genérica.
        """
        texto = cls._limpiar_texto(texto or "")

        # 1) Etiquetas exactas del informe del equipo: máxima prioridad.
        if tipo == "distancia":
            exactos = [
                # Etiqueta literal del informe: Dist. Car. Fem. = 55 cm.
                # Se aceptan variantes con/sin espacios, puntos, guiones, barras, pipes o salto OCR.
                r"(?is)\bD\s*i\s*s\s*t\s*\.?\s*(?:\||:|;|,|-|\s)*C\s*a\s*r\s*\.?\s*(?:\||:|;|,|-|\s)*F\s*e\s*m\s*\.?\s*(?:=|:|-|\||\s)*([0-9]{2,3}(?:[,.][0-9]+)?)\s*(cm|m)?",
                r"(?is)\bDist\.?\s*Car\.?\s*Fem\.?\s*(?:=|:|-|\||\s)*([0-9]{2,3}(?:[,.][0-9]+)?)\s*(cm|m)?",
                r"(?is)\bDist\.?\s*Car\.?\s*Fem\.?[^0-9\n]{0,60}([0-9]{2,3}(?:[,.][0-9]+)?)\s*(cm|m)?",
                r"(?is)\bDistancia\s*Car[oó]tid[ao]\s*[-/]?\s*Femoral\b[^0-9\n]{0,60}([0-9]{2,3}(?:[,.][0-9]+)?)\s*(cm|m)?",
                r"(?is)\bDistancia\s*Car[oó]tido\s*[-/]?\s*Femoral\s*REAL\b[^0-9\n]{0,60}([0-9]{2,3}(?:[,.][0-9]+)?)\s*(cm|m)?",
            ]
            lo, hi = 20, 200
        elif tipo == "tiempo":
            exactos = [
                r"\bTCF\b\s*(?:=|:|-)?\s*([\d.,]+)\s*(mseg|ms|msec|seg|s)?",
                r"\bTiempo\s*de\s*Tr[aá]nsito\s*CF\b\s*(?:=|:|-)?\s*([\d.,]+)\s*(mseg|ms|msec|seg|s)?",
            ]
            lo, hi = 10, 300
        else:
            exactos = [
                r"(?:Vel\.?\s*Car[oó]tida\s*[-/]?\s*Femoral|Vel\.?\s*Car[oó]tido\s*[-/]?\s*Femoral|Velocidad\s*Car[oó]tida\s*[-/]?\s*Femoral)[^\d\n|]{0,40}([\d.,]+)\s*(?:m\s*/\s*s|m/s)?",
                r"(?:VOP|PWV)\s*(?:Car[oó]tida|Car[oó]tido|Carotid)\s*[-/]?\s*Femoral[^\d\n|]{0,40}([\d.,]+)\s*(?:m\s*/\s*s|m/s)?",
            ]
            lo, hi = 3, 25

        for pat in exactos:
            for m in re.finditer(pat, texto, flags=re.IGNORECASE | re.DOTALL):
                val = cls._norm_num(m.group(1))
                if val is None:
                    continue
                unidad = ""
                if tipo in ("distancia", "tiempo") and m.lastindex and m.lastindex >= 2 and m.group(2):
                    unidad = m.group(2).lower()
                if tipo == "distancia" and unidad == "m":
                    val = val * 100
                if tipo == "tiempo" and unidad in ("s", "seg"):
                    val = val * 1000
                if lo <= val <= hi:
                    return round(val, 2)

        # 2) Búsqueda genérica como respaldo, manteniendo exclusión de radial/braquial.
        if tipo == "distancia":
            patrones = [
                r"(?:distancia\s*car[oó]tido\s*[-/]?\s*femoral|distance|dcf|d\s*cf|c\s*[-/]?f\s*distance|carotid\s*[-/]?femoral\s*distance|car[oó]tido\s*[-/]?femoral\s*distancia)[^\d\n|]{0,40}([\d.,]+)\s*(cm|m)?",
                r"(?:carotid|car[oó]tido|car[oó]tida).{0,50}(?:femoral).{0,50}([\d.,]+)\s*(cm|m)?",
            ]
            lo, hi = 20, 200
        elif tipo == "tiempo":
            patrones = [
                r"(?:tiempo\s*(?:de)?\s*tr[aá]nsito\s*(?:car[oó]tido\s*[-/]?\s*femoral)?|transit\s*time|pulse\s*transit\s*time|ptt|tt|delta\s*t|[Δ∆]\s*t)[^\d\n|]{0,40}([\d.,]+)\s*(ms|mseg|msec|s|seg)?",
                r"(?:carotid|car[oó]tido|car[oó]tida).{0,80}(?:femoral).{0,80}(?:tiempo|transit|tt|ptt).{0,40}([\d.,]+)\s*(ms|mseg|msec|s|seg)?",
            ]
            lo, hi = 10, 300
        else:  # vop cf
            patrones = [
                r"(?:cf\s*[-/]?\s*pwv|cfpwv|pwvcf|pwv\s*cf|vop\s*cf|vop\s*c\s*[-/]?\s*f)[^\d\n|]{0,40}([\d.,]+)\s*(?:m\s*/\s*s|m/s)?",
                r"(?:velocidad\s+de\s+onda\s+de\s+pulso|pulse\s+wave\s+velocity|PWV|VOP).{0,80}(?:car[oó]tido|car[oó]tida|carotid|femoral|cf).{0,40}([\d.,]+)\s*(?:m\s*/\s*s|m/s)?",
                r"(?:car[oó]tido|car[oó]tida|carotid).{0,40}(?:femoral).{0,80}(?:VOP|PWV|velocidad|vel\.)[^\d\n|]{0,40}([\d.,]+)\s*(?:m\s*/\s*s|m/s)?",
            ]
            lo, hi = 3, 25

        candidatos = []
        for pat in patrones:
            for m in re.finditer(pat, texto, flags=re.IGNORECASE | re.DOTALL):
                val = cls._norm_num(m.group(1))
                if val is None:
                    continue
                unidad = m.group(2).lower() if (tipo in ("distancia", "tiempo") and m.lastindex and m.lastindex >= 2 and m.group(2)) else ""
                if tipo == "distancia" and unidad == "m":
                    val = val * 100
                if tipo == "tiempo" and unidad in ("s", "seg"):
                    val = val * 1000
                ven = cls._ventana(texto, m.start(), m.end(), 160)
                if tipo in ("distancia", "tiempo"):
                    if any(x in ven for x in ["radial", "braquial", "brachial", "perifer", "peripheral"]):
                        continue
                    score = 2 if cls._es_cf(ven) else 1
                else:
                    if not cls._es_cf(ven):
                        continue
                    score = 3
                if lo <= val <= hi:
                    candidatos.append((score, m.start(), round(val, 2)))
        if not candidatos:
            # Respaldo ultra tolerante solo para la etiqueta del equipo: Dist. Car. Fem.
            # Recorre línea por línea para evitar capturar otros números del informe.
            if tipo == "distancia":
                for lin in cls._lineas_reales(texto):
                    lin_norm = cls._normalizar_ocr(lin)
                    if re.search(r"(?i)dist\.?\s*car\.?\s*fem\.?", lin_norm) or (re.search(r"(?i)dist", lin_norm) and re.search(r"(?i)car", lin_norm) and re.search(r"(?i)fem", lin_norm)):
                        nums = re.findall(r"(?<![A-Za-z])([0-9]{2,3}(?:[,.][0-9]+)?)(?![A-Za-z])", lin_norm)
                        for num in nums:
                            val = cls._norm_num(num)
                            if val is not None and 20 <= val <= 200:
                                return round(val, 2)
            return None
        candidatos.sort(key=lambda x: (-x[0], x[1]))
        return candidatos[0][2]

    @staticmethod
    def _pixmap_to_png_bytes(page, crop, zoom=2.8):
        """Renderiza una región del PDF a PNG bytes."""
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), clip=crop, alpha=False)
        return pix.tobytes("png")

    @staticmethod
    def _render_page_to_pil(page, zoom=2.5):
        """Renderiza página completa a PIL.Image para localizar la curva cuando el PDF es imagen."""
        if not _PIL_OK:
            return None
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        return img

    @classmethod
    def _capturar_por_pixeles_curva_cf(cls, page):
        """Localiza el panel CF por colores de las curvas roja/verde cuando el texto no es seleccionable.

        Muchos informes guardan el sector de VOP CF como una imagen. En esos casos no hay
        texto buscable, por lo que se renderiza la página y se busca el grupo de píxeles
        rojo/verde de las curvas. Luego se expande el recorte para incluir título, VOP,
        TCF, Dist. Car. Fem. y cantidad de pulsos.
        """
        if not _PIL_OK:
            return None
        try:
            img = cls._render_page_to_pil(page, zoom=2.8)
            if img is None:
                return None
            import numpy as _np
            arr = _np.asarray(img)
            r = arr[:, :, 0].astype(int)
            g = arr[:, :, 1].astype(int)
            b = arr[:, :, 2].astype(int)
            # Curvas típicas: una roja y una verde. Umbrales amplios para capturas de baja calidad.
            mask_red = (r > 135) & (g < 120) & (b < 120) & ((r - g) > 45)
            mask_green = (g > 95) & (r < 140) & (b < 140) & ((g - r) > 25)
            mask = mask_red | mask_green
            ys, xs = _np.where(mask)
            if len(xs) < 80:
                return None

            # Elegir el componente/cluster principal de píxeles de color dentro de una zona razonable.
            # Se evita capturar otros gráficos grandes del informe usando el primer cluster denso.
            # Histograma por bloques para localizar la mayor concentración de curva.
            h, w = mask.shape
            block = max(20, min(h, w) // 30)
            bx = xs // block
            by = ys // block
            keys, counts = _np.unique(by * 10000 + bx, return_counts=True)
            key = int(keys[int(_np.argmax(counts))])
            cy = (key // 10000) * block + block // 2
            cx = (key % 10000) * block + block // 2

            # Tomar píxeles coloreados cercanos al cluster dominante.
            radio_x = max(260, int(w * 0.35))
            radio_y = max(180, int(h * 0.25))
            near = (abs(xs - cx) < radio_x) & (abs(ys - cy) < radio_y)
            if near.sum() < 50:
                near = _np.ones_like(xs, dtype=bool)
            x0, x1 = xs[near].min(), xs[near].max()
            y0, y1 = ys[near].min(), ys[near].max()

            # Expandir para incluir todo el panel como en la imagen propuesta.
            pad_left = int(0.36 * (x1 - x0 + 1)) + 90
            pad_right = int(1.15 * (x1 - x0 + 1)) + 180
            pad_top = int(0.30 * (y1 - y0 + 1)) + 70
            pad_bottom = int(0.60 * (y1 - y0 + 1)) + 110
            crop = (
                max(0, int(x0 - pad_left)),
                max(0, int(y0 - pad_top)),
                min(w, int(x1 + pad_right)),
                min(h, int(y1 + pad_bottom)),
            )

            # Si el recorte queda demasiado grande, priorizar sector superior-izquierdo del panel.
            max_w = int(w * 0.65)
            max_h = int(h * 0.55)
            cx0, cy0, cx1, cy1 = crop
            if (cx1 - cx0) > max_w:
                cx1 = min(w, cx0 + max_w)
            if (cy1 - cy0) > max_h:
                cy1 = min(h, cy0 + max_h)
            crop_img = img.crop((cx0, cy0, cx1, cy1))
            buf = BytesIO()
            crop_img.save(buf, format="PNG")
            return buf.getvalue()
        except Exception:
            return None

    @staticmethod
    def _score_imagen_cf_pil(img):
        """Puntúa si una imagen/crop parece contener el panel de curva CF.
        Usa presencia simultánea de curva roja y verde y tamaño compatible.
        """
        if not _PIL_OK or img is None:
            return 0
        try:
            import numpy as _np
            im = img.convert("RGB")
            # Reducir para acelerar sin perder patrón de color.
            if max(im.size) > 1200:
                im.thumbnail((1200, 1200))
            arr = _np.asarray(im)
            if arr.size == 0:
                return 0
            r = arr[:, :, 0].astype(int)
            g = arr[:, :, 1].astype(int)
            b = arr[:, :, 2].astype(int)
            mask_red = (r > 130) & (g < 135) & (b < 135) & ((r - g) > 35)
            mask_green = (g > 85) & (r < 155) & (b < 155) & ((g - r) > 18)
            red_n = int(mask_red.sum())
            green_n = int(mask_green.sum())
            h, w = arr.shape[:2]
            area = max(1, h * w)
            # Panel esperado: no microscópico, con ambas curvas y proporción razonable.
            color_density = (red_n + green_n) / area
            score = 0
            if red_n > 30 and green_n > 30:
                score += 100
            score += min(80, int(color_density * 20000))
            if 1.0 <= (w / max(1, h)) <= 2.4:
                score += 20
            if w >= 180 and h >= 120:
                score += 20
            return score
        except Exception:
            return 0

    @classmethod
    def _recortar_panel_desde_imagen_pil(cls, img):
        """Recorta el panel CF dentro de una imagen de página o imagen embebida.
        Si no puede ubicar el panel, devuelve la imagen completa.
        """
        if not _PIL_OK or img is None:
            return None
        try:
            import numpy as _np
            im = img.convert("RGB")
            arr = _np.asarray(im)
            r = arr[:, :, 0].astype(int)
            g = arr[:, :, 1].astype(int)
            b = arr[:, :, 2].astype(int)
            mask_red = (r > 130) & (g < 135) & (b < 135) & ((r - g) > 35)
            mask_green = (g > 85) & (r < 155) & (b < 155) & ((g - r) > 18)
            mask = mask_red | mask_green
            ys, xs = _np.where(mask)
            if len(xs) < 60:
                return im
            x0, x1 = int(xs.min()), int(xs.max())
            y0, y1 = int(ys.min()), int(ys.max())
            w, h = im.size
            # Expandir mucho más hacia izquierda/derecha/abajo para incluir título, TCF y Dist. Car. Fem.
            dx = x1 - x0 + 1
            dy = y1 - y0 + 1
            cx0 = max(0, x0 - int(dx * 0.45) - 120)
            cy0 = max(0, y0 - int(dy * 0.55) - 80)
            cx1 = min(w, x1 + int(dx * 1.10) + 220)
            cy1 = min(h, y1 + int(dy * 0.85) + 140)
            crop = im.crop((cx0, cy0, cx1, cy1))
            return crop
        except Exception:
            return img

    @classmethod
    def _capturar_imagen_embebida_cf(cls, doc):
        """Busca imágenes embebidas dentro del PDF y devuelve el panel CF.
        Esto resuelve PDFs donde el sector propuesto es una imagen incrustada y no texto/capa vectorial.
        """
        if not (_FITZ_OK and _PIL_OK):
            return None
        try:
            best_img = None
            best_score = 0
            for page in doc:
                for info in page.get_images(full=True) or []:
                    xref = info[0]
                    try:
                        base = doc.extract_image(xref)
                        data = base.get("image")
                        if not data:
                            continue
                        img = Image.open(BytesIO(data)).convert("RGB")
                        score = cls._score_imagen_cf_pil(img)
                        if score > best_score:
                            best_score = score
                            best_img = img.copy()
                    except Exception:
                        continue
            if best_img is not None and best_score >= 100:
                panel = cls._recortar_panel_desde_imagen_pil(best_img)
                buf = BytesIO()
                panel.save(buf, format="PNG")
                return buf.getvalue()
        except Exception:
            return None
        return None

    @classmethod
    def _capturar_pagina_pdf_respaldo(cls, doc):
        """Respaldo seguro: si no se localiza el panel CF, captura una página completa.

        Esto evita que el informe quede sin imagen cuando el sector gráfico está rasterizado,
        no tiene texto seleccionable, cambió de posición o el detector de colores no encontró
        el recuadro. La prioridad sigue siendo capturar el panel carótido-femoral; este
        respaldo garantiza que al menos quede incorporada la imagen visible del PDF original
        para revisión clínica.
        """
        if doc is None or len(doc) == 0 or not _FITZ_OK:
            return None
        try:
            # Preferir una página con anclas CF si el texto existe; si no, usar la primera.
            idx = 0
            claves = (
                "carot", "femor", "cfpwv", "vop", "pwv",
                "dist. car", "dist car", "tcf", "vel. car"
            )
            for i, page in enumerate(doc):
                try:
                    txt = (page.get_text("text") or "").lower()
                    txt_norm = (txt.replace("ó", "o").replace("í", "i")
                                  .replace("á", "a").replace("é", "e").replace("ú", "u"))
                    if any(k in txt_norm for k in claves):
                        idx = i
                        break
                except Exception:
                    continue
            page = doc[idx]
            pr = page.rect
            # Página completa, con buena resolución pero tamaño razonable.
            return cls._pixmap_to_png_bytes(page, fitz.Rect(pr.x0, pr.y0, pr.x1, pr.y1), zoom=2.1)
        except Exception:
            return None

    @classmethod
    def capturar_curva_cf_pdf(cls, file_obj):
        """Devuelve PNG bytes del sector gráfico VOP carótido-femoral del PDF.

        Orden corregido y más estable:
        1) Primero usa anclas textuales del propio sector (Vel. Carótida-Femoral, TCF,
           Dist. Car. Fem.) y recorta un panel amplio alrededor.
        2) Si el texto no existe porque el informe es imagen, busca imágenes embebidas
           con curvas roja/verde compatibles.
        3) Luego analiza la página renderizada por píxeles de curvas roja/verde.
        4) Como respaldo, captura el cuadrante superior izquierdo de la primera página.
        """
        nombre = getattr(file_obj, "name", "") or ""
        if not nombre.lower().endswith(".pdf") or not _FITZ_OK:
            return None
        doc = None
        try:
            file_obj.seek(0)
            data = file_obj.read()
            doc = fitz.open(stream=data, filetype="pdf")
            claves = [
                "Vel. Carótida-Femoral", "Vel. Carotida-Femoral",
                "Vel Carotida Femoral", "Velocidad Carotida Femoral",
                "Velocidad Carótida Femoral", "carotida-femoral", "carótida-femoral",
                "carotido-femoral", "carótido-femoral", "VOP CF", "cfPWV",
                "TCF", "T C F", "Dist. Car. Fem.", "Dist Car Fem",
                "Dist. Car Fem", "Dist Car. Fem", "Car. Fem"
            ]
            claves_norm = [
                k.lower().replace("ó", "o").replace("í", "i").replace("á", "a")
                 .replace("é", "e").replace("ú", "u") for k in claves
            ]

            # 1) Captura guiada por texto legible. Prioritaria para no traer otro gráfico.
            for page in doc:
                rects = []
                for key in claves:
                    try:
                        for rr in page.search_for(key):
                            rects.append(fitz.Rect(rr))
                    except Exception:
                        pass
                if not rects:
                    for b in (page.get_text("blocks") or []):
                        txt = str(b[4]).lower()
                        norm = (txt.replace("ó", "o").replace("í", "i")
                                   .replace("á", "a").replace("é", "e").replace("ú", "u"))
                        if any(k in norm for k in claves_norm):
                            rects.append(fitz.Rect(b[:4]))
                if rects:
                    r = rects[0]
                    for rr in rects[1:]:
                        r |= rr
                    pr = page.rect
                    # Recorte amplio: incluye curva, título, VOP, TCF, Dist. Car. Fem. y pulsos.
                    crop = fitz.Rect(
                        max(pr.x0, r.x0 - max(145, pr.width * 0.18)),
                        max(pr.y0, r.y0 - max(145, pr.height * 0.16)),
                        min(pr.x1, r.x1 + max(520, pr.width * 0.48)),
                        min(pr.y1, r.y1 + max(300, pr.height * 0.30)),
                    )
                    # Garantizar tamaño de panel aunque el ancla textual sea pequeña.
                    if crop.width < pr.width * 0.52:
                        crop.x1 = min(pr.x1, crop.x0 + pr.width * 0.64)
                    if crop.height < pr.height * 0.32:
                        crop.y1 = min(pr.y1, crop.y0 + pr.height * 0.42)
                    return cls._pixmap_to_png_bytes(page, crop, zoom=3.4)

            # 2) Captura directa de imágenes embebidas si el sector no tiene texto seleccionable.
            png_emb = cls._capturar_imagen_embebida_cf(doc)
            if png_emb:
                return png_emb

            # 3) Captura por detección de curvas roja/verde en página renderizada.
            for page in doc:
                png = cls._capturar_por_pixeles_curva_cf(page)
                if png:
                    return png

            # 4) Fallback final: página completa del PDF.
            # Es preferible incluir la página visible completa antes que dejar el informe sin captura.
            respaldo = cls._capturar_pagina_pdf_respaldo(doc)
            if respaldo:
                return respaldo
            return None
        except Exception:
            return None
        finally:
            try:
                if doc is not None:
                    doc.close()
            except Exception:
                pass

    @classmethod
    def parsear(cls, file_obj) -> dict:
        texto = cls.extraer_texto(file_obj)
        curva_png = cls.capturar_curva_cf_pdf(file_obj)
        out = {"_texto_crudo": texto, "_curva_cf_png": curva_png}

        # Nombre: usar extractor reforzado + respaldo por nombre de archivo.
        out["nombre"] = cls._extraer_nombre_paciente(texto, file_obj)

        for campo in ("documento", "edad", "sexo"):
            v = None
            for pat in cls.PATRONES[campo]:
                m = re.search(pat, texto, flags=re.IGNORECASE)
                if m:
                    v = m.group(1).strip()
                    break
            out[campo] = v

        if out.get("sexo"):
            out["sexo"] = "Femenino" if str(out["sexo"]).upper().startswith("F") else "Masculino"
        if out.get("edad"):
            try:
                out["edad"] = int(out["edad"])
            except ValueError:
                out["edad"] = None

        out["pas"], out["pad"] = cls._extraer_pa(texto)
        out["distancia"] = cls._buscar_numero_cf(texto, "distancia")
        out["tiempo"] = cls._buscar_numero_cf(texto, "tiempo")
        out["vop"] = cls._buscar_numero_cf(texto, "vop")

        if out.get("vop") is None and out.get("distancia") and out.get("tiempo"):
            out["vop_recalculada"] = VascularEngine.calcular_vop(out["distancia"], out["tiempo"])
        else:
            out["vop_recalculada"] = None

        if out.get("nombre"):
            out["nombre"] = re.sub(r"\s+", " ", out["nombre"]).strip(" :-")

        # Validaciones finales: nunca aceptar PAD >= PAS ni VOP radial/periférica como VOP de referencia.
        if out.get("pas") and out.get("pad") and out["pad"] >= out["pas"]:
            out["pas"], out["pad"] = None, None
        return out

# ---------------------------------------------------------------------------
# UTILIDADES
# ---------------------------------------------------------------------------
def safe_latin1(t):
    if t is None:
        return ""
    rep = {"“": '"', "”": '"', "‘": "'", "’": "'", "–": "-", "—": "-",
           "•": "-", "…": "...", "≥": ">=", "≤": "<=", "±": "+/-"}
    for k, v in rep.items():
        t = t.replace(k, v)
    t = unicodedata.normalize("NFKD", t)
    return t.encode("latin-1", "ignore").decode("latin-1")


# ---------------------------------------------------------------------------
# GRÁFICA DIDÁCTICA PROFESIONAL
# ---------------------------------------------------------------------------
def construir_grafico_didactico(edad, sexo, vop, p10, p25, p50, p75, p90, color_paciente, edad_vasc):
    """Genera lámina didáctica con 4 paneles:
    1) Curva de rigidez arterial por edad y sexo.
    2) Comparativa por percentiles.
    3) Fenotipo vascular único.
    4) Comparación edad cronológica vs edad vascular.

    La edad vascular se muestra como gráfico comparativo independiente,
    sin modificar ni mezclar el fenotipo vascular asignado por percentiles de VOP.
    """
    plt.rcParams.update({
        "font.family": "DejaVu Sans",
        "axes.edgecolor": "#34495E",
        "axes.titlesize": 10,
        "axes.titleweight": "bold",
        "axes.labelsize": 8.5,
    })
    fig = plt.figure(figsize=(10.6, 7.0), dpi=180)
    gs = GridSpec(2, 2, width_ratios=[1.45, 1.0], height_ratios=[1.18, 0.90], wspace=0.30, hspace=0.42)

    # Panel 1: curva de rigidez arterial
    ax = fig.add_subplot(gs[0, 0])
    edades = np.linspace(20, 80, 200)
    base = 5.5 if sexo == "Femenino" else 6.0
    cp50 = base + 0.08 * (edades - 20)
    cp10 = cp50 * 0.80
    cp25 = cp50 * 0.90
    cp75 = cp50 * 1.15
    cp90 = cp50 * 1.30

    ax.fill_between(edades, 0, cp10, color=COLOR_SUPERNOVA, alpha=0.12, label="SUPERNOVA (<p10)")
    ax.fill_between(edades, cp10, cp90, color=COLOR_NORMAL, alpha=0.10, label="HVA (p10-p90)")
    ax.fill_between(edades, cp90, cp90 * 1.6, color=COLOR_EVA, alpha=0.15, label="EVA (>p90)")
    ax.axhline(10, color="#8B0000", lw=1.2, ls="--", alpha=0.72)
    ax.text(21, 10.18, "VOP > 10 m/s = estimación vascular de riesgo elevada por rigidez arterial",
            fontsize=6.8, color="#8B0000", fontweight="bold")
    ax.plot(edades, cp50, color="#34495E", lw=1.9, label="p50")
    ax.plot(edades, cp25, color="#34495E", lw=0.9, ls=":", alpha=0.7, label="p25/p75")
    ax.plot(edades, cp75, color="#34495E", lw=0.9, ls=":", alpha=0.7)
    ax.plot(edades, cp10, color=COLOR_SUPERNOVA, lw=1.1, ls="--", alpha=0.85)
    ax.plot(edades, cp90, color=COLOR_EVA, lw=1.1, ls="--", alpha=0.85)
    ax.scatter(edad, vop, color=color_paciente, s=135, zorder=6, edgecolor="white", linewidth=1.8, label="Paciente")
    ax.annotate(f"Paciente\n{edad} a / {vop} m/s",
                xy=(edad, vop), xytext=(min(edad + 4, 72), vop + 1.0),
                fontsize=7.5, fontweight="bold", color=color_paciente,
                arrowprops=dict(arrowstyle="->", color=color_paciente, lw=1.1))
    ax.set_title(f"Curva de rigidez arterial - Sexo {sexo}")
    ax.set_xlabel("Edad cronológica (años)")
    ax.set_ylabel("VOP carotídeo-femoral (m/s)")
    ax.set_xlim(20, 82)
    ax.set_ylim(2.5, max(16, vop + 3))
    ax.grid(True, ls=":", alpha=0.35)
    ax.legend(loc="upper left", fontsize=6.3, framealpha=0.92)

    # Panel 2: percentiles
    ax2 = fig.add_subplot(gs[0, 1])
    labels = ["p10", "p25", "p50", "p75", "p90", "Paciente"]
    valores = [p10, p25, p50, p75, p90, vop]
    colores = [COLOR_SUPERNOVA, "#5DADE2", COLOR_NORMAL, "#E59866", COLOR_EVA, color_paciente]
    bars = ax2.barh(labels, valores, color=colores, edgecolor="white")
    for bar, val in zip(bars, valores):
        ax2.text(val + 0.12, bar.get_y() + bar.get_height() / 2,
                 f"{val}", va="center", fontsize=7.2, fontweight="bold")
    ax2.set_title("Comparativa por percentiles")
    ax2.set_xlabel("VOP (m/s)")
    ax2.invert_yaxis()
    ax2.set_xlim(0, max(valores) * 1.28)
    ax2.grid(True, axis="x", ls=":", alpha=0.35)

    # Panel 3: fenotipo vascular único
    ax3 = fig.add_subplot(gs[1, 0])
    if vop < p10:
        fenotipo_graf = "SUPERNOVA"
        c_fen = COLOR_SUPERNOVA
        rango_txt = "VOP < p10"
    elif vop > p90:
        fenotipo_graf = "EVA"
        c_fen = COLOR_EVA
        rango_txt = "VOP > p90"
    else:
        fenotipo_graf = "HVA"
        c_fen = COLOR_NORMAL
        rango_txt = "p10 <= VOP <= p90"
    ax3.axis("off")
    ax3.set_title("Fenotipo vascular único")
    ax3.text(0.5, 0.65, fenotipo_graf, ha="center", va="center",
             fontsize=24, fontweight="bold", color=c_fen, transform=ax3.transAxes)
    ax3.text(0.5, 0.45, rango_txt, ha="center", va="center",
             fontsize=10, fontweight="bold", color="#34495E", transform=ax3.transAxes)
    ax3.text(0.5, 0.27, f"VOP CF medida: {vop} m/s", ha="center", va="center",
             fontsize=9.5, color="#34495E", transform=ax3.transAxes)
    if vop > 10:
        ax3.text(0.5, 0.10, "Estimación vascular de riesgo elevada por VOP > 10 m/s",
                 ha="center", va="center", fontsize=9.2, fontweight="bold",
                 color=COLOR_EVA, transform=ax3.transAxes)

    # Panel 4: edad cronológica vs edad vascular
    ax4 = fig.add_subplot(gs[1, 1])
    delta = edad_vasc - edad
    if delta > 0:
        c_edad = COLOR_EVA
        delta_txt = f"+{delta} años"
    elif delta < 0:
        c_edad = COLOR_SUPERNOVA
        delta_txt = f"{delta} años"
    else:
        c_edad = COLOR_NORMAL
        delta_txt = "0 años"

    barras = ax4.bar(["Cronológica", "Vascular"], [edad, edad_vasc],
                     color=["#566573", c_edad], width=0.55, edgecolor="white")
    ymax = max(edad, edad_vasc, 30) + 18
    ax4.set_ylim(0, ymax)
    ax4.set_ylabel("Años")
    ax4.set_title("Edad cronológica vs edad vascular")
    ax4.grid(True, axis="y", ls=":", alpha=0.35)
    for bar, val in zip(barras, [edad, edad_vasc]):
        ax4.text(bar.get_x() + bar.get_width() / 2, val + 1.2, f"{val} a",
                 ha="center", fontsize=9, fontweight="bold")
    ax4.text(0.5, ymax * 0.88, f"Diferencia: {delta_txt}", ha="center",
             fontsize=9.2, fontweight="bold", color=c_edad)
    ax4.text(0.5, ymax * 0.78, "Comparación gráfica de edades",
             ha="center", fontsize=6.8, color="#5D6D7E")

    fig.suptitle("Lámina Didáctica - Evaluación de Rigidez Arterial",
                 fontsize=12.5, fontweight="bold", color="#1F4E78", y=0.985)
    buf = BytesIO()
    fig.savefig(buf, format="png", dpi=180, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    buf.seek(0)
    return buf

# ---------------------------------------------------------------------------
# GENERADOR DE PDF PROFESIONAL
# ---------------------------------------------------------------------------
class PDFReport(FPDF):
    """PDF integrado en 3 hojas A4.

    Hoja 1: datos del paciente, resultados, diagnóstico y captura carótido-femoral.
    Hoja 2: lámina didáctica e interpretación clínica.
    Hoja 3: recomendaciones, cierre médico, firma y sello digital con espacio suficiente.
    """
    def __init__(self, profesional="", firma_path=None, sello_path=None):
        super().__init__(orientation="P", unit="mm", format="A4")
        self.profesional = profesional or "Profesional Responsable"
        self.firma_path = firma_path
        self.sello_path = sello_path
        # Se usa salto automático moderado. La distribución principal se controla manualmente
        # para estabilizar el informe integrado en 3 hojas.
        self.set_auto_page_break(auto=True, margin=14)
        self.set_margins(left=12, top=10, right=12)
        self.alias_nb_pages()

    def header(self):
        self.set_fill_color(*COLOR_BG_HEADER)
        self.rect(0, 0, 210, 18, "F")
        self.set_text_color(255, 255, 255)
        self.set_font("Arial", "B", 11.5)
        self.set_xy(12, 4)
        self.cell(0, 5, safe_latin1("LABORATORIO VASCULAR NO INVASIVO"), 0, 1, "L")
        self.set_font("Arial", "I", 7.8)
        self.set_x(12)
        self.cell(0, 4, safe_latin1("Evaluación de Rigidez Arterial y Fenotipado Vascular - EVA/SUPERNOVA"), 0, 1, "L")
        self.set_xy(150, 4)
        self.set_font("Arial", "", 7.2)
        self.cell(48, 4, safe_latin1(f"Folio: VAS-{datetime.datetime.now().strftime('%Y%m%d%H%M')}"), 0, 1, "R")
        self.set_x(150)
        self.cell(48, 4, safe_latin1(f"Fecha: {datetime.date.today().strftime('%d/%m/%Y')}"), 0, 1, "R")
        self.set_text_color(0, 0, 0)
        self.set_y(22)

    def footer(self):
        self.set_y(-12)
        self.set_draw_color(*COLOR_BG_HEADER)
        self.set_line_width(0.25)
        self.line(12, self.get_y(), 198, self.get_y())
        self.set_font("Arial", "I", 6.8)
        self.set_text_color(90, 90, 90)
        self.cell(0, 3.6, safe_latin1("Documento de uso clínico. Interpretación por personal médico calificado."), 0, 1, "C")
        self.cell(0, 3.6, safe_latin1(f"Página {self.page_no()} / {{nb}}"), 0, 0, "C")
        self.set_text_color(0, 0, 0)

    def section_title(self, t):
        self.set_fill_color(*COLOR_BG_HEADER)
        self.set_text_color(255, 255, 255)
        self.set_font("Arial", "B", 8.8)
        self.cell(0, 5.4, safe_latin1("  " + t.upper()), 0, 1, "L", fill=True)
        self.set_text_color(0, 0, 0)
        self.ln(0.8)

    def patient_info(self, datos):
        self.section_title("Datos del Paciente")
        self.set_fill_color(*COLOR_BG_SECTION)
        filas = [
            ("Nombre", datos.get("nombre", "-"), "Documento", datos.get("documento", "-")),
            ("Edad", f"{datos.get('edad', '-')} años", "Sexo", datos.get("sexo", "-")),
            ("Médico solicitante", datos.get("medico_solicitante", "-"),
             "Fecha del estudio", datetime.date.today().strftime("%d/%m/%Y")),
        ]
        for f in filas:
            self.set_font("Arial", "B", 8.2)
            self.cell(34, 5.8, safe_latin1(f[0]), 1, 0, "L", fill=True)
            self.set_font("Arial", "", 8.2)
            self.cell(60, 5.8, safe_latin1(str(f[1])[:42]), 1, 0, "L")
            self.set_font("Arial", "B", 8.2)
            self.cell(39, 5.8, safe_latin1(f[2]), 1, 0, "L", fill=True)
            self.set_font("Arial", "", 8.2)
            self.cell(53, 5.8, safe_latin1(str(f[3])[:38]), 1, 1, "L")
        self.ln(2)

    def resultados_table(self, r):
        self.section_title("Resultados Hemodinámicos")
        self.set_font("Arial", "B", 8.1)
        self.set_fill_color(*COLOR_BG_HEADER)
        self.set_text_color(255, 255, 255)
        for h, w in [("Parámetro", 68), ("Resultado", 44), ("Referencia", 74)]:
            self.cell(w, 5.4, safe_latin1(h), 1, 0, "C", fill=True)
        self.ln()
        self.set_text_color(0, 0, 0)
        filas = [
            ("VOP carótido-femoral", f"{r['vop']} m/s", f"p50: {r['p50']} m/s"),
            ("Presión Arterial", f"{r['pas']}/{r['pad']} mmHg", "< 130/80 mmHg"),
            ("Presión de Pulso", f"{r['pp']} mmHg", "< 50 mmHg"),
            ("Presión Arterial Media", f"{r['pam']} mmHg", "70 - 105 mmHg"),
            ("Percentil 10", f"{r['p10']} m/s", "SUPERNOVA si VOP < p10"),
            ("Percentil 90", f"{r['p90']} m/s", "EVA si VOP > p90"),
            ("Lesión Órgano Blanco", r["lob"], "VOP > 10 m/s"),
            ("Estimación vascular de riesgo", r["riesgo"], "Estratificación clínica"),
        ]
        fill = False
        for f in filas:
            self.set_fill_color(245, 247, 250)
            self.set_font("Arial", "", 7.9)
            self.cell(68, 5.2, safe_latin1(f[0]), 1, 0, "L", fill=fill)
            self.set_font("Arial", "B", 7.9)
            self.cell(44, 5.2, safe_latin1(f[1]), 1, 0, "C", fill=fill)
            self.set_font("Arial", "", 7.9)
            self.cell(74, 5.2, safe_latin1(f[2]), 1, 1, "L", fill=fill)
            fill = not fill
        self.ln(2)

    def diagnostico_box(self, fenotipo, riesgo, vop=None):
        self.section_title("Conclusión Diagnóstica")
        if "EVA" in fenotipo:
            fill = COLOR_BG_ALERT; txt = (155, 30, 20)
        elif "SUPERNOVA" in fenotipo:
            fill = (217, 232, 252); txt = (28, 80, 140)
        else:
            fill = COLOR_BG_OK; txt = (28, 110, 60)
        tiene_lob = vop is not None and vop > 10
        alto = 18 if tiene_lob else 14
        y0 = self.get_y()
        self.set_fill_color(*fill)
        self.set_draw_color(*txt)
        self.set_text_color(*txt)
        self.set_line_width(0.3)
        self.rect(12, y0, 186, alto, "DF")
        self.set_xy(15, y0 + 2.5)
        self.set_font("Arial", "B", 9.8)
        self.cell(0, 5, safe_latin1(f"Fenotipo vascular único: {fenotipo}"), 0, 1, "L")
        self.set_x(15)
        self.set_font("Arial", "B", 8.6)
        if tiene_lob:
            self.cell(0, 4.4, safe_latin1(f"Estimación vascular de riesgo: {riesgo} - elevada por VOP mayor a 10 m/seg"), 0, 1, "L")
        else:
            self.cell(0, 4.4, safe_latin1(f"Estimación vascular de riesgo: {riesgo}"), 0, 1, "L")
        if tiene_lob:
            self.set_x(15)
            self.set_text_color(155, 30, 20)
            self.set_font("Arial", "B", 8.1)
            self.cell(0, 4.2, safe_latin1(f"LOB por rigidez vascular: VOP {vop} m/s > 10 m/s."), 0, 1, "L")
        self.set_y(y0 + alto + 3)
        self.set_text_color(0, 0, 0)
        self.set_draw_color(0, 0, 0)

    def insert_imported_cf_curve(self, image_bytes):
        if not image_bytes:
            return
        self.section_title("Curva Carótido-Femoral Importada del Estudio Original")
        self.set_font("Arial", "", 7.4)
        self.multi_cell(0, 3.8, safe_latin1(
            "Captura del estudio original. TCF: tiempo de tránsito carótido-femoral. Dist. Car. Fem.: distancia carótido-femoral."
        ))
        self.ln(1)
        bio = BytesIO(image_bytes) if isinstance(image_bytes, (bytes, bytearray)) else image_bytes
        bio.seek(0)
        try:
            # En el informe de 3 hojas se prioriza conservar esta imagen en página 1.
            if self.get_y() > 190:
                self.add_page()
            self.image(bio, x=22, w=150)
        except Exception:
            import tempfile, os
            with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tf:
                bio.seek(0)
                tf.write(bio.read())
                tmp_path = tf.name
            try:
                self.image(tmp_path, x=22, w=150)
            finally:
                try:
                    os.unlink(tmp_path)
                except OSError:
                    pass
        self.ln(2)

    def insert_chart(self, image_bytes):
        self.section_title("Lámina Didáctica - Análisis Gráfico")
        image_bytes.seek(0)
        try:
            self.image(image_bytes, x=14, w=180)
        except Exception:
            import tempfile, os
            with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tf:
                image_bytes.seek(0)
                tf.write(image_bytes.read())
                tmp_path = tf.name
            try:
                self.image(tmp_path, x=14, w=180)
            finally:
                try:
                    os.unlink(tmp_path)
                except OSError:
                    pass
        self.ln(2)

    def interpretacion(self, fenotipo, vop, edad, edad_vasc, riesgo):
        self.section_title("Interpretación Clínica")
        self.set_font("Arial", "", 8.6)
        lob_txt = ""
        if vop > 10:
            lob_txt = "\n\nVOP > 10 m/s: lesión de órgano blanco por rigidez arterial."

        if "EVA" in fenotipo:
            def_fen = "Envejecimiento Vascular Acelerado: VOP por encima del percentil 90 para sexo y edad."
        elif "SUPERNOVA" in fenotipo:
            def_fen = "SUPERNOVA / Envejecimiento Vascular Supranormal: VOP carótido-femoral por debajo del percentil 10 para sexo y edad."
        else:
            def_fen = "HVA / Envejecimiento Vascular Saludable: VOP entre el percentil 10 y el percentil 90 para sexo y edad."

        txt = (
            f"La Velocidad de Onda de Pulso (VOP) carótido-femoral medida fue de {vop} m/s. "
            f"El diagnóstico de rigidez vascular se expresa como un fenotipo único, definido por la posición de la VOP frente a percentiles por edad y sexo.\n\n"
            f"Fenotipo vascular único: {fenotipo}.\n"
            f"{def_fen}\n\n"
            f"Estimación vascular de riesgo: {riesgo}."
            f"{lob_txt}"
        )
        self.multi_cell(0, 4.6, safe_latin1(txt))
        self.ln(2)

    def recomendaciones_clinicas(self, recs):
        self.section_title("Recomendaciones Clínicas")
        self.set_font("Arial", "", 8.4)
        for i, r in enumerate((recs or [])[:5], 1):
            self.set_x(12)
            self.multi_cell(0, 4.3, safe_latin1(f"{i}. {r}"))
        self.ln(2)

    def cierre_medico(self, fenotipo, riesgo, vop=None):
        self.section_title("Cierre Médico Integrado")
        self.set_font("Arial", "", 8.6)
        if vop is not None and vop > 10:
            riesgo_txt = (
                f"Estimación vascular de riesgo consignado: {riesgo}. "
                f"La estimación vascular de riesgo se considera elevada por VOP mayor a 10 m/seg."
            )
        else:
            riesgo_txt = f"Estimación vascular de riesgo consignado: {riesgo}."
        txt = (
            f"Informe integrado de rigidez vascular con fenotipo final: {fenotipo}. "
            f"La conducta clínica debe integrarse con presión arterial, antecedentes, daño de órgano blanco y criterio médico tratante. "
            f"{riesgo_txt}"
        )
        self.multi_cell(0, 4.6, safe_latin1(txt))
        self.ln(4)

    def referencias_bibliograficas(self):
        """Referencias compactas y seguras para fpdf/fpdf2.

        Corrección clave:
        - Antes de cada multi_cell se reinicia X al margen izquierdo.
        - Se usa ancho explícito para evitar el error:
          FPDFException: Not enough horizontal space to render a single character.
        - Se limita la cantidad de referencias para preservar el informe de 3 hojas.
        """
        self.section_title("Referencias Bibliográficas")
        self.set_font("Arial", "", 6.4)

        x0 = self.l_margin
        w = self.w - self.l_margin - self.r_margin

        nota = "Nota metodológica: " + NOTA_METODOLOGICA_RIESGO
        self.set_x(x0)
        self.multi_cell(w, 3.2, safe_latin1(nota))
        self.ln(0.6)

        for i, ref in enumerate(REFERENCIAS_BIBLIOGRAFICAS[:6], 1):
            # Si el espacio inferior queda reservado para firma/sello, no forzar más texto.
            if self.get_y() > 195:
                break
            self.set_x(x0)
            ref_txt = safe_latin1(f"{i}. {ref}")
            self.multi_cell(w, 3.1, ref_txt)
        self.ln(1.0)

    def firma(self):
        """Firma y sello con espacio reservado en página 3."""
        # Reservar sector inferior amplio. Si el contenido llega bajo, mover a zona segura.
        if self.get_y() > 205:
            self.set_y(205)
        else:
            self.ln(8)
        y0 = self.get_y()

        # Caja visual de firma, útil para evitar superposiciones y mantener formato profesional.
        self.set_draw_color(210, 210, 210)
        self.set_fill_color(250, 250, 250)
        self.rect(108, y0, 88, 48, "D")

        try:
            if self.firma_path and os.path.exists(self.firma_path):
                self.image(self.firma_path, x=114, y=y0 + 4, w=42)
            if self.sello_path and os.path.exists(self.sello_path):
                self.image(self.sello_path, x=160, y=y0 + 3, w=30)
            if self.firma_path or self.sello_path:
                self.set_y(y0 + 30)
            else:
                self.set_y(y0 + 27)
        except Exception:
            self.set_y(y0 + 27)

        self.set_draw_color(60, 60, 60)
        self.line(112, self.get_y(), 192, self.get_y())
        self.set_font("Arial", "B", 8.8)
        self.set_xy(112, self.get_y() + 1.5)
        self.cell(80, 4.4, safe_latin1(self.profesional), 0, 1, "C")
        self.set_font("Arial", "I", 7.8)
        self.set_x(112)
        self.cell(80, 4.2, safe_latin1("Médico responsable del estudio"), 0, 1, "C")


def construir_pdf(datos, resultados, fenotipo, riesgo, recs, chart_buf, profesional, curva_cf_png=None, firma_path=None, sello_path=None):
    pdf = PDFReport(profesional=profesional, firma_path=firma_path, sello_path=sello_path)

    # Hoja 1: resumen clínico + evidencia original CF.
    pdf.add_page()
    pdf.patient_info(datos)
    pdf.resultados_table(resultados)
    pdf.diagnostico_box(fenotipo, riesgo, vop=resultados.get("vop"))
    pdf.insert_imported_cf_curve(curva_cf_png)

    # Hoja 2: gráfico didáctico + interpretación.
    pdf.add_page()
    pdf.insert_chart(chart_buf)
    pdf.interpretacion(fenotipo, resultados["vop"], resultados["edad"],
                       resultados["edad_vasc"], riesgo)

    # Hoja 3: recomendaciones + cierre + firma y sello con espacio suficiente.
    pdf.add_page()
    pdf.recomendaciones_clinicas(recs)
    pdf.cierre_medico(fenotipo, riesgo, vop=resultados.get("vop"))
    pdf.referencias_bibliograficas()
    pdf.firma()

    # Seguridad: limitar el informe integrado a 3 páginas.
    try:
        while getattr(pdf, "page", 0) > 3:
            pdf.pages.pop(pdf.page, None)
            pdf.page -= 1
    except Exception:
        pass

    out = pdf.output(dest="S")
    if isinstance(out, str):
        return out.encode("latin-1")
    return bytes(out)

# ---------------------------------------------------------------------------
# INTERFAZ STREAMLIT
# ---------------------------------------------------------------------------
def main():
    if "auth" not in st.session_state:
        st.session_state.auth = False

    if not st.session_state.auth:
        st.title("🔐 Acceso Médico")
        if _USERS is None:
            st.error("No se pudo inicializar la gestión de usuarios.")
            st.code(f"Detalle del error: {_USER_ERR}")
            st.markdown(
                "**Diagnóstico rápido:**\n\n"
                "1. Verifique que `users.py` esté en la misma carpeta que `app.py`.\n"
                "2. Verifique su versión de Python: requiere **Python 3.8 o superior**.\n"
                "   En la consola ejecute `python --version`.\n"
                "3. Si la carpeta es de solo lectura (ej. Google Drive sincronizado), "
                "mueva `app.py` y `users.py` a una carpeta local con permisos de escritura."
            )
            return

        tab_login, tab_reg, tab_rec = st.tabs(
            ["Iniciar sesión", "Registrarse", "Recuperar contraseña"]
        )

        # --- Tab 1: Login ---
        with tab_login:
            user = st.text_input("Usuario", key="login_user")
            pw = st.text_input("Contraseña", type="password", key="login_pw")
            if st.button("Ingresar", key="btn_login"):
                ok, role_or_msg = _USERS.verificar(user, pw)
                if ok:
                    st.session_state.auth = True
                    st.session_state.user_role = role_or_msg
                    st.session_state.username = user.strip().lower()
                    st.rerun()
                else:
                    st.error(role_or_msg)

        # --- Tab 2: Registrarse ---
        with tab_reg:
            st.caption("Cree una cuenta médica nueva. La pregunta de "
                       "seguridad le servirá para recuperar la contraseña.")
            r_user = st.text_input("Nuevo usuario (alfanumérico, mín. 3)",
                                   key="reg_user")
            r_nombre = st.text_input("Nombre completo", key="reg_nombre")
            r_matricula = st.text_input("Matrícula profesional (opcional)",
                                        key="reg_matricula")
            r_pw = st.text_input("Contraseña (mín. 8 caracteres)",
                                 type="password", key="reg_pw")
            r_pw2 = st.text_input("Repita la contraseña",
                                  type="password", key="reg_pw2")
            r_q = st.text_input("Pregunta de seguridad",
                                placeholder="Ej.: Nombre de su primera mascota",
                                key="reg_q")
            r_a = st.text_input("Respuesta a la pregunta",
                                key="reg_a")
            st.markdown("**Firma y sello digital (opcional)**")
            st.caption("Estos archivos quedan asociados exclusivamente a este usuario y solo se insertan en sus propios informes PDF.")
            r_firma = st.file_uploader("Cargar firma digital", type=["png", "jpg", "jpeg"], key="reg_firma")
            r_sello = st.file_uploader("Cargar sello digital", type=["png", "jpg", "jpeg"], key="reg_sello")
            if st.button("Crear cuenta", key="btn_reg"):
                if r_pw != r_pw2:
                    st.error("Las contraseñas no coinciden.")
                else:
                    ok, msg = _USERS.registrar(
                        r_user, r_pw, r_q, r_a,
                        role="medico",
                        nombre_completo=r_nombre,
                        matricula=r_matricula,
                    )
                    if ok:
                        try:
                            _guardar_asset_usuario(r_user, r_firma, "firma")
                            _guardar_asset_usuario(r_user, r_sello, "sello")
                        except Exception as exc:
                            st.warning(f"La cuenta fue creada, pero no se pudo guardar firma/sello: {exc}")
                        st.success(msg)
                    else:
                        st.error(msg)

        # --- Tab 3: Recuperar contraseña ---
        with tab_rec:
            st.caption("Ingrese su usuario para ver su pregunta de seguridad. "
                       "Luego responda y elija una contraseña nueva.")
            rc_user = st.text_input("Usuario", key="rec_user")
            if rc_user and _USERS.existe(rc_user):
                pregunta = _USERS.obtener_pregunta(rc_user)
                st.info(f"Pregunta de seguridad: **{pregunta}**")
                rc_a = st.text_input("Su respuesta", key="rec_a")
                rc_pw = st.text_input("Nueva contraseña (mín. 8)",
                                      type="password", key="rec_pw")
                rc_pw2 = st.text_input("Repita la nueva contraseña",
                                       type="password", key="rec_pw2")
                if st.button("Restablecer contraseña", key="btn_rec"):
                    if rc_pw != rc_pw2:
                        st.error("Las contraseñas no coinciden.")
                    else:
                        ok, msg = _USERS.recuperar(rc_user, rc_a, rc_pw)
                        if ok:
                            st.success(msg)
                        else:
                            st.error(msg)
            elif rc_user:
                st.warning("Usuario no encontrado.")
        return

    st.sidebar.title(f"Bienvenido, {st.session_state.user_role.capitalize()}")
    profesional = st.sidebar.text_input("Profesional responsable",
                                        value="Dr. / Dra. ____________________")
    usuario_actual = st.session_state.get("username", "")
    perfil_actual = _perfil_usuario_actual(usuario_actual)
    with st.sidebar.expander("Firma y sello digital del usuario"):
        if perfil_actual.get("firma_path"):
            st.image(perfil_actual["firma_path"], caption="Firma cargada", width=120)
        else:
            st.caption("Sin firma cargada para este usuario.")
        if perfil_actual.get("sello_path"):
            st.image(perfil_actual["sello_path"], caption="Sello cargado", width=100)
        else:
            st.caption("Sin sello cargado para este usuario.")
        up_firma = st.file_uploader("Actualizar firma", type=["png", "jpg", "jpeg"], key="up_firma_usuario")
        up_sello = st.file_uploader("Actualizar sello", type=["png", "jpg", "jpeg"], key="up_sello_usuario")
        if st.button("Guardar firma/sello", key="btn_guardar_assets_usuario"):
            try:
                if up_firma is not None:
                    _guardar_asset_usuario(usuario_actual, up_firma, "firma")
                if up_sello is not None:
                    _guardar_asset_usuario(usuario_actual, up_sello, "sello")
                st.success("Firma/sello guardados para este usuario.")
                st.rerun()
            except Exception as exc:
                st.error(f"No se pudo guardar firma/sello: {exc}")
        perfil_actual = _perfil_usuario_actual(usuario_actual)
    if st.sidebar.button("Cerrar Sesión"):
        st.session_state.auth = False
        st.rerun()

    menu = ["Nuevo Estudio", "Historial y Exportación"]
    if st.session_state.get("user_role") == "admin":
        menu.append("Administrar Usuarios")
    choice = st.sidebar.selectbox("Menú", menu)

    # --- Panel de administración de usuarios ---
    if choice == "Administrar Usuarios":
        st.header("👥 Administración de Usuarios")
        if _USERS is None:
            st.error("Gestión de usuarios no disponible.")
            return
        st.subheader("Usuarios registrados")
        st.write(_USERS.lista_usuarios())

        st.subheader("Resetear contraseña de un usuario")
        target = st.selectbox("Usuario objetivo", _USERS.lista_usuarios(),
                              key="adm_target")
        adm_pw = st.text_input("Su contraseña de admin (confirmación)",
                               type="password", key="adm_pw")
        new_pw = st.text_input("Nueva contraseña (mín. 8 caracteres)",
                               type="password", key="adm_newpw")
        if st.button("Aplicar reset", key="btn_admreset"):
            ok, msg = _USERS.admin_reset(
                st.session_state.get("username", "admin"),
                adm_pw, target, new_pw,
            )
            if ok:
                st.success(msg)
            else:
                st.error(msg)
        return

    if choice == "Nuevo Estudio":
        st.header("📋 Registro de Evaluación Vascular")

        # ------- Importador PDF  -------
        with st.expander("📤 Importar mediciones desde archivo PDF o TXT", expanded=True):
            st.caption("Suba el informe PDF o el archivo TXT sin formato del equipo de medición. "
                       "Los campos se completarán automáticamente cuando el texto sea reconocible.")
            pdf_in = st.file_uploader("Seleccionar archivo PDF o TXT", type=["pdf", "txt"],
                                      key="importacion_upload")
            if pdf_in is not None:
                es_txt = str(getattr(pdf_in, "name", "")).lower().endswith(".txt")
                if _PDF_BACKEND is None and not es_txt:
                    st.error("Falta lector de PDF. Instale: `pip install pdfplumber`. Los archivos TXT pueden importarse sin lector PDF.")
                else:
                    try:
                        # Solo parsear si es un archivo nuevo (cambio de file_id)
                        new_fid = getattr(pdf_in, "file_id", pdf_in.name)
                        if st.session_state.get("importacion_fid") != new_fid:
                            datos = GenericReportParser.parsear(pdf_in)
                            st.session_state["importacion_datos"] = datos
                            st.session_state["importacion_fid"] = new_fid
                            # ESCRIBIR DIRECTO A session_state con las keys del form.
                            # Primero se limpian valores previos para no arrastrar 50/70/0 u otros datos de otro archivo.
                            mapeo = {
                                "nombre": "f_nombre",
                                "documento": "f_documento",
                                "edad": "f_edad",
                                "sexo": "f_sexo",
                                "pas": "f_pas",
                                "pad": "f_pad",
                                "distancia": "f_distancia",
                                "tiempo": "f_tiempo",
                                "vop": "f_vop_medida",
                            }
                            vacios = {
                                "f_nombre": "", "f_documento": "", "f_edad": 50,
                                "f_sexo": "Masculino", "f_pas": 0, "f_pad": 0,
                                "f_distancia": 0.0, "f_tiempo": 0.0, "f_vop_medida": 0.0,
                            }
                            for kform, v0 in vacios.items():
                                st.session_state[kform] = v0
                            for kparser, kform in mapeo.items():
                                v = datos.get(kparser)
                                if v is None or v == "":
                                    continue
                                st.session_state[kform] = v
                            st.rerun()

                        datos = st.session_state["importacion_datos"]
                        detectados = {k: v for k, v in datos.items()
                                      if not k.startswith("_") and v is not None}
                        if detectados:
                            st.success(f"Se detectaron {len(detectados)} campos del archivo. "
                                       "Verifique antes de generar el diagnóstico.")
                            st.json(detectados)
                            # Mostrar campos faltantes
                            todos = ["nombre","documento","edad","sexo","pas","pad",
                                     "distancia","tiempo","vop"]
                            faltantes = [c for c in todos if not datos.get(c)]
                            if faltantes:
                                st.warning(
                                    "Campos no detectados automáticamente desde el archivo original (complételos "
                                    "manualmente en el formulario): " + ", ".join(faltantes)
                                )
                            st.info("Criterio aplicado: la VOP de referencia es la carotídeo-femoral/cfPWV. La VOP radial o periférica se ignora para el diagnóstico.")
                        else:
                            st.warning("No se reconocieron campos. Verifique manualmente.")
                        if datos.get("_curva_cf_png"):
                            st.image(datos["_curva_cf_png"], caption="Sector carótido-femoral capturado del PDF original", use_container_width=False)
                        elif str(getattr(pdf_in, "name", "")).lower().endswith(".pdf"):
                            st.warning(
                                "No se pudo capturar automáticamente el sector gráfico carótido-femoral. "
                                "Verifique dependencias en requirements.txt: PyMuPDF y Pillow. "
                                "Si el PDF es una imagen no detectable, la app intentará incorporar una captura de página completa como respaldo."
                            )
                        st.markdown("**Texto extraído del archivo (primeros 4000 caracteres)**")
                        st.text_area(
                            "Texto extraído",
                            value=datos.get("_texto_crudo", "")[:4000],
                            height=220,
                            key="txt_extraido_importacion",
                            label_visibility="collapsed",
                        )
                    except Exception as exc:
                        st.error(f"Error al leer el archivo: {exc}")

        ram = st.session_state.get("importacion_datos", {}) or {}

        # Inicializar defaults en session_state SOLO si aun no existen
        defaults = {
            "f_nombre": "", "f_documento": "", "f_edad": 50,
            "f_sexo": "Masculino", "f_medsol": "",
            # No usar valores fisiológicos ficticios como default: si el PDF no importa, queda 0 y se obliga a revisar.
            "f_pas": 0, "f_pad": 0,
            "f_distancia": 0.0, "f_tiempo": 0.0,
            "f_vop_medida": 0.0,
        }
        for k, v in defaults.items():
            st.session_state.setdefault(k, v)

        # La VOP cf medida del archivo queda como valor primario.
        # Si el PDF no la detecta, el usuario puede cargarla manualmente abajo.
        usar_vop_pdf = True

        with st.form("estudio_form", clear_on_submit=False):
            col1, col2 = st.columns(2)
            with col1:
                st.text_input("Nombre Completo del Paciente", key="f_nombre")
                st.text_input("Documento / Identificación", key="f_documento")
                st.number_input("Edad", min_value=1, max_value=120, step=1, key="f_edad")
                st.radio("Sexo", ["Masculino", "Femenino"], key="f_sexo",
                         horizontal=True)
                st.text_input("Médico solicitante", key="f_medsol")
            with col2:
                st.number_input("Presión Sistólica (PAS, mmHg)",
                                min_value=0, max_value=260, step=1, key="f_pas")
                st.number_input("Presión Diastólica (PAD, mmHg)",
                                min_value=0, max_value=160, step=1, key="f_pad")
                st.number_input("Distancia Carótido-Femoral REAL importada del archivo (cm) = Dist. Car. Fem.",
                                min_value=0.0, max_value=200.0, step=0.5,
                                format="%.1f", key="f_distancia")
                st.number_input("Tiempo de Tránsito CF importado del archivo (ms) — TCF",
                                min_value=0.0, max_value=500.0, step=0.5,
                                format="%.1f", key="f_tiempo")
                st.number_input("VOP CARÓTIDO-FEMORAL medida del archivo original (m/s)",
                                min_value=0.0, max_value=25.0, step=0.1,
                                format="%.2f", key="f_vop_medida",
                                help="Dato primario importado desde el archivo original. Debe ser VOP carotídeo-femoral/cfPWV; no se usa VOP radial.")
                if ram.get("vop"):
                    st.caption(f"VOP cf medida importada del archivo: **{ram['vop']} m/s** (valor primario)")
                elif ram.get("vop_recalculada"):
                    st.caption(f"No se detectó VOP medida en el archivo. Respaldo calculado desde distancia/tiempo: **{ram['vop_recalculada']} m/s**")

            submitted = st.form_submit_button("Calcular y Generar Diagnóstico")

        # Recuperar valores del session_state al momento del submit
        nombre = st.session_state["f_nombre"]
        documento = st.session_state["f_documento"]
        edad = int(st.session_state["f_edad"])
        sexo = st.session_state["f_sexo"]
        medico_solicitante = st.session_state["f_medsol"]
        pas = int(st.session_state["f_pas"])
        pad = int(st.session_state["f_pad"])
        distancia = float(st.session_state["f_distancia"])
        tiempo = float(st.session_state["f_tiempo"])
        vop_medida = float(st.session_state.get("f_vop_medida", 0.0) or 0.0)

        if submitted:
            # Eco diagnostico: muestra los valores capturados
            with st.expander("Verificacion de datos capturados", expanded=False):
                st.write({
                    "edad": edad, "sexo": sexo, "pas": pas, "pad": pad,
                    "distancia_cm": distancia, "tiempo_ms": tiempo,
                    "vop_medida_cf": vop_medida,
                })
            errores = []
            if distancia <= 0:
                errores.append("No se importó la distancia carótido-femoral real desde el archivo. Revise el archivo original o cargue el valor manualmente.")
            if tiempo <= 0:
                errores.append("No se importó el tiempo de tránsito CF real desde el archivo. Revise el archivo original o cargue el valor manualmente.")
            if pas <= pad:
                errores.append("La PAS debe ser mayor que la PAD. Revise la importación del archivo: PAS/PAD parecen invertidas o mal leídas.")
            if vop_medida and not (3 <= vop_medida <= 25):
                errores.append("La VOP cf medida debe estar entre 3 y 25 m/s.")
            if not (1 <= edad <= 120):
                errores.append("Edad fuera de rango (1-120).")
            if errores:
                for er in errores:
                    st.error(er)
                st.stop()

            e = VascularEngine()
            vop_calc = e.calcular_vop(distancia, tiempo)
            # La VOP cf medida del archivo/campo manual es el dato primario.
            # El recálculo desde distancia/tiempo se muestra solo como control técnico.
            if vop_medida > 0:
                vop = round(vop_medida, 2)
                st.info(f"Se usa la VOP cf medida del archivo original: {vop} m/s. "
                        f"Control por distancia/tiempo: {vop_calc} m/s")
            elif ram.get("vop_recalculada"):
                vop = float(ram["vop_recalculada"])
                st.warning(f"No se detectó VOP cf medida. Se usa VOP recalculada desde distancia/tiempo: {vop} m/s")
            else:
                vop = vop_calc
                st.warning(f"No se detectó VOP cf medida en el archivo. Se usa recálculo desde distancia/tiempo: {vop} m/s")
            p10, p25, p50, p75, p90 = e.obtener_percentiles(edad, sexo)
            fenotipo, color, _ = e.clasificar_fenotipo(vop, p10, p90)
            edad_vasc = e.edad_vascular(vop, sexo)
            pp = e.presion_pulso(pas, pad)
            pam = e.presion_arterial_media(pas, pad)
            riesgo = e.estimacion_vascular_riesgo_global(vop, pas, pad, edad)
            lob = "SI (VOP > 10 m/s)" if vop > 10 else "NO"
            recs = e.recomendaciones(fenotipo, riesgo, vop, pas, pad)

            st.markdown(f"### Resultado: <span style='color:{color}'>{fenotipo}</span>",
                        unsafe_allow_html=True)
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("VOP medida", f"{vop} m/s")
            c2.metric("Percentil 50", f"{p50} m/s")
            c3.metric("Fenotipo vascular", fenotipo.split(" (")[0])
            c4.metric("Estimación vascular", riesgo)
            c5, c6, c7 = st.columns(3)
            c5.metric("Presion de Pulso", f"{pp} mmHg")
            c6.metric("PAM", f"{pam} mmHg")
            c7.metric("LOB (VOP>10)", lob)

            try:
                chart_buf = construir_grafico_didactico(
                    edad, sexo, vop, p10, p25, p50, p75, p90, color, edad_vasc)
                try:
                    st.image(chart_buf, use_container_width=True)
                except TypeError:
                    chart_buf.seek(0)
                    st.image(chart_buf, use_column_width=True)
                chart_buf.seek(0)

                datos = {"nombre": nombre or "Sin nombre",
                         "documento": documento,
                         "edad": edad, "sexo": sexo,
                         "medico_solicitante": medico_solicitante}
                res = {"vop": vop, "pas": pas, "pad": pad, "pp": pp, "pam": pam,
                       "p10": p10, "p50": p50, "p90": p90, "lob": lob,
                       "edad": edad, "edad_vasc": edad_vasc, "riesgo": riesgo}
                perfil_pdf = _perfil_usuario_actual(st.session_state.get("username", ""))
                pdf_bytes = construir_pdf(datos, res, fenotipo, riesgo, recs,
                                          chart_buf, profesional, ram.get("_curva_cf_png"),
                                          firma_path=perfil_pdf.get("firma_path"),
                                          sello_path=perfil_pdf.get("sello_path"))

                st.download_button(
                    "Descargar Informe PDF Profesional",
                    data=pdf_bytes,
                    file_name=f"Informe_Vascular_{(nombre or 'paciente').replace(' ', '_')}_"
                              f"{datetime.date.today()}.pdf",
                    mime="application/pdf",
                )
            except Exception as exc:
                st.error(f"Error generando el informe: {exc}")
                with st.expander("Detalle tecnico"):
                    st.code(traceback.format_exc())
                st.stop()

            usuario_sesion = st.session_state.get("username", "")
            rol_sesion = st.session_state.get("user_role", "medico")
            fuente_vop = "VOP cf medida" if vop_medida > 0 else "VOP cf recalculada desde distancia/tiempo"
            registro = {
                "timestamp": datetime.datetime.now().isoformat(timespec="seconds"),
                "fecha_estudio": str(datetime.date.today()),
                "usuario": usuario_sesion,
                "usuario_id": _usuario_seguro(usuario_sesion),
                "rol": rol_sesion,
                "paciente": nombre,
                "documento": documento,
                "edad": edad,
                "sexo": sexo,
                "medico_solicitante": medico_solicitante,
                "vop_cf_ms": vop,
                "distancia_cf_cm": distancia,
                "tiempo_transito_cf_ms": tiempo,
                "pas_mmhg": pas,
                "pad_mmhg": pad,
                "pam_mmhg": pam,
                "pp_mmhg": pp,
                "p10_ms": p10,
                "p25_ms": p25,
                "p50_ms": p50,
                "p75_ms": p75,
                "p90_ms": p90,
                "edad_vascular": edad_vasc,
                "fenotipo_vascular_unico": fenotipo,
                "estimacion_vascular_riesgo": riesgo,
                "lob_vop_mayor_10": "SI" if vop > 10 else "NO",
                "estimacion_vascular_elevada_por_vop_mayor_10": "SI" if vop > 10 else "NO",
                "fuente_vop": fuente_vop,
                "archivo_importado": st.session_state.get("archivo_importado_nombre", ""),
            }
            try:
                _guardar_registro_paciente(registro)
                st.success("Registro guardado en la base histórica del usuario.")
            except Exception as exc:
                st.warning(f"El informe se generó, pero no se pudo guardar el registro histórico: {exc}")

    elif choice == "Historial y Exportación":
        rol = st.session_state.get("user_role", "medico")
        usuario = st.session_state.get("username", "")
        if rol == "admin":
            st.header("Gestión de Datos - Administrador")
            st.caption("Vista global: incluye pacientes registrados por todos los usuarios.")
        else:
            st.header("Mis pacientes - Exportación Excel")
            st.caption("Vista individual: incluye solo los pacientes cargados por el usuario actual.")

        df = _df_registros(usuario, rol)
        if not df.empty:
            try:
                st.dataframe(df, use_container_width=True)
            except TypeError:
                st.dataframe(df)

            excel_data = _excel_bytes_registros(df, "Pacientes_VOP")
            if rol == "admin":
                label = "Excel: exportar pacientes de TODOS los usuarios"
                fname = f"Base_VOP_Todos_los_Usuarios_{datetime.date.today()}.xlsx"
            else:
                label = "Excel: exportar mis pacientes"
                fname = f"Base_VOP_{_usuario_seguro(usuario)}_{datetime.date.today()}.xlsx"
            st.download_button(
                label=label,
                data=excel_data,
                file_name=fname,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

            if rol == "admin":
                st.subheader("Resumen por usuario")
                resumen = (
                    df.groupby("usuario", dropna=False)
                    .agg(pacientes=("paciente", "count"), vop_promedio=("vop_cf_ms", "mean"))
                    .reset_index()
                )
                try:
                    st.dataframe(resumen, use_container_width=True)
                except TypeError:
                    st.dataframe(resumen)
                st.download_button(
                    label="Excel: exportar resumen por usuario",
                    data=_excel_bytes_registros(resumen, "Resumen_Usuarios"),
                    file_name=f"Resumen_VOP_Usuarios_{datetime.date.today()}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
        else:
            st.info("No hay registros disponibles para exportar.")


if __name__ == "__main__":
    main()

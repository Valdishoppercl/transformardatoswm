import streamlit as st
import pandas as pd
import zipfile
import io
import re
from openpyxl import load_workbook

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Transformador de Indicadores WM", page_icon="📊", layout="wide")

# --- FUNCIONES DE APOYO ---
def parse_filters_text(text: str):
    if not isinstance(text, str): return {}
    t = text.replace("\n", " ").replace("\r", " ")
    def m_int(pattern):
        m = re.search(pattern, t, re.I)
        return int(m.group(1)) if m else None
    return {
        "Local ID": m_int(r"(?:local|sala|nodo)\s*(?:es|=|:)?\s*(\d{1,6})"),
        "Semana":   m_int(r"semana\s*(?:es|=|:)?\s*(\d{1,2})"),
        "Año":      m_int(r"año\s*(?:es|=|:)?\s*(\d{4})"),
    }

def to_number(v):
    if v is None or v == "" or pd.isna(v): return pd.NA
    if isinstance(v, (int, float)): return float(v)
    s = str(v).strip().replace("$", "").replace("%", "").replace(" ", "").replace(".", "").replace(",", ".")
    try: return float(s)
    except: return pd.NA

def dia_to_date(day_val, year_val):
    if pd.isna(day_val) or str(day_val).strip().lower() == "total":
        return pd.NaT
    # Intentar conversión directa
    dt = pd.to_datetime(day_val, dayfirst=True, errors="coerce")
    # Si el Excel lo trae como año 1900, forzar el año correcto del metadato
    if pd.notna(dt) and year_val and dt.year == 1900:
        try: return dt.replace(year=int(year_val))
        except: return dt
    return dt

# --- INTERFAZ STREAMLIT ---
st.title("📊 Transformación de Indicadores WM")
st.markdown("Sube tu archivo ZIP y obtén los reportes procesados y agrupados.")

uploaded_file = st.file_uploader("Cargar archivo ZIP con reportes Excel", type="zip")

if uploaded_file is not None:
    if st.button("🚀 Iniciar Transformación"):
        try:
            by_key = {}
            PERCENT_COLS = ["Armado a Tiempo","OTEA","NPS","NSG","Completitud","Same Day","N2H","Contactos Perfectos","Participación","Variación %","Reclamos","Desviación","TEP"]
            FINAL_SCHEMA = ["Local ID","Año","Mes","Semana","Dia","Pedidos Facturados"] + PERCENT_COLS
            
            with zipfile.ZipFile(uploaded_file, 'r') as z:
                all_files = [f for f in z.namelist() if f.lower().endswith(".xlsx") and "indicadores" in f.lower()]
                
                if not all_files:
                    st.error("No se encontraron archivos válidos.")
                else:
                    progress_bar = st.progress(0)
                    for i, fn in enumerate(all_files):
                        with z.open(fn) as f:
                            content = f.read()
                            # 1. Metadatos con Openpyxl
                            meta_wb = load_workbook(io.BytesIO(content), data_only=True, read_only=True)
                            ws = meta_wb["Export"] if "Export" in meta_wb.sheetnames else meta_wb.worksheets[0]
                            meta = {}
                            for r in range(1, 50):
                                val = ws.cell(r, 1).value
                                if isinstance(val, str) and "filtros aplicados:" in val.lower():
                                    meta = parse_filters_text(val)
                                    break
                            meta_wb.close()

                            # 2. Datos con Pandas
                            df = pd.read_excel(io.BytesIO(content), sheet_name="Export", header=0)
                            
                            # Identificar columnas por posición real (evita el error de Unnamed)
                            # Tradicionalmente: Col 0=Año, 1=Mes, 2=Semana, 3=Dia
                            df.columns.values[0] = "Año"
                            df.columns.values[1] = "Mes"
                            df.columns.values[2] = "Semana"
                            df.columns.values[3] = "Dia"

                            # Limpiar filas basura
                            df = df[df["Dia"].notna()].copy()
                            df = df[df["Dia"].astype(str).str.lower() != "total"]
                            
                            # Convertir Año a numérico antes de procesar fechas
                            df["Año"] = pd.to_numeric(df["Año"], errors="coerce").fillna(meta.get("Año", 0))
                            
                            # PROCESAR FECHAS (DIA)
                            local_id = meta.get("Local ID", "XX")
                            df["Dia"] = df.apply(lambda row: dia_to_date(row["Dia"], row["Año"]), axis=1)
                            
                            # Eliminar si después de la conversión el Día sigue siendo nulo
                            df = df[df["Dia"].notna()].copy()
                            
                            df.insert(0, "Local ID", local_id)

                            # Limpiar resto de columnas numéricas
                            for c in df.columns:
                                if c not in ["Local ID", "Año", "Mes", "Semana", "Dia"]:
                                    df[c] = df[c].apply(to_number)
                            
                            for c in FINAL_SCHEMA:
                                if c not in df.columns: df[c] = pd.NA
                            
                            # Normalizar porcentajes
                            for c in PERCENT_COLS:
                                if c in df.columns:
                                    s = pd.to_numeric(df[c], errors="coerce")
                                    if s.dropna().gt(1).any(): df[c] = s / 100.0

                            df = df[FINAL_SCHEMA].copy()

                            # Agrupar
                            key = (str(local_id), str(int(meta.get("Año", 0))), str(meta.get("Semana", "XX")))
                            if key not in by_key:
                                by_key[key] = df
                            else:
                                by_key[key] = pd.concat([by_key[key], df], ignore_index=True).drop_duplicates(subset=["Dia", "Local ID"])
                        
                        progress_bar.progress((i + 1) / len(all_files))

            # --- GENERAR SALIDA ---
            if by_key:
                out_buf = io.BytesIO()
                with zipfile.ZipFile(out_buf, "w") as new_z:
                    for (loc, yr, wk), fdf in by_key.items():
                        excel_buf = io.BytesIO()
                        # Formato de fecha explícito para Excel
                        with pd.ExcelWriter(excel_buf, engine='openpyxl', datetime_format="DD/MM/YYYY") as writer:
                            fdf.to_excel(writer, index=False, sheet_name="Datos")
                        new_z.writestr(f"Local_{loc}_S{wk}_{yr}.xlsx", excel_buf.getvalue())
                
                st.success("✅ ¡Proceso Exitoso!")
                st.download_button("📥 Descargar ZIP", out_buf.getvalue(), "Reportes_WM.zip", "application/zip")
                
        except Exception as e:
            st.error(f"Error en el proceso: {e}")


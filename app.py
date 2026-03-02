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
    if not isinstance(text, str):
        return {}
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
    if v is None or v == "" or pd.isna(v):
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace("$", "").replace("%", "").replace(" ", "").replace(".", "").replace(",", ".")
    try:
        return float(s)
    except:
        return 0.0

def dia_to_date(day_val, year_val):
    if pd.isna(day_val) or str(day_val).strip().lower() == "total":
        return pd.NaT
    dt = pd.to_datetime(day_val, dayfirst=True, errors="coerce")
    if pd.notna(dt) and year_val and dt.year == 1900:
        try: return dt.replace(year=int(year_val))
        except: return dt
    return dt

def infer_week_label(df):
    s = pd.to_numeric(df["Semana"], errors="coerce").dropna()
    if s.empty: return "YY"
    mn, mx = int(s.min()), int(s.max())
    return f"{mn:02d}" if mn == mx else f"{mn:02d}-{mx:02d}"

def infer_year_label(df):
    y = pd.to_numeric(df["Año"], errors="coerce").dropna()
    if y.empty: return "YYYY"
    mn, mx = int(y.min()), int(y.max())
    return str(mn) if mn == mx else f"{mn}-{mx}"

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
                    st.error("No se encontraron archivos válidos dentro del ZIP.")
                else:
                    progress_bar = st.progress(0)
                    for i, fn in enumerate(all_files):
                        with z.open(fn) as f:
                            content = f.read()
                            # 1. Metadatos con Openpyxl
                            meta_wb = load_workbook(io.BytesIO(content), data_only=True, read_only=True)
                            ws = meta_wb["Export"] if "Export" in meta_wb.sheetnames else meta_wb.worksheets[0]
                            meta = {}
                            for r in range(1, 100):
                                v = ws.cell(r, 1).value
                                if isinstance(v, str) and v.strip().lower().startswith("filtros aplicados:"):
                                    meta = parse_filters_text(v)
                                    break
                            meta_wb.close()
                            
                            # 2. Transformar con Pandas
                            df = pd.read_excel(io.BytesIO(content), sheet_name="Export", header=0)
                            
                            # Forzamos los nombres de las primeras 4 columnas por posición (Blindaje)
                            df.columns.values[0] = "Año"
                            df.columns.values[1] = "Mes"
                            df.columns.values[2] = "Semana"
                            df.columns.values[3] = "Dia"
                            
                            # Limpieza de filas de "Total" y vacías
                            df = df[df["Dia"].notna()].copy()
                            df = df[~df["Dia"].astype(str).str.lower().str.contains("total", na=False)]
                            
                            # Procesar Año y Local ID
                            anio_meta = meta.get("Año", 2026)
                            df["Año"] = pd.to_numeric(df["Año"], errors="coerce").fillna(anio_meta)
                            local_id = meta.get("Local ID", "Desconocido")
                            df.insert(0, "Local ID", local_id)
                            
                            # PROCESAR FECHAS (DIA) - Usamos apply para pasar el año de la fila
                            df["Dia"] = df.apply(lambda row: dia_to_date(row["Dia"], row["Año"]), axis=1)
                            
                            # Convertir resto de columnas a número
                            for c in df.columns:
                                if c not in ["Local ID", "Año", "Mes", "Semana", "Dia"]:
                                    df[c] = df[c].apply(to_number)
                            
                            # Asegurar que existan todas las columnas del esquema
                            for c in FINAL_SCHEMA:
                                if c not in df.columns: df[c] = 0.0
                            
                            # Normalizar porcentajes
                            for c in PERCENT_COLS:
                                if df[c].gt(1).any(): df[c] = df[c] / 100.0
                            
                            df = df[FINAL_SCHEMA].copy()
                            
                            # Agrupar
                            week_label = infer_week_label(df)
                            year_label = infer_year_label(df)
                            key = (str(local_id), year_label, week_label)
                            
                            if key not in by_key:
                                by_key[key] = df
                            else:
                                by_key[key] = pd.concat([by_key[key], df], ignore_index=True).drop_duplicates(subset=["Dia", "Local ID"])
                        
                        progress_bar.progress((i + 1) / len(all_files))

            # --- GENERAR ZIP DE SALIDA ---
            if by_key:
                output_zip_buffer = io.BytesIO()
                with zipfile.ZipFile(output_zip_buffer, "w", zipfile.ZIP_DEFLATED) as new_zip:
                    for (loc, yr, wk), final_df in by_key.items():
                        excel_buf = io.BytesIO()
                        with pd.ExcelWriter(excel_buf, engine='openpyxl', datetime_format="DD/MM/YYYY") as writer:
                            final_df.to_excel(writer, index=False, sheet_name="Datos")
                        new_zip.writestr(f"Local{loc}_Año{yr}_Semanas{wk}.xlsx", excel_buf.getvalue())
                
                st.success(f"✅ ¡Hecho! Se procesaron {len(by_key)} grupos de locales.")
                st.download_button(
                    label="📥 Descargar Resultados (ZIP)",
                    data=output_zip_buffer.getvalue(),
                    file_name="Reportes_Procesados_WM.zip",
                    mime="application/zip"
                )
                # Vista previa para confirmar que "Dia" tiene datos
                st.subheader("Vista previa del último proceso:")
                st.dataframe(df.head(5))
        except Exception as e:
            st.error(f"Se produjo un error durante el proceso: {e}")

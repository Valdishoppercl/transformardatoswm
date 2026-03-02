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
    if v is None or v == "" or pd.isna(v): return 0.0
    if isinstance(v, (int, float)): return float(v)
    s = str(v).strip().replace("$", "").replace("%", "").replace(" ", "").replace(".", "").replace(",", ".")
    try: return float(s)
    except: return 0.0

def dia_to_date(day_val, year_val):
    if pd.isna(day_val): return pd.NaT
    # Limpieza de valores que no son fechas
    s_val = str(day_val).strip().lower()
    if s_val in ["dia", "total", "none", ""]: return pd.NaT
    
    dt = pd.to_datetime(day_val, dayfirst=True, errors="coerce")
    # Si el año es 1900, aplicar el año detectado en metadatos
    if pd.notna(dt) and year_val and dt.year == 1900:
        try: return dt.replace(year=int(year_val))
        except: return dt
    return dt

# --- INTERFAZ STREAMLIT ---
st.title("📊 Transformación de Indicadores WM")
st.markdown("Sube tu archivo ZIP para procesar y agrupar los indicadores correctamente.")

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
                            # 1. Leer Metadatos
                            meta_wb = load_workbook(io.BytesIO(content), data_only=True, read_only=True)
                            ws = meta_wb["Export"] if "Export" in meta_wb.sheetnames else meta_wb.worksheets[0]
                            meta = {}
                            for r in range(1, 50):
                                v = ws.cell(r, 1).value
                                if isinstance(v, str) and "filtros aplicados:" in v.lower():
                                    meta = parse_filters_text(v)
                                    break
                            meta_wb.close()
                            
                            # 2. Cargar con Pandas buscando la tabla real
                            df_raw = pd.read_excel(io.BytesIO(content), header=None)
                            
                            # Buscar la fila donde realmente empiezan los datos (evita filas extras)
                            start_idx = 0
                            for idx, row in df_raw.iterrows():
                                if any(str(val).strip().lower() == "kpi" for val in row):
                                    start_idx = idx + 1
                                    break
                            
                            # Re-leer desde la fila detectada
                            df = pd.read_excel(io.BytesIO(content), skiprows=start_idx)
                            
                            # Forzar nombres por posición para evitar errores de nombres
                            df.columns.values[0] = "Año"
                            df.columns.values[1] = "Mes"
                            df.columns.values[2] = "Semana"
                            df.columns.values[3] = "Dia"
                            
                            # LIMPIEZA CRÍTICA: Eliminar filas de texto extra o totales
                            df = df[df["Dia"].notna()].copy()
                            df = df[~df["Dia"].astype(str).str.lower().str.contains("total|dia|mes|semana", na=False)]
                            
                            # Procesar Año y Local ID
                            anio_meta = meta.get("Año", 2026)
                            df["Año"] = pd.to_numeric(df["Año"], errors="coerce").fillna(anio_meta)
                            local_id = meta.get("Local ID", "Desconocido")
                            df.insert(0, "Local ID", local_id)
                            
                            # PROCESAR FECHAS (DIA)
                            df["Dia"] = df.apply(lambda row: dia_to_date(row["Dia"], row["Año"]), axis=1)
                            
                            # Asegurar que no queden filas sin fecha válida
                            df = df[df["Dia"].notna()].copy()

                            # Convertir indicadores a números
                            for c in df.columns:
                                if c not in ["Local ID", "Año", "Mes", "Semana", "Dia"]:
                                    df[c] = df[c].apply(to_number)
                                    if c in PERCENT_COLS and df[c].gt(1).any():
                                        df[c] = df[c] / 100.0
                            
                            # Garantizar el esquema final
                            for c in FINAL_SCHEMA:
                                if c not in df.columns: df[c] = 0.0
                            
                            df_final = df[FINAL_SCHEMA].copy()
                            
                            # Agrupar por Local y Semana (usando etiquetas de metadatos)
                            key = (str(local_id), str(anio_meta), str(meta.get("Semana", "XX")))
                            
                            if key not in by_key:
                                by_key[key] = df_final
                            else:
                                by_key[key] = pd.concat([by_key[key], df_final]).drop_duplicates(subset=["Dia", "Local ID"])
                        
                        progress_bar.progress((i + 1) / len(all_files))

            # --- GENERAR ZIP ---
            if by_key:
                out_zip = io.BytesIO()
                with zipfile.ZipFile(out_zip, "w") as new_z:
                    for (loc, yr, wk), fdf in by_key.items():
                        excel_buf = io.BytesIO()
                        with pd.ExcelWriter(excel_buf, engine='openpyxl', datetime_format="DD/MM/YYYY") as writer:
                            fdf.to_excel(writer, index=False)
                        new_z.writestr(f"Local_{loc}_Año{yr}_S{wk}.xlsx", excel_buf.getvalue())
                
                st.success(f"✅ Procesados {len(by_key)} grupos de datos.")
                st.download_button("📥 Descargar Resultados", out_zip.getvalue(), "Reportes_WM.zip", "application/zip")
                
                st.subheader("Vista previa del proceso (sin filas extra):")
                st.dataframe(df_final.head(10)) # Aquí verás que el 'Dia' ahora tiene fecha
            else:
                st.warning("No se encontraron datos válidos.")
        except Exception as e:
            st.error(f"❌ Error: {e}")

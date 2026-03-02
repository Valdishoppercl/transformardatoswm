import streamlit as st
import pandas as pd
import zipfile
import io
import re
from openpyxl import load_workbook

st.set_page_config(page_title="Transformador WM", page_icon="📊", layout="wide")

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
    dt = pd.to_datetime(day_val, dayfirst=True, errors='coerce')
    if pd.notna(dt) and year_val and dt.year == 1900:
        try: return dt.replace(year=int(year_val))
        except: return dt
    return dt

# --- INTERFAZ ---
st.title("📊 Transformación de Indicadores WM")
uploaded_file = st.file_uploader("Cargar archivo ZIP", type="zip")

if uploaded_file is not None:
    if st.button("🚀 Iniciar Transformación"):
        try:
            by_key = {}
            PERCENT_COLS = ["Armado a Tiempo","OTEA","NPS","NSG","Completitud","Same Day","N2H","Contactos Perfectos","Participación","Variación %","Reclamos","Desviación","TEP"]
            FINAL_SCHEMA = ["Local ID","Año","Mes","Semana","Dia","Pedidos Facturados"] + PERCENT_COLS
            
            with zipfile.ZipFile(uploaded_file, 'r') as z:
                all_files = [f for f in z.namelist() if f.lower().endswith(".xlsx") and not f.startswith('__') and not '/.' in f]
                
                if not all_files:
                    st.error("No se encontraron archivos .xlsx válidos.")
                else:
                    progress_bar = st.progress(0)
                    for i, fn in enumerate(all_files):
                        with z.open(fn) as f:
                            content = f.read()
                            # 1. Metadatos
                            meta_wb = load_workbook(io.BytesIO(content), data_only=True, read_only=True)
                            ws = meta_wb.active
                            meta = {}
                            for r in range(1, 100):
                                val = ws.cell(r, 1).value
                                if val and "filtros aplicados:" in str(val).lower():
                                    meta = parse_filters_text(str(val))
                                    break
                            meta_wb.close()

                            # 2. Cargar Datos (lectura inicial para encontrar la tabla)
                            df_raw = pd.read_excel(io.BytesIO(content), header=None)
                            
                            # Buscar la fila del encabezado (donde diga KPI o Año)
                            start_row = 0
                            for idx, row in df_raw.iterrows():
                                row_values = row.astype(str).str.lower().tolist()
                                if any("kpi" in s or "año" in s for s in row_values):
                                    start_row = idx + 1
                                    break
                            
                            # Re-leer con el encabezado correcto
                            df = pd.read_excel(io.BytesIO(content), skiprows=start_row)
                            if df.empty: continue

                            # Forzar nombres de columnas por posición
                            df.columns.values[0] = "Año"
                            df.columns.values[1] = "Mes"
                            df.columns.values[2] = "Semana"
                            df.columns.values[3] = "Dia"

                            # --- SOLUCIÓN AL ERROR DE 'SERIES' ---
                            # Usamos .str.contains para manejar la columna como texto correctamente
                            df = df[df["Dia"].notna()].copy()
                            df = df[~df["Dia"].astype(str).str.lower().str.contains("total", na=False)]

                            l_id = meta.get("Local ID", "Sin_ID")
                            anio_def = meta.get("Año", 2024)
                            sem_def = meta.get("Semana", "XX")

                            df.insert(0, "Local ID", l_id)
                            df["Año"] = pd.to_numeric(df["Año"], errors="coerce").fillna(anio_def)
                            df["Dia"] = df.apply(lambda row: dia_to_date(row["Dia"], row["Año"]), axis=1)

                            # Asegurar columnas numéricas
                            for col in FINAL_SCHEMA:
                                if col not in df.columns:
                                    df[col] = 0.0
                                elif col not in ["Local ID", "Año", "Mes", "Semana", "Dia"]:
                                    df[col] = df[col].apply(to_number)
                                    if col in PERCENT_COLS and df[col].gt(1).any():
                                        df[col] = df[col] / 100.0

                            df_final = df[FINAL_SCHEMA].copy()

                            key = (str(l_id), str(sem_def))
                            if key not in by_key:
                                by_key[key] = df_final
                            else:
                                by_key[key] = pd.concat([by_key[key], df_final]).drop_duplicates()
                        
                        progress_bar.progress((i + 1) / len(all_files))

            if by_key:
                out_buf = io.BytesIO()
                with zipfile.ZipFile(out_buf, "w") as new_z:
                    for (loc, sem), fdf in by_key.items():
                        if not fdf.empty:
                            excel_buf = io.BytesIO()
                            with pd.ExcelWriter(excel_buf, engine='openpyxl') as writer:
                                fdf.to_excel(writer, index=False)
                            new_z.writestr(f"Local_{loc}_Semana_{sem}.xlsx", excel_buf.getvalue())
                
                st.success(f"✅ Procesados {len(by_key)} grupos de datos.")
                st.download_button("📥 Descargar ZIP", out_buf.getvalue(), "Reportes_WM.zip", "application/zip")
                st.subheader("Vista previa de datos procesados:")
                st.dataframe(df_final.head(10)) 
            else:
                st.warning("No se encontraron datos para procesar.")

        except Exception as e:
            st.error(f"❌ Error crítico: {e}")


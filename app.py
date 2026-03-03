import streamlit as st
import pandas as pd
import zipfile
import io
import re
from openpyxl import load_workbook

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Transformador de Indicadores WM", page_icon="📊", layout="wide")

# --- MAPEO DE MESES ---
MESES_MAP = {
    'enero': 1, 'febrero': 2, 'marzo': 3, 'abril': 4, 'mayo': 5, 'junio': 6,
    'julio': 7, 'agosto': 8, 'septiembre': 9, 'octubre': 10, 'noviembre': 11, 'diciembre': 12
}

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

def construir_fecha(row):
    try:
        anio = int(to_number(row['Año']))
        mes_str = str(row['Mes']).strip().lower()
        mes = MESES_MAP.get(mes_str, 1)
        
        # Extraer solo el número del día (por si viene como "lunes 1")
        dia_match = re.search(r'(\d+)', str(row['Dia']))
        dia = int(dia_match.group(1)) if dia_match else 1
        
        return pd.Timestamp(year=anio, month=mes, day=dia)
    except:
        return pd.NaT

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
                all_files = [fn for fn in z.namelist() if fn.lower().endswith(".xlsx") and "indicadores" in fn.lower()]
                
                if not all_files:
                    st.error("No se encontraron archivos válidos.")
                else:
                    progress_bar = st.progress(0)
                    for i, fn in enumerate(all_files):
                        with z.open(fn) as f:
                            content = f.read()
                            # 1. Metadatos
                            meta_wb = load_workbook(io.BytesIO(content), data_only=True, read_only=True)
                            ws = meta_wb["Export"] if "Export" in meta_wb.sheetnames else meta_wb.worksheets[0]
                            meta = {}
                            for r in range(1, 40):
                                v = ws.cell(r, 1).value
                                if v and "filtros aplicados:" in str(v).lower():
                                    meta = parse_filters_text(str(v))
                                    break
                            meta_wb.close()

                            # 2. Búsqueda segura del encabezado (Soluciona el error 'Series')
                            df_raw = pd.read_excel(io.BytesIO(content), header=None)
                            header_idx = 0
                            for idx, row in df_raw.iterrows():
                                row_text = " ".join(row.dropna().astype(str)).lower()
                                if "kpi" in row_text or "año" in row_text:
                                    header_idx = idx
                                    break
                            
                            # 3. Lectura de datos reales
                            df = pd.read_excel(io.BytesIO(content), skiprows=header_idx)
                            if df.empty: continue

                            # Forzar nombres por posición
                            df.columns.values[0] = "Año"
                            df.columns.values[1] = "Mes"
                            df.columns.values[2] = "Semana"
                            df.columns.values[3] = "Dia"

                            # Limpiar filas basura (encabezados de texto)
                            df = df[df["Dia"].notna()].copy()
                            df = df[~df["Dia"].astype(str).str.lower().str.contains("total|dia|mes|semana", na=False)]

                            # Inyectar Local ID
                            local_id = meta.get("Local ID", "Sin_ID")
                            df.insert(0, "Local ID", local_id)

                            # RECONSTRUCCIÓN DE FECHA
                            df["Dia"] = df.apply(construir_fecha, axis=1)
                            
                            # Si después de reconstruir sigue sin fecha, eliminar esa fila
                            df = df[df["Dia"].notna()].copy()

                            # Normalizar números e indicadores
                            for c in df.columns:
                                if c not in ["Local ID", "Año", "Mes", "Semana", "Dia"]:
                                    df[c] = df[c].apply(to_number)
                                    if c in PERCENT_COLS and df[c].gt(1).any():
                                        df[c] = df[c] / 100.0
                            
                            for c in FINAL_SCHEMA:
                                if c not in df.columns: df[c] = 0.0

                            df_final = df[FINAL_SCHEMA].copy()
                            
                            # Agrupar por Local y Semana
                            key = (str(local_id), str(meta.get("Año", 2026)), str(meta.get("Semana", "XX")))
                            if key not in by_key:
                                by_key[key] = df_final
                            else:
                                by_key[key] = pd.concat([by_key[key], df_final]).drop_duplicates(subset=["Dia", "Local ID"])
                        
                        progress_bar.progress((i + 1) / len(all_files))

            if by_key:
                out_zip = io.BytesIO()
                with zipfile.ZipFile(out_zip, "w") as new_z:
                    for (loc, yr, wk), fdf in by_key.items():
                        excel_buf = io.BytesIO()
                        with pd.ExcelWriter(excel_buf, engine='openpyxl', datetime_format="DD/MM/YYYY") as writer:
                            fdf.to_excel(writer, index=False)
                        new_z.writestr(f"Local{loc}_Año{yr}_Semanas{wk}.xlsx", excel_buf.getvalue())
                
                st.success(f"✅ Procesados {len(by_key)} locales con datos.")
                st.download_button("📥 Descargar Resultados", out_zip.getvalue(), "Reportes_WM.zip", "application/zip")
                st.dataframe(df_final.head(10)) 
            else:
                st.warning("No se encontraron tablas de datos válidas.")
        except Exception as e:
            st.error(f"❌ Error crítico: {e}")


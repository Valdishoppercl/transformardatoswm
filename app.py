import streamlit as st



import pandas as pd



import zipfile



import io



import os



import re



import shutil



from openpyxl import load_workbook







# --- CONFIGURACIÓN DE PÁGINA ---



st.set_page_config(page_title="Transformador de Indicadores WM", page_icon="📊", layout="wide")







# --- FUNCIONES DE APOYO (TU LÓGICA ORIGINAL) ---



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



        return pd.NA



    if isinstance(v, (int, float)):



        return float(v)



    s = str(v).strip().replace("$", "").replace("%", "").replace(" ", "")



    s = s.replace(".", "").replace(",", ".")



    try:



        return float(s)



    except:



        return pd.NA







def dia_to_date(day_val, year_val):



    if day_val is None or day_val == "" or pd.isna(day_val):



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



            



            # Procesar el ZIP en memoria



            with zipfile.ZipFile(uploaded_file, 'r') as z:



                # Filtrar solo archivos válidos



                all_files = [f for f in z.namelist() if f.lower().endswith(".xlsx") and "indicadores operaciones sod" in f.lower()]



                



                if not all_files:



                    st.error("No se encontraron archivos con el nombre 'indicadores operaciones sod' dentro del ZIP.")



                else:



                    progress_bar = st.progress(0)



                    for i, fn in enumerate(all_files):



                        with z.open(fn) as f:



                            # 1. Leer Metadatos (necesitamos cargarlo dos veces: una con openpyxl para filtros y otra con pandas)



                            content = f.read()



                            meta_wb = load_workbook(io.BytesIO(content), data_only=True, read_only=True)



                            ws = meta_wb["Export"] if "Export" in meta_wb.sheetnames else meta_wb.worksheets[0]



                            meta = {}



                            for r in range(1, 100): # Escaneo rápido de filas superiores



                                v = ws.cell(r, 1).value



                                if isinstance(v, str) and v.strip().lower().startswith("filtros aplicados:"):



                                    meta = parse_filters_text(v)



                                    break



                            meta_wb.close()



                            



                            # 2. Transformar con Pandas



                            df = pd.read_excel(io.BytesIO(content), sheet_name="Export", header=0)



                            df = df.rename(columns={"KPI": "Año", "Unnamed: 1": "Mes", "Unnamed: 2": "Semana", "Unnamed: 3": "Dia"})



                            



                            # Limpieza



                            df = df[df["Dia"].notna() & (df["Dia"].astype(str).str.lower() != "total")]



                            df["Año"] = pd.to_numeric(df["Año"], errors="coerce")



                            df = df[df["Año"].notna()].copy()



                            



                            local_id = meta.get("Local ID", "Desconocido")



                            df.insert(0, "Local ID", local_id)



                            df["Dia"] = [dia_to_date(d, y) for d, y in zip(df["Dia"], df["Año"])]



                            



                            for c in df.columns:



                                if c not in ["Local ID", "Año", "Mes", "Semana", "Dia"]:



                                    df[c] = df[c].apply(to_number)



                            



                            for c in FINAL_SCHEMA:



                                if c not in df.columns: df[c] = pd.NA



                            



                            # Ajuste de porcentajes



                            for c in PERCENT_COLS:



                                s = pd.to_numeric(df[c], errors="coerce")



                                if s.dropna().gt(1).any(): df[c] = s / 100.0



                            



                            df = df[FINAL_SCHEMA].copy()



                            



                            # Agrupar por Local/Año/Semana



                            week_label = infer_week_label(df)



                            year_label = infer_year_label(df)



                            key = (str(local_id), year_label, week_label)



                            



                            if key not in by_key:



                                by_key[key] = df



                            else:



                                by_key[key] = pd.concat([by_key[key], df], ignore_index=True).drop_duplicates(subset=["Local ID", "Año", "Mes", "Semana", "Dia"])



                        



                        progress_bar.progress((i + 1) / len(all_files))







            # --- GENERAR ZIP DE SALIDA ---



            if by_key:



                output_zip_buffer = io.BytesIO()



                with zipfile.ZipFile(output_zip_buffer, "w", zipfile.ZIP_DEFLATED) as new_zip:



                    for (loc, yr, wk), final_df in by_key.items():



                        excel_buf = io.BytesIO()



                        with pd.ExcelWriter(excel_buf, engine='openpyxl', datetime_format="DD/MM/YYYY") as writer:



                            final_df.to_excel(writer, index=False, sheet_name="Sheet1")



                        



                        file_name = f"Local{loc}_Año{yr}_Semanas{wk}.xlsx"



                        new_zip.writestr(file_name, excel_buf.getvalue())



                



                st.success(f"✅ ¡Hecho! Se procesaron {len(by_key)} grupos de locales.")



                st.download_button(



                    label="📥 Descargar Resultados (ZIP)",



                    data=output_zip_buffer.getvalue(),



                    file_name="Reportes_Procesados_WM.zip",



                    mime="application/zip"



                )



        except Exception as e:



            st.error(f"Se produjo un error durante el proceso: {e}")


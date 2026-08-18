import io
import pandas as pd
import requests
import streamlit as st
from docx import Document
from docx.shared import Pt
from fpdf import FPDF

# Configuração da página
st.set_page_config(page_title="Portal PNCP - Rio das Pedras", layout="wide")

st.title("🏛️ Contratações de Rio das Pedras/SP")

# --- BARRA LATERAL ---
tipo_consulta = st.sidebar.selectbox("Selecione:", ["Contratos", "Atas de Registro de Preços"])
data_inicio = st.sidebar.date_input("Data Inicial", value=pd.to_datetime("2026-01-01"))
data_fim = st.sidebar.date_input("Data Final", value=pd.to_datetime("2026-08-18"))

if st.sidebar.button("Gerar Relatório"):
    url = "https://pncp.gov.br/api/consulta/v1/contratos" if tipo_consulta == "Contratos" else "https://pncp.gov.br/api/consulta/v1/atas"
    
    with st.spinner("Consultando..."):
        params = {"cnpj": "44826840000183", "dataInicial": data_inicio.strftime("%Y%m%d"), "dataFinal": data_fim.strftime("%Y%m%d")}
        resp = requests.get(url, params=params, timeout=30)
        
        if resp.status_code == 200:
            df = pd.DataFrame(resp.json() if isinstance(resp.json(), list) else resp.json().get("data", []))
            st.dataframe(df, use_container_width=True)
            
            # --- BOTÕES DE DOWNLOAD ---
            st.markdown("### 📥 Baixar Relatórios")
            cols = st.columns(4)
            
            # 1. EXCEL
            buffer_xlsx = io.BytesIO()
            df.to_excel(buffer_xlsx, index=False)
            cols[0].download_button("Excel (.xlsx)", buffer_xlsx.getvalue(), "dados.xlsx")
            
            # 2. CSV
            cols[1].download_button("CSV (.csv)", df.to_csv(index=False), "dados.csv")
            
            # 3. WORD
            doc = Document()
            doc.add_heading(f"Relatório {tipo_consulta}", 0)
            for _, row in df.head(20).iterrows():
                doc.add_paragraph(str(row.to_dict()))
            buffer_docx = io.BytesIO()
            doc.save(buffer_docx)
            cols[2].download_button("Word (.docx)", buffer_docx.getvalue(), "relatorio.docx")
            
            # 4. PDF (Simples)
            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("Arial", size=10)
            pdf.cell(200, 10, txt=f"Relatorio {tipo_consulta}", ln=True, align='C')
            for index, row in df.head(20).iterrows():
                pdf.cell(200, 10, txt=str(row.values[0:2]), ln=True)
            
            buffer_pdf = io.BytesIO(pdf.output(dest='S').encode('latin-1'))
            cols[3].download_button("PDF (.pdf)", buffer_pdf.getvalue(), "relatorio.pdf")
            
        else:
            st.error("Erro na API.")

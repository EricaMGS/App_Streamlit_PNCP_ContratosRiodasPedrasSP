import io
import pandas as pd
import requests
import streamlit as st
from docx import Document
from fpdf import FPDF

# Configuração da página
st.set_page_config(page_title="Portal PNCP - Rio das Pedras/SP", layout="wide")

st.title("🏛️ Contratações de Rio das Pedras/SP")
st.markdown("Consulta integrada de **Contratos**, **Atas** e **Editais**.")

# --- BARRA LATERAL ---
st.sidebar.header("Parâmetros da Consulta")
tipo_consulta = st.sidebar.selectbox("Selecione o tipo de dado:", 
                                     ["Contratos", "Atas de Registro de Preços", "Editais e Avisos de Contratação"])
data_inicio = st.sidebar.date_input("Data Inicial", value=pd.to_datetime("2026-01-01"))
data_fim = st.sidebar.date_input("Data Final", value=pd.to_datetime("2026-08-18"))

# --- GERENCIAMENTO DE ESTADO ---
# Inicializa o session_state para salvar os dados da consulta
if 'df_resultado' not in st.session_state:
    st.session_state.df_resultado = None

if st.sidebar.button("Gerar Relatório"):
    endpoints = {
        "Contratos": "https://pncp.gov.br/api/consulta/v1/contratos",
        "Atas de Registro de Preços": "https://pncp.gov.br/api/consulta/v1/atas",
        "Editais e Avisos de Contratação": "https://pncp.gov.br/api/consulta/v1/contratacoes/publicacao"
    }
    
    with st.spinner("Consultando..."):
        params = {"cnpj": "44826840000183", "dataInicial": data_inicio.strftime("%Y%m%d"), "dataFinal": data_fim.strftime("%Y%m%d")}
        resp = requests.get(endpoints[tipo_consulta], params=params, timeout=45)
        
        if resp.status_code == 200:
            data = resp.json()
            lote = data if isinstance(data, list) else data.get("data", data.get("items", []))
            st.session_state.df_resultado = pd.DataFrame(lote) if lote else pd.DataFrame()
        else:
            st.error(f"Erro na API: {resp.status_code}")

# --- EXIBIÇÃO PERSISTENTE ---
if st.session_state.df_resultado is not None and not st.session_state.df_resultado.empty:
    df = st.session_state.df_resultado
    st.success(f"Dados carregados: {len(df)} registros.")
    st.dataframe(df, use_container_width=True)

    st.markdown("### 📥 Opções de Exportação (Disponíveis)")
    cols = st.columns(4)
    nome = tipo_consulta.replace(" ", "_")

    # 1. EXCEL
    buffer_xlsx = io.BytesIO()
    df.to_excel(buffer_xlsx, index=False)
    cols[0].download_button("📊 Excel (.xlsx)", buffer_xlsx.getvalue(), f"{nome}.xlsx")

    # 2. CSV
    cols[1].download_button("📄 CSV (.csv)", df.to_csv(index=False), f"{nome}.csv")

    # 3. WORD
    doc = Document()
    doc.add_heading(f"Relatório: {tipo_consulta}", 0)
    for _, row in df.head(50).iterrows():
        doc.add_paragraph(str(row.to_dict()))
    buffer_docx = io.BytesIO()
    doc.save(buffer_docx)
    cols[2].download_button("📝 Word (.docx)", buffer_docx.getvalue(), f"{nome}.docx")

    # 4. PDF
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=10)
    pdf.cell(200, 10, txt=f"Relatório {tipo_consulta}", ln=True, align='C')
    buffer_pdf = io.BytesIO(pdf.output(dest='S').encode('latin-1'))
    cols[3].download_button("📕 PDF (.pdf)", buffer_pdf.getvalue(), f"{nome}.pdf")
elif st.session_state.df_resultado is not None and st.session_state.df_resultado.empty:
    st.warning("Nenhum dado encontrado.")

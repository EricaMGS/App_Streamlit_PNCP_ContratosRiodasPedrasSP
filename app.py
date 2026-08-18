import io
import pandas as pd
import requests
import streamlit as st
from docx import Document
from fpdf import FPDF

st.set_page_config(page_title="Portal PNCP - Rio das Pedras/SP", layout="wide")

st.title("🏛️ Contratações de Rio das Pedras/SP")
st.markdown("Consulta integrada de Contratos, Atas e Editais direto do PNCP.")

# --- BARRA LATERAL ---
st.sidebar.header("Parâmetros da Consulta")
tipo_consulta = st.sidebar.selectbox("Selecione:", [
    "Contratos", 
    "Atas de Registro de Preços", 
    "Editais e Avisos de Contratação"
])

data_inicio = st.sidebar.date_input("Data Inicial", value=pd.to_datetime("2026-01-01"))
data_fim = st.sidebar.date_input("Data Final", value=pd.to_datetime("2026-08-18"))

# Validação do limite de 365 dias do PNCP
if (data_fim - data_inicio).days > 365:
    st.sidebar.error("⚠️ O período não pode ser maior que 365 dias.")
    st.stop()

# Resetar dados se o usuário mudar o tipo de consulta
if 'tipo_anterior' not in st.session_state:
    st.session_state.tipo_anterior = tipo_consulta

if st.session_state.tipo_anterior != tipo_consulta:
    st.session_state.df_resultado = None
    st.session_state.tipo_anterior = tipo_consulta

if 'df_resultado' not in st.session_state:
    st.session_state.df_resultado = None

if st.sidebar.button("Gerar Relatório"):
    endpoints = {
        "Contratos": "https://pncp.gov.br/api/consulta/v1/contratos",
        "Atas de Registro de Preços": "https://pncp.gov.br/api/consulta/v1/atas",
        "Editais e Avisos de Contratação": "https://pncp.gov.br/api/consulta/v1/contratacoes/publicacao"
    }
    
    with st.spinner("Consultando dados no PNCP..."):
        cnpj_alvo = "44826840000183"
        params = {
            "cnpj": cnpj_alvo, 
            "dataInicial": data_inicio.strftime("%Y%m%d"), 
            "dataFinal": data_fim.strftime("%Y%m%d")
        }
        
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
            "Accept": "application/json"
        }
        
        try:
            resp = requests.get(endpoints[tipo_consulta], params=params, headers=headers, timeout=45)
            
            if resp.status_code == 200:
                data = resp.json()
                lote = data if isinstance(data, list) else data.get("data", data.get("items", []))
                
                if lote:
                    df_temp = pd.DataFrame(lote)
                    
                    # --- FILTRAGEM RÍGIDA PARA RIO DAS PEDRAS ---
                    if "orgaoEntidade" in df_temp.columns:
                        df_temp = df_temp[df_temp["orgaoEntidade"].astype(str).str.contains(cnpj_alvo, na=False)]
                    
                    st.session_state.df_resultado = df_temp
                    
                    if df_temp.empty:
                        st.warning("Nenhum registro encontrado especificamente para o CNPJ de Rio das Pedras neste período.")
                else:
                    st.session_state.df_resultado = pd.DataFrame()
                    st.warning("Nenhum registro retornado pelo servidor.")
            else:
                st.error(f"Erro na API (Status {resp.status_code}): {resp.text}")
                if resp.status_code == 400:
                    st.info("💡 A rota de Editais pode exigir parâmetros adicionais na API do PNCP. Tente consultar 'Contratos' ou 'Atas'.")
                st.session_state.df_resultado = None
        except Exception as e:
            st.error(f"Erro de conexão: {e}")

# --- EXIBIÇÃO PERSISTENTE DOS DADOS E BOTÕES ---
if st.session_state.df_resultado is not None and not st.session_state.df_resultado.empty:
    df = st.session_state.df_resultado
    st.success(f"Exibindo {len(df)} registros para Rio das Pedras/SP.")
    st.dataframe(df, use_container_width=True)
    
    st.markdown("### 📥 Opções de Exportação")
    cols = st.columns(4)
    nome = tipo_consulta.replace(" ", "_")

    # 1. Excel
    buffer_xlsx = io.BytesIO()
    df.to_excel(buffer_xlsx, index=False)
    cols[0].download_button("📊 Excel (.xlsx)", buffer_xlsx.getvalue(), f"{nome}_Rio_Das_Pedras.xlsx")
    
    # 2. CSV
    cols[1].download_button("📄 CSV (.csv)", df.to_csv(index=False), f"{nome}_Rio_Das_Pedras.csv")
    
    # 3. Word
    doc = Document()
    doc.add_heading(f"Relatório {tipo_consulta}", 0)
    doc.add_paragraph(f"Município: Rio das Pedras/SP")
    doc.add_paragraph(f"Total de registros: {len(df)}")
    for _, row in df.head(50).iterrows():
        doc.add_paragraph(str(row.to_dict()))
    buffer_docx = io.BytesIO()
    doc.save(buffer_docx)
    cols[2].download_button("📝 Word (.docx)", buffer_docx.getvalue(), f"Relatorio_{nome}.docx")
    
    # 4. PDF
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=10)
    pdf.cell(200, 10, txt=f"Relatorio {tipo_consulta} - Rio das Pedras", ln=True, align='C')
    for _, row in df.head(30).iterrows():
        texto_limpo = str(row.values[0] if len(row) > 0 else "").encode('latin-1', 'replace').decode('latin-1')
        pdf.cell(200, 10, txt=texto_limpo, ln=True)
    buffer_pdf = io.BytesIO(pdf.output(dest='S').encode('latin-1'))
    cols[3].download_button("📕 PDF (.pdf)", buffer_pdf.getvalue(), f"Relatorio_{nome}.pdf")

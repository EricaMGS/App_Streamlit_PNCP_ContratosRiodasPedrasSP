import io
import pandas as pd
import requests
import streamlit as st
from docx import Document
from fpdf import FPDF

# Configuração da página
st.set_page_config(page_title="Portal PNCP - Rio das Pedras/SP", layout="wide")

st.title("🏛️ Contratações de Rio das Pedras/SP")
st.markdown("Consulta integrada de **Contratos**, **Atas** e **Editais** direto do Portal Nacional de Contratações Públicas.")

# --- BARRA LATERAL ---
st.sidebar.header("Parâmetros da Consulta")

# Opções incluindo Editais e Avisos
tipo_consulta = st.sidebar.selectbox(
    "Selecione o tipo de dado:",
    ["Contratos", "Atas de Registro de Preços", "Editais e Avisos de Contratação"]
)

# Definindo datas (com limite implícito de 365 dias para evitar o erro 422)
data_inicio = st.sidebar.date_input("Data Inicial", value=pd.to_datetime("2026-01-01"))
data_fim = st.sidebar.date_input("Data Final", value=pd.to_datetime("2026-08-18"))

# Validação do limite de 365 dias exigido pelo PNCP
if (data_fim - data_inicio).days > 365:
    st.sidebar.error("⚠️ O período selecionado não pode ser maior que 365 dias.")
    st.stop()

# Mapeamento de endpoints
endpoints = {
    "Contratos": "https://pncp.gov.br/api/consulta/v1/contratos",
    "Atas de Registro de Preços": "https://pncp.gov.br/api/consulta/v1/atas",
    "Editais e Avisos de Contratação": "https://pncp.gov.br/api/consulta/v1/contratacoes/publicacao"
}

if st.sidebar.button("Gerar Relatório"):
    url_escolhida = endpoints[tipo_consulta]
    
    with st.spinner(f"Consultando {tipo_consulta} no PNCP..."):
        params = {
            "cnpj": "44826840000183",
            "dataInicial": data_inicio.strftime("%Y%m%d"),
            "dataFinal": data_fim.strftime("%Y%m%d"),
            "pagina": 1
        }
        
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
            "Accept": "application/json"
        }
        
        try:
            resp = requests.get(url_escolhida, params=params, headers=headers, timeout=45)
            
            if resp.status_code == 200:
                data = resp.json()
                
                # O PNCP pode retornar uma lista ou um dicionário contendo "data"
                lote = data if isinstance(data, list) else data.get("data", data.get("items", []))
                
                if not lote:
                    st.warning(f"Nenhum registro de '{tipo_consulta}' encontrado para este período.")
                else:
                    df = pd.DataFrame(lote)
                    st.success(f"Consulta finalizada! {len(df)} registros encontrados.")
                    
                    st.dataframe(df, use_container_width=True)
                    
                    # --- BOTÕES DE DOWNLOAD ---
                    st.markdown("### 📥 Baixar Relatórios")
                    cols = st.columns(4)
                    
                    nome_arquivo_base = tipo_consulta.replace(" ", "_")
                    
                    # 1. EXCEL
                    buffer_xlsx = io.BytesIO()
                    df.to_excel(buffer_xlsx, index=False)
                    cols[0].download_button("📊 Excel (.xlsx)", buffer_xlsx.getvalue(), f"{nome_arquivo_base}.xlsx")
                    
                    # 2. CSV
                    cols[1].download_button("📄 CSV (.csv)", df.to_csv(index=False), f"{nome_arquivo_base}.csv")
                    
                    # 3. WORD
                    doc = Document()
                    doc.add_heading(f"Relatório: {tipo_consulta}", 0)
                    doc.add_paragraph(f"Período: {data_inicio.strftime('%d/%m/%Y')} a {data_fim.strftime('%d/%m/%Y')}")
                    
                    # Criar tabela simples no Word para amostra (máx 50 linhas para não travar)
                    amostra = df.head(50)
                    if not amostra.empty:
                        colunas_exibir = amostra.columns[:4] # Pegamos apenas as primeiras 4 colunas para caber
                        tabela = doc.add_table(rows=1, cols=len(colunas_exibir))
                        tabela.style = 'Table Grid'
                        
                        # Cabeçalhos
                        for i, col in enumerate(colunas_exibir):
                            tabela.rows[0].cells[i].text = str(col)
                            
                        # Linhas
                        for _, row in amostra.iterrows():
                            cells = tabela.add_row().cells
                            for i, col in enumerate(colunas_exibir):
                                cells[i].text = str(row[col])[:40] # Limite de caracteres para não quebrar o layout
                    
                    buffer_docx = io.BytesIO()
                    doc.save(buffer_docx)
                    cols[2].download_button("📝 Word (.docx)", buffer_docx.getvalue(), f"{nome_arquivo_base}.docx")
                    
                    # 4. PDF (Simples)
                    pdf = FPDF()
                    pdf.add_page()
                    pdf.set_font("Arial", size=10)
                    pdf.cell(200, 10, txt=f"Relatorio de {tipo_consulta}", ln=True, align='C')
                    pdf.cell(200, 10, txt=f"Periodo: {data_inicio.strftime('%d/%m/%Y')} a {data_fim.strftime('%d/%m/%Y')}", ln=True, align='C')
                    
                    # Tabela PDF simples (Apenas primeira coluna para não transbordar a página)
                    if not df.empty:
                        coluna_alvo = df.columns[0]
                        for _, row in df.head(40).iterrows():
                            # Remove caracteres acentuados que quebram o FPDF
                            texto_limpo = str(row[coluna_alvo]).encode('latin-1', 'replace').decode('latin-1')
                            pdf.cell(200, 10, txt=texto_limpo, ln=True)
                            
                    buffer_pdf = io.BytesIO(pdf.output(dest='S').encode('latin-1'))
                    cols[3].download_button("📕 PDF (.pdf)", buffer_pdf.getvalue(), f"{nome_arquivo_base}.pdf")
                    
            else:
                # Tratamento explícito de erros da API (como o 400 ou 422)
                st.error(f"O servidor do PNCP recusou a consulta. Código do erro: {resp.status_code}")
                st.code(resp.text)
                if resp.status_code == 400 and tipo_consulta == "Editais e Avisos de Contratação":
                    st.info("💡 A API de Editais mudou recentemente e pode estar exigindo filtros adicionais (como a modalidade). Caso o problema persista, sugerimos utilizar a consulta de Contratos ou Atas.")

        except requests.exceptions.ReadTimeout:
            st.error("Tempo de resposta esgotado. O portal do PNCP está demorando muito para responder. Tente consultar um período de tempo mais curto.")
        except Exception as e:
            st.error(f"Erro inesperado de conexão: {e}")

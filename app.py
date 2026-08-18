import io
import pandas as pd
import requests
import streamlit as st
from docx import Document
from docx.shared import Pt, RGBColor
from datetime import timedelta

# Configuração da página
st.set_page_config(page_title="Portal PNCP - Rio das Pedras/SP", layout="wide")

st.title("🏛️ Contratações de Rio das Pedras/SP")

# --- BARRA LATERAL ---
st.sidebar.header("Parâmetros da Consulta")
tipo_consulta = st.sidebar.selectbox(
    "Selecione o tipo de dado:",
    ["Contratos", "Atas de Registro de Preços"] # Removido temporariamente por erro de parâmetro da API
)

data_inicio = st.sidebar.date_input("Data Inicial", value=pd.to_datetime("2026-01-01"))
# Lógica de segurança: Data fim não pode exceder 365 dias da inicial
data_fim = st.sidebar.date_input("Data Final", value=pd.to_datetime("2026-08-18"))

# Validação de segurança para o erro 422
if (data_fim - data_inicio).days > 365:
    st.sidebar.error("O período não pode ser maior que 365 dias.")
    st.stop()

endpoints = {
    "Contratos": "https://pncp.gov.br/api/consulta/v1/contratos",
    "Atas de Registro de Preços": "https://pncp.gov.br/api/consulta/v1/atas",
}

if st.sidebar.button("Gerar Relatório Consolidado"):
    url_escolhida = endpoints[tipo_consulta]
    
    with st.spinner("Consultando dados..."):
        params = {
            "cnpj": "44826840000183",
            "dataInicial": data_inicio.strftime("%Y%m%d"),
            "dataFinal": data_fim.strftime("%Y%m%d"),
            "pagina": 1
        }
        
        try:
            response = requests.get(url_escolhida, params=params, timeout=30)
            if response.status_code == 200:
                data = response.json()
                lote = data.get("data", data.get("items", [])) if isinstance(data, dict) else data
                
                if not lote:
                    st.warning("Nenhum registro encontrado para este período.")
                else:
                    df = pd.DataFrame(lote)
                    st.dataframe(df, use_container_width=True)
                    
                    # (Coloque aqui o restante da lógica de download que você já tem)
                    st.success("Dados carregados!")
            else:
                st.error(f"Erro do servidor PNCP: {response.status_code} - {response.text}")
        except Exception as e:
            st.error(f"Erro ao conectar: {e}")

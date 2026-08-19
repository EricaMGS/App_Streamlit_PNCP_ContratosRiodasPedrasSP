import io
import datetime
import time
import pandas as pd
import requests
import streamlit as st
from docx import Document
from fpdf import FPDF

# ============================================================
# CONFIGURAÇÃO
# ============================================================

st.set_page_config(
    page_title="Portal PNCP - Rio das Pedras/SP",
    layout="wide"
)

st.title("🏛️ Contratações de Rio das Pedras/SP")

st.markdown(
    "Consulta integrada de Contratos, Atas e Editais "
    "direto do Portal Nacional de Contratações Públicas."
)

# ============================================================
# DADOS DO MUNICÍPIO
# ============================================================

CNPJ_RIO_DAS_PEDRAS = "44826840000183"

# Código IBGE de Rio das Pedras/SP
CODIGO_IBGE_RIO_DAS_PEDRAS = "3544003"

UF = "SP"

BASE_URL = "[https://pncp.gov.br/api/consulta/v1](https://pncp.gov.br/api/consulta/v1)"

# ============================================================
# BARRA LATERAL
# ============================================================

st.sidebar.header("Parâmetros da Consulta")

tipo_consulta = st.sidebar.selectbox(
    "Selecione:",
    [
        "Contratos",
        "Atas de Registro de Preços",
        "Editais e Avisos de Contratações"
    ]
)

# ============================================================
# MODALIDADE — EDITAIS
# ============================================================

modalidade_codigo = None

if tipo_consulta == "Editais e Avisos de Contratações":

    modalidade_opcoes = {
        "Pregão - Eletrônico (6)": 6,
        "Dispensa de Licitação (8)": 8,
        "Inexigibilidade (9)": 9,
        "Concorrência - Eletrônica (2)": 2
    }

    mod_escolhida = st.sidebar.selectbox(
        "Modalidade:",
        list(modalidade_opcoes.keys())
    )

    modalidade_codigo = modalidade_opcoes[mod_escolhida]


# ============================================================
# DATAS
# ============================================================

data_inicio = st.sidebar.date_input(
    "Data Inicial",
    value=pd.to_datetime("2026-01-01")
)

data_fim = st.sidebar.date_input(
    "Data Final",
    value=datetime.date.today()
)

# ============================================================
# VALIDAÇÕES
# ============================================================

if data_fim < data_inicio:
    st.sidebar.error(
        "⚠️ A Data Final não pode ser anterior à Data Inicial."
    )
    st.stop()

if (data_fim - data_inicio).days > 365:
    st.sidebar.error(
        "⚠️ O período não pode ser maior que 365 dias."
    )
    st.stop()


# ============================================================
# SESSION STATE
# ============================================================

if "tipo_anterior" not in st.session_state:
    st.session_state.tipo_anterior = tipo_consulta

if

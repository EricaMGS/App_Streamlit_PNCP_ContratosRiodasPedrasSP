import io
import datetime
import time
from concurrent.futures import ThreadPoolExecutor
import pandas as pd
import requests
import streamlit as st
from docx import Document
from docx.shared import Pt, RGBColor
from fpdf import FPDF

# ============================================================
# FUNÇÃO DE FORMATAÇÃO BR
# ============================================================
def formatar_moeda_br(valor):
    """Formata número para o padrão brasileiro: R$ 15.300.296,03"""
    valor_fmt = "{:,.2f}".format(valor)
    return f"R$ {valor_fmt.replace(',', 'X').replace('.', ',').replace('X', '.')}"

# ============================================================
# CONFIGURAÇÃO
# ============================================================
st.set_page_config(page_title="Portal PNCP - Rio das Pedras/SP", layout="wide")
st.title("Contratações de Rio das Pedras/SP")
st.markdown("Consulta integrada de Contratos, Atas e Editais direto do Portal Nacional de Contratações Públicas.")

CNPJ_RIO_DAS_PEDRAS = "44826840000183"
CODIGO_IBGE_RIO_DAS_PEDRAS = "3544004"
UF = "SP"
BASE_URL = "https://pncp.gov.br/api/consulta/v1"

# ============================================================
# SIDEBAR
# ============================================================
st.sidebar.header("Parâmetros da Consulta")
tipo_consulta = st.sidebar.selectbox("Selecione:", ["Contratos", "Atas de Registro de Preços", "Editais e Avisos de Contratações"])

modalidade_codigo = None
if tipo_consulta == "Editais e Avisos de Contratações":
    modalidade_opcoes = {"Pregão - Eletrônico (6)": 6, "Dispensa de Licitação (8)": 8, "Inexigibilidade (9)": 9, "Concorrência - Eletrônica (2)": 2}
    modalidade_codigo = modalidade_opcoes[st.sidebar.selectbox("Modalidade:", list(modalidade_opcoes.keys()))]

data_inicio = st.sidebar.date_input("Data Inicial", value=pd.to_datetime("2026-01-01"))
data_fim = st.sidebar.date_input("Data Final", value=datetime.date.today())

# Inicialização do estado
if "df_resultado" not in st.session_state: st.session_state.df_resultado = None

# ============================================================
# FUNÇÕES DE CONSULTA
# ============================================================
def consultar_pncp(url, params, max_tentativas=5):
    for tentativa in range(1, max_tentativas + 1):
        try:
            resp = requests.get(url, params=params, timeout=(15, 90))
            if resp.status_code == 200: return resp.json()
            if resp.status_code == 204: return []
            if resp.status_code in [429, 500, 502, 503, 504] and tentativa < max_tentativas:
                time.sleep(2 ** tentativa)
                continue
            raise Exception(f"Erro {resp.status_code}")
        except:
            if tentativa == max_tentativas: raise
    return []

def consultar_detalhes_contrato(id_contrato):
    url = f"{BASE_URL}/contratos/{id_contrato}"
    try:
        contrato = requests.get(url, timeout=15).json()
        aditivos = contrato.get('termosAditivos', []) + contrato.get('termosApostilamentos', [])
        for item in aditivos: item['tipoDocumentoNome'] = 'Termo'
        return aditivos
    except: return []

def consultar_paginas_rapido(url, params):
    todos = extrair_registros(consultar_pncp(url, {**params, "pagina": 1}))
    def buscar(p): return extrair_registros(consultar_pncp(url, {**params, "pagina": p}))
    with ThreadPoolExecutor(max_workers=5) as exec:
        res = list(exec.map(buscar, range(2, 11)))
    for r in res: todos.extend(r)
    return todos

def extrair_registros(data):
    if isinstance(data, list): return data
    if isinstance(data, dict):
        for c in ["data", "items", "content", "dados"]:
            if c in data and isinstance(data[c], list): return data[c]
    return []

def tratar_dataframe(df):
    if df.empty: return df
    for col in df.columns:
        if not df[col].dropna().empty and isinstance(df[col].dropna().iloc[0], dict):
            df[col] = df[col].apply(lambda x: x.get('nome') or x.get('razaoSocial') or x.get('descricao') or str(x))
    return df

def obter_dados_registro(row, tipo):
    id_pncp = row.get('numeroControlePNCP', 'N/D')
    return str(id_pncp), str(row.get('processo', 'N/D')), "Detalhes", str(row.get('objetoContrato', 'N/D'))

# ============================================================
# EXECUÇÃO E INTERFACE
# ============================================================
if st.sidebar.button("🔎 Gerar Consulta", type="primary"):
    endpoints = {"Contratos": f"{BASE_URL}/contratos", "Atas de Registro de Preços": f"{BASE_URL}/atas", "Editais e Avisos de Contratações": f"{BASE_URL}/contratacoes/publicacao"}
    params = {"dataInicial": data_inicio.strftime("%Y%m%d"), "dataFinal": data_fim.strftime("%Y%m%d"), "tamanhoPagina": 50}
    if tipo_consulta == "Contratos": params["cnpjOrgao"] = CNPJ_RIO_DAS_PEDRAS
    st.session_state.df_resultado = tratar_dataframe(pd.DataFrame(consultar_paginas_rapido(endpoints[tipo_consulta], params)))

if st.session_state.df_resultado is not None and not st.session_state.df_resultado.empty:
    df = st.session_state.df_resultado
    st.success(f"📊 {len(df)} registros encontrados.")
    col_m1, col_m2 = st.columns(2)
    col_m1.metric("Total de Registros", len(df))
    col_m2.metric("Valor Total Envolvido", formatar_moeda_br(pd.to_numeric(df.get('valorGlobal', 0), errors='coerce').sum()))

    st.markdown("### 📥 Opções de Exportação")
    c1, c2, c3, c4 = st.columns(4)
    c1.download_button("📊 Excel", df.to_csv(index=False), "dados.csv")
    
    # Aditivos e Gráficos
    if tipo_consulta == "Contratos":
        st.markdown("### 🔍 Aditivos")
        id_sel = st.selectbox("Contrato:", df.apply(lambda x: f"{x.get('numeroControlePNCP')} - {x.get('processo')}", axis=1))
        if st.button("Buscar Aditivos"):
            ads = consultar_detalhes_contrato(id_sel.split(" - ")[0])
            for a in ads: st.info(f"**Tipo:** {a.get('tipoDocumentoNome')} | {a.get('objeto')}")

    st.markdown("### 📈 Análise Gráfica: Volume de Registros por Mês/Ano")
    col_data = next((c for c in ["dataPublicacao", "dataInclusao"] if c in df.columns), None)
    if col_data:
        df['mes_ano'] = pd.to_datetime(df[col_data], errors='coerce').dt.to_period('M').astype(str)
        aba1, aba2 = st.tabs(["🔢 Quantidade de Registros", "💰 Volume Financeiro (R$)"])
        with aba1: st.bar_chart(df['mes_ano'].value_counts().sort_index())
        with aba2: st.bar_chart(df.groupby('mes_ano')['valorGlobal'].sum().sort_index())

    st.dataframe(df, use_container_width=True)

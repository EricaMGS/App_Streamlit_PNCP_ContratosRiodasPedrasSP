import io
import datetime
import time
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
    # Troca a vírgula pelo caractere temporário 'X', o ponto pela vírgula, e o 'X' pelo ponto
    return f"R$ {valor_fmt.replace(',', 'X').replace('.', ',').replace('X', '.')}"

# ============================================================
# CONFIGURAÇÃO
# ============================================================

st.set_page_config(
    page_title="Portal PNCP - Rio das Pedras/SP",
    layout="wide"
)

st.title("Contratações de Rio das Pedras/SP")

st.markdown(
    "Consulta integrada de Contratos, Atas e Editais "
    "direto do Portal Nacional de Contratações Públicas."
)

# ============================================================
# DADOS DO MUNICÍPIO
# ============================================================

CNPJ_RIO_DAS_PEDRAS = "44826840000183"
CODIGO_IBGE_RIO_DAS_PEDRAS = "3544004"
UF = "SP"
BASE_URL = "https://pncp.gov.br/api/consulta/v1"

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
    st.sidebar.error("⚠️ A Data Final não pode ser anterior à Data Inicial.")
    st.stop()

if (data_fim - data_inicio).days > 365:
    st.sidebar.error("⚠️ O período não pode ser maior que 365 dias.")
    st.stop()

# ============================================================
# SESSION STATE (INICIALIZAÇÃO CORRETA)
# ============================================================

if "df_resultado" not in st.session_state:
    st.session_state.df_resultado = None

if "tipo_anterior" not in st.session_state:
    st.session_state.tipo_anterior = tipo_consulta

if st.session_state.tipo_anterior != tipo_consulta:
    st.session_state.df_resultado = None
    st.session_state.tipo_anterior = tipo_consulta

# ============================================================
# FUNÇÕES DE CONSULTA COM BACKOFF EXPONENCIAL
# ============================================================

def consultar_pncp(url, params, max_tentativas=5):
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/151.0 Safari/537.36",
        "Accept": "application/json",
        "Connection": "keep-alive"
    }
    
    for tentativa in range(1, max_tentativas + 1):
        try:
            resp = requests.get(url, params=params, headers=headers, timeout=(15, 90))
            if resp.status_code == 200:
                try:
                    return resp.json()
                except ValueError:
                    raise Exception("Resposta da API não é um JSON válido.")
            if resp.status_code == 204:
                return []
            
            # Tratamento robusto para erros de gateway e servidor (502, 500, etc.)
            if resp.status_code in [429, 500, 502, 503, 504]:
                if tentativa < max_tentativas:
                    espera = 2 ** tentativa
                    st.warning(f"⚠️ Servidor instável (Erro {resp.status_code}). Nova tentativa em {espera}s ({tentativa}/{max_tentativas}).")
                    time.sleep(espera)
                    continue
            
            raise Exception(f"API retornou HTTP {resp.status_code}: {resp.text[:200]}")
            
        except (requests.exceptions.Timeout, requests.exceptions.ConnectionError) as e:
            if tentativa < max_tentativas:
                espera = 2 ** tentativa
                st.warning(f"🌐 Problema de rede. Tentativa {tentativa}/{max_tentativas} em {espera}s...")
                time.sleep(espera)
                continue
            raise e
            
    raise Exception("Falha de conexão com o PNCP após várias tentativas.")

def consultar_detalhes_contrato(id_contrato):
    """Busca documentos vinculados a um contrato específico (Aditivos/Apostilamentos)."""
    url = f"{BASE_URL}/contratos/{id_contrato}/documentos"
    try:
        resp = requests.get(url, timeout=10)
        if resp.status_code == 200:
            return resp.json()
        return []
    except:
        return []

def extrair_registros(data):
    if isinstance(data, list): return data
    if isinstance(data, dict):
        for chave in ["data", "items", "content", "dados"]:
            if chave in data and isinstance(data[chave], list):
                return data[chave]
    return []

def tratar_dataframe(df):
    if df.empty: return df
    df_tratado = df.copy()
    for col in df_tratado.columns:
        sample = df_tratado[col].dropna()
        if not sample.empty and isinstance(sample.iloc[0], dict):
            df_tratado[col] = df_tratado[col].apply(
                lambda x: x.get('nome') or x.get('razaoSocial') or x.get('descricao') or str(x) if isinstance(x, dict) else str(x)
            )
    return df_tratado

def consultar_paginas(url, params, max_paginas=100):
    todos_registros = []
    for pagina in range(1, max_paginas + 1):
        params_pagina = params.copy()
        params_pagina["pagina"] = pagina
        data = consultar_pncp(url, params_pagina)
        registros = extrair_registros(data)
        if not registros: break
        todos_registros.extend(registros)
        if len(registros) < params_pagina.get("tamanhoPagina", 50): break
        time.sleep(0.5)
    return todos_registros

def obter_dados_registro(row, tipo):
    id_pncp = row.get('numeroControlePNCP', row.get('numeroControlePNCPAta', row.get('numeroControlePNCPCompra', 'N/D')))
    if tipo == "Atas de Registro de Preços":
        processo = row.get('numeroAtaRegistroPreco', 'N/D')
        info_extra = f"Vigência: Início: {row.get('vigenciaInicio', 'N/D')} | Fim: {row.get('vigenciaFim', 'N/D')}"
    elif tipo == "Contratos":
        processo = row.get('processo', 'N/D')
        valor = row.get('valorGlobal', row.get('valorInicial', 0))
        valor_fmt = f"R$ {float(valor):,.2f}" if str(valor).replace('.','',1).isdigit() else str(valor)
        info_extra = f"Fornecedor: {row.get('nomeRazaoSocialFornecedor', 'N/D')} | Valor: {valor_fmt}"
    else:
        processo = row.get('processo', 'N/D')
        valor = row.get('valorTotalHomologado', row.get('valorTotalEstimado', 0))
        valor_fmt = f"R$ {float(valor):,.2f}" if str(valor).replace('.','',1).isdigit() else str(valor)
        info_extra = f"Responsável: {row.get('usuarioNome', 'N/D')} | Valor: {valor_fmt}"
    return str(id_pncp), str(processo), info_extra, str(row.get('objetoContrato' if tipo=="Contratos" else 'objetoCompra', 'N/D'))

# ============================================================
# INTERFACE
# ============================================================

if st.sidebar.button("🔎 Gerar Consulta", type="primary"):
    endpoints = {"Contratos": f"{BASE_URL}/contratos", "Atas de Registro de Preços": f"{BASE_URL}/atas", "Editais e Avisos de Contratações": f"{BASE_URL}/contratacoes/publicacao"}
    endpoint = endpoints[tipo_consulta]
    params = {"dataInicial": data_inicio.strftime("%Y%m%d"), "dataFinal": data_fim.strftime("%Y%m%d"), "pagina": 1, "tamanhoPagina": 50}
    if tipo_consulta == "Editais e Avisos de Contratações":
        params.update({"codigoModalidadeContratacao": modalidade_codigo, "uf": UF, "codigoMunicipioIbge": CODIGO_IBGE_RIO_DAS_PEDRAS, "cnpj": CNPJ_RIO_DAS_PEDRAS})
    elif tipo_consulta == "Contratos": params["cnpjOrgao"] = CNPJ_RIO_DAS_PEDRAS
    elif tipo_consulta == "Atas de Registro de Preços": params["cnpj"] = CNPJ_RIO_DAS_PEDRAS

    try:
        with st.spinner("🔄 Buscando e tratando dados no PNCP..."):
            registros = consultar_paginas(endpoint, params)
            df_temp = tratar_dataframe(pd.DataFrame(registros))
            st.session_state.df_resultado = df_temp
    except Exception as e:
        st.error(f"❌ Erro: {str(e)}")

if st.session_state.df_resultado is not None and not st.session_state.df_resultado.empty:
    df = st.session_state.df_resultado
    col_m1, col_m2 = st.columns(2)
    col_m1.metric("Total de Registros", len(df))
    coluna_valor = next((c for c in ["valorGlobal", "valorInicial", "valorTotalHomologado", "valorTotalEstimado"] if c in df.columns), None)
    if coluna_valor:
        col_m2.metric("Valor Total Envolvido", formatar_moeda_br(pd.to_numeric(df[coluna_valor], errors='coerce').sum()))
    
    st.markdown("---")
    st.dataframe(df, use_container_width=True, hide_index=True)

    # Lógica de Exportação e Aditivos (Mantida igual ao código anterior)
    # ...

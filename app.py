import io
import datetime
import time
import pandas as pd
import requests
import streamlit as st
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
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
# SESSION STATE
# ============================================================

if "tipo_anterior" not in st.session_state:
    st.session_state.tipo_anterior = tipo_consulta

if st.session_state.tipo_anterior != tipo_consulta:
    st.session_state.df_resultado = None
    st.session_state.tipo_anterior = tipo_consulta

if "df_resultado" not in st.session_state:
    st.session_state.df_resultado = None

# ============================================================
# FUNÇÃO DE CONSULTA COM RETENTATIVAS
# ============================================================

def consultar_pncp(url, params, max_tentativas=3):
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/151.0 Safari/537.36",
        "Accept": "application/json",
        "Connection": "keep-alive"
    }
    ultima_excecao = None

    for tentativa in range(1, max_tentativas + 1):
        try:
            resp = requests.get(url, params=params, headers=headers, timeout=(15, 90))
            if resp.status_code == 200:
                try:
                    return resp.json()
                except ValueError:
                    raise Exception("O PNCP respondeu, mas não enviou um JSON válido.")
            if resp.status_code == 204:
                return []
            if resp.status_code in [429, 500, 502, 503, 504]:
                if tentativa < max_tentativas:
                    espera = tentativa * 3
                    st.warning(f"⚠️ O Portal PNCP está demorando. Nova tentativa em {espera}s ({tentativa}/{max_tentativas}).")
                    time.sleep(espera)
                    continue
            raise Exception(f"API retornou HTTP {resp.status_code}: {resp.text[:500]}")
        except requests.exceptions.Timeout as e:
            ultima_excecao = e
            if tentativa < max_tentativas:
                espera = tentativa * 3
                st.warning(f"⏳ Timeout no PNCP. Nova tentativa em {espera}s ({tentativa}/{max_tentativas}).")
                time.sleep(espera)
            continue
        except requests.exceptions.ConnectionError as e:
            ultima_excecao = e
            if tentativa < max_tentativas:
                espera = tentativa * 3
                st.warning(f"🌐 Erro de conexão. Nova tentativa em {espera}s ({tentativa}/{max_tentativas}).")
                time.sleep(espera)
            continue
        except Exception:
            raise
    raise ultima_excecao

# ============================================================
# EXTRAIR REGISTROS
# ============================================================

def extrair_registros(data):
    if isinstance(data, list):
        return data
    if isinstance(data, dict):
        for chave in ["data", "items", "content", "dados"]:
            if chave in data and isinstance(data[chave], list):
                return data[chave]
    return []

# ============================================================
# PAGINAÇÃO
# ============================================================

def consultar_paginas(url, params, max_paginas=100):
    todos_registros = []
    for pagina in range(1, max_paginas + 1):
        params_pagina = params.copy()
        params_pagina["pagina"] = pagina
        data = consultar_pncp(url, params_pagina)
        registros = extrair_registros(data)
        if not registros:
            break
        todos_registros.extend(registros)
        tamanho = params_pagina.get("tamanhoPagina", 50)
        if len(registros) < tamanho:
            break
        time.sleep(0.5)
    return todos_registros

# ============================================================
# BOTÃO CONSULTAR
# ============================================================

if st.sidebar.button("🔎 Gerar Relatório", type="primary"):
    st.info("ℹ️ Consultando diretamente o PNCP. O processo pode levar alguns instantes...")

    endpoints = {
        "Contratos": f"{BASE_URL}/contratos",
        "Atas de Registro de Preços": f"{BASE_URL}/atas",
        "Editais e Avisos de Contratações": f"{BASE_URL}/contratacoes/publicacao"
    }

    endpoint = endpoints[tipo_consulta]
    tamanho_pagina = 50 if tipo_consulta == "Editais e Avisos de Contratações" else 100

    if tipo_consulta == "Editais e Avisos de Contratações":
        params = {
            "dataInicial": data_inicio.strftime("%Y%m%d"),
            "dataFinal": data_fim.strftime("%Y%m%d"),
            "dataPublicacaoInicial": data_inicio.strftime("%Y%m%d"),
            "dataPublicacaoFinal": data_fim.strftime("%Y%m%d"),
            "pagina": 1,
            "tamanhoPagina": tamanho_pagina,
            "codigoModalidadeContratacao": modalidade_codigo,
            "uf": UF,
            "codigoMunicipioIbge": CODIGO_IBGE_RIO_DAS_PEDRAS,
            "cnpj": CNPJ_RIO_DAS_PEDRAS
        }
    else:
        params = {
            "dataInicial": data_inicio.strftime("%Y%m%d"),
            "dataFinal": data_fim.strftime("%Y%m%d"),
            "pagina": 1,
            "tamanhoPagina": tamanho_pagina
        }
        if tipo_consulta == "Contratos":
            params["cnpjOrgao"] = CNPJ_RIO_DAS_PEDRAS
        elif tipo_consulta == "Atas de Registro de Preços":
            params["cnpj"] = CNPJ_RIO_DAS_PEDRAS

    try:
        with st.spinner("🔄 Buscando dados no PNCP..."):
            registros = consultar_paginas(endpoint, params)

        if registros:
            df_temp = pd.DataFrame(registros)
            if tipo_consulta == "Contratos":
                possiveis_colunas = ["cnpjOrgao", "orgaoEntidade", "orgao", "unidadeOrgao"]
                encontrou_cnpj = False
                for coluna in possiveis_colunas:
                    if coluna in df_temp.columns:
                        serie = df_temp[coluna].astype(str).str.replace(r"\D", "", regex=True)
                        mask = serie.str.contains(CNPJ_RIO_DAS_PEDRAS, na=False)
                        if mask.any():
                            df_temp = df_temp[mask]
                            encontrou_cnpj = True
                            break
                if not encontrou_cnpj:
                    st.warning("⚠️ O PNCP retornou registros, mas sem coluna CNPJ padronizada.")

            st.session_state.df_resultado = df_temp
            if df_temp.empty:
                st.warning("⚠️ Nenhum registro encontrado para Rio das Pedras/SP no período.")
            else:
                st.success(f"✅ Consulta concluída! {len(df_temp)} registros encontrados.")
        else:
            st.session_state.df_resultado = pd.DataFrame()
            st.warning("ℹ️ Nenhum registro retornado pelo PNCP.")

    except requests.exceptions.Timeout:
        st.session_state.df_resultado = None
        st.error("⏱️ Tempo limite excedido ao consultar o PNCP.")
    except requests.exceptions.ConnectionError:
        st.session_state.df_resultado = None
        st.error("🌐 Erro de conexão com o PNCP.")
    except Exception as e:
        st.session_state.df_resultado = None
        st.error(f"❌ Erro: {str(e)}")

# ============================================================
# EXIBIÇÃO PERSISTENTE
# ============================================================

if st.session_state.df_resultado is not None and not st.session_state.df_resultado.empty:
    df = st.session_state.df_resultado
    st.success(f"📊 Exibindo {len(df)} registros para Rio das Pedras/SP.")
    st.dataframe(df, use_container_width=True, hide_index=True)

    st.markdown("### 📥 Opções de Exportação")
    cols = st.columns(4)
    nome = tipo_consulta.replace(" ", "_").replace("/", "_")

    # Excel
    buffer_xlsx = io.BytesIO()
    df.to_excel(buffer_xlsx, index=False)
    cols[0].download_button("📊 Excel (.xlsx)", buffer_xlsx.getvalue(), f"{nome}_Rio_Das_Pedras.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # CSV
    cols[1].download_button("📄 CSV (.csv)", df.to_csv(index=False, encoding="utf-8-sig"), f"{nome}_Rio_Das_Pedras.csv", mime="text/csv")

    # ============================================================
    # WORD FORMATADO E LIMPO
    # ============================================================
    doc = Document()
    
    # Título do Documento
    p_titulo = doc.add_paragraph()
    r_titulo = p_titulo.add_run(f"Relatório Executivo: {tipo_consulta}")
    r_titulo.bold = True
    r_titulo.font.size = Pt(16)
    r_titulo.font.color.rgb = RGBColor(0, 51, 102)
    
    doc.add_paragraph(f"Município: Prefeitura Municipal de Rio das Pedras / SP")
    doc.add_paragraph(f"Período: {data_inicio.strftime('%d/%m/%Y')} a {data_fim.strftime('%d/%m/%Y')}")
    doc.add_paragraph(f"Total de Registros: {len(df)}")
    doc.add_heading("Detalhamento dos Registros", level=2)

    for idx, row in df.head(50).iterrows():
        p_reg = doc.add_paragraph()
        p_reg.add_run(f"Item #{idx + 1}\n").bold = True
        
        # Extração inteligente de campos comuns limpos
        processo = row.get('processo', 'N/D')
        fornecedor = row.get('nomeRazaoSocialFornecedor', 'N/D')
        objeto = row.get('objetoContrato', row.get('objetoCompra', 'N/D'))
        valor = row.get('valorGlobal', row.get('valorTotalEstimado', 0))
        
        try:
            valor_fmt = f"R$ {float(valor):,.2f}" if pd.notna(valor) else "N/D"
        except:
            valor_fmt = str(valor)

        p_reg.add_run(f"• Processo: {processo}\n")
        p_reg.add_run(f"• Fornecedor: {fornecedor}\n")
        p_reg.add_run(f"• Valor: {valor_fmt}\n")
        p_reg.add_run(f"• Objeto: {objeto}\n")
        p_reg.add_paragraph("-" * 40)

    buffer_docx = io.BytesIO()
    doc.save(buffer_docx)
    buffer_docx.seek(0)
    cols[2].download_button("📝 Word (.docx)", buffer_docx.getvalue(), f"Relatorio_{nome}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

    # ============================================================
    # PDF FORMATADO E LIMPO
    # ============================================================
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    
    # Cabeçalho do PDF
    pdf.set_font("Arial", 'B', 14)
    pdf.cell(0, 10, txt=f"Relatorio: {tipo_consulta}", ln=True, align="C")
    
    pdf.set_font("Arial", size=10)
    pdf.cell(0, 6, txt="Municipio: Prefeitura Municipal de Rio das Pedras / SP", ln=True)
    pdf.cell(0, 6, txt=f"Periodo: {data_inicio.strftime('%d/%m/%Y')} a {data_fim.strftime('%d/%m/%Y')} | Total: {len(df)} registros", ln=True)
    pdf.ln(5)

    pdf.set_font("Arial", 'B', 10)
    pdf.cell(0, 8, txt="Principais Registros:", ln=True)
    pdf.set_font("Arial", size=9)

    for idx, row in df.head(30).iterrows():
        processo = str(row.get('processo', 'N/D'))
        fornecedor = str(row.get('nomeRazaoSocialFornecedor', 'N/D'))
        objeto = str(row.get('objetoContrato', row.get('objetoCompra', 'N/D')))
        valor = row.get('valorGlobal', row.get('valorTotalEstimado', 0))
        
        try:
            valor_fmt = f"R$ {float(valor):,.2f}" if pd.notna(valor) else "N/D"
        except:
            valor_fmt = str(valor)

        bloco = f"[{idx+1}] Proc: {processo} | Fornecedor: {fornecedor} | Valor: {valor_fmt}\nObjeto: {objeto}"
        
        bloco_limpo = bloco.encode("latin-1", "replace").decode("latin-1")
        pdf.multi_cell(0, 5, txt=bloco_limpo)
        pdf.ln(3)
    
    pdf_bytes = pdf.output(dest="S")
    if isinstance(pdf_bytes, str):
        pdf_bytes = pdf_bytes.encode("latin-1")
    cols[3].download_button("📕 PDF (.pdf)", pdf_bytes, f"Relatorio_{nome}.pdf", mime="application/pdf")

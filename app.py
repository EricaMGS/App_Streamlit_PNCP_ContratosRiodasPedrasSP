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
# FUNÇÕES DE CONSULTA
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
    if isinstance(data, list):
        return data
    if isinstance(data, dict):
        for chave in ["data", "items", "content", "dados"]:
            if chave in data and isinstance(data[chave], list):
                return data[chave]
    return []

def tratar_dataframe(df):
    if df.empty:
        return df
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
        if not registros:
            break
        todos_registros.extend(registros)
        tamanho = params_pagina.get("tamanhoPagina", 50)
        if len(registros) < tamanho:
            break
        time.sleep(0.5)
    return todos_registros

def obter_dados_registro(row, tipo):
    id_pncp = row.get('numeroControlePNCP', row.get('numeroControlePNCPAta', row.get('numeroControlePNCPCompra', 'N/D')))
    
    if tipo == "Atas de Registro de Preços":
        processo = row.get('numeroAtaRegistroPreco', 'N/D')
        vigencia = f"Início: {row.get('vigenciaInicio', 'N/D')} | Fim: {row.get('vigenciaFim', 'N/D')}"
        objeto = row.get('objetoContratacao', 'N/D')
        info_extra = f"Vigência: {vigencia}"
    elif tipo == "Contratos":
        processo = row.get('processo', 'N/D')
        fornecedor = row.get('nomeRazaoSocialFornecedor', 'N/D')
        objeto = row.get('objetoContrato', 'N/D')
        valor = row.get('valorGlobal', row.get('valorInicial', 0))
        try:
            val_float = float(valor)
            valor_fmt = f"R$ {val_float:,.2f}"
        except:
            valor_fmt = str(valor)
        info_extra = f"Fornecedor: {fornecedor} | Valor: {valor_fmt}"
    else: # Editais
        processo = row.get('processo', 'N/D')
        fornecedor = row.get('usuarioNome', 'N/D')
        objeto = row.get('objetoCompra', 'N/D')
        valor = row.get('valorTotalHomologado', row.get('valorTotalEstimado', 0))
        try:
            val_float = float(valor)
            valor_fmt = f"R$ {val_float:,.2f}"
        except:
            valor_fmt = str(valor)
        info_extra = f"Responsável: {fornecedor} | Valor: {valor_fmt}"

    return str(id_pncp), str(processo), info_extra, str(objeto)

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
        with st.spinner("🔄 Buscando e tratando dados no PNCP..."):
            registros = consultar_paginas(endpoint, params)

        if registros:
            df_temp = pd.DataFrame(registros)
            df_temp = tratar_dataframe(df_temp)
            
            if tipo_consulta == "Contratos":
                possiveis_colunas = ["cnpjOrgao", "orgaoEntidade", "orgao", "unidadeOrgao"]
                for coluna in possiveis_colunas:
                    if coluna in df_temp.columns:
                        serie = df_temp[coluna].astype(str).str.replace(r"\D", "", regex=True)
                        mask = serie.str.contains(CNPJ_RIO_DAS_PEDRAS, na=False)
                        if mask.any():
                            df_temp = df_temp[mask]
                            break

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
# EXIBIÇÃO PERSISTENTE NO DASHBOARD
# ============================================================

if st.session_state.df_resultado is not None and not st.session_state.df_resultado.empty:
    df = st.session_state.df_resultado
    st.success(f"📊 Exibindo {len(df)} registros para Rio das Pedras/SP.")
    
    # KPIs
    col_m1, col_m2 = st.columns(2)
    col_m1.metric("Total de Registros Encontrados", len(df))
    
    coluna_valor = next((c for c in ["valorGlobal", "valorInicial", "valorTotalHomologado", "valorTotalEstimado"] if c in df.columns), None)
    if coluna_valor:
        total_valor = pd.to_numeric(df[coluna_valor], errors='coerce').sum()
        col_m2.metric("Valor Total Envolvido", f"R$ {total_valor:,.2f}")
    else:
        col_m2.metric("Status da Consulta", "Concluída com Sucesso")

    st.markdown("---")

    # OPÇÕES DE EXPORTAÇÃO NO TOPO
    st.markdown("### 📥 Opções de Exportação")
    cols = st.columns(4)
    nome = tipo_consulta.replace(" ", "_").replace("/", "_")

    # Excel
    buffer_xlsx = io.BytesIO()
    df.to_excel(buffer_xlsx, index=False)
    cols[0].download_button("📊 Excel (.xlsx)", buffer_xlsx.getvalue(), f"{nome}_Rio_Das_Pedras.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # CSV
    cols[1].download_button("📄 CSV (.csv)", df.to_csv(index=False, encoding="utf-8-sig"), f"{nome}_Rio_Das_Pedras.csv", mime="text/csv")

    # Word
    doc = Document()
    p_titulo = doc.add_paragraph()
    r_titulo = p_titulo.add_run(f"Relatório Executivo: {tipo_consulta}")
    r_titulo.bold = True
    r_titulo.font.size = Pt(16)
    r_titulo.font.color.rgb = RGBColor(0, 51, 102)
    doc.add_paragraph("Município: Prefeitura Municipal de Rio das Pedras / SP")
    doc.add_paragraph(f"Período: {data_inicio.strftime('%d/%m/%Y')} a {data_fim.strftime('%d/%m/%Y')}")
    doc.add_paragraph(f"Total de Registros: {len(df)}")
    doc.add_heading("Detalhamento dos Registros", level=2)

    for idx, row in df.head(50).iterrows():
        p_reg = doc.add_paragraph()
        p_reg.add_run(f"Item #{idx + 1}\n").bold = True
        id_pncp, processo, info_extra, objeto = obter_dados_registro(row, tipo_consulta)
        p_reg.add_run(f"• ID Contratação PNCP: {id_pncp}\n")
        p_reg.add_run(f"• Processo/Ref: {processo}\n")
        p_reg.add_run(f"• Detalhes: {info_extra}\n")
        p_reg.add_run(f"• Objeto: {objeto}\n")
        doc.add_paragraph("-" * 40)

    buffer_docx = io.BytesIO()
    doc.save(buffer_docx)
    buffer_docx.seek(0)
    cols[2].download_button("📝 Word (.docx)", buffer_docx.getvalue(), f"Relatorio_{nome}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

    # PDF
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
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
        id_pncp, processo, info_extra, objeto = obter_dados_registro(row, tipo_consulta)
        bloco = f"[{idx+1}] ID PNCP: {id_pncp} | Proc: {processo} | {info_extra}\nObjeto: {objeto}"
        bloco_limpo = bloco.encode("latin-1", "replace").decode("latin-1")
        pdf.multi_cell(0, 5, txt=bloco_limpo)
        pdf.ln(3)
    
    pdf_bytes = pdf.output(dest="S")
    if isinstance(pdf_bytes, str):
        pdf_bytes = pdf_bytes.encode("latin-1")
    cols[3].download_button("📕 PDF (.pdf)", pdf_bytes, f"Relatorio_{nome}.pdf", mime="application/pdf")

    st.markdown("---")

    # CONSULTA DE ADITIVOS (APARECE APENAS SE FOR CONTRATOS)
    if tipo_consulta == "Contratos":
        st.markdown("### 🔍 Consultar Aditivos / Documentos por Contrato")
        lista_contratos = df.apply(lambda x: f"{x.get('numeroControlePNCP')} - Proc: {x.get('processo')}", axis=1).tolist()
        contrato_selecionado = st.selectbox("Selecione um contrato:", lista_contratos)
        
        if st.button("Buscar Aditivos do Contrato"):
            id_escolhido = contrato_selecionado.split(" - ")[0]
            with st.spinner("Buscando aditivos no PNCP..."):
                aditivos = consultar_detalhes_contrato(id_escolhido)
                if aditivos:
                    st.success(f"Encontrados {len(aditivos)} documentos vinculados:")
                    for doc in aditivos:
                        tipo_doc = doc.get('tipoDocumentoNome', 'Outro')
                        st.info(f"**Tipo:** {tipo_doc} | **Data:** {doc.get('dataPublicacao', 'N/D')}\n\n{doc.get('objeto', '')}")
                else:
                    st.warning("Nenhum documento/aditivo encontrado para este contrato.")
        st.markdown("---")

    # GRÁFICOS
    st.markdown("### 📈 Análise Gráfica")
    coluna_data = next((c for c in ["dataPublicacao", "dataAssinatura", "dataInclusao"] if c in df.columns), None)
    if coluna_data:
        try:
            df_grafico = df.copy()
            df_grafico['mes_ano'] = pd.to_datetime(df_grafico[coluna_data], errors='coerce').dt.to_period('M').astype(str)
            contagem_mes = df_grafico['mes_ano'].value_counts().sort_index()
            if not contagem_mes.empty:
                st.bar_chart(contagem_mes)
            else:
                st.info("ℹ️ Dados insuficientes para gerar gráfico por período.")
        except Exception:
            st.info("ℹ️ Não foi possível gerar o gráfico temporal automaticamente.")
    else:
        st.info("ℹ️ Coluna de data não encontrada para exibição do gráfico temporal.")

    st.markdown("---")
    
    # TABELA
    st.markdown("### 📋 Tabela de Dados Detalhada")
    st.dataframe(df, use_container_width=True, hide_index=True)

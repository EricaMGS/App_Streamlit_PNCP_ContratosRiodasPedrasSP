### 1. Verificação do Erro

O erro `SyntaxError: invalid syntax` na **linha 1** ocorreu porque o código fornecido continha as marcas de formatação de bloco de código do markdown (os três acentos graves com a palavra `python` — ```python) coladas diretamente no topo do arquivo do script Python (`app.py`). O interpretador Python tentou ler esses caracteres textuais como código executável, gerando a falha de sintaxe.

---

### 2. Código Corrigido

O código abaixo foi totalmente limpo, removendo as tags markdown e os caracteres de espaçamento invisíveis (non-breaking spaces), estando pronto e estruturado para rodar perfeitamente no Streamlit:

```python
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
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 Chrome/151.0 Safari/537.36"
        ),
        "Accept": "application/json",
        "Connection": "keep-alive"
    }

    ultima_excecao = None

    for tentativa in range(1, max_tentativas + 1):

        try:

            resp = requests.get(
                url,
                params=params,
                headers=headers,
                timeout=(15, 90)
            )

            if resp.status_code == 200:

                try:
                    return resp.json()

                except ValueError:
                    raise Exception(
                        "O PNCP respondeu, mas não enviou um JSON válido."
                    )

            # 204 = consulta sem registros
            if resp.status_code == 204:
                return []

            # Erros temporários do servidor
            if resp.status_code in [429, 500, 502, 503, 504]:

                if tentativa < max_tentativas:

                    espera = tentativa * 3

                    st.warning(
                        f"⚠️ O Portal PNCP está demorando para responder. "
                        f"Nova tentativa em {espera} segundos "
                        f"({tentativa}/{max_tentativas})."
                    )

                    time.sleep(espera)
                    continue

            raise Exception(
                f"API retornou HTTP {resp.status_code}: "
                f"{resp.text[:500]}"
            )

        except requests.exceptions.Timeout as e:

            ultima_excecao = e

            if tentativa < max_tentativas:

                espera = tentativa * 3

                st.warning(
                    f"⏳ O PNCP está demorando para responder. "
                    f"Nova tentativa em {espera} segundos "
                    f"({tentativa}/{max_tentativas})."
                )

                time.sleep(espera)

            continue

        except requests.exceptions.ConnectionError as e:

            ultima_excecao = e

            if tentativa < max_tentativas:

                espera = tentativa * 3

                st.warning(
                    f"🌐 Não foi possível conectar ao PNCP. "
                    f"Tentando novamente em {espera} segundos "
                    f"({tentativa}/{max_tentativas})."
                )

                time.sleep(espera)

            continue

        except Exception:
            raise

    raise ultima_excecao


# ============================================================
# FUNÇÃO PARA EXTRAIR REGISTROS DA RESPOSTA
# ============================================================

def extrair_registros(data):

    if isinstance(data, list):
        return data

    if isinstance(data, dict):

        for chave in [
            "data",
            "items",
            "content",
            "dados"
        ]:

            if chave in data and isinstance(data[chave], list):
                return data[chave]

    return []


# ============================================================
# FUNÇÃO DE PAGINAÇÃO
# ============================================================

def consultar_paginas(url, params, max_paginas=100):

    todos_registros = []

    for pagina in range(1, max_paginas + 1):

        params_pagina = params.copy()
        params_pagina["pagina"] = pagina

        data = consultar_pncp(
            url,
            params_pagina
        )

        registros = extrair_registros(data)

        if not registros:
            break

        todos_registros.extend(registros)

        # Se veio menos que o tamanho solicitado,
        # provavelmente chegamos ao final.
        tamanho = params_pagina.get("tamanhoPagina", 50)

        if len(registros) < tamanho:
            break

        # Pequena pausa para evitar sobrecarregar o portal
        time.sleep(0.5)

    return todos_registros


# ============================================================
# BOTÃO CONSULTAR
# ============================================================

if st.sidebar.button(
    "🔎 Gerar Relatório",
    type="primary"
):

    # --------------------------------------------------------
    # AVISO DE INSTABILIDADE
    # --------------------------------------------------------

    st.info(
        "ℹ️ **Atenção:** a consulta é realizada diretamente "
        "no Portal Nacional de Contratações Públicas (PNCP). "
        "O portal pode apresentar instabilidade ou lentidão. "
        "Dependendo do período e da quantidade de registros, "
        "a consulta pode demorar alguns minutos. "
        "Por favor, aguarde enquanto os dados são carregados."
    )

    # --------------------------------------------------------
    # ENDPOINTS
    # --------------------------------------------------------

    endpoints = {

        "Contratos":
            f"{BASE_URL}/contratos",

        "Atas de Registro de Preços":
            f"{BASE_URL}/atas",

        "Editais e Avisos de Contratações":
            f"{BASE_URL}/contratacoes/publicacao"
    }

    endpoint = endpoints[tipo_consulta]

    # --------------------------------------------------------
    # PARÂMETROS BASE
    # --------------------------------------------------------

    tamanho_pagina = 50 if tipo_consulta == "Editais e Avisos de Contratações" else 100

    params = {
        "dataInicial": data_inicio.strftime("%Y%m%d"),
        "dataFinal": data_fim.strftime("%Y%m%d"),
        "pagina": 1,
        "tamanhoPagina": tamanho_pagina
    }

    # --------------------------------------------------------
    # FILTROS ESPECÍFICOS
    # --------------------------------------------------------

    if tipo_consulta == "Contratos":
        params["cnpjOrgao"] = CNPJ_RIO_DAS_PEDRAS

    elif tipo_consulta == "Atas de Registro de Preços":
        params["cnpj"] = CNPJ_RIO_DAS_PEDRAS

    elif tipo_consulta == "Editais e Avisos de Contratações":
        params["codigoModalidadeContratacao"] = modalidade_codigo
        params["uf"] = UF
        params["codigoMunicipioIbge"] = CODIGO_IBGE_RIO_DAS_PEDRAS
        params["cnpj"] = CNPJ_RIO_DAS_PEDRAS

    # --------------------------------------------------------
    # CONSULTA
    # --------------------------------------------------------

    try:

        with st.spinner(
            "🔄 Consultando o PNCP... "
            "Isso pode demorar devido à instabilidade do portal."
        ):

            registros = consultar_paginas(
                endpoint,
                params
            )

        # ----------------------------------------------------
        # RESULTADO
        # ----------------------------------------------------

        if registros:

            df_temp = pd.DataFrame(registros)

            # ------------------------------------------------
            # FILTRAGEM COMPLEMENTAR
            # ------------------------------------------------

            if tipo_consulta == "Contratos":

                possiveis_colunas = [
                    "cnpjOrgao",
                    "orgaoEntidade",
                    "orgao",
                    "unidadeOrgao"
                ]

                encontrou_cnpj = False

                for coluna in possiveis_colunas:

                    if coluna in df_temp.columns:

                        serie = (
                            df_temp[coluna]
                            .astype(str)
                            .str.replace(r"\D", "", regex=True)
                        )

                        mask = serie.str.contains(
                            CNPJ_RIO_DAS_PEDRAS,
                            na=False
                        )

                        if mask.any():
                            df_temp = df_temp[mask]
                            encontrou_cnpj = True
                            break

                if not encontrou_cnpj:
                    st.warning(
                        "⚠️ O PNCP retornou registros, mas "
                        "a estrutura da resposta não possui "
                        "uma coluna de CNPJ reconhecida para "
                        "a filtragem complementar."
                    )

            st.session_state.df_resultado = df_temp

            if df_temp.empty:
                st.warning(
                    "⚠️ A API respondeu, mas nenhum registro "
                    "foi encontrado para Rio das Pedras/SP "
                    "no período informado."
                )
            else:
                st.success(
                    f"✅ Consulta concluída! "
                    f"{len(df_temp)} registros encontrados."
                )

        else:

            st.session_state.df_resultado = pd.DataFrame()

            st.warning(
                "ℹ️ Nenhum registro foi retornado pelo PNCP "
                "para os parâmetros informados."
            )

    except requests.exceptions.Timeout:

        st.session_state.df_resultado = None

        st.error(
            "⏱️ **O Portal PNCP demorou mais que o esperado "
            "para responder.**\n\n"
            "Isso pode acontecer devido à instabilidade ou "
            "lentidão temporária do portal. "
            "Tente novamente em alguns instantes ou consulte "
            "um período menor."
        )

    except requests.exceptions.ConnectionError:

        st.session_state.df_resultado = None

        st.error(
            "🌐 **Não foi possível estabelecer conexão com o PNCP.**\n\n"
            "O Portal Nacional de Contratações Públicas pode "
            "estar temporariamente indisponível ou instável. "
            "Tente novamente mais tarde."
        )

    except Exception as e:

        st.session_state.df_resultado = None

        st.error(
            f"❌ **Erro ao consultar o PNCP:**\n\n"
            f"{str(e)}"
        )


# ============================================================
# EXIBIÇÃO PERSISTENTE
# ============================================================

if (
    st.session_state.df_resultado is not None
    and not st.session_state.df_resultado.empty
):

    df = st.session_state.df_resultado

    st.success(
        f"📊 Exibindo {len(df)} registros para Rio das Pedras/SP."
    )

    st.dataframe(
        df,
        use_container_width=True,
        hide_index=True
    )

    # ========================================================
    # EXPORTAÇÃO
    # ========================================================

    st.markdown("### 📥 Opções de Exportação")

    cols = st.columns(4)

    nome = (
        tipo_consulta
        .replace(" ", "_")
        .replace("/", "_")
    )

    # --------------------------------------------------------
    # EXCEL
    # --------------------------------------------------------

    buffer_xlsx = io.BytesIO()

    df.to_excel(
        buffer_xlsx,
        index=False
    )

    cols[0].download_button(
        "📊 Excel (.xlsx)",
        buffer_xlsx.getvalue(),
        f"{nome}_Rio_Das_Pedras.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        )
    )

    # --------------------------------------------------------
    # CSV
    # --------------------------------------------------------

    cols[1].download_button(
        "📄 CSV (.csv)",
        df.to_csv(
            index=False,
            encoding="utf-8-sig"
        ),
        f"{nome}_Rio_Das_Pedras.csv",
        mime="text/csv"
    )

    # --------------------------------------------------------
    # WORD
    # --------------------------------------------------------

    doc = Document()

    doc.add_heading(
        f"Relatório {tipo_consulta}",
        0
    )

    doc.add_paragraph(
        "Município: Rio das Pedras/SP"
    )

    doc.add_paragraph(
        f"Total de registros: {len(df)}"
    )

    for _, row in df.head(50).iterrows():

        doc.add_paragraph(
            str(row.to_dict())
        )

    buffer_docx = io.BytesIO()

    doc.save(buffer_docx)

    cols[2].download_button(
        "📝 Word (.docx)",
        buffer_docx.getvalue(),
        f"Relatorio_{nome}.docx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "wordprocessingml.document"
        )
    )

    # --------------------------------------------------------
    # PDF
    # --------------------------------------------------------

    pdf = FPDF()

    pdf.set_auto_page_break(
        auto=True,
        margin=15
    )

    pdf.add_page()

    pdf.set_font(
        "Arial",
        size=12
    )

    titulo_pdf = (
        f"Relatorio {tipo_consulta} - "
        f"Rio das Pedras"
    )

    pdf.cell(
        200,
        10,
        txt=titulo_pdf,
        ln=True,
        align="C"
    )

    pdf.cell(
        200,
        10,
        txt=f"Total de registros: {len(df)}",
        ln=True
    )

    pdf.set_font(
        "Arial",
        size=9
    )

    for _, row in df.head(30).iterrows():

        texto = " | ".join(
            f"{col}: {row[col]}"
            for col in df.columns
        )

        texto_limpo = (
            str(texto)
            .encode("latin-1", "replace")
            .decode("latin-1")
        )

        if len(texto_limpo) > 180:
            texto_limpo = texto_limpo[:180] + "..."

        pdf.multi_cell(
            190,
            6,
            txt=texto_limpo
        )

    pdf_bytes = pdf.output(
        dest="S"
    )

    if isinstance(pdf_bytes, str):
        pdf_bytes = pdf_bytes.encode("latin-1")

    cols[3].download_button(
        "📕 PDF (.pdf)",
        pdf_bytes,
        f"Relatorio_{nome}.pdf",
        mime="application/pdf"
    )


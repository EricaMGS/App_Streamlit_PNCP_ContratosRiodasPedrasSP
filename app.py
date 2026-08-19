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
# CONFIGURAÇÃO DO STREAMLIT
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
CODIGO_IBGE_RIO_DAS_PEDRAS = "3544003"
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
# MODALIDADE DOS EDITAIS
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
    value=datetime.date(2026, 1, 1)
)

data_fim = st.sidebar.date_input(
    "Data Final",
    value=datetime.date.today()
)


# ============================================================
# VALIDAÇÃO DAS DATAS
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

if "df_resultado" not in st.session_state:
    st.session_state.df_resultado = None


if "tipo_anterior" not in st.session_state:
    st.session_state.tipo_anterior = tipo_consulta


if st.session_state.tipo_anterior != tipo_consulta:

    st.session_state.df_resultado = None

    st.session_state.tipo_anterior = tipo_consulta


# ============================================================
# FUNÇÃO PARA EXTRAIR REGISTROS
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

            valor = data.get(chave)

            if isinstance(valor, list):
                return valor

    return []


# ============================================================
# CONSULTA AO PNCP COM RETENTATIVAS
# ============================================================

def consultar_pncp(
    url,
    params,
    tentativa_maxima=3
):

    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 "
            "Chrome/151.0.0.0 Safari/537.36"
        ),
        "Accept": "application/json",
        "Connection": "keep-alive"
    }

    ultimo_erro = None

    for tentativa in range(
        1,
        tentativa_maxima + 1
    ):

        try:

            resposta = requests.get(
                url,
                params=params,
                headers=headers,
                timeout=(15, 90)
            )

            # ==================================================
            # SUCESSO
            # ==================================================

            if resposta.status_code == 200:

                try:

                    return resposta.json()

                except ValueError:

                    raise Exception(
                        "O PNCP respondeu, mas a resposta "
                        "não está em formato JSON válido."
                    )


            # ==================================================
            # SEM CONTEÚDO
            # ==================================================

            if resposta.status_code == 204:

                return []


            # ==================================================
            # ERROS TEMPORÁRIOS
            # ==================================================

            if resposta.status_code in [
                429,
                500,
                502,
                503,
                504
            ]:

                ultimo_erro = Exception(
                    f"Servidor PNCP retornou HTTP "
                    f"{resposta.status_code}"
                )

                if tentativa < tentativa_maxima:

                    espera = tentativa * 5

                    st.warning(
                        f"⚠️ O Portal PNCP está temporariamente "
                        f"instável ou sobrecarregado. "
                        f"Tentando novamente em {espera} segundos "
                        f"({tentativa}/{tentativa_maxima})..."
                    )

                    time.sleep(espera)

                    continue

                raise ultimo_erro


            # ==================================================
            # OUTROS ERROS HTTP
            # ==================================================

            raise Exception(
                f"Erro na API do PNCP. "
                f"Status HTTP {resposta.status_code}: "
                f"{resposta.text[:500]}"
            )


        # ======================================================
        # TIMEOUT
        # ======================================================

        except requests.exceptions.Timeout as erro:

            ultimo_erro = erro

            if tentativa < tentativa_maxima:

                espera = tentativa * 5

                st.warning(
                    f"⏳ O PNCP está demorando para responder. "
                    f"Nova tentativa em {espera} segundos "
                    f"({tentativa}/{tentativa_maxima})..."
                )

                time.sleep(espera)

                continue


        # ======================================================
        # ERRO DE CONEXÃO
        # ======================================================

        except requests.exceptions.ConnectionError as erro:

            ultimo_erro = erro

            if tentativa < tentativa_maxima:

                espera = tentativa * 5

                st.warning(
                    f"🌐 Não foi possível conectar ao PNCP. "
                    f"Nova tentativa em {espera} segundos "
                    f"({tentativa}/{tentativa_maxima})..."
                )

                time.sleep(espera)

                continue


        # ======================================================
        # OUTROS ERROS
        # ======================================================

        except Exception:

            raise


    if ultimo_erro:

        raise ultimo_erro


    raise Exception(
        "Não foi possível consultar o PNCP."
    )


# ============================================================
# CONSULTA PAGINADA
# ============================================================

def consultar_paginas(
    url,
    parametros
):

    todos_registros = []

    max_paginas = 100

    for pagina in range(
        1,
        max_paginas + 1
    ):

        parametros_pagina = parametros.copy()

        parametros_pagina["pagina"] = pagina

        resposta = consultar_pncp(
            url,
            parametros_pagina
        )

        registros = extrair_registros(
            resposta
        )

        if not registros:

            break

        todos_registros.extend(
            registros
        )

        tamanho_pagina = parametros_pagina.get(
            "tamanhoPagina",
            50
        )

        if len(registros) < tamanho_pagina:

            break

        time.sleep(0.5)

    return todos_registros


# ============================================================
# BOTÃO GERAR RELATÓRIO
# ============================================================

if st.sidebar.button(
    "🔎 Gerar Relatório",
    type="primary"
):

    # ========================================================
    # AVISO AO USUÁRIO
    # ========================================================

    st.info(
        "ℹ️ **A consulta é feita diretamente no Portal PNCP.** "
        "O portal pode apresentar instabilidade ou lentidão. "
        "Dependendo do período e da quantidade de registros, "
        "a consulta pode demorar alguns minutos. "
        "O sistema fará novas tentativas automaticamente "
        "se o PNCP demorar para responder."
    )


    # ========================================================
    # ENDPOINTS
    # ========================================================

    endpoints = {

        "Contratos":
            f"{BASE_URL}/contratos",

        "Atas de Registro de Preços":
            f"{BASE_URL}/atas",

        "Editais e Avisos de Contratações":
            f"{BASE_URL}/contratacoes/publicacao"
    }


    endpoint = endpoints[
        tipo_consulta
    ]


    # ========================================================
    # TAMANHO DA PÁGINA
    #
    # IMPORTANTE:
    # O endpoint de Editais/Avisos estava retornando:
    #
    # HTTP 400
    # "Tamanho de página inválido"
    #
    # Por isso:
    # - Contratos = 100
    # - Atas = 100
    # - Editais = 50
    # ========================================================

    if tipo_consulta == "Editais e Avisos de Contratações":

        tamanho_pagina = 50

    else:

        tamanho_pagina = 100


    # ========================================================
    # PARÂMETROS BÁSICOS
    # ========================================================

    parametros = {

        "dataInicial":
            data_inicio.strftime("%Y%m%d"),

        "dataFinal":
            data_fim.strftime("%Y%m%d"),

        "pagina":
            1,

        "tamanhoPagina":
            tamanho_pagina
    }


    # ========================================================
    # CONTRATOS
    # ========================================================

    if tipo_consulta == "Contratos":

        parametros[
            "cnpjOrgao"
        ] = CNPJ_RIO_DAS_PEDRAS


    # ========================================================
    # ATAS
    # ========================================================

    elif tipo_consulta == "Atas de Registro de Preços":

        parametros[
            "cnpj"
        ] = CNPJ_RIO_DAS_PEDRAS


    # ========================================================
    # EDITAIS E AVISOS
    # ========================================================

    elif tipo_consulta == "Editais e Avisos de Contratações":

        parametros[
            "codigoModalidadeContratacao"
        ] = modalidade_codigo

        parametros[
            "uf"
        ] = UF

        parametros[
            "codigoMunicipioIbge"
        ] = CODIGO_IBGE_RIO_DAS_PEDRAS

        parametros[
            "cnpj"
        ] = CNPJ_RIO_DAS_PEDRAS


    # ========================================================
    # CONSULTA
    # ========================================================

    try:

        with st.spinner(
            "🔄 Consultando o Portal PNCP... "
            "Aguarde, esta consulta pode demorar."
        ):

            registros = consultar_paginas(
                endpoint,
                parametros
            )


        # ====================================================
        # PROCESSAMENTO DOS RESULTADOS
        # ====================================================

        if registros:

            df_temp = pd.DataFrame(
                registros
            )


            # =================================================
            # FILTRAGEM COMPLEMENTAR DOS CONTRATOS
            # =================================================

            if tipo_consulta == "Contratos":

                colunas_cnpj = [
                    "cnpjOrgao",
                    "orgaoEntidade",
                    "orgao",
                    "unidadeOrgao"
                ]

                encontrou_cnpj = False


                for coluna in colunas_cnpj:

                    if coluna in df_temp.columns:

                        valores = (
                            df_temp[coluna]
                            .astype(str)
                            .str.replace(
                                r"\D",
                                "",
                                regex=True
                            )
                        )


                        mascara = valores.str.contains(
                            CNPJ_RIO_DAS_PEDRAS,
                            na=False
                        )


                        if mascara.any():

                            df_temp = df_temp[
                                mascara
                            ]

                            encontrou_cnpj = True

                            break


                if not encontrou_cnpj:

                    st.warning(
                        "⚠️ O PNCP retornou registros, mas "
                        "a estrutura da resposta não possui "
                        "uma coluna de CNPJ reconhecida para "
                        "a filtragem complementar."
                    )


            # =================================================
            # SALVAR RESULTADO
            # =================================================

            st.session_state.df_resultado = (
                df_temp
            )


            # =================================================
            # RESULTADO VAZIO
            # =================================================

            if df_temp.empty:

                st.warning(
                    "⚠️ Nenhum registro encontrado para "
                    "Rio das Pedras/SP no período informado."
                )


            # =================================================
            # RESULTADO OK
            # =================================================

            else:

                st.success(
                    f"✅ Consulta concluída! "
                    f"{len(df_temp)} registro(s) encontrado(s)."
                )


        # ====================================================
        # NENHUM REGISTRO
        # ====================================================

        else:

            st.session_state.df_resultado = (
                pd.DataFrame()
            )

            st.warning(
                "ℹ️ O PNCP não retornou registros para "
                "os parâmetros informados."
            )


    # ========================================================
    # TIMEOUT FINAL
    # ========================================================

    except requests.exceptions.Timeout:

        st.session_state.df_resultado = None

        st.error(
            "⏱️ **O Portal PNCP demorou demais para responder.**\n\n"
            "O sistema tentou realizar a consulta mais de uma vez, "
            "mas o portal não respondeu dentro do tempo esperado.\n\n"
            "Isso pode acontecer devido à instabilidade ou "
            "lentidão temporária do PNCP.\n\n"
            "**Tente novamente ou utilize um período menor.**"
        )


    # ========================================================
    # ERRO DE CONEXÃO
    # ========================================================

    except requests.exceptions.ConnectionError:

        st.session_state.df_resultado = None

        st.error(
            "🌐 **Não foi possível conectar ao Portal PNCP.**\n\n"
            "O portal pode estar temporariamente indisponível "
            "ou apresentando instabilidade.\n\n"
            "Tente novamente em alguns minutos."
        )


    # ========================================================
    # ERRO GERAL
    # ========================================================

    except Exception as erro:

        st.session_state.df_resultado = None

        st.error(
            "❌ **Erro ao consultar o PNCP**"
        )

        st.code(
            str(erro)
        )


# ============================================================
# EXIBIÇÃO DOS RESULTADOS
# ============================================================

if (
    st.session_state.df_resultado is not None
    and not st.session_state.df_resultado.empty
):

    df = st.session_state.df_resultado


    st.success(
        f"📊 Exibindo {len(df)} registro(s) "
        f"para Rio das Pedras/SP."
    )


    st.dataframe(
        df,
        use_container_width=True,
        hide_index=True
    )


    # ========================================================
    # EXPORTAÇÃO
    # ========================================================

    st.markdown(
        "### 📥 Opções de Exportação"
    )


    col1, col2, col3, col4 = st.columns(4)


    nome = (
        tipo_consulta
        .replace(" ", "_")
        .replace("/", "_")
    )


    # ========================================================
    # EXCEL
    # ========================================================

    buffer_excel = io.BytesIO()

    df.to_excel(
        buffer_excel,
        index=False
    )

    buffer_excel.seek(0)


    col1.download_button(
        label="📊 Excel (.xlsx)",
        data=buffer_excel.getvalue(),
        file_name=(
            f"{nome}_Rio_Das_Pedras.xlsx"
        ),
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        )
    )


    # ========================================================
    # CSV
    # ========================================================

    csv_data = df.to_csv(
        index=False,
        encoding="utf-8-sig"
    )


    col2.download_button(
        label="📄 CSV (.csv)",
        data=csv_data,
        file_name=(
            f"{nome}_Rio_Das_Pedras.csv"
        ),
        mime="text/csv"
    )


    # ========================================================
    # WORD
    # ========================================================

    documento = Document()


    documento.add_heading(
        f"Relatório - {tipo_consulta}",
        0
    )


    documento.add_paragraph(
        "Município: Rio das Pedras/SP"
    )


    documento.add_paragraph(
        f"Total de registros: {len(df)}"
    )


    documento.add_paragraph(
        f"Período: "
        f"{data_inicio.strftime('%d/%m/%Y')} "
        f"a "
        f"{data_fim.strftime('%d/%m/%Y')}"
    )


    documento.add_heading(
        "Registros",
        level=1
    )


    for _, linha in df.head(50).iterrows():

        texto = " | ".join(
            f"{coluna}: {linha[coluna]}"
            for coluna in df.columns
        )

        documento.add_paragraph(
            texto
        )


    buffer_word = io.BytesIO()


    documento.save(
        buffer_word
    )


    buffer_word.seek(0)


    col3.download_button(
        label="📝 Word (.docx)",
        data=buffer_word.getvalue(),
        file_name=(
            f"Relatorio_{nome}.docx"
        ),
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "wordprocessingml.document"
        )
    )


    # ========================================================
    # PDF
    # ========================================================

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


    titulo = (
        f"Relatorio {tipo_consulta} "
        f"- Rio das Pedras/SP"
    )


    titulo_limpo = (
        titulo
        .encode("latin-1", "replace")
        .decode("latin-1")
    )


    pdf.cell(
        0,
        10,
        txt=titulo_limpo,
        ln=True,
        align="C"
    )


    pdf.set_font(
        "Arial",
        size=9
    )


    pdf.cell(
        0,
        8,
        txt=(
            f"Total de registros: {len(df)}"
        ),
        ln=True
    )


    pdf.cell(
        0,
        8,
        txt=(
            f"Periodo: "
            f"{data_inicio.strftime('%d/%m/%Y')} "
            f"a "
            f"{data_fim.strftime('%d/%m/%Y')}"
        ),
        ln=True
    )


    pdf.ln(5)


    for _, linha in df.head(30).iterrows():

        for coluna in df.columns:

            valor = str(
                linha[coluna]
            )

            texto = (
                f"{coluna}: {valor}"
            )


            texto_limpo = (
                texto
                .encode("latin-1", "replace")
                .decode("latin-1")
            )


            if len(texto_limpo) > 180:

                texto_limpo = (
                    texto_limpo[:180]
                    + "..."
                )


            pdf.multi_cell(
                0,
                5,
                txt=texto_limpo
            )


        pdf.ln(3)


    pdf_bytes = pdf.output(
        dest="S"
    )


    if isinstance(
        pdf_bytes,
        str
    ):

        pdf_bytes = pdf_bytes.encode(
            "latin-1"
        )


    col4.download_button(
        label="📕 PDF (.pdf)",
        data=pdf_bytes,
        file_name=(
            f"Relatorio_{nome}.pdf"
        ),
        mime="application/pdf"
    )
```

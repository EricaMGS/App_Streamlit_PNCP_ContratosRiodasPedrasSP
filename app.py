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
    """Formata número para o padrão brasileiro."""
    try:
        valor = float(valor)
        valor_fmt = "{:,.2f}".format(valor)
        return f"R$ {valor_fmt.replace(',', 'X').replace('.', ',').replace('X', '.')}"
    except (ValueError, TypeError):
        return "N/D"


def valor_numerico(valor):
    """Converte valores numéricos do PNCP de forma segura."""
    if valor is None:
        return None

    if isinstance(valor, (int, float)):
        return float(valor)

    texto = str(valor).strip()

    if not texto:
        return None

    try:
        return float(texto)
    except ValueError:
        pass

    try:
        texto = texto.replace("R$", "").replace(" ", "")

        if "," in texto and "." in texto:
            texto = texto.replace(".", "").replace(",", ".")
        elif "," in texto:
            texto = texto.replace(",", ".")

        return float(texto)
    except (ValueError, TypeError):
        return None


def texto_valido(valor):
    """Retorna texto válido ou N/D."""
    if valor is None:
        return "N/D"

    if isinstance(valor, float) and pd.isna(valor):
        return "N/D"

    texto = str(valor).strip()

    if texto.lower() in ["", "none", "nan", "null", "n/d", "nd"]:
        return "N/D"

    return texto


def obter_primeiro_valor(row, campos, padrao="N/D"):
    """
    Procura o primeiro campo disponível no registro.
    Permite trabalhar com diferentes nomes de campos retornados pelo PNCP.
    """
    for campo in campos:
        if campo in row:
            valor = row.get(campo)

            if isinstance(valor, dict):
                valor = (
                    valor.get("nome")
                    or valor.get("razaoSocial")
                    or valor.get("descricao")
                    or valor.get("valor")
                )

            if valor is not None and str(valor).strip() not in [
                "",
                "None",
                "nan",
                "null",
            ]:
                return valor

    return padrao


def formatar_data(valor):
    """Converte datas do PNCP para DD/MM/AAAA."""
    if valor is None:
        return "N/D"

    texto = str(valor).strip()

    if not texto or texto.lower() in ["none", "nan", "null", "n/d"]:
        return "N/D"

    try:
        data = pd.to_datetime(valor, errors="coerce")

        if pd.isna(data):
            return texto

        return data.strftime("%d/%m/%Y")
    except Exception:
        return texto


# ============================================================
# DADOS PRINCIPAIS DO REGISTRO
# ============================================================

def obter_dados_registro(row, tipo):
    """
    Extrai os principais dados para Word e PDF.

    A função foi ampliada principalmente para ATAS DE REGISTRO
    DE PREÇOS, pois os nomes dos campos podem variar conforme
    o retorno do PNCP.
    """

    # --------------------------------------------------------
    # IDENTIFICAÇÃO PNCP
    # --------------------------------------------------------

    id_pncp = obter_primeiro_valor(
        row,
        [
            "numeroControlePNCP",
            "numeroControlePNCPAta",
            "numeroControlePNCPCompra",
            "numeroControlePncp",
            "numeroControlePNCPAta",
        ],
    )

    # --------------------------------------------------------
    # ATA DE REGISTRO DE PREÇOS
    # --------------------------------------------------------

    if tipo == "Atas de Registro de Preços":

        numero_ata = obter_primeiro_valor(
            row,
            [
                "numeroAtaRegistroPreco",
                "numeroAta",
                "numeroRegistroPreco",
                "numeroAtaRegistroPrecos",
                "numero",
            ],
        )

        processo = obter_primeiro_valor(
            row,
            [
                "processo",
                "numeroProcesso",
                "processoAdministrativo",
                "numeroProcessoAdministrativo",
            ],
        )

        objeto = obter_primeiro_valor(
            row,
            [
                "objetoCompra",
                "objetoAta",
                "objeto",
                "descricaoObjeto",
                "descricao",
            ],
        )

        fornecedor = obter_primeiro_valor(
            row,
            [
                "nomeRazaoSocialFornecedor",
                "razaoSocialFornecedor",
                "nomeFornecedor",
                "fornecedorNome",
                "razaoSocial",
                "nomeRazaoSocial",
            ],
        )

        cnpj_fornecedor = obter_primeiro_valor(
            row,
            [
                "niFornecedor",
                "cnpjFornecedor",
                "numeroDocumentoFornecedor",
                "documentoFornecedor",
                "cpfCnpjFornecedor",
            ],
        )

        valor = obter_primeiro_valor(
            row,
            [
                "valorTotal",
                "valorGlobal",
                "valorTotalAta",
                "valorAta",
                "valorTotalEstimado",
                "valorTotalHomologado",
                "valor",
            ],
            padrao=None,
        )

        valor_num = valor_numerico(valor)
        valor_formatado = (
            formatar_moeda_br(valor_num)
            if valor_num is not None
            else "N/D"
        )

        data_assinatura = obter_primeiro_valor(
            row,
            [
                "dataAssinatura",
                "dataAssinaturaAta",
                "dataCelebracao",
                "dataFormalizacao",
            ],
        )

        vigencia_inicio = obter_primeiro_valor(
            row,
            [
                "vigenciaInicio",
                "dataInicioVigencia",
                "dataVigenciaInicio",
                "inicioVigencia",
            ],
        )

        vigencia_fim = obter_primeiro_valor(
            row,
            [
                "vigenciaFim",
                "dataFimVigencia",
                "dataVigenciaFim",
                "fimVigencia",
            ],
        )

        situacao = obter_primeiro_valor(
            row,
            [
                "situacao",
                "situacaoAta",
                "status",
                "situacaoRegistro",
            ],
        )

        orgao = obter_primeiro_valor(
            row,
            [
                "orgaoEntidade",
                "nomeOrgao",
                "razaoSocialOrgao",
                "orgao",
            ],
        )

        unidade = obter_primeiro_valor(
            row,
            [
                "unidadeOrgao",
                "nomeUnidade",
                "unidade",
            ],
        )

        return {
            "id_pncp": texto_valido(id_pncp),
            "numero": texto_valido(numero_ata),
            "processo": texto_valido(processo),
            "objeto": texto_valido(objeto),
            "fornecedor": texto_valido(fornecedor),
            "cnpj_fornecedor": texto_valido(cnpj_fornecedor),
            "valor": valor_formatado,
            "data_assinatura": formatar_data(data_assinatura),
            "vigencia_inicio": formatar_data(vigencia_inicio),
            "vigencia_fim": formatar_data(vigencia_fim),
            "situacao": texto_valido(situacao),
            "orgao": texto_valido(orgao),
            "unidade": texto_valido(unidade),
        }

    # --------------------------------------------------------
    # CONTRATOS
    # --------------------------------------------------------

    elif tipo == "Contratos":

        processo = obter_primeiro_valor(
            row,
            [
                "processo",
                "numeroProcesso",
                "processoAdministrativo",
            ],
        )

        objeto = obter_primeiro_valor(
            row,
            [
                "objetoContrato",
                "objetoCompra",
                "objeto",
                "descricaoObjeto",
            ],
        )

        fornecedor = obter_primeiro_valor(
            row,
            [
                "nomeRazaoSocialFornecedor",
                "razaoSocialFornecedor",
                "nomeFornecedor",
                "fornecedorNome",
                "razaoSocial",
            ],
        )

        valor = obter_primeiro_valor(
            row,
            [
                "valorGlobal",
                "valorInicial",
                "valorTotal",
                "valorContrato",
            ],
            padrao=None,
        )

        valor_num = valor_numerico(valor)

        return {
            "id_pncp": texto_valido(id_pncp),
            "numero": texto_valido(
                obter_primeiro_valor(
                    row,
                    ["numeroContrato", "numeroContratoPncp", "numero"],
                )
            ),
            "processo": texto_valido(processo),
            "objeto": texto_valido(objeto),
            "fornecedor": texto_valido(fornecedor),
            "cnpj_fornecedor": texto_valido(
                obter_primeiro_valor(
                    row,
                    [
                        "niFornecedor",
                        "cnpjFornecedor",
                        "numeroDocumentoFornecedor",
                    ],
                )
            ),
            "valor": (
                formatar_moeda_br(valor_num)
                if valor_num is not None
                else "N/D"
            ),
            "data_assinatura": formatar_data(
                obter_primeiro_valor(
                    row,
                    [
                        "dataAssinatura",
                        "dataCelebracao",
                    ],
                )
            ),
            "vigencia_inicio": formatar_data(
                obter_primeiro_valor(
                    row,
                    [
                        "dataVigenciaInicio",
                        "vigenciaInicio",
                    ],
                )
            ),
            "vigencia_fim": formatar_data(
                obter_primeiro_valor(
                    row,
                    [
                        "dataVigenciaFim",
                        "vigenciaFim",
                    ],
                )
            ),
            "situacao": texto_valido(
                obter_primeiro_valor(
                    row,
                    [
                        "situacao",
                        "status",
                    ],
                )
            ),
            "orgao": texto_valido(
                obter_primeiro_valor(
                    row,
                    [
                        "orgaoEntidade",
                        "nomeOrgao",
                        "orgao",
                    ],
                )
            ),
            "unidade": texto_valido(
                obter_primeiro_valor(
                    row,
                    [
                        "unidadeOrgao",
                        "nomeUnidade",
                        "unidade",
                    ],
                )
            ),
        }

    # --------------------------------------------------------
    # EDITAIS
    # --------------------------------------------------------

    else:

        processo = obter_primeiro_valor(
            row,
            [
                "processo",
                "numeroProcesso",
                "processoAdministrativo",
            ],
        )

        objeto = obter_primeiro_valor(
            row,
            [
                "objetoCompra",
                "objeto",
                "descricaoObjeto",
            ],
        )

        valor = obter_primeiro_valor(
            row,
            [
                "valorTotalHomologado",
                "valorTotalEstimado",
                "valorEstimado",
                "valorTotal",
            ],
            padrao=None,
        )

        valor_num = valor_numerico(valor)

        return {
            "id_pncp": texto_valido(id_pncp),
            "numero": texto_valido(
                obter_primeiro_valor(
                    row,
                    [
                        "numeroCompra",
                        "numeroEdital",
                        "numero",
                    ],
                )
            ),
            "processo": texto_valido(processo),
            "objeto": texto_valido(objeto),
            "fornecedor": texto_valido(
                obter_primeiro_valor(
                    row,
                    [
                        "nomeRazaoSocialFornecedor",
                        "razaoSocialFornecedor",
                        "nomeFornecedor",
                    ],
                )
            ),
            "cnpj_fornecedor": texto_valido(
                obter_primeiro_valor(
                    row,
                    [
                        "niFornecedor",
                        "cnpjFornecedor",
                    ],
                )
            ),
            "valor": (
                formatar_moeda_br(valor_num)
                if valor_num is not None
                else "N/D"
            ),
            "data_assinatura": formatar_data(
                obter_primeiro_valor(
                    row,
                    [
                        "dataPublicacao",
                        "dataPublicacaoPncp",
                        "dataAberturaProposta",
                    ],
                )
            ),
            "vigencia_inicio": "N/D",
            "vigencia_fim": "N/D",
            "situacao": texto_valido(
                obter_primeiro_valor(
                    row,
                    [
                        "situacaoCompra",
                        "situacao",
                        "status",
                    ],
                )
            ),
            "orgao": texto_valido(
                obter_primeiro_valor(
                    row,
                    [
                        "orgaoEntidade",
                        "nomeOrgao",
                        "orgao",
                    ],
                )
            ),
            "unidade": texto_valido(
                obter_primeiro_valor(
                    row,
                    [
                        "unidadeOrgao",
                        "nomeUnidade",
                        "unidade",
                    ],
                )
            ),
        }


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
# CONSULTA PNCP
# ============================================================

def consultar_pncp(url, params, max_tentativas=5):

    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 Chrome/151.0 Safari/537.36"
        ),
        "Accept": "application/json",
        "Connection": "keep-alive"
    }

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
                        "Resposta da API não é um JSON válido."
                    )

            if resp.status_code == 204:
                return []

            if resp.status_code in [
                429,
                500,
                502,
                503,
                504
            ]:

                if tentativa < max_tentativas:

                    espera = 2 ** tentativa

                    time.sleep(espera)

                    continue

            raise Exception(
                f"API retornou HTTP {resp.status_code}: "
                f"{resp.text[:300]}"
            )

        except (
            requests.exceptions.Timeout,
            requests.exceptions.ConnectionError
        ) as e:

            if tentativa < max_tentativas:

                espera = 2 ** tentativa

                time.sleep(espera)

                continue

            raise e

    raise Exception(
        "Falha de conexão com o PNCP após várias tentativas."
    )


# ============================================================
# DETALHES DO CONTRATO
# ============================================================

def consultar_detalhes_contrato(id_contrato):

    url = f"{BASE_URL}/contratos/{id_contrato}"

    try:

        resp = requests.get(
            url,
            timeout=15
        )

        if resp.status_code == 200:

            contrato = resp.json()

            aditivos = contrato.get(
                "termosAditivos",
                []
            )

            apostilamentos = contrato.get(
                "termosApostilamentos",
                []
            )

            lista_final = []

            for item in aditivos:

                item["tipoDocumentoNome"] = (
                    "Termo Aditivo"
                )

                lista_final.append(item)

            for item in apostilamentos:

                item["tipoDocumentoNome"] = (
                    "Termo de Apostilamento"
                )

                lista_final.append(item)

            return lista_final

        return []

    except Exception:
        return []


# ============================================================
# EXTRAÇÃO
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

            if (
                chave in data
                and isinstance(data[chave], list)
            ):
                return data[chave]

    return []


# ============================================================
# TRATAMENTO DATAFRAME
# ============================================================

def tratar_dataframe(df):

    if df.empty:
        return df

    df_tratado = df.copy()

    for col in df_tratado.columns:

        sample = df_tratado[col].dropna()

        if (
            not sample.empty
            and isinstance(sample.iloc[0], dict)
        ):

            df_tratado[col] = df_tratado[col].apply(
                lambda x:
                    x.get("nome")
                    or x.get("razaoSocial")
                    or x.get("descricao")
                    or str(x)
                    if isinstance(x, dict)
                    else str(x)
            )

    return df_tratado


# ============================================================
# PAGINAÇÃO
# ============================================================

def consultar_paginas_rapido(
    url,
    params,
    max_paginas=20
):

    params_primeira = params.copy()

    params_primeira["pagina"] = 1

    data_inicial = consultar_pncp(
        url,
        params_primeira
    )

    registros = extrair_registros(
        data_inicial
    )

    if not registros:
        return []

    todos_registros = list(registros)

    tamanho = params.get(
        "tamanhoPagina",
        50
    )

    if len(registros) < tamanho:
        return todos_registros

    def buscar_pagina(p):

        params_p = params.copy()

        params_p["pagina"] = p

        try:

            dados = consultar_pncp(
                url,
                params_p
            )

            return extrair_registros(
                dados
            )

        except Exception:
            return []

    with ThreadPoolExecutor(
        max_workers=5
    ) as executor:

        resultados = list(
            executor.map(
                buscar_pagina,
                range(
                    2,
                    max_paginas + 1
                )
            )
        )

    for res in resultados:

        if res:

            todos_registros.extend(res)

            if len(res) < tamanho:
                break

    return todos_registros


# ============================================================
# CONSULTA
# ============================================================

if st.sidebar.button(
    "🔎 Gerar Consulta",
    type="primary"
):

    endpoints = {

        "Contratos":
            f"{BASE_URL}/contratos",

        "Atas de Registro de Preços":
            f"{BASE_URL}/atas",

        "Editais e Avisos de Contratações":
            f"{BASE_URL}/contratacoes/publicacao"
    }

    endpoint = endpoints[tipo_consulta]

    tamanho_pagina = (
        50
        if tipo_consulta
        == "Editais e Avisos de Contratações"
        else 100
    )

    params = {

        "dataInicial":
            data_inicio.strftime("%Y%m%d"),

        "dataFinal":
            data_fim.strftime("%Y%m%d"),

        "pagina": 1,

        "tamanhoPagina":
            tamanho_pagina
    }

    if tipo_consulta == "Editais e Avisos de Contratações":

        params.update({

            "codigoModalidadeContratacao":
                modalidade_codigo,

            "uf":
                UF,

            "codigoMunicipioIbge":
                CODIGO_IBGE_RIO_DAS_PEDRAS,

            "cnpj":
                CNPJ_RIO_DAS_PEDRAS
        })

    elif tipo_consulta == "Contratos":

        params["cnpjOrgao"] = (
            CNPJ_RIO_DAS_PEDRAS
        )

    elif tipo_consulta == "Atas de Registro de Preços":

        params["cnpj"] = (
            CNPJ_RIO_DAS_PEDRAS
        )

    try:

        with st.spinner(
            "🔄 Buscando dados no PNCP..."
        ):

            registros = consultar_paginas_rapido(
                endpoint,
                params
            )

            df_temp = pd.DataFrame(
                registros
            )

            df_temp = tratar_dataframe(
                df_temp
            )

            # ------------------------------------------------
            # FILTRO EXTRA PARA CONTRATOS
            # ------------------------------------------------

            if (
                tipo_consulta == "Contratos"
                and not df_temp.empty
            ):

                possiveis_colunas = [
                    "cnpjOrgao",
                    "orgaoEntidade",
                    "orgao",
                    "unidadeOrgao"
                ]

                for coluna in possiveis_colunas:

                    if coluna in df_temp.columns:

                        serie = (
                            df_temp[coluna]
                            .astype(str)
                            .str.replace(
                                r"\D",
                                "",
                                regex=True
                            )
                        )

                        mask = serie.str.contains(
                            CNPJ_RIO_DAS_PEDRAS,
                            na=False
                        )

                        if mask.any():

                            df_temp = (
                                df_temp[mask]
                            )

                            break

            st.session_state.df_resultado = (
                df_temp
            )

    except Exception as e:

        st.session_state.df_resultado = None

        st.error(
            f"❌ Erro: {str(e)}"
        )


# ============================================================
# EXIBIÇÃO
# ============================================================

if (
    st.session_state.df_resultado
    is not None
    and not st.session_state.df_resultado.empty
):

    df = st.session_state.df_resultado

    st.success(
        f"📊 Exibindo {len(df)} registros "
        "para Rio das Pedras/SP."
    )

    # --------------------------------------------------------
    # MÉTRICAS
    # --------------------------------------------------------

    col_m1, col_m2 = st.columns(2)

    col_m1.metric(
        "Total de Registros",
        len(df)
    )

    coluna_valor = next(
        (
            c
            for c in [
                "valorGlobal",
                "valorInicial",
                "valorTotal",
                "valorTotalHomologado",
                "valorTotalEstimado",
                "valorAta"
            ]
            if c in df.columns
        ),
        None
    )

    if coluna_valor:

        total_valor = pd.to_numeric(
            df[coluna_valor],
            errors="coerce"
        ).sum()

        col_m2.metric(
            "Valor Total Envolvido",
            formatar_moeda_br(
                total_valor
            )
        )

    else:

        col_m2.metric(
            "Status da Consulta",
            "Concluída com Sucesso"
        )

    st.markdown("---")


    # ========================================================
    # EXPORTAÇÃO
    # ========================================================

    st.markdown(
        "### 📥 Opções de Exportação"
    )

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

    csv_data = df.to_csv(
        index=False,
        encoding="utf-8-sig"
    )

    cols[1].download_button(
        "📄 CSV (.csv)",
        csv_data,
        f"{nome}_Rio_Das_Pedras.csv",
        mime="text/csv"
    )


    # ========================================================
    # WORD
    # ========================================================

    doc = Document()

    # Margens
    section = doc.sections[0]

    section.top_margin = Pt(40)
    section.bottom_margin = Pt(40)
    section.left_margin = Pt(45)
    section.right_margin = Pt(45)

    # Título
    p_titulo = doc.add_paragraph()

    p_titulo.alignment = 1

    r_titulo = p_titulo.add_run(
        f"RELATÓRIO DE {tipo_consulta.upper()}"
    )

    r_titulo.bold = True
    r_titulo.font.size = Pt(16)
    r_titulo.font.color.rgb = RGBColor(
        0,
        51,
        102
    )

    p_subtitulo = doc.add_paragraph()

    p_subtitulo.alignment = 1

    r = p_subtitulo.add_run(
        "Prefeitura Municipal de Rio das Pedras/SP"
    )

    r.bold = True
    r.font.size = Pt(11)

    # Informações gerais
    doc.add_paragraph(
        f"Período da consulta: "
        f"{data_inicio.strftime('%d/%m/%Y')} "
        f"a "
        f"{data_fim.strftime('%d/%m/%Y')}"
    )

    doc.add_paragraph(
        f"Total de registros encontrados: "
        f"{len(df)}"
    )

    doc.add_paragraph("")

    doc.add_heading(
        "Principais Registros",
        level=2
    )

    # --------------------------------------------------------
    # DADOS DOS REGISTROS
    # --------------------------------------------------------

    for pos, (_, row) in enumerate(
        df.head(50).iterrows(),
        start=1
    ):

        dados = obter_dados_registro(
            row,
            tipo_consulta
        )

        tabela = doc.add_table(
            rows=0,
            cols=2
        )

        tabela.style = "Table Grid"

        campos_word = [
            ("Registro", f"{pos}"),
            ("Número", dados["numero"]),
            ("Controle PNCP", dados["id_pncp"]),
            ("Processo", dados["processo"]),
            ("Objeto", dados["objeto"]),
            ("Fornecedor", dados["fornecedor"]),
            (
                "CNPJ/CPF Fornecedor",
                dados["cnpj_fornecedor"]
            ),
            ("Valor", dados["valor"]),
            (
                "Data de Assinatura",
                dados["data_assinatura"]
            ),
            (
                "Início da Vigência",
                dados["vigencia_inicio"]
            ),
            (
                "Fim da Vigência",
                dados["vigencia_fim"]
            ),
            ("Situação", dados["situacao"]),
            ("Órgão", dados["orgao"]),
            ("Unidade", dados["unidade"]),
        ]

        for nome_campo, valor_campo in campos_word:

            cells = tabela.add_row().cells

            cells[0].text = nome_campo
            cells[1].text = str(
                valor_campo
            )

            for run in cells[0].paragraphs[0].runs:

                run.bold = True
                run.font.size = Pt(9)

            for run in cells[1].paragraphs[0].runs:

                run.font.size = Pt(9)

        doc.add_paragraph("")

    # Rodapé
    section = doc.sections[0]

    footer = section.footer

    p_footer = footer.paragraphs[0]

    p_footer.alignment = 1

    p_footer.add_run(
        "Consulta realizada no Portal Nacional "
        "de Contratações Públicas – PNCP"
    ).font.size = Pt(8)

    buffer_docx = io.BytesIO()

    doc.save(buffer_docx)

    buffer_docx.seek(0)

    cols[2].download_button(
        "📝 Word (.docx)",
        buffer_docx.getvalue(),
        f"Relatorio_{nome}.docx",
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

    # --------------------------------------------------------
    # TÍTULO
    # --------------------------------------------------------

    pdf.set_font(
        "Arial",
        "B",
        15
    )

    titulo_pdf = (
        f"RELATORIO DE {tipo_consulta.upper()}"
    )

    titulo_pdf = (
        titulo_pdf
        .encode("latin-1", "replace")
        .decode("latin-1")
    )

    pdf.cell(
        0,
        10,
        txt=titulo_pdf,
        ln=True,
        align="C"
    )

    pdf.set_font(
        "Arial",
        "B",
        10
    )

    pdf.cell(
        0,
        7,
        txt=(
            "Prefeitura Municipal de "
            "Rio das Pedras / SP"
        ),
        ln=True,
        align="C"
    )

    pdf.ln(5)

    pdf.set_font(
        "Arial",
        size=9
    )

    periodo_pdf = (
        f"Periodo: "
        f"{data_inicio.strftime('%d/%m/%Y')} "
        f"a "
        f"{data_fim.strftime('%d/%m/%Y')}"
    )

    pdf.cell(
        0,
        6,
        txt=periodo_pdf,
        ln=True
    )

    pdf.cell(
        0,
        6,
        txt=f"Total de registros: {len(df)}",
        ln=True
    )

    pdf.ln(5)

    # --------------------------------------------------------
    # REGISTROS
    # --------------------------------------------------------

    for pos, (_, row) in enumerate(
        df.head(50).iterrows(),
        start=1
    ):

        dados = obter_dados_registro(
            row,
            tipo_consulta
        )

        pdf.set_font(
            "Arial",
            "B",
            10
        )

        pdf.cell(
            0,
            7,
            txt=f"Registro {pos}",
            ln=True
        )

        pdf.set_font(
            "Arial",
            size=8
        )

        campos_pdf = [
            (
                "Numero",
                dados["numero"]
            ),
            (
                "Controle PNCP",
                dados["id_pncp"]
            ),
            (
                "Processo",
                dados["processo"]
            ),
            (
                "Objeto",
                dados["objeto"]
            ),
            (
                "Fornecedor",
                dados["fornecedor"]
            ),
            (
                "CNPJ/CPF Fornecedor",
                dados["cnpj_fornecedor"]
            ),
            (
                "Valor",
                dados["valor"]
            ),
            (
                "Data de Assinatura",
                dados["data_assinatura"]
            ),
            (
                "Inicio da Vigencia",
                dados["vigencia_inicio"]
            ),
            (
                "Fim da Vigencia",
                dados["vigencia_fim"]
            ),
            (
                "Situacao",
                dados["situacao"]
            ),
            (
                "Orgao",
                dados["orgao"]
            ),
            (
                "Unidade",
                dados["unidade"]
            ),
        ]

        for nome_campo, valor_campo in campos_pdf:

            texto = (
                f"{nome_campo}: "
                f"{texto_valido(valor_campo)}"
            )

            texto = (
                texto
                .encode("latin-1", "replace")
                .decode("latin-1")
            )

            pdf.multi_cell(
                0,
                5,
                txt=texto
            )

        pdf.ln(4)

    # --------------------------------------------------------
    # RODAPÉ
    # --------------------------------------------------------

    pdf.set_font(
        "Arial",
        "I",
        7
    )

    rodape = (
        "Consulta realizada no Portal Nacional "
        "de Contratacoes Publicas - PNCP"
    )

    pdf.cell(
        0,
        5,
        txt=rodape,
        ln=True,
        align="C"
    )

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

    cols[3].download_button(
        "📕 PDF (.pdf)",
        pdf_bytes,
        f"Relatorio_{nome}.pdf",
        mime="application/pdf"
    )


    st.markdown("---")


    # ========================================================
    # CONSULTA DE ADITIVOS
    # ========================================================

    if tipo_consulta == "Contratos":

        st.markdown(
            "### 🔍 Consultar Aditivos / "
            "Documentos por Contrato"
        )

        lista_contratos = df.apply(
            lambda x:
                f"{x.get('numeroControlePNCP')} "
                f"- Proc: {x.get('processo')}",
            axis=1
        ).tolist()

        contrato_selecionado = st.selectbox(
            "Selecione um contrato:",
            lista_contratos
        )

        numero_aditivo = st.text_input(
            "Digite o número do aditivo (opcional):",
            placeholder="Ex: 01/2026"
        )

        if st.button(
            "Buscar Aditivos do Contrato"
        ):

            id_escolhido = (
                contrato_selecionado
                .split(" - ")[0]
            )

            with st.spinner(
                "Buscando aditivos no PNCP..."
            ):

                aditivos = (
                    consultar_detalhes_contrato(
                        id_escolhido
                    )
                )

                if aditivos:

                    if numero_aditivo:

                        aditivos = [
                            d
                            for d in aditivos
                            if numero_aditivo
                            in str(
                                d.get(
                                    "numero",
                                    ""
                                )
                            )
                        ]

                    if aditivos:

                        st.success(
                            f"Encontrados "
                            f"{len(aditivos)} "
                            "documentos vinculados:"
                        )

                        for doc in aditivos:

                            tipo_doc = doc.get(
                                "tipoDocumentoNome",
                                "Outro"
                            )

                            num_doc = doc.get(
                                "numero",
                                "S/N"
                            )

                            st.info(
                                f"**Tipo:** "
                                f"{tipo_doc} | "
                                f"**Nº:** "
                                f"{num_doc} | "
                                f"**Data:** "
                                f"{doc.get('dataPublicacao', 'N/D')}"
                                f"\n\n"
                                f"{doc.get('objeto', '')}"
                            )

                    else:

                        st.warning(
                            "Nenhum aditivo encontrado "
                            "com este número para este contrato."
                        )

                else:

                    st.warning(
                        "Nenhum documento/aditivo "
                        "encontrado para este contrato."
                    )

        st.markdown("---")


    # ========================================================
    # GRÁFICOS
    # ========================================================

    st.markdown(
        "### 📈 Análise Gráfica"
    )

    coluna_data = next(
        (
            c
            for c in [
                "dataPublicacao",
                "dataAssinatura",
                "dataInclusao",
                "dataPublicacaoPncp"
            ]
            if c in df.columns
        ),
        None
    )

    coluna_valor_grafico = next(
        (
            c
            for c in [
                "valorGlobal",
                "valorInicial",
                "valorTotal",
                "valorTotalHomologado",
                "valorTotalEstimado",
                "valorAta"
            ]
            if c in df.columns
        ),
        None
    )

    if coluna_data:

        try:

            df_grafico = df.copy()

            df_grafico["mes_ano"] = (
                pd.to_datetime(
                    df_grafico[coluna_data],
                    errors="coerce"
                )
                .dt
                .to_period("M")
                .astype(str)
            )

            aba1, aba2 = st.tabs(
                [
                    "🔢 Quantidade de Registros",
                    "💰 Volume Financeiro (R$)"
                ]
            )

            with aba1:

                st.markdown(
                    f"#### Quantidade de "
                    f"{tipo_consulta} por Mês/Ano"
                )

                st.bar_chart(
                    df_grafico[
                        "mes_ano"
                    ]
                    .value_counts()
                    .sort_index()
                )

            with aba2:

                if coluna_valor_grafico:

                    st.markdown(
                        f"#### Volume Financeiro de "
                        f"{tipo_consulta} por Mês/Ano"
                    )

                    df_grafico[
                        coluna_valor_grafico
                    ] = pd.to_numeric(
                        df_grafico[
                            coluna_valor_grafico
                        ],
                        errors="coerce"
                    ).fillna(0)

                    st.bar_chart(
                        df_grafico
                        .groupby("mes_ano")[
                            coluna_valor_grafico
                        ]
                        .sum()
                        .sort_index()
                    )

                else:

                    st.info(
                        "ℹ️ Dados financeiros "
                        "indisponíveis."
                    )

        except Exception as e:

            st.info(
                f"ℹ️ Erro ao gerar gráficos: {e}"
            )

    else:

        st.info(
            "ℹ️ Coluna de data não encontrada "
            "para gerar o gráfico."
        )


    # ========================================================
    # TABELA
    # ========================================================

    st.markdown(
        "### 📋 Tabela de Dados Detalhada"
    )

    st.dataframe(
        df,
        use_container_width=True,
        hide_index=True
    )


# ============================================================
# RESULTADO VAZIO
# ============================================================

elif (
    st.session_state.df_resultado
    is not None
    and st.session_state.df_resultado.empty
):

    st.warning(
        "⚠️ Nenhum registro encontrado para "
        "Rio das Pedras/SP no período selecionado."
    )

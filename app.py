import io
import datetime
import time
from concurrent.futures import ThreadPoolExecutor

import pandas as pd
import requests
import streamlit as st

from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT

from fpdf import FPDF


# ============================================================
# FUNÇÃO DE FORMATAÇÃO BR
# ============================================================

def formatar_moeda_br(valor):
    """Formata número para o padrão brasileiro: R$ 15.300.296,03"""
    try:
        valor = float(valor)
    except (TypeError, ValueError):
        valor = 0

    valor_fmt = "{:,.2f}".format(valor)

    return (
        f"R$ "
        f"{valor_fmt.replace(',', 'X').replace('.', ',').replace('X', '.')}"
    )


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
# FUNÇÕES DE CONSULTA
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

            if resp.status_code in [429, 500, 502, 503, 504]:

                if tentativa < max_tentativas:

                    espera = 2 ** tentativa

                    st.warning(
                        f"⚠️ Servidor instável "
                        f"(Erro {resp.status_code}). "
                        f"Nova tentativa em {espera}s "
                        f"({tentativa}/{max_tentativas})."
                    )

                    time.sleep(espera)

                    continue

            raise Exception(
                f"API retornou HTTP {resp.status_code}: "
                f"{resp.text[:200]}"
            )

        except (
            requests.exceptions.Timeout,
            requests.exceptions.ConnectionError
        ) as e:

            if tentativa < max_tentativas:

                espera = 2 ** tentativa

                st.warning(
                    f"🌐 Problema de rede. "
                    f"Tentativa {tentativa}/{max_tentativas} "
                    f"em {espera}s..."
                )

                time.sleep(espera)

                continue

            raise e

    raise Exception(
        "Falha de conexão com o PNCP após várias tentativas."
    )


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

                item["tipoDocumentoNome"] = "Termo Aditivo"

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

            return extrair_registros(dados)

        except Exception:

            return []

    with ThreadPoolExecutor(
        max_workers=5
    ) as executor:

        resultados = list(
            executor.map(
                buscar_pagina,
                range(2, max_paginas + 1)
            )
        )

    for res in resultados:

        if res:

            todos_registros.extend(res)

            if len(res) < tamanho:
                break

    return todos_registros


def obter_dados_registro(row, tipo):

    id_pncp = row.get(
        "numeroControlePNCP",
        row.get(
            "numeroControlePNCPAta",
            row.get(
                "numeroControlePNCPCompra",
                "N/D"
            )
        )
    )

    if tipo == "Atas de Registro de Preços":

        processo = row.get(
            "numeroAtaRegistroPreco",
            "N/D"
        )

        info_extra = (
            f"Vigência: "
            f"Início: {row.get('vigenciaInicio', 'N/D')} | "
            f"Fim: {row.get('vigenciaFim', 'N/D')}"
        )

    elif tipo == "Contratos":

        processo = row.get(
            "processo",
            "N/D"
        )

        valor = row.get(
            "valorGlobal",
            row.get(
                "valorInicial",
                0
            )
        )

        try:
            valor_fmt = formatar_moeda_br(valor)

        except Exception:
            valor_fmt = str(valor)

        info_extra = (
            f"Fornecedor: "
            f"{row.get('nomeRazaoSocialFornecedor', 'N/D')} | "
            f"Valor: {valor_fmt}"
        )

    else:

        processo = row.get(
            "processo",
            "N/D"
        )

        valor = row.get(
            "valorTotalHomologado",
            row.get(
                "valorTotalEstimado",
                0
            )
        )

        try:
            valor_fmt = formatar_moeda_br(valor)

        except Exception:
            valor_fmt = str(valor)

        info_extra = (
            f"Responsável: "
            f"{row.get('usuarioNome', 'N/D')} | "
            f"Valor: {valor_fmt}"
        )

    return (
        str(id_pncp),
        str(processo),
        info_extra,
        str(
            row.get(
                "objetoContrato"
                if tipo == "Contratos"
                else "objetoCompra",
                "N/D"
            )
        )
    )


# ============================================================
# PREPARAR DADOS PRINCIPAIS PARA WORD E PDF
# ============================================================

def preparar_dados_relatorio(df, tipo):

    dados = []

    for _, row in df.iterrows():

        id_pncp, processo, info_extra, objeto = (
            obter_dados_registro(
                row,
                tipo
            )
        )

        if tipo == "Contratos":

            fornecedor = row.get(
                "nomeRazaoSocialFornecedor",
                "N/D"
            )

            valor = row.get(
                "valorGlobal",
                row.get(
                    "valorInicial",
                    0
                )
            )

            dados.append({
                "ID PNCP": id_pncp,
                "Processo": processo,
                "Fornecedor": str(fornecedor),
                "Valor": formatar_moeda_br(valor),
                "Objeto": objeto
            })

        elif tipo == "Atas de Registro de Preços":

            dados.append({
                "ID PNCP": id_pncp,
                "Ata": processo,
                "Vigência": info_extra.replace(
                    "Vigência: ",
                    ""
                ),
                "Objeto": objeto
            })

        else:

            modalidade = row.get(
                "modalidadeNome",
                "N/D"
            )

            valor = row.get(
                "valorTotalHomologado",
                row.get(
                    "valorTotalEstimado",
                    0
                )
            )

            dados.append({
                "ID PNCP": id_pncp,
                "Processo": processo,
                "Modalidade": str(modalidade),
                "Valor": formatar_moeda_br(valor),
                "Objeto": objeto
            })

    return dados


# ============================================================
# GERAR WORD
# ============================================================

def gerar_word(df, tipo, data_inicio, data_fim):

    doc = Document()

    # Margens
    section = doc.sections[0]

    section.top_margin = Inches(0.6)
    section.bottom_margin = Inches(0.6)
    section.left_margin = Inches(0.6)
    section.right_margin = Inches(0.6)

    # --------------------------------------------------------
    # TÍTULO
    # --------------------------------------------------------

    p = doc.add_paragraph()

    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    run = p.add_run(
        "RELATÓRIO EXECUTIVO\n"
        "CONTRATAÇÕES PÚBLICAS"
    )

    run.bold = True
    run.font.size = Pt(18)
    run.font.color.rgb = RGBColor(
        0,
        51,
        102
    )

    p2 = doc.add_paragraph()

    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER

    run = p2.add_run(
        "Prefeitura Municipal de Rio das Pedras / SP"
    )

    run.bold = True
    run.font.size = Pt(12)

    # --------------------------------------------------------
    # INFORMAÇÕES DA CONSULTA
    # --------------------------------------------------------

    doc.add_paragraph()

    tabela_info = doc.add_table(
        rows=4,
        cols=2
    )

    tabela_info.alignment = WD_TABLE_ALIGNMENT.CENTER

    informacoes = [
        ("Tipo de Consulta", tipo),
        (
            "Período",
            f"{data_inicio.strftime('%d/%m/%Y')} "
            f"a {data_fim.strftime('%d/%m/%Y')}"
        ),
        (
            "Total de Registros",
            str(len(df))
        ),
        (
            "Data do Relatório",
            datetime.date.today().strftime(
                "%d/%m/%Y"
            )
        )
    ]

    for i, (campo, valor) in enumerate(
        informacoes
    ):

        tabela_info.cell(
            i,
            0
        ).text = campo

        tabela_info.cell(
            i,
            1
        ).text = valor

        tabela_info.cell(
            i,
            0
        ).vertical_alignment = (
            WD_CELL_VERTICAL_ALIGNMENT.CENTER
        )

        tabela_info.cell(
            i,
            1
        ).vertical_alignment = (
            WD_CELL_VERTICAL_ALIGNMENT.CENTER
        )

        for run in tabela_info.cell(
            i,
            0
        ).paragraphs[0].runs:

            run.bold = True

    doc.add_paragraph()

    # --------------------------------------------------------
    # VALOR TOTAL
    # --------------------------------------------------------

    colunas_valor = [
        "valorGlobal",
        "valorInicial",
        "valorTotalHomologado",
        "valorTotalEstimado"
    ]

    coluna_valor = next(
        (
            c for c in colunas_valor
            if c in df.columns
        ),
        None
    )

    if coluna_valor:

        total_valor = pd.to_numeric(
            df[coluna_valor],
            errors="coerce"
        ).fillna(0).sum()

        p_valor = doc.add_paragraph()

        p_valor.alignment = WD_ALIGN_PARAGRAPH.CENTER

        run = p_valor.add_run(
            f"VALOR TOTAL ENVOLVIDO: "
            f"{formatar_moeda_br(total_valor)}"
        )

        run.bold = True
        run.font.size = Pt(13)
        run.font.color.rgb = RGBColor(
            0,
            102,
            51
        )

    # --------------------------------------------------------
    # DADOS PRINCIPAIS
    # --------------------------------------------------------

    doc.add_heading(
        "Principais Registros",
        level=2
    )

    dados = preparar_dados_relatorio(
        df,
        tipo
    )

    # Limita o relatório executivo a 50 registros
    dados = dados[:50]

    if dados:

        colunas = list(
            dados[0].keys()
        )

        tabela = doc.add_table(
            rows=1,
            cols=len(colunas)
        )

        tabela.alignment = (
            WD_TABLE_ALIGNMENT.CENTER
        )

        tabela.style = "Table Grid"

        # Cabeçalho
        for i, coluna in enumerate(
            colunas
        ):

            cell = tabela.rows[0].cells[i]

            cell.text = coluna

            for run in cell.paragraphs[0].runs:

                run.bold = True
                run.font.size = Pt(9)
                run.font.color.rgb = (
                    RGBColor(
                        255,
                        255,
                        255
                    )
                )

        # Dados
        for item in dados:

            row_cells = tabela.add_row().cells

            for i, coluna in enumerate(
                colunas
            ):

                valor = str(
                    item.get(
                        coluna,
                        ""
                    )
                )

                # Evita objetos gigantes no relatório
                if len(valor) > 350:

                    valor = (
                        valor[:347]
                        + "..."
                    )

                row_cells[i].text = valor

                for run in row_cells[i].paragraphs[0].runs:

                    run.font.size = Pt(8)

    # --------------------------------------------------------
    # OBSERVAÇÃO
    # --------------------------------------------------------

    doc.add_paragraph()

    p_obs = doc.add_paragraph()

    run = p_obs.add_run(
        "Observação: "
        "Este relatório apresenta os principais "
        "dados retornados pelo Portal Nacional "
        "de Contratações Públicas (PNCP). "
        "Para análise completa, recomenda-se "
        "consultar também a planilha Excel ou "
        "o arquivo CSV disponibilizado."
    )

    run.font.size = Pt(9)
    run.italic = True

    # --------------------------------------------------------
    # RODAPÉ
    # --------------------------------------------------------

    footer = section.footer

    p_footer = footer.paragraphs[0]

    p_footer.alignment = (
        WD_ALIGN_PARAGRAPH.CENTER
    )

    run = p_footer.add_run(
        "Portal PNCP - Rio das Pedras/SP"
    )

    run.font.size = Pt(8)
    run.font.color.rgb = RGBColor(
        100,
        100,
        100
    )

    # Salvar
    buffer = io.BytesIO()

    doc.save(buffer)

    buffer.seek(0)

    return buffer.getvalue()


# ============================================================
# GERAR PDF
# ============================================================

class RelatorioPDF(FPDF):

    def header(self):

        self.set_font(
            "Arial",
            "B",
            14
        )

        self.cell(
            0,
            8,
            "RELATORIO EXECUTIVO",
            ln=True,
            align="C"
        )

        self.set_font(
            "Arial",
            "B",
            11
        )

        self.cell(
            0,
            6,
            "CONTRATACOES PUBLICAS",
            ln=True,
            align="C"
        )

        self.set_font(
            "Arial",
            "",
            9
        )

        self.cell(
            0,
            5,
            "Prefeitura Municipal de Rio das Pedras / SP",
            ln=True,
            align="C"
        )

        self.ln(4)

        self.line(
            10,
            self.get_y(),
            200,
            self.get_y()
        )

        self.ln(5)

    def footer(self):

        self.set_y(-15)

        self.set_font(
            "Arial",
            "",
            8
        )

        self.cell(
            0,
            10,
            f"Pagina {self.page_no()}",
            align="C"
        )


def limpar_texto_pdf(texto):

    if texto is None:
        return ""

    texto = str(texto)

    return (
        texto
        .encode(
            "latin-1",
            "replace"
        )
        .decode("latin-1")
    )


def gerar_pdf(
    df,
    tipo,
    data_inicio,
    data_fim
):

    pdf = RelatorioPDF(
        orientation="L",
        unit="mm",
        format="A4"
    )

    pdf.set_auto_page_break(
        auto=True,
        margin=15
    )

    pdf.add_page()

    # --------------------------------------------------------
    # RESUMO
    # --------------------------------------------------------

    pdf.set_font(
        "Arial",
        "B",
        10
    )

    pdf.cell(
        40,
        7,
        "Tipo de Consulta:"
    )

    pdf.set_font(
        "Arial",
        "",
        10
    )

    pdf.cell(
        80,
        7,
        limpar_texto_pdf(tipo)
    )

    pdf.set_font(
        "Arial",
        "B",
        10
    )

    pdf.cell(
        30,
        7,
        "Periodo:"
    )

    pdf.set_font(
        "Arial",
        "",
        10
    )

    periodo = (
        f"{data_inicio.strftime('%d/%m/%Y')} "
        f"a {data_fim.strftime('%d/%m/%Y')}"
    )

    pdf.cell(
        70,
        7,
        periodo
    )

    pdf.ln(8)

    pdf.set_font(
        "Arial",
        "B",
        10
    )

    pdf.cell(
        40,
        7,
        "Total de Registros:"
    )

    pdf.set_font(
        "Arial",
        "",
        10
    )

    pdf.cell(
        30,
        7,
        str(len(df))
    )

    # Valor total
    colunas_valor = [
        "valorGlobal",
        "valorInicial",
        "valorTotalHomologado",
        "valorTotalEstimado"
    ]

    coluna_valor = next(
        (
            c for c in colunas_valor
            if c in df.columns
        ),
        None
    )

    if coluna_valor:

        total_valor = pd.to_numeric(
            df[coluna_valor],
            errors="coerce"
        ).fillna(0).sum()

        pdf.set_font(
            "Arial",
            "B",
            10
        )

        pdf.cell(
            45,
            7,
            "Valor Total:"
        )

        pdf.set_font(
            "Arial",
            "",
            10
        )

        pdf.cell(
            60,
            7,
            limpar_texto_pdf(
                formatar_moeda_br(
                    total_valor
                )
            )
        )

    pdf.ln(10)

    # --------------------------------------------------------
    # TABELA PRINCIPAL
    # --------------------------------------------------------

    pdf.set_font(
        "Arial",
        "B",
        11
    )

    pdf.cell(
        0,
        7,
        "PRINCIPAIS REGISTROS",
        ln=True
    )

    pdf.ln(2)

    dados = preparar_dados_relatorio(
        df,
        tipo
    )

    dados = dados[:50]

    if not dados:

        pdf.set_font(
            "Arial",
            "",
            10
        )

        pdf.cell(
            0,
            8,
            "Nenhum registro disponivel.",
            ln=True
        )

    else:

        # ----------------------------------------------------
        # Configuração das colunas
        # ----------------------------------------------------

        if tipo == "Contratos":

            colunas = [
                ("ID PNCP", 38),
                ("Processo", 25),
                ("Fornecedor", 48),
                ("Valor", 30),
                ("Objeto", 106)
            ]

        elif tipo == "Atas de Registro de Preços":

            colunas = [
                ("ID PNCP", 40),
                ("Ata", 35),
                ("Vigencia", 45),
                ("Objeto", 127)
            ]

        else:

            colunas = [
                ("ID PNCP", 38),
                ("Processo", 25),
                ("Modalidade", 40),
                ("Valor", 30),
                ("Objeto", 114)
            ]

        # Cabeçalho
        pdf.set_font(
            "Arial",
            "B",
            8
        )

        for titulo, largura in colunas:

            pdf.cell(
                largura,
                8,
                limpar_texto_pdf(titulo),
                border=1,
                align="C"
            )

        pdf.ln()

        # Linhas
        pdf.set_font(
            "Arial",
            "",
            7
        )

        for item in dados:

            altura = 14

            for coluna, largura in colunas:

                valor = str(
                    item.get(
                        coluna,
                        ""
                    )
                )

                if len(valor) > 150:

                    valor = (
                        valor[:147]
                        + "..."
                    )

                valor = limpar_texto_pdf(
                    valor
                )

                pdf.cell(
                    largura,
                    altura,
                    valor,
                    border=1,
                    align="L"
                )

            pdf.ln()

    # --------------------------------------------------------
    # OBSERVAÇÃO
    # --------------------------------------------------------

    pdf.ln(6)

    pdf.set_font(
        "Arial",
        "I",
        8
    )

    observacao = (
        "Observacao: este relatorio apresenta "
        "os principais dados retornados pelo "
        "Portal Nacional de Contratacoes Publicas "
        "(PNCP). Para analise completa, consulte "
        "tambem os arquivos Excel e CSV."
    )

    pdf.multi_cell(
        0,
        4,
        limpar_texto_pdf(
            observacao
        )
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

    return pdf_bytes


# ============================================================
# INTERFACE
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

    endpoint = endpoints[
        tipo_consulta
    ]

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

        "pagina":
            1,

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

        params[
            "cnpjOrgao"
        ] = CNPJ_RIO_DAS_PEDRAS

    elif tipo_consulta == "Atas de Registro de Preços":

        params[
            "cnpj"
        ] = CNPJ_RIO_DAS_PEDRAS

    try:

        with st.spinner(
            "🔄 Buscando e tratando dados em paralelo no PNCP..."
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

                            df_temp = df_temp[
                                mask
                            ]

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
# EXIBIÇÃO DOS RESULTADOS
# ============================================================

if (
    st.session_state.df_resultado is not None
    and not st.session_state.df_resultado.empty
):

    df = st.session_state.df_resultado

    st.success(
        f"📊 Exibindo {len(df)} registros "
        f"para Rio das Pedras/SP."
    )

    col_m1, col_m2 = st.columns(2)

    col_m1.metric(
        "Total de Registros",
        len(df)
    )

    coluna_valor = next(
        (
            c for c in [
                "valorGlobal",
                "valorInicial",
                "valorTotalHomologado",
                "valorTotalEstimado"
            ]
            if c in df.columns
        ),
        None
    )

    if coluna_valor:

        total_valor = pd.to_numeric(
            df[coluna_valor],
            errors="coerce"
        ).fillna(0).sum()

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
    # OPÇÕES DE EXPORTAÇÃO
    # ========================================================

    st.markdown(
        "### 📥 Opções de Exportação"
    )

    st.caption(
        "Excel e CSV contêm a base completa. "
        "Word e PDF apresentam um relatório executivo "
        "com os principais dados."
    )

    cols = st.columns(4)

    nome = (
        tipo_consulta
        .replace(" ", "_")
        .replace("/", "_")
    )


    # ========================================================
    # EXCEL
    # ========================================================

    buffer_xlsx = io.BytesIO()

    with pd.ExcelWriter(
        buffer_xlsx,
        engine="openpyxl"
    ) as writer:

        df.to_excel(
            writer,
            index=False,
            sheet_name="Dados"
        )

        worksheet = writer.sheets[
            "Dados"
        ]

        # Ajuste automático das larguras
        for coluna in worksheet.columns:

            tamanho = 0

            letra = coluna[0].column_letter

            for celula in coluna:

                try:

                    tamanho = max(
                        tamanho,
                        len(str(celula.value))
                    )

                except Exception:
                    pass

            worksheet.column_dimensions[
                letra
            ].width = min(
                tamanho + 2,
                60
            )

    buffer_xlsx.seek(0)

    cols[0].download_button(
        "📊 Excel (.xlsx)",
        buffer_xlsx.getvalue(),
        f"{nome}_Rio_Das_Pedras.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        )
    )


    # ========================================================
    # CSV
    # ========================================================

    csv_bytes = df.to_csv(
        index=False,
        encoding="utf-8-sig"
    ).encode(
        "utf-8-sig"
    )

    cols[1].download_button(
        "📄 CSV (.csv)",
        csv_bytes,
        f"{nome}_Rio_Das_Pedras.csv",
        mime="text/csv"
    )


    # ========================================================
    # WORD
    # ========================================================

    word_bytes = gerar_word(
        df,
        tipo_consulta,
        data_inicio,
        data_fim
    )

    cols[2].download_button(
        "📝 Word (.docx)",
        word_bytes,
        f"Relatorio_{nome}.docx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "wordprocessingml.document"
        )
    )


    # ========================================================
    # PDF
    # ========================================================

    pdf_bytes = gerar_pdf(
        df,
        tipo_consulta,
        data_inicio,
        data_fim
    )

    cols[3].download_button(
        "📕 PDF (.pdf)",
        pdf_bytes,
        f"Relatorio_{nome}.pdf",
        mime="application/pdf"
    )


    # ========================================================
    # GRÁFICOS
    # ========================================================

    st.markdown("---")

    st.markdown(
        "### 📈 Análise Gráfica"
    )

    coluna_data = next(
        (
            c for c in [
                "dataPublicacao",
                "dataAssinatura",
                "dataInclusao"
            ]
            if c in df.columns
        ),
        None
    )

    coluna_valor = next(
        (
            c for c in [
                "valorGlobal",
                "valorInicial",
                "valorTotalHomologado",
                "valorTotalEstimado"
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

                if coluna_valor:

                    st.markdown(
                        f"#### Volume Financeiro de "
                        f"{tipo_consulta} por Mês/Ano"
                    )

                    df_grafico[
                        coluna_valor
                    ] = pd.to_numeric(
                        df_grafico[
                            coluna_valor
                        ],
                        errors="coerce"
                    ).fillna(0)

                    st.bar_chart(
                        df_grafico
                        .groupby("mes_ano")[
                            coluna_valor
                        ]
                        .sum()
                        .sort_index()
                    )

                else:

                    st.info(
                        "ℹ️ Dados financeiros indisponíveis."
                    )

        except Exception as e:

            st.info(
                f"ℹ️ Erro ao gerar gráficos: {e}"
            )


    # ========================================================
    # TABELA
    # ========================================================

    st.markdown("---")

    st.markdown(
        "### 📋 Tabela de Dados Detalhada"
    )

    st.dataframe(
        df,
        use_container_width=True,
        hide_index=True
    )


# ============================================================
# NENHUM REGISTRO
# ============================================================

elif (
    st.session_state.df_resultado is not None
    and st.session_state.df_resultado.empty
):

    st.warning(
        "⚠️ Nenhum registro encontrado "
        "para Rio das Pedras/SP no período selecionado."
    )

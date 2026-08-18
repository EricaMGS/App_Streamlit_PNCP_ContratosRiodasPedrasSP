import requests
import pandas as pd
import time
from pathlib import Path
from datetime import date, timedelta


# ============================================================
# CONFIGURAÇÕES
# ============================================================

CNPJ_ORGAO = "44826840000183"

# API de CONSULTA
URL_CONSULTA = (
    "https://pncp.gov.br/api/consulta/v1"
)

# API de TABELAS/DOMÍNIOS
URL_PNCP = (
    "https://pncp.gov.br/api/pncp/v1"
)

URL_CONTRATACOES = (
    f"{URL_CONSULTA}/contratacoes/publicacao"
)

URL_MODALIDADES = (
    f"{URL_PNCP}/modalidades"
)

DIRETORIO_DADOS = (
    Path(__file__).parent / "dados"
)

ARQUIVO_SAIDA = (
    DIRETORIO_DADOS / "compras.parquet"
)

ANOS = [2024, 2025, 2026]

TAMANHO_PAGINA = 500

TIMEOUT = 60

# Períodos de 180 dias para evitar consultas muito grandes
DIAS_POR_PERIODO = 180

# Pequena pausa entre requisições
PAUSA = 0.2


# ============================================================
# SESSÃO HTTP
# ============================================================

session = requests.Session()

session.headers.update({
    "Accept": "*/*",
    "User-Agent": "Painel-PNCP-Rio-das-Pedras/1.0"
})


# ============================================================
# BUSCAR MODALIDADES
# ============================================================

def buscar_modalidades():

    print()
    print("=" * 70)
    print("🔎 BUSCANDO MODALIDADES NO PNCP")
    print("=" * 70)

    try:

        resposta = session.get(
            URL_MODALIDADES,
            params={
                "statusAtivo": "true"
            },
            timeout=TIMEOUT
        )

        print(
            f"HTTP {resposta.status_code}"
        )

        if resposta.status_code != 200:

            print(
                "❌ Erro ao consultar modalidades."
            )

            print(
                resposta.text[:2000]
            )

            return []

        dados = resposta.json()

        # A API pode devolver lista diretamente
        # ou uma estrutura contendo "data".
        if isinstance(dados, list):

            modalidades = dados

        elif isinstance(dados, dict):

            modalidades = dados.get(
                "data",
                []
            )

        else:

            modalidades = []

        if not modalidades:

            print(
                "❌ Nenhuma modalidade retornada."
            )

            return []

        print(
            f"✅ {len(modalidades)} modalidades encontradas."
        )

        for modalidade in modalidades:

            codigo = modalidade.get("id")
            nome = modalidade.get("nome")

            print(
                f"   {codigo} - {nome}"
            )

        return modalidades

    except Exception as erro:

        print(
            f"❌ Erro ao buscar modalidades: {erro}"
        )

        return []


# ============================================================
# CRIAR PERÍODOS DE CONSULTA
# ============================================================

def gerar_periodos(ano):

    inicio = date(
        ano,
        1,
        1
    )

    # Para 2026, não consulta datas futuras.
    if ano == date.today().year:

        fim = date.today()

    else:

        fim = date(
            ano,
            12,
            31
        )

    periodos = []

    atual = inicio

    while atual <= fim:

        final_periodo = min(
            atual + timedelta(
                days=DIAS_POR_PERIODO - 1
            ),
            fim
        )

        periodos.append(
            (
                atual.strftime("%Y%m%d"),
                final_periodo.strftime("%Y%m%d")
            )
        )

        atual = (
            final_periodo
            + timedelta(days=1)
        )

    return periodos


# ============================================================
# CONSULTAR UMA MODALIDADE
# ============================================================

def consultar_modalidade(
    codigo_modalidade,
    nome_modalidade,
    ano,
    data_inicial,
    data_final
):

    registros_modalidade = []

    pagina = 1

    while True:

        params = {
            "dataInicial": data_inicial,
            "dataFinal": data_final,

            "codigoModalidadeContratacao": (
                codigo_modalidade
            ),

            # Filtro pelo CNPJ da Prefeitura/órgão
            "cnpj": CNPJ_ORGAO,

            "pagina": pagina,

            "tamanhoPagina": TAMANHO_PAGINA
        }

        try:

            resposta = session.get(
                URL_CONTRATACOES,
                params=params,
                timeout=TIMEOUT
            )

            print(
                f"      Página {pagina} | "
                f"HTTP {resposta.status_code}"
            )

            # ------------------------------------------------
            # ERROS HTTP
            # ------------------------------------------------

            if resposta.status_code != 200:

                print(
                    "      ❌ Erro retornado pelo PNCP:"
                )

                print(
                    resposta.text[:1000]
                )

                break

            # ------------------------------------------------
            # JSON
            # ------------------------------------------------

            try:

                dados = resposta.json()

            except ValueError:

                print(
                    "      ❌ Resposta não é JSON."
                )

                break

            # ------------------------------------------------
            # REGISTROS
            # ------------------------------------------------

            registros = dados.get(
                "data",
                []
            )

            quantidade = len(
                registros
            )

            print(
                f"      📄 Registros: {quantidade}"
            )

            if not registros:

                break

            registros_modalidade.extend(
                registros
            )

            # ------------------------------------------------
            # PAGINAÇÃO
            # ------------------------------------------------

            if quantidade < TAMANHO_PAGINA:

                break

            pagina += 1

            time.sleep(PAUSA)

        except requests.exceptions.Timeout:

            print(
                "      ❌ Timeout."
            )

            break

        except requests.exceptions.RequestException as erro:

            print(
                f"      ❌ Erro de conexão: {erro}"
            )

            break

        except Exception as erro:

            print(
                f"      ❌ Erro inesperado: {erro}"
            )

            break

    return registros_modalidade


# ============================================================
# EXTRAÇÃO PRINCIPAL
# ============================================================

def extrair_dados_pncp():

    todas_licitacoes = []

    modalidades = buscar_modalidades()

    if not modalidades:

        print()
        print(
            "❌ Não foi possível obter as modalidades."
        )

        return []

    print()
    print("=" * 70)
    print("🏛️ INICIANDO EXTRAÇÃO DO PNCP")
    print("=" * 70)

    print(
        f"CNPJ: {CNPJ_ORGAO}"
    )

    print(
        f"Anos: {ANOS}"
    )

    print(
        f"Tamanho da página: {TAMANHO_PAGINA}"
    )

    # ========================================================
    # ANOS
    # ========================================================

    for ano in ANOS:

        periodos = gerar_periodos(
            ano
        )

        print()
        print("=" * 70)
        print(
            f"📅 ANO {ano} "
            f"({len(periodos)} períodos)"
        )
        print("=" * 70)

        # ====================================================
        # PERÍODOS
        # ====================================================

        for numero_periodo, periodo in enumerate(
            periodos,
            start=1
        ):

            data_inicial, data_final = periodo

            print()
            print(
                f"📆 Período "
                f"{numero_periodo}/{len(periodos)}: "
                f"{data_inicial} → {data_final}"
            )

            # =================================================
            # MODALIDADES
            # =================================================

            for modalidade in modalidades:

                codigo = modalidade.get(
                    "id"
                )

                nome = modalidade.get(
                    "nome",
                    "Modalidade desconhecida"
                )

                if codigo is None:

                    continue

                print()
                print(
                    f"   🔎 {nome} "
                    f"(código {codigo})"
                )

                registros = consultar_modalidade(
                    codigo_modalidade=codigo,
                    nome_modalidade=nome,
                    ano=ano,
                    data_inicial=data_inicial,
                    data_final=data_final
                )

                if registros:

                    print(
                        f"      ✅ "
                        f"{len(registros)} registros adicionados."
                    )

                    todas_licitacoes.extend(
                        registros
                    )

                time.sleep(PAUSA)

    # ========================================================
    # RESULTADO
    # ========================================================

    print()
    print("=" * 70)
    print(
        f"📊 TOTAL BRUTO DE REGISTROS: "
        f"{len(todas_licitacoes)}"
    )
    print("=" * 70)

    return todas_licitacoes


# ============================================================
# SALVAR PARQUET
# ============================================================

def salvar_dados(lista_compras):

    DIRETORIO_DADOS.mkdir(
        parents=True,
        exist_ok=True
    )

    # --------------------------------------------------------
    # PROTEÇÃO CONTRA RESULTADO VAZIO
    # --------------------------------------------------------

    if not lista_compras:

        print()
        print(
            "⚠️ Nenhum registro foi encontrado."
        )

        print(
            "⚠️ O arquivo Parquet existente "
            "NÃO será sobrescrito."
        )

        return False

    # --------------------------------------------------------
    # DATAFRAME
    # --------------------------------------------------------

    df = pd.DataFrame(
        lista_compras
    )

    print()
    print(
        f"📊 DataFrame criado com "
        f"{len(df)} registros."
    )

    print(
        f"📋 Colunas: "
        f"{len(df.columns)}"
    )

    # --------------------------------------------------------
    # REMOVER DUPLICADOS
    # --------------------------------------------------------

    quantidade_antes = len(df)

    if "numeroControlePNCP" in df.columns:

        df = df.drop_duplicates(
            subset=[
                "numeroControlePNCP"
            ]
        )

    else:

        df = df.drop_duplicates()

    quantidade_depois = len(df)

    print(
        f"🧹 Duplicidades removidas: "
        f"{quantidade_antes - quantidade_depois}"
    )

    # --------------------------------------------------------
    # SALVAR
    # --------------------------------------------------------

    try:

        df.to_parquet(
            ARQUIVO_SAIDA,
            index=False,
            engine="pyarrow"
        )

    except Exception as erro:

        print()
        print(
            f"❌ Erro ao salvar Parquet: {erro}"
        )

        return False

    # --------------------------------------------------------
    # RESULTADO
    # --------------------------------------------------------

    print()
    print("=" * 70)
    print("✅ PARQUET SALVO COM SUCESSO")
    print("=" * 70)

    print(
        f"📁 Arquivo: {ARQUIVO_SAIDA}"
    )

    print(
        f"📊 Registros: {len(df)}"
    )

    print(
        f"📋 Colunas: {len(df.columns)}"
    )

    print(
        f"💾 Tamanho: "
        f"{ARQUIVO_SAIDA.stat().st_size:,} bytes"
    )

    print()
    print("📋 COLUNAS:")

    for coluna in df.columns:

        print(
            f"   - {coluna}"
        )

    return True


# ============================================================
# EXECUÇÃO
# ============================================================

if __name__ == "__main__":

    print()
    print("🏛️ EXTRATOR PNCP")
    print(
        f"CNPJ: {CNPJ_ORGAO}"
    )

    dados = extrair_dados_pncp()

    sucesso = salvar_dados(
        dados
    )

    print()

    if sucesso:

        print(
            "🎉 PROCESSO CONCLUÍDO COM SUCESSO!"
        )

    else:

        print(
            "⚠️ PROCESSO TERMINOU SEM NOVOS DADOS."
        )

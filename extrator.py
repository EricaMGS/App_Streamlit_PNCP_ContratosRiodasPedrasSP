import requests
import pandas as pd
import time
from pathlib import Path


# ============================================================
# CONFIGURAÇÕES
# ============================================================

CNPJ_ORGAO = "44826840000183"

URL_API = "https://pncp.gov.br/api/consulta/v1"

DIRETORIO_DADOS = Path(__file__).parent / "dados"
ARQUIVO_SAIDA = DIRETORIO_DADOS / "compras.parquet"

ANOS = [2024, 2025, 2026]

TAMANHO_PAGINA = 50

TIMEOUT = 60


# ============================================================
# BUSCAR MODALIDADES
# ============================================================

def buscar_modalidades():

    url = f"{URL_API}/modalidades"

    print()
    print("=" * 60)
    print("🔎 Buscando modalidades de contratação no PNCP")
    print("=" * 60)

    try:

        resposta = requests.get(
            url,
            params={"statusAtivo": "true"},
            timeout=TIMEOUT
        )

        print(
            f"HTTP {resposta.status_code}"
        )

        if resposta.status_code != 200:

            print("❌ Não foi possível obter as modalidades.")

            print(
                resposta.text[:2000]
            )

            return []

        dados = resposta.json()

        # Algumas APIs retornam diretamente uma lista,
        # outras podem colocar a lista dentro de "data".
        if isinstance(dados, list):

            modalidades = dados

        elif isinstance(dados, dict):

            modalidades = dados.get("data", [])

        else:

            modalidades = []

        if not modalidades:

            print(
                "❌ A API não retornou modalidades."
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

    except requests.exceptions.RequestException as erro:

        print(
            f"❌ Erro de conexão ao buscar modalidades: {erro}"
        )

        return []

    except Exception as erro:

        print(
            f"❌ Erro inesperado ao buscar modalidades: {erro}"
        )

        return []


# ============================================================
# CONSULTAR CONTRATAÇÕES
# ============================================================

def extrair_dados_pncp():

    todas_licitacoes = []

    modalidades = buscar_modalidades()

    if not modalidades:

        print(
            "❌ Nenhuma modalidade disponível."
        )

        return []

    url = f"{URL_API}/contratacoes/publicacao"

    for ano in ANOS:

        print()
        print("=" * 60)
        print(f"📅 CONSULTANDO ANO {ano}")
        print("=" * 60)

        for modalidade in modalidades:

            codigo_modalidade = modalidade.get("id")
            nome_modalidade = modalidade.get("nome")

            if codigo_modalidade is None:

                continue

            print()
            print(
                f"🔎 Modalidade: "
                f"{nome_modalidade} "
                f"(código {codigo_modalidade})"
            )

            pagina = 1

            while True:

                params = {
                    "dataInicial": f"01/01/{ano}",
                    "dataFinal": f"31/12/{ano}",
                    "codigoModalidadeContratacao": codigo_modalidade,
                    "cnpjOrgao": CNPJ_ORGAO,
                    "pagina": pagina,
                    "tamanhoPagina": TAMANHO_PAGINA
                }

                try:

                    print(
                        f"   📡 Página {pagina}..."
                    )

                    resposta = requests.get(
                        url,
                        params=params,
                        timeout=TIMEOUT
                    )

                    print(
                        f"      HTTP {resposta.status_code}"
                    )

                    # ------------------------------------------------
                    # ERRO DA API
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
                            "      ❌ Resposta não é JSON válido."
                        )

                        break

                    # ------------------------------------------------
                    # REGISTROS
                    # ------------------------------------------------

                    registros = dados.get(
                        "data",
                        []
                    )

                    print(
                        f"      📄 Registros: "
                        f"{len(registros)}"
                    )

                    if not registros:

                        break

                    # Adiciona os registros
                    todas_licitacoes.extend(
                        registros
                    )

                    # ------------------------------------------------
                    # FIM DA PAGINAÇÃO
                    # ------------------------------------------------

                    if len(registros) < TAMANHO_PAGINA:

                        break

                    pagina += 1

                    time.sleep(0.5)

                except requests.exceptions.Timeout:

                    print(
                        "      ❌ Timeout na consulta."
                    )

                    break

                except requests.exceptions.ConnectionError as erro:

                    print(
                        f"      ❌ Erro de conexão: {erro}"
                    )

                    break

                except requests.exceptions.RequestException as erro:

                    print(
                        f"      ❌ Erro HTTP: {erro}"
                    )

                    break

                except Exception as erro:

                    print(
                        f"      ❌ Erro inesperado: {erro}"
                    )

                    break

            # Pequena pausa entre modalidades
            time.sleep(0.3)

    print()
    print("=" * 60)
    print(
        f"📊 TOTAL BRUTO DE REGISTROS: "
        f"{len(todas_licitacoes)}"
    )
    print("=" * 60)

    return todas_licitacoes


# ============================================================
# SALVAR PARQUET
# ============================================================

def salvar_dados(lista_compras):

    DIRETORIO_DADOS.mkdir(
        parents=True,
        exist_ok=True
    )

    if not lista_compras:

        print()
        print(
            "⚠️ Nenhum registro foi encontrado."
        )

        print(
            "⚠️ O arquivo Parquet existente NÃO será sobrescrito."
        )

        return False

    # --------------------------------------------------------
    # DataFrame
    # --------------------------------------------------------

    df = pd.DataFrame(
        lista_compras
    )

    print()
    print(
        f"📊 DataFrame criado: "
        f"{len(df)} registros"
    )

    # --------------------------------------------------------
    # Remove duplicidades
    # --------------------------------------------------------

    quantidade_antes = len(df)

    colunas_identificacao = [
        coluna
        for coluna in [
            "numeroControlePNCP",
            "numeroControlePncp",
            "numeroCompra",
            "anoCompra"
        ]
        if coluna in df.columns
    ]

    if colunas_identificacao:

        df = df.drop_duplicates(
            subset=colunas_identificacao
        )

    else:

        df = df.drop_duplicates()

    quantidade_depois = len(df)

    print(
        f"🧹 Duplicidades removidas: "
        f"{quantidade_antes - quantidade_depois}"
    )

    # --------------------------------------------------------
    # Salva
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

    print()
    print(
        "✅ PARQUET SALVO COM SUCESSO"
    )

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
    print("📋 Colunas encontradas:")

    for coluna in df.columns:

        print(
            f"   - {coluna}"
        )

    return True


# ============================================================
# EXECUÇÃO PRINCIPAL
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

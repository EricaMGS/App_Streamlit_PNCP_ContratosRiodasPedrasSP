import requests
import pandas as pd
import time
from pathlib import Path


# ============================================================
# CONFIGURAÇÕES
# ============================================================

CNPJ_ORGAO = "44826840000183"

URL_API = (
    "https://pncp.gov.br/api/consulta/v1/contratacoes/publicacao"
)

DIRETORIO_DADOS = Path(__file__).parent / "dados"
ARQUIVO_SAIDA = DIRETORIO_DADOS / "compras.parquet"

ANOS = [2024, 2025, 2026]

TAMANHO_PAGINA = 50

TIMEOUT = 60


# ============================================================
# MODALIDADES DO PNCP
# ============================================================
#
# 1  = Leilão - Eletrônico
# 2  = Diálogo Competitivo
# 3  = Concurso
# 4  = Concorrência - Eletrônica
# 5  = Concorrência - Presencial
# 6  = Pregão - Eletrônico
# 7  = Pregão - Presencial
# 8  = Dispensa de Licitação
# 9  = Inexigibilidade
# 10 = Manifestação de Interesse
# 11 = Pré-qualificação
# 12 = Credenciamento
# 13 = Leilão - Presencial
#
# Fonte: Manual das APIs de Consulta do PNCP
# ============================================================

MODALIDADES = {
    1: "Leilão - Eletrônico",
    2: "Diálogo Competitivo",
    3: "Concurso",
    4: "Concorrência - Eletrônica",
    5: "Concorrência - Presencial",
    6: "Pregão - Eletrônico",
    7: "Pregão - Presencial",
    8: "Dispensa de Licitação",
    9: "Inexigibilidade",
    10: "Manifestação de Interesse",
    11: "Pré-qualificação",
    12: "Credenciamento",
    13: "Leilão - Presencial",
}


# ============================================================
# EXTRATOR
# ============================================================

def extrair_dados_pncp():

    todas_licitacoes = []

    print()
    print("=" * 70)
    print("🏛️ EXTRATOR PNCP")
    print("=" * 70)

    print(
        f"CNPJ: {CNPJ_ORGAO}"
    )

    print(
        f"Anos: {ANOS}"
    )

    print(
        f"Modalidades: {len(MODALIDADES)}"
    )

    # --------------------------------------------------------
    # ANOS
    # --------------------------------------------------------

    for ano in ANOS:

        print()
        print("=" * 70)
        print(
            f"📅 CONSULTANDO ANO {ano}"
        )
        print("=" * 70)

        # ----------------------------------------------------
        # MODALIDADES
        # ----------------------------------------------------

        for codigo_modalidade, nome_modalidade in MODALIDADES.items():

            print()
            print(
                f"🔎 {nome_modalidade}"
            )

            print(
                f"   Código: {codigo_modalidade}"
            )

            pagina = 1

            while True:

                params = {
                    # IMPORTANTE:
                    # PNCP exige AAAAMMDD
                    "dataInicial": f"{ano}0101",
                    "dataFinal": f"{ano}1231",

                    # Modalidade obrigatória
                    "codigoModalidadeContratacao": (
                        codigo_modalidade
                    ),

                    # O parâmetro correto é "cnpj"
                    "cnpj": CNPJ_ORGAO,

                    "pagina": pagina,
                    "tamanhoPagina": TAMANHO_PAGINA,
                }

                try:

                    print(
                        f"   📡 Página {pagina}..."
                    )

                    resposta = requests.get(
                        URL_API,
                        params=params,
                        timeout=TIMEOUT,
                        headers={
                            "accept": "*/*"
                        }
                    )

                    print(
                        f"      HTTP {resposta.status_code}"
                    )

                    # ------------------------------------------------
                    # ERRO HTTP
                    # ------------------------------------------------

                    if resposta.status_code != 200:

                        print(
                            "      ❌ Erro da API:"
                        )

                        print(
                            resposta.text[:2000]
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

                        print(
                            resposta.text[:1000]
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
                        f"      📄 Registros encontrados: "
                        f"{len(registros)}"
                    )

                    # ------------------------------------------------
                    # SEM MAIS REGISTROS
                    # ------------------------------------------------

                    if not registros:

                        break

                    # ------------------------------------------------
                    # ADICIONA OS REGISTROS
                    # ------------------------------------------------

                    todas_licitacoes.extend(
                        registros
                    )

                    # ------------------------------------------------
                    # PAGINAÇÃO
                    # ------------------------------------------------

                    if len(registros) < TAMANHO_PAGINA:

                        break

                    pagina += 1

                    time.sleep(0.5)

                except requests.exceptions.Timeout:

                    print(
                        "      ❌ Timeout."
                    )

                    break

                except requests.exceptions.ConnectionError as erro:

                    print(
                        f"      ❌ Erro de conexão: {erro}"
                    )

                    break

                except requests.exceptions.RequestException as erro:

                    print(
                        f"      ❌ Erro na requisição: {erro}"
                    )

                    break

                except Exception as erro:

                    print(
                        f"      ❌ Erro inesperado: {erro}"
                    )

                    break

            # Pausa entre modalidades
            time.sleep(0.3)

    # ========================================================
    # RESULTADO FINAL
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
    # NÃO SOBRESCREVE COM ARQUIVO VAZIO
    # --------------------------------------------------------

    if not lista_compras:

        print()
        print(
            "⚠️ Nenhum registro foi encontrado."
        )

        print(
            "⚠️ O Parquet existente NÃO será sobrescrito."
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
        f"📋 Total de colunas: "
        f"{len(df.columns)}"
    )

    # --------------------------------------------------------
    # REMOVE DUPLICADOS
    # --------------------------------------------------------

    quantidade_antes = len(df)

    if "numeroControlePNCP" in df.columns:

        df = df.drop_duplicates(
            subset=["numeroControlePNCP"]
        )

    else:

        df = df.drop_duplicates()

    quantidade_depois = len(df)

    print(
        f"🧹 Duplicidades removidas: "
        f"{quantidade_antes - quantidade_depois}"
    )

    # --------------------------------------------------------
    # SALVA
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
            f"❌ Erro ao salvar Parquet:"
        )

        print(
            erro
        )

        return False

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
    print("📋 COLUNAS ENCONTRADAS:")

    for coluna in df.columns:

        print(
            f"   - {coluna}"
        )

    return True


# ============================================================
# EXECUÇÃO
# ============================================================

if __name__ == "__main__":

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

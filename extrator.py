import requests
import pandas as pd
import time
from pathlib import Path

# ============================================================
# CONFIGURAÇÕES
# ============================================================

CNPJ_ORGAO = "44826840000183"

URL_API = "https://pncp.gov.br/api/consulta/v1/contratacoes/publicacao"

DIRETORIO_DADOS = Path(__file__).parent / "dados"
ARQUIVO_SAIDA = DIRETORIO_DADOS / "compras.parquet"

ANOS = [2024, 2025, 2026]

TAMANHO_PAGINA = 50


# ============================================================
# CONSULTA À API DO PNCP
# ============================================================

def extrair_dados_pncp():

    todas_licitacoes = []

    for ano in ANOS:

        print()
        print("=" * 60)
        print(f"🔎 Consultando PNCP - ano {ano}")
        print("=" * 60)

        pagina = 1

        while True:

            params = {
                "dataInicial": f"01/01/{ano}",
                "dataFinal": f"31/12/{ano}",
                "cnpjOrgao": CNPJ_ORGAO,
                "pagina": pagina,
                "tamanhoPagina": TAMANHO_PAGINA
            }

            try:

                print(
                    f"📡 Página {pagina}..."
                )

                resposta = requests.get(
                    URL_API,
                    params=params,
                    timeout=60
                )

                print(
                    f"   HTTP {resposta.status_code}"
                )

                # ------------------------------------------------
                # Verifica erro HTTP
                # ------------------------------------------------

                if resposta.status_code != 200:

                    print("❌ A API retornou um erro.")

                    print(
                        "Resposta do servidor:"
                    )

                    print(
                        resposta.text[:2000]
                    )

                    break

                # ------------------------------------------------
                # Converte resposta para JSON
                # ------------------------------------------------

                try:

                    dados = resposta.json()

                except ValueError:

                    print(
                        "❌ A resposta da API não é um JSON válido."
                    )

                    print(
                        resposta.text[:2000]
                    )

                    break

                # ------------------------------------------------
                # Obtém registros
                # ------------------------------------------------

                registros = dados.get("data", [])

                print(
                    f"   📄 Registros encontrados: "
                    f"{len(registros)}"
                )

                # ------------------------------------------------
                # Não existem mais registros
                # ------------------------------------------------

                if not registros:

                    print(
                        f"   ✅ Fim do ano {ano}."
                    )

                    break

                # ------------------------------------------------
                # Adiciona registros
                # ------------------------------------------------

                todas_licitacoes.extend(registros)

                # ------------------------------------------------
                # Se veio menos que o limite, acabou
                # ------------------------------------------------

                if len(registros) < TAMANHO_PAGINA:

                    print(
                        f"   ✅ Última página do ano {ano}."
                    )

                    break

                pagina += 1

                # Pequena pausa para não sobrecarregar a API
                time.sleep(0.5)

            except requests.exceptions.Timeout:

                print(
                    "❌ A requisição demorou demais (timeout)."
                )

                break

            except requests.exceptions.ConnectionError as erro:

                print(
                    f"❌ Erro de conexão: {erro}"
                )

                break

            except requests.exceptions.RequestException as erro:

                print(
                    f"❌ Erro na requisição: {erro}"
                )

                break

            except Exception as erro:

                print(
                    f"❌ Erro inesperado: {erro}"
                )

                break

    print()
    print("=" * 60)
    print(
        f"📊 TOTAL DE REGISTROS: "
        f"{len(todas_licitacoes)}"
    )
    print("=" * 60)

    return todas_licitacoes


# ============================================================
# SALVA OS DADOS EM PARQUET
# ============================================================

def salvar_dados(lista_compras):

    DIRETORIO_DADOS.mkdir(
        parents=True,
        exist_ok=True
    )

    # --------------------------------------------------------
    # Não cria um Parquet "falso" vazio silenciosamente
    # --------------------------------------------------------

    if not lista_compras:

        print()
        print(
            "⚠️ Nenhuma licitação foi encontrada."
        )

        print(
            "⚠️ O arquivo Parquet NÃO será sobrescrito."
        )

        return False

    # --------------------------------------------------------
    # Converte para DataFrame
    # --------------------------------------------------------

    df = pd.DataFrame(lista_compras)

    print()
    print(
        f"📊 DataFrame criado com "
        f"{len(df)} registros."
    )

    print(
        f"📋 Quantidade de colunas: "
        f"{len(df.columns)}"
    )

    # --------------------------------------------------------
    # Mostra as colunas encontradas
    # --------------------------------------------------------

    print()
    print("📋 Colunas encontradas:")

    for coluna in df.columns:

        print(f"   - {coluna}")

    # --------------------------------------------------------
    # Salva Parquet
    # --------------------------------------------------------

    try:

        df.to_parquet(
            ARQUIVO_SAIDA,
            index=False
        )

    except Exception as erro:

        print()
        print(
            f"❌ Erro ao salvar Parquet: {erro}"
        )

        return False

    print()
    print(
        f"✅ Arquivo salvo com sucesso:"
    )

    print(
        f"   {ARQUIVO_SAIDA}"
    )

    print(
        f"   Registros: {len(df)}"
    )

    print(
        f"   Tamanho: {ARQUIVO_SAIDA.stat().st_size:,} bytes"
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

    sucesso = salvar_dados(dados)

    print()

    if sucesso:

        print(
            "🎉 Processo concluído com sucesso!"
        )

    else:

        print(
            "⚠️ Processo terminou sem gerar novos dados."
        )

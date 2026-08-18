from datetime import date, timedelta
from pathlib import Path
import time
import pandas as pd
import requests

# ============================================================
# CONFIGURAÇÕES
# ============================================================

CNPJ_ORGAO = "44826840000183"

# API de CONSULTA oficial do PNCP
URL_CONSULTA = "https://pncp.gov.br/api/consulta/v1"
URL_CONTRATACOES = f"{URL_CONSULTA}/contratacoes/publicacao"

DIRETORIO_DADOS = Path(__file__).parent / "dados"
ARQUIVO_SAIDA = DIRETORIO_DADOS / "compras.parquet"

ANOS = [2024, 2025, 2026]

# Tamanho de página padrão seguro para a API do PNCP
TAMANHO_PAGINA = 50
TIMEOUT = 60
DIAS_POR_PERIODO = 180
PAUSA = 1.0

# ============================================================
# SESSÃO HTTP
# ============================================================

session = requests.Session()
session.headers.update(
    {"Accept": "application/json", "User-Agent": "Painel-PNCP-Rio-das-Pedras/1.0"}
)

# ============================================================
# CRIAR PERÍODOS DE CONSULTA
# ============================================================


def gerar_periodos(ano):
    inicio = date(ano, 1, 1)
    if ano == date.today().year:
        fim = date.today()
    else:
        fim = date(ano, 12, 31)

    periodos = []
    atual = inicio

    while atual <= fim:
        final_periodo = min(
            atual + timedelta(days=DIAS_POR_PERIODO - 1),
            fim,
        )
        # O PNCP geralmente espera datas no formato AAAA-MM-DD
        periodos.append(
            (atual.strftime("%Y-%m-%d"), final_periodo.strftime("%Y-%m-%d"))
        )
        atual = final_periodo + timedelta(days=1)

    return periodos


# ============================================================
# CONSULTAR CONTRATAÇÕES POR PERÍODO
# ============================================================


def consultar_periodo(data_inicial, data_final):
    registros_periodo = []
    pagina = 1

    while True:
        params = {
            "dataInicial": data_inicial,
            "dataFinal": data_final,
            "cnpj": CNPJ_ORGAO,
            "pagina": pagina,
            "tamanhoPagina": TAMANHO_PAGINA,
        }

        try:
            resposta = session.get(
                URL_CONTRATACOES, params=params, timeout=TIMEOUT
            )

            print(f"      Página {pagina} | HTTP {resposta.status_code}")

            if resposta.status_code == 429:
                print("      ⚠️ Rate limit atingido. Aguardando 10 segundos...")
                time.sleep(10)
                continue

            if resposta.status_code != 200:
                print("      ❌ Erro retornado pelo PNCP:")
                print(resposta.text[:1000])
                break

            try:
                dados = resposta.json()
            except ValueError:
                print("      ❌ Resposta não é JSON.")
                break

            # Tratamento robusto para capturar os registros independentemente da chave do JSON
            registros = []
            if isinstance(dados, list):
                registros = dados
            elif isinstance(dados, dict):
                # O PNCP costuma usar 'resultado', 'content', 'data' ou 'itens'
                for chave in ["resultado", "content", "data", "itens", "elements"]:
                    if chave in dados and isinstance(dados[chave], list):
                        registros = dados[chave]
                        break
                # Se não encontrou nas chaves padrão, tenta pegar qualquer lista do dicionário
                if not registros:
                    for v in dados.values():
                        if isinstance(v, list):
                            registros = v
                            break

            quantidade = len(registros)
            print(f"      📄 Registros encontrados: {quantidade}")

            if not registros:
                break

            registros_periodo.extend(registros)

            if quantidade < TAMANHO_PAGINA:
                break

            pagina += 1
            time.sleep(PAUSA)

        except requests.exceptions.Timeout:
            print("      ❌ Timeout.")
            break
        except requests.exceptions.RequestException as erro:
            print(f"      ❌ Erro de conexão: {erro}")
            break
        except Exception as erro:
            print(f"      ❌ Erro inesperado: {erro}")
            break

    return registros_periodo


# ============================================================
# EXTRAÇÃO PRINCIPAL
# ============================================================


def extrair_dados_pncp():
    todas_licitacoes = []

    print()
    print("=" * 70)
    print("🏛️ INICIANDO EXTRAÇÃO DO PNCP")
    print("=" * 70)
    print(f"CNPJ: {CNPJ_ORGAO}")
    print(f"Anos: {ANOS}")

    for ano in ANOS:
        periodos = gerar_periodos(ano)

        print()
        print("=" * 70)
        print(f"📅 ANO {ano} ({len(periodos)} períodos)")
        print("=" * 70)

        for numero_periodo, periodo in enumerate(periodos, start=1):
            data_inicial, data_final = periodo

            print()
            print(
                f"📆 Período {numero_periodo}/{len(periodos)}: {data_inicial} → {data_final}"
            )

            registros = consultar_periodo(data_inicial, data_final)

            if registros:
                print(f"      ✅ {len(registros)} registros adicionados.")
                todas_licitacoes.extend(registros)

            time.sleep(PAUSA)

    print()
    print("=" * 70)
    print(f"📊 TOTAL BRUTO DE REGISTROS: {len(todas_licitacoes)}")
    print("=" * 70)

    return todas_licitacoes


# ============================================================
# SALVAR PARQUET
# ============================================================


def salvar_dados(lista_compras):
    DIRETORIO_DADOS.mkdir(parents=True, exist_ok=True)

    if not lista_compras:
        print()
        print("⚠️ Nenhum registro foi encontrado.")
        print("⚠️ O arquivo Parquet existente NÃO será sobrescrito.")
        return False

    df = pd.DataFrame(lista_compras)

    print()
    print(f"📊 DataFrame criado com {len(df)} registros.")
    print(f"📋 Colunas: {len(df.columns)}")

    quantidade_antes = len(df)
    if "numeroControlePNCP" in df.columns:
        df = df.drop_duplicates(subset=["numeroControlePNCP"])
    else:
        df = df.drop_duplicates()
    quantidade_depois = len(df)

    print(f"🧹 Duplicidades removidas: {quantidade_antes - quantidade_depois}")

    try:
        df.to_parquet(ARQUIVO_SAIDA, index=False, engine="pyarrow")
    except Exception as erro:
        print()
        print(f"❌ Erro ao salvar Parquet: {erro}")
        return False

    print()
    print("=" * 70)
    print("✅ PARQUET SALVO COM SUCESSO")
    print("=" * 70)
    print(f"📁 Arquivo: {ARQUIVO_SAIDA}")
    print(f"📊 Registros: {len(df)}")
    print(f"📋 Colunas: {len(df.columns)}")
    print(f"💾 Tamanho: {ARQUIVO_SAIDA.stat().st_size:,} bytes")
    print()
    print("📋 COLUNAS:")
    for coluna in df.columns:
        print(f"   - {coluna}")

    return True


# ============================================================
# EXECUÇÃO
# ============================================================

if __name__ == "__main__":
    print()
    print("🏛️ EXTRATOR PNCP")
    print(f"CNPJ: {CNPJ_ORGAO}")

    dados = extrair_dados_pncp()
    sucesso = salvar_dados(dados)

    print()
    if sucesso:
        print("🎉 PROCESSO CONCLUÍDO COM SUCESSO!")
    else:
        print("⚠️ PROCESSO TERMINOU SEM NOVOS DADOS.")

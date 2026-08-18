import requests
import pandas as pd
import time
import os
from datetime import datetime

CNPJ_ORGAO = "44826840000183" 
DATA_INICIAL = "20230101" 
DATA_FINAL = datetime.now().strftime("%Y%m%d")

DIRETORIO_DADOS = "dados"
ARQUIVO_SAIDA = f"{DIRETORIO_DADOS}/contratos.parquet"

def extrair_dados_pncp():
    url = f"https://pncp.gov.br/api/pncp/v1/orgaos/{CNPJ_ORGAO}/contratos"
    pagina = 1
    tamanho_pagina = 50
    todos_contratos = []

    while True:
        params = {"dataInicial": DATA_INICIAL, "dataFinal": DATA_FINAL, "pagina": pagina, "tamanhoPagina": tamanho_pagina}
        try:
            resposta = requests.get(url, params=params)
            if resposta.status_code != 200:
                break
            dados = resposta.json()
            registros = dados.get("data", [])
            
            if not registros:
                break
            todos_contratos.extend(registros)
            if len(registros) < tamanho_pagina:
                break
            pagina += 1
            time.sleep(0.5)
        except Exception:
            break
    return todos_contratos

def salvar_dados(lista_contratos):
    if not lista_contratos:
        return
    os.makedirs(DIRETORIO_DADOS, exist_ok=True)
    df = pd.DataFrame(lista_contratos)
    df.to_parquet(ARQUIVO_SAIDA, index=False)

if __name__ == "__main__":
    dados = extrair_dados_pncp()
    salvar_dados(dados)

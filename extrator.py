import requests
import pandas as pd
import time
import os

CNPJ_ORGAO = "44826840000183" 
DIRETORIO_DADOS = "dados"
ARQUIVO_SAIDA = f"{DIRETORIO_DADOS}/compras.parquet" 

def extrair_dados_pncp():
    # Usando a API de Consulta que você encontrou!
    url = "https://pncp.gov.br/api/consulta/v1/licitacoes"
    todas_licitacoes = []
    
    # Vamos buscar os dados de 2024, 2025 e 2026
    anos = [2024, 2025, 2026]

    for ano in anos:
        pagina = 1
        while True:
            params = {
                "cnpjOrgao": CNPJ_ORGAO,
                "ano": ano,
                "pagina": pagina,
                "tamanhoPagina": 50
            }
            try:
                resposta = requests.get(url, params=params)
                if resposta.status_code != 200:
                    break
                
                dados = resposta.json()
                registros = dados.get("data", [])
                
                if not registros:
                    break
                
                todas_licitacoes.extend(registros)
                
                # Se vieram menos de 50 registros, acabou a páginação deste ano
                if len(registros) < 50:
                    break
                
                pagina += 1
                time.sleep(0.5)
            except Exception:
                break
                
    return todas_licitacoes

def salvar_dados(lista_compras):
    os.makedirs(DIRETORIO_DADOS, exist_ok=True)
    if not lista_compras:
        df = pd.DataFrame(columns=['numeroCompra', 'anoCompra', 'dataPublicacaoPncp', 'valorTotalEstimado', 'objetoCompra'])
    else:
        df = pd.DataFrame(lista_compras)
        
    df.to_parquet(ARQUIVO_SAIDA, index=False)

if __name__ == "__main__":
    dados = extrair_dados_pncp()
    salvar_dados(dados)

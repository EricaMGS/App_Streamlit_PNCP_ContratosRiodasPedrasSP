import pandas as pd
import requests
import streamlit as st

st.set_page_config(
    page_title="Consulta PNCP - Rio das Pedras/SP", layout="wide"
)

st.title("📊 Consulta de Contratações Públicas - Rio das Pedras/SP")
st.markdown("Dados oficiais extraídos diretamente do PNCP.")

# Filtros na Barra Lateral
st.sidebar.header("Filtros")
data_inicio = st.sidebar.date_input(
    "Data Inicial", value=pd.to_datetime("2026-01-01")
)
data_fim = st.sidebar.date_input(
    "Data Final", value=pd.to_datetime("2026-12-31")
)

if st.sidebar.button("Gerar Relatório"):
  with st.spinner("Consultando API do PNCP..."):
    cnpj_prefeitura = "44826840000183"

    # O PNCP exige estritamente o formato AAAAMMDD
    d_inicio_str = data_inicio.strftime("%Y%m%d")
    d_fim_str = data_fim.strftime("%Y%m%d")

    # Endpoint geral de contratações (que não exige modalidade obrigatória)
    url = "https://pncp.gov.br/api/consulta/v1/contratacoes"

    params = {
        "cnpj": cnpj_prefeitura,
        "dataInicial": d_inicio_str,
        "dataFinal": d_fim_str,
        "pagina": 1,
    }

    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML,"
            " like Gecko) Chrome/120.0.0.0 Safari/537.36"
        ),
        "Accept": "application/json",
    }

    try:
      response = requests.get(url, params=params, headers=headers, timeout=15)

      if response.status_code == 200:
        dados = response.json()
        
        # Tratamento seguro para diferentes estruturas de retorno do JSON
        if isinstance(dados, list):
            lista_contratacoes = dados
        elif isinstance(dados, dict):
            lista_contratacoes = dados.get("data", dados.get("items", []))
        else:
            lista_contratacoes = []

        df = pd.DataFrame(lista_contratacoes)

        if not df.empty:
          st.success(
              f"Sucesso! Encontrados {len(df)} registros para Rio das"
              " Pedras/SP."
          )
          st.dataframe(df)

          csv = df.to_csv(index=False).encode("utf-8")
          st.download_button(
              label="📥 Baixar Relatório em CSV",
              data=csv,
              file_name=(
                  f"contratacoes_rio_das_pedras_{d_inicio_str}_{d_fim_str}.csv"
              ),
              mime="text/csv",
          )
        else:
          st.warning(
              "Nenhum contrato encontrado no intervalo de datas selecionado."
          )
      else:
        st.error(f"Erro na API do PNCP (Código: {response.status_code}) - Detalhes: {response.text}")
    except requests.exceptions.Timeout:
      st.error(
          "O servidor do PNCP demorou muito para responder. Tente novamente"
          " mais tarde."
      )
    except Exception as e:
      st.error(f"Erro de conexão: {e}")

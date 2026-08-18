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
data_inicio = st.sidebar.date_input("Data Inicial", value=pd.to_datetime("2026-01-01"))
data_fim = st.sidebar.date_input("Data Final", value=pd.to_datetime("2026-12-31"))

if st.sidebar.button("Gerar Relatório"):
  with st.spinner("Consultando API do PNCP..."):
    # CNPJ da Prefeitura de Rio das Pedras/SP sem formatação
    cnpj_prefeitura = "44826840000183"

    # Endpoint oficial de contratações por órgão e período do PNCP
    # Formato de data exigido pela API do PNCP na URL: AAAAMMDD
    d_inicio_str = data_inicio.strftime("%Y%m%d")
    d_fim_str = data_fim.strftime("%Y%m%d")

    url = f"https://pncp.gov.br/api/consulta/v1/orgaos/{cnpj_prefeitura}/contratacoes"

    params = {
        "data_inicial": d_inicio_str,
        "data_final": d_fim_str,
        "pagina": 1,
    }

    try:
      response = requests.get(url, params=params)

      if response.status_code == 200:
        dados = response.json()
        # O PNCP costuma retornar uma lista direta ou dentro de uma chave 'resultado' / 'items'
        # Dependendo da versão exata do endpoint, ajustamos o encapsulamento:
        lista_contratacoes = dados if isinstance(dados, list) else dados.get("items", [])
        
        df = pd.DataFrame(lista_contratacoes)

        if not df.empty:
          st.success(f"Sucesso! Encontrados {len(df)} registros para Rio das Pedras/SP.")
          st.dataframe(df)

          # Botão de Download
          csv = df.to_csv(index=False).encode("utf-8")
          st.download_button(
              label="📥 Baixar Relatório em CSV",
              data=csv,
              file_name=f"contratacoes_rio_das_pedras_{d_inicio_str}_{d_fim_str}.csv",
              mime="text/csv",
          )
        else:
          st.warning("Nenhum contrato encontrado no intervalo de datas selecionado.")
      else:
        st.error(f"Erro na API do PNCP (Código: {response.status_code})")
    except Exception as e:
      st.error(f"Erro de conexão: {e}")

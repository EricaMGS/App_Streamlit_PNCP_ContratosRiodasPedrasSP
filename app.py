import pandas as pd
import requests
import streamlit as st

st.set_page_config(page_title="Consulta PNCP - Rio das Pedras/SP", layout="wide")

st.title("📊 Consulta de Contratos e Compras - Rio das Pedras/SP")
st.markdown("Dados oficiais extraídos diretamente do PNCP.")

# Filtros na Barra Lateral
st.sidebar.header("Filtros de Período")
data_inicio = st.sidebar.date_input("Data Inicial", value=pd.to_datetime("2026-01-01"))
data_fim = st.sidebar.date_input("Data Final", value=pd.to_datetime("2026-03-31"))

if st.sidebar.button("Gerar Relatório"):
    with st.spinner("Consultando dados oficiais no PNCP..."):
        cnpj_prefeitura = "44826840000183"

        d_inicio_str = data_inicio.strftime("%Y%m%d")
        d_fim_str = data_fim.strftime("%Y%m%d")

        # URL corrigida para o endpoint correto de consulta por órgão
        url = "https://pncp.gov.br/api/consulta/v1/contratos"

        # CORREÇÃO: O parâmetro esperado pelo PNCP nesta rota é 'cnpj'
        params = {
            "cnpj": cnpj_prefeitura,
            "dataInicial": d_inicio_str,
            "dataFinal": d_fim_str,
            "pagina": 1,
        }

        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
            "Accept": "application/json",
        }

        try:
            response = requests.get(url, params=params, headers=headers, timeout=25)

            if response.status_code == 200:
                dados = response.json()

                # Tratamento do retorno da API corrigido
                if isinstance(dados, list):
                    lista_contratos = dados
                elif isinstance(dados, dict):
                    # Uso correto do .get() para evitar erro se a chave não existir
                    lista_contratos = dados.get("data", dados.get("items", []))
                else:
                    lista_contratos = []

                if lista_contratos:
                    df = pd.DataFrame(lista_contratos)
                    st.success(f"Sucesso! Encontrados {len(df)} registros.")
                    st.dataframe(df)

                    csv = df.to_csv(index=False).encode("utf-8")
                    st.download_button(
                        label="📥 Baixar Relatório em CSV",
                        data=csv,
                        file_name=f"contratos_rio_das_pedras_{d_inicio_str}_{d_fim_str}.csv",
                        mime="text/csv",
                    )
                else:
                    st.warning("Nenhum contrato encontrado no intervalo de datas selecionado.")
            
            elif response.status_code in [502, 503, 504]:
                st.warning(f"O servidor do PNCP está instável (Erro {response.status_code}). Tente novamente em alguns segundos.")
            else:
                st.error(f"Erro na API do PNCP (Código: {response.status_code}) - Detalhes: {response.text}")

        except requests.exceptions.Timeout:
            st.error("O servidor do PNCP demorou muito para responder. Tente um período menor.")
        except Exception as e:
            st.error(f"Erro técnico: {e}")

import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="Painel PNCP - Rio das Pedras", page_icon="🏛️", layout="wide")

@st.cache_data(ttl="1h")
def carregar_dados():
    caminho_arquivo = "dados/compras.parquet"
    if not os.path.exists(caminho_arquivo):
        return pd.DataFrame()
    return pd.read_parquet(caminho_arquivo)

st.title("🏛️ Painel de Licitações (PNCP)")
st.markdown("Monitoramento de Editais - Rio das Pedras/SP")
st.divider()

df_compras = carregar_dados()

if df_compras.empty:
    st.warning("⚠️ Os dados ainda não carregaram ou a base está vazia. Rode o GitHub Actions e limpe o cache.")
else:
    st.success(f"✅ Sucesso! Foram encontradas {len(df_compras)} licitações.")
    
    # Exibe a base de dados completa para você analisar quais colunas a API nova trouxe
    st.dataframe(df_compras, use_container_width=True)

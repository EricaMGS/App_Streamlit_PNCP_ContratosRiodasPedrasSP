import streamlit as st
import pandas as pd
from pathlib import Path

# Configuração da página
st.set_page_config(
    page_title="Painel PNCP - Rio das Pedras",
    page_icon="🏛️",
    layout="wide"
)


@st.cache_data(ttl="1h")
def carregar_dados():
    caminho_arquivo = Path(__file__).parent / "dados" / "compras.parquet"

    if not caminho_arquivo.exists():
        return pd.DataFrame()

    try:
        return pd.read_parquet(caminho_arquivo)
    except Exception as e:
        st.error(f"Erro ao ler o arquivo Parquet: {e}")
        return pd.DataFrame()


# Título
st.title("🏛️ Painel de Licitações (PNCP)")
st.markdown("Monitoramento de Editais - Rio das Pedras/SP")
st.divider()

# Carrega os dados
df_compras = carregar_dados()

# Verifica se encontrou dados
if df_compras.empty:
    st.warning(
        "⚠️ Os dados ainda não carregaram ou a base está vazia."
    )

    caminho = Path(__file__).parent / "dados" / "compras.parquet"

    st.write("### 🔎 Diagnóstico")
    st.write(f"Arquivo esperado: `{caminho}`")
    st.write(f"Arquivo existe? **{caminho.exists()}**")

else:
    st.success(
        f"✅ Sucesso! Foram encontradas {len(df_compras)} licitações."
    )

    st.write("### 📊 Dados encontrados")
    st.dataframe(
        df_compras,
        use_container_width=True
    )

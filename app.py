import streamlit as st
import pandas as pd
from pathlib import Path

st.set_page_config(
    page_title="Painel PNCP - Rio das Pedras",
    page_icon="🏛️",
    layout="wide"
)


def carregar_dados():

    caminho_arquivo = (
        Path(__file__).parent
        / "dados"
        / "compras.parquet"
    )

    if not caminho_arquivo.exists():
        return pd.DataFrame()

    try:
        return pd.read_parquet(
            caminho_arquivo,
            engine="pyarrow"
        )

    except Exception as erro:

        st.error(
            f"❌ Erro ao ler o arquivo Parquet: {erro}"
        )

        return pd.DataFrame()


st.title("🏛️ Painel de Licitações (PNCP)")

st.markdown(
    "Monitoramento de Editais - Rio das Pedras/SP"
)

st.divider()

df_compras = carregar_dados()
st.write("### 🔬 Diagnóstico do Parquet")

st.write(
    f"Linhas: **{len(df_compras)}**"
)

st.write(
    f"Colunas: **{len(df_compras.columns)}**"
)

st.write(
    "Nomes das colunas:"
)

st.write(
    list(df_compras.columns)
)

if not df_compras.empty:
    st.write("### Primeiros registros")
    st.dataframe(
        df_compras.head(10),
        use_container_width=True
    )

if df_compras.empty:

    st.warning(
        "⚠️ Os dados ainda não carregaram "
        "ou a base está vazia."
    )

    caminho = (
        Path(__file__).parent
        / "dados"
        / "compras.parquet"
    )

    st.write("### 🔎 Diagnóstico")

    st.write(
        f"Arquivo esperado: `{caminho}`"
    )

    st.write(
        f"Arquivo existe? **{caminho.exists()}**"
    )

    if caminho.exists():

        tamanho = caminho.stat().st_size

        st.write(
            f"Tamanho do arquivo: **{tamanho:,} bytes**"
        )

else:

    st.success(
        f"✅ Sucesso! Foram encontradas "
        f"{len(df_compras)} licitações."
    )

    st.write(
        f"### 📊 {len(df_compras)} registros"
    )

    st.dataframe(
        df_compras,
        use_container_width=True
    )

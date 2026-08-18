import streamlit as st
import pandas as pd
import plotly.express as px
import os

st.set_page_config(page_title="Painel PNCP", page_icon="🏛️", layout="wide")

@st.cache_data(ttl="1h")
def carregar_dados():
    caminho_arquivo = "dados/compras.parquet"
    if not os.path.exists(caminho_arquivo):
        return pd.DataFrame()
    df = pd.read_parquet(caminho_arquivo)
    if 'dataPublicacaoPncp' in df.columns:
        df['dataPublicacaoPncp'] = pd.to_datetime(df['dataPublicacaoPncp'], errors='coerce')
    return df

st.title("🏛️ Painel de Licitações (PNCP)")
st.markdown("Monitoramento de Editais e Compras Públicas - Rio das Pedras/SP")
st.divider()

df_compras = carregar_dados()

if df_compras.empty:
    st.warning("⚠️ Os dados de compras ainda não foram carregados. Aguarde a execução do GitHub Actions.")
else:
    st.subheader("Visão Geral")
    col1, col2, col3 = st.columns(3)
    col1.metric("Total de Licitações/Compras", f"{len(df_compras)}")
    
    if 'valorTotalEstimado' in df_compras.columns:
        valor_total = df_compras['valorTotalEstimado'].sum()
        col2.metric("Valor Total Estimado", f"R$ {valor_total:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        
    st.divider()
    st.subheader("Evolução das Licitações")
    
    if 'dataPublicacaoPncp' in df_compras.columns and 'valorTotalEstimado' in df_compras.columns:
        df_agrupado = df_compras.copy()
        df_agrupado['Mes_Ano'] = df_agrupado['dataPublicacaoPncp'].dt.to_period('M').astype(str)
        df_grafico = df_agrupado.groupby('Mes_Ano').agg(Valor_Total=('valorTotalEstimado', 'sum'), Qtd_Compras=('numeroCompra', 'count')).reset_index().dropna()
        
        c1, c2 = st.columns(2)
        with c1:
            fig_valor = px.bar(df_grafico, x='Mes_Ano', y='Valor_Total', title="Volume Estimado por Mês", color_discrete_sequence=['#1f77b4'])
            fig_valor.update_traces(hovertemplate="Mês: %{x}<br>Valor: R$ %{y:,.2f}<extra></extra>")
            fig_valor.update_layout(plot_bgcolor="rgba(0,0,0,0)", margin=dict(t=40, l=0, r=0, b=0))
            st.plotly_chart(fig_valor, use_container_width=True)
            
        with c2:
            fig_qtd = px.line(df_grafico, x='Mes_Ano', y='Qtd_Compras', title="Editais Publicados por Mês", markers=True, color_discrete_sequence=['#ff7f0e'])
            fig_qtd.update_layout(plot_bgcolor="rgba(0,0,0,0)", margin=dict(t=40, l=0, r=0, b=0))
            st.plotly_chart(fig_qtd, use_container_width=True)
            
    st.divider()
    colunas_tabela = ['anoCompra', 'numeroCompra', 'dataPublicacaoPncp', 'objetoCompra', 'valorTotalEstimado']
    colunas_presentes = [col for col in colunas_tabela if col in df_compras.columns]
    st.dataframe(df_compras[colunas_presentes], use_container_width=True, hide_index=True)

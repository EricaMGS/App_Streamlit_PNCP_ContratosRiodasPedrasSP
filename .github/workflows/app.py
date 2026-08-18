import streamlit as st
import pandas as pd
import plotly.express as px
import os

st.set_page_config(page_title="Painel PNCP", page_icon="🏛️", layout="wide")

@st.cache_data(ttl="1h")
def carregar_dados():
    caminho_arquivo = "dados/contratos.parquet"
    if not os.path.exists(caminho_arquivo):
        return pd.DataFrame()
    df = pd.read_parquet(caminho_arquivo)
    if 'dataVigenciaInicio' in df.columns:
        df['dataVigenciaInicio'] = pd.to_datetime(df['dataVigenciaInicio'], errors='coerce')
    return df

st.title("🏛️ Painel de Contratos Públicos (PNCP)")
st.divider()

df_contratos = carregar_dados()

if df_contratos.empty:
    st.warning("⚠️ Os dados ainda estão sendo processados pela primeira vez no GitHub Actions.")
else:
    st.subheader("Visão Geral")
    col1, col2, col3 = st.columns(3)
    col1.metric("Total de Contratos", f"{len(df_contratos)}")
    
    if 'valorInicial' in df_contratos.columns:
        valor_total = df_contratos['valorInicial'].sum()
        col2.metric("Volume Financeiro", f"R$ {valor_total:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        
    st.divider()
    st.subheader("Evolução das Contratações")
    
    if 'dataAssinatura' in df_contratos.columns and 'valorInicial' in df_contratos.columns:
        df_contratos['dataAssinatura'] = pd.to_datetime(df_contratos['dataAssinatura'], errors='coerce')
        df_contratos['Mes_Ano'] = df_contratos['dataAssinatura'].dt.to_period('M').astype(str)
        df_agrupado = df_contratos.groupby('Mes_Ano').agg(Valor_Total=('valorInicial', 'sum'), Qtd_Contratos=('id', 'count')).reset_index().dropna()
        
        c1, c2 = st.columns(2)
        with c1:
            fig_valor = px.bar(df_agrupado, x='Mes_Ano', y='Valor_Total', title="Volume Financeiro por Mês", color_discrete_sequence=['#1f77b4'])
            fig_valor.update_traces(hovertemplate="Mês: %{x}Valor: R$ %{y:,.2f}")
            fig_valor.update_layout(plot_bgcolor="rgba(0,0,0,0)", margin=dict(t=40, l=0, r=0, b=0))
            st.plotly_chart(fig_valor, use_container_width=True)
            
        with c2:
            fig_qtd = px.line(df_agrupado, x='Mes_Ano', y='Qtd_Contratos', title="Contratos por Mês", markers=True, color_discrete_sequence=['#ff7f0e'])
            fig_qtd.update_layout(plot_bgcolor="rgba(0,0,0,0)", margin=dict(t=40, l=0, r=0, b=0))
            st.plotly_chart(fig_qtd, use_container_width=True)
            
    st.divider()
    st.dataframe(df_contratos, use_container_width=True, hide_index=True)

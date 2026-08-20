# GRÁFICOS
    st.markdown("### 📈 Análise Gráfica")
    coluna_data = next((c for c in ["dataPublicacao", "dataAssinatura", "dataInclusao"] if c in df.columns), None)
    coluna_valor = next((c for c in ["valorGlobal", "valorInicial", "valorTotalHomologado", "valorTotalEstimado"] if c in df.columns), None)
    
    if coluna_data:
        try:
            df_grafico = df.copy()
            df_grafico['mes_ano'] = pd.to_datetime(df_grafico[coluna_data], errors='coerce').dt.to_period('M').astype(str)
            
            # Criar abas para separar os tipos de visualização
            aba1, aba2 = st.tabs(["🔢 Quantidade de Registros", "💰 Volume Financeiro (R$)"])
            
            with aba1:
                st.markdown(f"#### Quantidade de {tipo_consulta} por Mês/Ano")
                contagem_mes = df_grafico['mes_ano'].value_counts().sort_index()
                st.bar_chart(contagem_mes)
                
            with aba2:
                if coluna_valor:
                    st.markdown(f"#### Volume Financeiro de {tipo_consulta} por Mês/Ano")
                    df_grafico[coluna_valor] = pd.to_numeric(df_grafico[coluna_valor], errors='coerce').fillna(0)
                    soma_mes = df_grafico.groupby('mes_ano')[coluna_valor].sum().sort_index()
                    st.bar_chart(soma_mes)
                else:
                    st.info("ℹ️ Não há dados financeiros disponíveis para esta consulta.")
                    
        except Exception as e:
            st.info(f"ℹ️ Não foi possível gerar os gráficos: {e}")
    else:
        st.info("ℹ️ Coluna de data não encontrada para exibição gráfica.")

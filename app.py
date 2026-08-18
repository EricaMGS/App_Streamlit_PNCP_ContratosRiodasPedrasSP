import pandas as pd
import requests
import streamlit as st
import io

# Configuração da página para ser responsiva e profissional
st.set_page_config(page_title="Relatório PNCP - Rio das Pedras", layout="wide")

# Estilização CSS para deixar os cards e botões mais elegantes
st.markdown("""
    <style>
    .stMetric { background-color: #f0f2f6; padding: 15px; border-radius: 10px; }
    </style>
""", unsafe_allow_html=True)

st.title("🏛️ Dashboard: Contratos Rio das Pedras/SP")
st.markdown("Consulta oficial integrada ao Portal Nacional de Contratações Públicas (PNCP).")

# Barra lateral para filtros
st.sidebar.header("Configurações do Relatório")
data_inicio = st.sidebar.date_input("Data Inicial", value=pd.to_datetime("2026-01-01"))
data_fim = st.sidebar.date_input("Data Final", value=pd.to_datetime("2026-03-31"))

if st.sidebar.button("Gerar Dashboard Completo"):
    with st.spinner("Consultando dados no PNCP..."):
        # CNPJ da prefeitura
        cnpj = "44826840000183"
        d_inicio_str = data_inicio.strftime("%Y%m%d")
        d_fim_str = data_fim.strftime("%Y%m%d")
        
        url = "https://pncp.gov.br/api/consulta/v1/contratos"
        headers = {"User-Agent": "Mozilla/5.0", "Accept": "application/json"}
        
        all_data = []
        pagina = 1
        
        # Loop de Paginação (pega tudo, mesmo se passar de 500 registros)
        while True:
            params = {"cnpj": cnpj, "dataInicial": d_inicio_str, "dataFinal": d_fim_str, "pagina": pagina}
            response = requests.get(url, params=params, headers=headers, timeout=20)
            
            if response.status_code == 200:
                lote = response.json()
                lote = lote.get("data", lote.get("items", [])) if isinstance(lote, dict) else lote
                
                if not lote: break
                all_data.extend(lote)
                if len(lote) < 50: break
                pagina += 1
            else:
                break
        
        if all_data:
            df = pd.DataFrame(all_data)
            
            # --- DASHBOARD RESPONSIVO ---
            # Cards de resumo
            c1, c2 = st.columns(2)
            c1.metric("Total de Contratos", len(df))
            
            # Formatação do Excel
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='Contratos')
                workbook = writer.book
                worksheet = writer.sheets['Contratos']
                header_format = workbook.add_format({'bold': True, 'bg_color': '#003366', 'font_color': 'white'})
                for col_num, value in enumerate(df.columns.values):
                    worksheet.write(0, col_num, value, header_format)
            
            # Exibição da tabela responsiva
            st.subheader("Lista de Processos")
            st.dataframe(df, use_container_width=True)
            
            # Download button
            st.download_button(
                label="📥 Baixar Relatório Profissional (.xlsx)",
                data=buffer.getvalue(),
                file_name=f"Contratos_Rio_das_Pedras_{data_inicio}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("Nenhum contrato encontrado para o período selecionado.")

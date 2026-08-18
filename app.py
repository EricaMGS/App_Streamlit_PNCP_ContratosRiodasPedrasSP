import io
import pandas as pd
import requests
import streamlit as st
from docx import Document
from docx.shared import Pt, RGBColor

# Configuração da página responsiva
st.set_page_config(page_title="Portal PNCP - Rio das Pedras/SP", layout="wide")

st.markdown("""
    <style>
    .stMetric { background-color: #f0f2f6; padding: 15px; border-radius: 10px; }
    </style>
""", unsafe_allow_html=True)

st.title("🏛️ Contratações de Rio das Pedras/SP")
st.markdown("Consulta integrada de Editais, Atas e Contratos direto do Portal Nacional de Contratações Públicas (PNCP).")

# --- BARRA LATERAL DE CONFIGURAÇÕES ---
st.sidebar.header("Parâmetros da Consulta")

tipo_consulta = st.sidebar.selectbox(
    "Selecione o tipo de dado:",
    ["Contratos", "Atas de Registro de Preços", "Editais e Avisos de Contratação"],
)

data_inicio = st.sidebar.date_input("Data Inicial", value=pd.to_datetime("2025-01-01"))
data_fim = st.sidebar.date_input("Data Final", value=pd.to_datetime("2026-12-31"))

endpoints = {
    "Contratos": "https://pncp.gov.br/api/consulta/v1/contratos",
    "Atas de Registro de Preços": "https://pncp.gov.br/api/consulta/v1/atas",
    "Editais e Avisos de Contratação": "https://pncp.gov.br/api/consulta/v1/contratacoes/publicacao",
}

if st.sidebar.button("Gerar Relatório Consolidado"):
    url_escolhida = endpoints[tipo_consulta]

    with st.spinner(f"Buscando '{tipo_consulta}' no PNCP..."):
        cnpj_prefeitura = "44826840000183"
        d_inicio_str = data_inicio.strftime("%Y%m%d")
        d_fim_str = data_fim.strftime("%Y%m%d")

        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
            "Accept": "application/json",
        }

        all_data = []
        pagina = 1

        while pagina <= 10:
            params = {
                "cnpj": cnpj_prefeitura,
                "dataInicial": d_inicio_str,
                "dataFinal": d_fim_str,
                "pagina": pagina,
            }

            try:
                response = requests.get(url_escolhida, params=params, headers=headers, timeout=30)
                if response.status_code == 200:
                    lote = response.json()
                    lote = lote.get("data", lote.get("items", [])) if isinstance(lote, dict) else lote
                    if not lote:
                        break
                    all_data.extend(lote)
                    if len(lote) < 50:
                        break
                    pagina += 1
                else:
                    break
            except Exception:
                break

        if all_data:
            df = pd.DataFrame(all_data)

            if "orgaoEntidade" in df.columns:
                df = df[df["orgaoEntidade"].astype(str).str.contains(cnpj_prefeitura, na=False)]

            if not df.empty:
                st.success(f"Consulta realizada com sucesso! Categoria: {tipo_consulta}")

                c1, c2 = st.columns(2)
                c1.metric(f"Total de Registros ({tipo_consulta})", len(df))

                col_valor = None
                for col in ["valorGlobal", "valorTotalEstimado", "valorHomologado", "valorInicial"]:
                    if col in df.columns:
                        col_valor = col
                        break

                total_valor = 0
                if col_valor:
                    total_valor = pd.to_numeric(df[col_valor], errors="coerce").sum()
                    c2.metric("Valor Consolidado", f"R$ {total_valor:,.2f}")
                else:
                    c2.metric("Município", "Rio das Pedras/SP")

                st.subheader(f"📋 Relação Consolidada: {tipo_consulta}")
                st.dataframe(df, use_container_width=True)

                # --- SEÇÃO DE DOWNLOADS ---
                st.markdown("### 📥 Opções de Exportação")
                col_dl1, col_dl2 = st.columns(2)

                # 1. EXCEL (.xlsx)
                buffer_excel = io.BytesIO()
                with pd.ExcelWriter(buffer_excel, engine="xlsxwriter") as writer:
                    df.to_excel(writer, index=False)
                
                with col_dl1:
                    st.download_button("📊 Baixar Planilha (.xlsx)", buffer_excel.getvalue(), 
                                       file_name=f"{tipo_consulta.replace(' ', '_')}_Rio_Das_Pedras.xlsx")

                # 2. WORD (.docx)
                doc = Document()
                doc.add_heading(f"Relatório: {tipo_consulta}", 0)
                doc.add_paragraph(f"Município: Rio das Pedras/SP")
                doc.add_paragraph(f"Total de registros: {len(df)}")
                
                buffer_word = io.BytesIO()
                doc.save(buffer_word)
                buffer_word.seek(0)

                with col_dl2:
                    st.download_button("📝 Baixar Relatório em Word (.docx)", buffer_word.getvalue(), 
                                       file_name=f"Relatorio_{tipo_consulta.replace(' ', '_')}_Rio_Das_Pedras.docx")
            else:
                st.warning(f"Nenhum registro de '{tipo_consulta}' encontrado.")
        else:
            st.warning("Não foi possível carregar os dados ou o servidor do PNCP está instável.")

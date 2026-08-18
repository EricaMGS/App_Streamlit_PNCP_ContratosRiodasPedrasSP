import io
import pandas as pd
import requests
import streamlit as st
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

# Configuração da página responsiva
st.set_page_config(page_title="Portal PNCP - Rio das Pedras/SP", layout="wide")

st.markdown("""
    <style>
    .stMetric { background-color: #f0f2f6; padding: 15px; border-radius: 10px; }
    </style>
""", unsafe_allow_html=True)

st.title("🏛️ Contratações de Rio das Pedras/SP")
st.markdown("Consulta integrada de Editais, Atas e Contratos direto do Portal Nacional de Contratações Públicas (PNCP). Com relatórios em Excel e Word.")

# --- BARRA LATERAL DE CONFIGURAÇÕES ---
st.sidebar.header("Parâmetros da Consulta")

tipo_consulta = st.sidebar.selectbox(
    "Selecione o tipo de dado:",
    [
        "Contratos",
        "Atas de Registro de Preços",
        "Editais e Avisos de Contratação",
    ],
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

                # --- SEÇÃO DE DOWNLOADS (EXCEL E WORD) ---
                st.markdown("### 📥 Opções de Exportação de Relatórios")
                col_dl1, col_dl2 = st.columns(2)

                # 1. GERAÇÃO DO EXCEL (.xlsx)
                buffer_excel = io.BytesIO()
                with pd.ExcelWriter(buffer_excel, engine="xlsxwriter") as writer:
                    sheet_name_safe = tipo_consulta.replace(" ", "_")[:31]
                    df.to_excel(writer, index=False, sheet_name=sheet_name_safe)
                    workbook = writer.book
                    worksheet = writer.sheets[sheet_name_safe]
                    header_format = workbook.add_format({"bold": True, "bg_color": "#003366", "font_color": "white"})
                    for col_num, value in enumerate(df.columns.values):
                        worksheet.write(0, col_num, value, header_format)

                with col_dl1:
                    st.download_button(
                        label="📊 Baixar Planilha (.xlsx)",
                        data=buffer_excel.getvalue(),
                        file_name=f"{tipo_consulta.replace(' ', '_')}_Rio_Das_Pedras_{d_inicio_str}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

                # 2. GERAÇÃO DO RELATÓRIO ESTILO WORD (.docx)
                doc = Document()
                
                # Cabeçalho / Título do documento
                p_title = doc.add_paragraph()
                run_title = p_title.add_run(f"Relatório Executivo: {tipo_consulta}")
                run_title.bold = True
                run_title.font.size = Pt(18)
                run_title.font.color.rgb = RGBColor(0, 51, 102)

                doc.add_paragraph(f"Município: Prefeitura Municipal de Rio das Pedras / SP")
                doc.add_paragraph(f"Período consultado: {data_inicio.strftime('%d/%m/%Y')} até {data_fim.strftime('%d/%m/%Y')}")
                doc.add_paragraph(f"Total de registros encontrados: {len(df)}")
                if col_valor:
                    doc.add_paragraph(f"Valor Consolidado: R$ {total_valor:,.2f}")

                doc.add_heading("Resumo dos Dados", level=2)
                doc.add_paragraph("Abaixo estão listados os principais registros obtidos diretamente do Portal Nacional de Contratações Públicas (PNCP):")

                # Adicionando uma tabela com as primeiras linhas no Word (para o relatório não ficar pesado)
                amostra_df = df.head(20) # Primeiras 20 linhas como amostra executiva
                table = doc.add_table(rows=1, cols=min(3, len(df.columns)))
                hdr_cells = table.rows[0].cells
                
                # Seleciona algumas colunas principais para exibir na tabela do Word
                colunas_amostra = df.columns[:3]
                for i, col_name in enumerate(colunas_amostra):
                    hdr_cells[i].text = str(col_name)

                for _, row in amostra_df.iterrows():
                    row_cells = table.add_row().cells
                    for i, col_name in enumerate(colunas_amostra):
                        row_cells[i].text = str(row[col_name])[:50] # Corta textos longos para caber bem no Word

                doc.add_paragraph()
                doc.add_paragraph("Relatório gerado automaticamente através da aplicação web de transparência pública.")

                buffer_word = io.BytesIO()
                doc.save(buffer_word)
                buffer_word.seek(0)

                with col_dl2:
                    st.download_button(
                        label="📝 Baixar Relatório em Word (.docx)",
                        data=buffer_word.getvalue(),
                        file_name=f"Relatorio_{tipo_consulta.replace(' ', '_')}_Rio_Das_Pedras.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    )
            else:
                st.warning(f"Nenhum registro de '{tipo_consulta}' foi encontrado para o período selecionado.")
        else:
            st.warning("Não foi possível carregar os dados ou o servidor do PNCP está instável.")

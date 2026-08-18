import io
import pandas as pd
import requests
import streamlit as st

# Configuração da página
st.set_page_config(
    page_title="Contratos - Rio das Pedras/SP", layout="wide"
)

st.title("🏛️ Dashboard: Contratos - Rio das Pedras/SP")
st.markdown(
    "Consulta oficial filtrada diretamente do Portal Nacional de Contratações"
    " Públicas (PNCP)."
)

# Barra lateral para filtros de período
st.sidebar.header("Filtros")
data_inicio = st.sidebar.date_input(
    "Data Inicial", value=pd.to_datetime("2026-01-01")
)
data_fim = st.sidebar.date_input("Data Final", value=pd.to_datetime("2026-03-31"))

if st.sidebar.button("Gerar Relatório de Rio das Pedras"):
  with st.spinner("Buscando e filtrando dados da Prefeitura..."):
    # CNPJ oficial da Prefeitura Municipal de Rio das Pedras/SP
    cnpj_alvo = "44826840000183"

    d_inicio_str = data_inicio.strftime("%Y%m%d")
    d_fim_str = data_fim.strftime("%Y%m%d")

    url = "https://pncp.gov.br/api/consulta/v1/contratos"
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML,"
            " like Gecko) Chrome/120.0.0.0 Safari/537.36"
        ),
        "Accept": "application/json",
    }

    all_data = []
    pagina = 1

    # Loop de paginação seguro
    while pagina <= 5:
      params = {
          "cnpj": cnpj_alvo,
          "dataInicial": d_inicio_str,
          "dataFinal": d_fim_str,
          "pagina": pagina,
      }

      try:
        response = requests.get(url, params=params, headers=headers, timeout=30)
        if response.status_code == 200:
          lote = response.json()
          lote = (
              lote.get("data", lote.get("items", []))
              if isinstance(lote, dict)
              else lote
          )

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

      # --- FILTRAGEM DE SEGURANÇA ---
      # Garante que apenas registros contendo o CNPJ de Rio das Pedras fiquem no relatório
      if "orgaoEntidade" in df.columns:
        df = df[
            df["orgaoEntidade"].astype(str).str.contains(cnpj_alvo, na=False)
        ]

      if not df.empty:
        # Métricas na tela
        col1, col2 = st.columns(2)
        col1.metric(
            "Contratos da Prefeitura de Rio das Pedras", len(df)
        )

        if "valorGlobal" in df.columns:
          total_valor = pd.to_numeric(df["valorGlobal"], errors="coerce").sum()
          col2.metric("Valor Global Somado", f"R$ {total_valor:,.2f}")

        # Exibição da tabela limpa
        st.subheader("📋 Processos e Contratos Encontrados")
        st.dataframe(df, use_container_width=True)

        # Geração do arquivo Excel formatado (.xlsx)
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
          df.to_excel(writer, index=False, sheet_name="Rio_das_Pedras")
          workbook = writer.book
          worksheet = writer.sheets["Rio_das_Pedras"]
          header_format = workbook.add_format(
              {"bold": True, "bg_color": "#003366", "font_color": "white"}
          )
          for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)

        # Botão de download
        st.download_button(
            label="📥 Baixar Relatório de Rio das Pedras (.xlsx)",
            data=buffer.getvalue(),
            file_name=(
                f"Contratos_Rio_Das_Pedras_{d_inicio_str}_a_{d_fim_str}.xlsx"
            ),
            mime=(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
        )
      else:
        st.warning(
            "Nenhum contrato da Prefeitura de Rio das Pedras foi encontrado"
            " para o período selecionado."
        )
    else:
      st.warning(
          "Não foi possível obter dados da API do PNCP neste momento."
          " Tente novamente em instantes."
      )

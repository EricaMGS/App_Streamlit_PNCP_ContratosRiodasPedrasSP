import io
import pandas as pd
import requests
import streamlit as st

# Configuração da página para ser responsiva e profissional
st.set_page_config(
    page_title="Relatório PNCP - Rio das Pedras", layout="wide"
)

# Estilização CSS para deixar os cards e botões mais elegantes
st.markdown(
    """
    <style>
    .stMetric { background-color: #f0f2f6; padding: 15px; border-radius: 10px; }
    </style>
""",
    unsafe_allow_html=True,
)

st.title("🏛️ Dashboard: Contratos Rio das Pedras/SP")
st.markdown(
    "Consulta oficial integrada ao Portal Nacional de Contratações Públicas"
    " (PNCP)."
)

# Barra lateral para filtros
st.sidebar.header("Configurações do Relatório")
data_inicio = st.sidebar.date_input(
    "Data Inicial", value=pd.to_datetime("2026-01-01")
)
data_fim = st.sidebar.date_input(
    "Data Final", value=pd.to_datetime("2026-03-31")
)

if st.sidebar.button("Gerar Dashboard Completo"):
  with st.spinner(
      "Consultando dados no PNCP (isso pode levar alguns segundos)..."
  ):
    # CNPJ da prefeitura
    cnpj = "44826840000183"
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

    # Loop de Paginação com tratamento para evitar travamentos
    while pagina <= 10:  # Limite de segurança de páginas para evitar loop infinito
      params = {
          "cnpj": cnpj,
          "dataInicial": d_inicio_str,
          "dataFinal": d_fim_str,
          "pagina": pagina,
      }

      try:
        # Timeout aumentado para 45 segundos para dar tempo do servidor responder
        response = requests.get(url, params=params, headers=headers, timeout=45)

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
      except requests.exceptions.ReadTimeout:
        # Se uma página específica der timeout, encerra a busca parcial e exibe o que já foi coletado
        break
      except Exception:
        break

    if all_data:
      df = pd.DataFrame(all_data)

      # --- DASHBOARD RESPONSIVO ---
      c1, c2 = st.columns(2)
      c1.metric("Total de Contratos Carregados", len(df))

      # Formatação do Excel
      buffer = io.BytesIO()
      with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Contratos")
        workbook = writer.book
        worksheet = writer.sheets["Contratos"]
        header_format = workbook.add_format(
            {"bold": True, "bg_color": "#003366", "font_color": "white"}
        )
        for col_num, value in enumerate(df.columns.values):
          worksheet.write(0, col_num, value, header_format)

      # Exibição da tabela responsiva
      st.subheader("Lista de Processos")
      st.dataframe(df, use_container_width=True)

      # Download button
      st.download_button(
          label="📥 Baixar Relatório Profissional (.xlsx)",
          data=buffer.getvalue(),
          file_name=f"Contratos_Rio_das_Pedras_{d_inicio_str}.xlsx",
          mime=(
              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
          ),
      )
    else:
      st.warning(
          "O servidor do PNCP demorou muito para responder ou nenhum contrato"
          " foi encontrado para o período. Tente reduzir o intervalo de datas"
          " (por exemplo, buscando mês a mês)."
      )

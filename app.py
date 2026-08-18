import io
import pandas as pd
import requests
import streamlit as st

# Configuração da página responsiva
st.set_page_config(
    page_title="Editais e Avisos - Rio das Pedras/SP", layout="wide"
)

st.markdown(
    """
    <style>
    .stMetric { background-color: #f0f2f6; padding: 15px; border-radius: 10px; }
    </style>
""",
    unsafe_allow_html=True,
)

st.title("📢 Editais e Avisos de Contratação - Rio das Pedras/SP")
st.markdown(
    "Painel oficial para consulta de editais, avisos e licitações publicadas"
    " no PNCP."
)

# Filtros na Barra Lateral
st.sidebar.header("Filtros de Período")
data_inicio = st.sidebar.date_input(
    "Data Inicial", value=pd.to_datetime("2025-01-01")
)
data_fim = st.sidebar.date_input(
    "Data Final", value=pd.to_datetime("2026-12-31")
)

if st.sidebar.button("Consultar Editais e Avisos"):
  with st.spinner("Buscando editais e avisos no PNCP..."):
    cnpj_prefeitura = "44826840000183"

    d_inicio_str = data_inicio.strftime("%Y%m%d")
    d_fim_str = data_fim.strftime("%Y%m%d")

    # Endpoint oficial de contratações/publicações (Editais e Avisos)
    url = "https://pncp.gov.br/api/consulta/v1/contratacoes/publicacao"

    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML,"
            " like Gecko) Chrome/120.0.0.0 Safari/537.36"
        ),
        "Accept": "application/json",
    }

    all_data = []
    pagina = 1

    # Loop de paginação para capturar todas as publicações do período
    while pagina <= 10:
      params = {
          "cnpj": cnpj_prefeitura,
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

      # Filtro de segurança por CNPJ
      if "orgaoEntidade" in df.columns:
        df = df[
            df["orgaoEntidade"].astype(str).str.contains(cnpj_prefeitura, na=False)
        ]

      if not df.empty:
        # Métricas Responsivas
        c1, c2 = st.columns(2)
        c1.metric("Total de Editais/Avisos Encontrados", len(df))

        if "valorTotalEstimado" in df.columns:
          total_est = pd.to_numeric(
              df["valorTotalEstimado"], errors="coerce"
          ).sum()
          c2.metric("Valor Total Estimado", f"R$ {total_est:,.2f}")
        else:
          c2.metric("Município", "Rio das Pedras/SP")

        # Exibição da tabela responsiva
        st.subheader("📋 Relação de Editais e Avisos Publicados")
        st.dataframe(df, use_container_width=True)

        # Geração do arquivo Excel profissional (.xlsx)
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
          df.to_excel(writer, index=False, sheet_name="Editais_PNCP")
          workbook = writer.book
          worksheet = writer.sheets["Editais_PNCP"]
          header_format = workbook.add_format(
              {"bold": True, "bg_color": "#003366", "font_color": "white"}
          )
          for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)

        st.download_button(
            label="📥 Baixar Relatório de Editais (.xlsx)",
            data=buffer.getvalue(),
            file_name=(
                f"Editais_Rio_Das_Pedras_{d_inicio_str}_a_{d_fim_str}.xlsx"
            ),
            mime=(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
        )
      else:
        st.warning(
            "Nenhum edital ou aviso encontrado para o CNPJ e período"
            " selecionados."
        )
    else:
      st.warning(
          "Não foram encontrados registros ou o servidor do PNCP está"
          " temporariamente instável. Tente ampliar o período (ex: a partir de"
          " 2025)."
      )

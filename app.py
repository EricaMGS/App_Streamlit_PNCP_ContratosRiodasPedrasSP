import io
import pandas as pd
import requests
import streamlit as st

# Configuração da página responsiva
st.set_page_config(
    page_title="Portal PNCP - Rio das Pedras/SP", layout="wide"
)

st.markdown(
    """
    <style>
    .stMetric { background-color: #f0f2f6; padding: 15px; border-radius: 10px; }
    </style>
""",
    unsafe_allow_html=True,
)

st.title("🏛️ Portal de Transparência - Rio das Pedras/SP")
st.markdown(
    "Consulta integrada de Editais, Atas e Contratos direto do PNCP."
)

# --- BARRA LATERAL DE CONFIGURAÇÕES ---
st.sidebar.header("Parâmetros da Consulta")

# Seletor de Tipo de Documento
tipo_consulta = st.sidebar.selectbox(
    "Selecione o tipo de dado:",
    [
        "Contratos",
        "Atas de Registro de Preços",
        "Editais e Avisos de Contratação",
    ],
)

data_inicio = st.sidebar.date_input(
    "Data Inicial", value=pd.to_datetime("2025-01-01")
)
data_fim = st.sidebar.date_input(
    "Data Final", value=pd.to_datetime("2026-12-31")
)

# Mapeamento das URLs da API do PNCP para cada opção
endpoints = {
    "Contratos": "https://pncp.gov.br/api/consulta/v1/contratos",
    "Atas de Registro de Preços": "https://pncp.gov.br/api/consulta/v1/atas",
    "Editais e Avisos de Contratação": (
        "https://pncp.gov.br/api/consulta/v1/contratacoes/publicacao"
    ),
}

if st.sidebar.button("Gerar Relatório Consolidado"):
  url_escolhida = endpoints[tipo_consulta]

  with st.spinner(f"Buscando '{tipo_consulta}' no PNCP..."):
    cnpj_prefeitura = "44826840000183"
    d_inicio_str = data_inicio.strftime("%Y%m%d")
    d_fim_str = data_fim.strftime("%Y%m%d")

    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML,"
            " like Gecko) Chrome/120.0.0.0 Safari/537.36"
        ),
        "Accept": "application/json",
    }

    all_data = []
    pagina = 1

    # Loop de paginação robusto para capturar os registros do período
    while pagina <= 10:
      params = {
          "cnpj": cnpj_prefeitura,
          "dataInicial": d_inicio_str,
          "dataFinal": d_fim_str,
          "pagina": pagina,
      }

      try:
        response = requests.get(
            url_escolhida, params=params, headers=headers, timeout=30
        )

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

      # Filtro de segurança por CNPJ na coluna de órgão
      if "orgaoEntidade" in df.columns:
        df = df[
            df["orgaoEntidade"].astype(str).str.contains(
                cnpj_prefeitura, na=False
            )
        ]

      if not df.empty:
        # --- EXIBIÇÃO DO CONSOLIDADO NA TELA ---
        st.success(
            f"Consulta realizada com sucesso! Categoria: {tipo_consulta}"
        )

        c1, c2 = st.columns(2)
        c1.metric(f"Total de Registros ({tipo_consulta})", len(df))

        # Identifica dinamicamente colunas de valor para somar no card gerencial
        col_valor = None
        for col in [
            "valorGlobal",
            "valorTotalEstimado",
            "valorHomologado",
            "valorInicial",
        ]:
          if col in df.columns:
            col_valor = col
            break

        if col_valor:
          total_valor = pd.to_numeric(df[col_valor], errors="coerce").sum()
          c2.metric("Valor Consolidado", f"R$ {total_valor:,.2f}")
        else:
          c2.metric("Município", "Rio das Pedras/SP")

        # Tabela interativa responsiva
        st.subheader(f"📋 Relação Consolidada: {tipo_consulta}")
        st.dataframe(df, use_container_width=True)

        # --- GERAÇÃO DO EXCEL FORMATADO (.xlsx) ---
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
          sheet_name_safe = tipo_consulta.replace(" ", "_")[:31]
          df.to_excel(writer, index=False, sheet_name=sheet_name_safe)
          workbook = writer.book
          worksheet = writer.sheets[sheet_name_safe]
          header_format = workbook.add_format(
              {"bold": True, "bg_color": "#003366", "font_color": "white"}
          )
          for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)

        # Botão de Download na interface
        st.download_button(
            label=f"📥 Baixar Relatório de {tipo_consulta} (.xlsx)",
            data=buffer.getvalue(),
            file_name=(
                f"{sheet_name_safe}_Rio_Das_Pedras_{d_inicio_str}_a_{d_fim_str}.xlsx"
            ),
            mime=(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
        )
      else:
        st.warning(
            f"Nenhum registro de '{tipo_consulta}' foi encontrado para o"
            " período selecionado."
        )
    else:
      st.warning(
          "Não foi possível carregar os dados ou o servidor do PNCP está"
          " instável. Tente novamente."
      )

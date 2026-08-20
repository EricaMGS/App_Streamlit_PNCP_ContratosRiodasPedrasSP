import io
import datetime
import time
import pandas as pd
import requests
import streamlit as st
from docx import Document
from docx.shared import Pt, RGBColor
from fpdf import FPDF

# ============================================================
# CONFIGURAÇÃO E DADOS
# ============================================================

st.set_page_config(page_title="Portal PNCP - Rio das Pedras/SP", layout="wide")
st.title("🏛️ Contratações de Rio das Pedras/SP")
CNPJ_RIO_DAS_PEDRAS = "44826840000183"
BASE_URL = "https://pncp.gov.br/api/consulta/v1"

# [ ... FUNÇÕES DE CONSULTA (consultar_pncp, extrair_registros, consultar_paginas) MANTIDAS IGUAIS ... ]
# (Inserir aqui as mesmas funções do código anterior)

def consultar_detalhes_contrato(id_contrato):
    """Busca documentos vinculados a um contrato específico (Aditivos/Apostilamentos)."""
    url = f"{BASE_URL}/contratos/{id_contrato}/documentos"
    try:
        resp = requests.get(url, timeout=10)
        if resp.status_code == 200:
            return resp.json()
        return []
    except:
        return []

# ============================================================
# INTERFACE E LÓGICA DE CONSULTA
# ============================================================
# [ ... (Manter toda a lógica de sidebar e busca inicial) ... ]

if st.session_state.df_resultado is not None and not st.session_state.df_resultado.empty:
    df = st.session_state.df_resultado
    
    # Exibir Métricas e Gráficos (como no código anterior)
    
    # ------------------------------------------------------------
    # NOVO: SELETOR PARA VER ADITIVOS
    # ------------------------------------------------------------
    st.markdown("---")
    st.markdown("### 🔍 Consultar Aditivos por Contrato")
    
    # Criar uma lista de seleção com base nos IDs disponíveis na busca
    lista_contratos = df.apply(lambda x: f"{x.get('numeroControlePNCP')} - Proc: {x.get('processo')}", axis=1).tolist()
    contrato_selecionado = st.selectbox("Selecione um contrato para verificar aditivos:", lista_contratos)
    
    if st.button("Buscar Aditivos do Contrato"):
        id_escolhido = contrato_selecionado.split(" - ")[0]
        with st.spinner("Buscando aditivos no PNCP..."):
            aditivos = consultar_detalhes_contrato(id_escolhido)
            
            if aditivos:
                st.write(f"Encontrados {len(aditivos)} documentos vinculados:")
                for doc in aditivos:
                    # Filtra apenas aditivos (geralmente tipo 2 ou 3 no PNCP)
                    tipo_doc = doc.get('tipoDocumentoNome', 'Outro')
                    st.info(f"**Tipo:** {tipo_doc} | **Data:** {doc.get('dataPublicacao', 'N/D')}\n\n{doc.get('objeto', '')}")
            else:
                st.warning("Nenhum documento/aditivo encontrado para este contrato.")

    st.markdown("---")
    st.dataframe(df, use_container_width=True, hide_index=True)

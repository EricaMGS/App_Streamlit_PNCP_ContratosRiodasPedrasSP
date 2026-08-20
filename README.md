
Portal PNCP - Rio das Pedras/SP
Este é um projeto de análise de dados governamentais desenvolvido para o município de Rio das Pedras/SP, permitindo a consulta integrada, análise e exportação de Contratos, Atas de Registro de Preços e Editais diretamente do Portal Nacional de Contratações Públicas (PNCP).

🚀 Funcionalidades
Consulta Direta à API: Integração em tempo real com a API oficial do PNCP.
Dashboards Interativos: Visualização de métricas (KPIs) e gráficos de evolução temporal das contratações.
Gestão de Aditivos: Ferramenta dedicada para consultar documentos vinculados (Aditivos e Apostilamentos) de cada contrato selecionado.
Exportação de Relatórios: Geração automática de documentos para auxiliar a fiscalização e controle:
📊 Excel (.xlsx)
📄 CSV (.csv)
📝 Word (.docx) - Formatado com detalhes executivos.
📕 PDF (.pdf) - Relatório limpo e profissional.
Robustez: Implementação de retentativas automáticas (retry logic) para lidar com instabilidades da API e timeouts.

🛠️ Tecnologias Utilizadas
Python: Linguagem base para processamento e lógica.
Streamlit: Framework para criação do dashboard web interativo.
Pandas: Manipulação e tratamento de dataframes.
Requests: Consumo da API do PNCP.
FPDF & python-docx: Geração de relatórios em PDF e Word.

📋 Como executar localmente
Clone o repositório:
Bash
git clone https://github.com/EricaMGS/contratacoespncpriodaspedrassp.git
cd seu-repositorio


Instale as dependências:
Bash
pip install -r requirements.txt

(Caso não tenha o arquivo, instale: pip install streamlit pandas requests fpdf python-docx openpyxl)
Execute o aplicativo:
Bash
streamlit run app.py

🏗️ Estrutura do Código
app.py: Arquivo principal contendo a lógica de interface, integração com API e geração de relatórios.
consultar_pncp(): Função robusta com controle de erros e retentativas para evitar falhas durante a carga de dados.
consultar_detalhes_contrato(): Módulo específico para buscar aditivos vinculados aos contratos.

⚖️ Aviso Legal
Este software é uma ferramenta de auxílio à transparência pública. A precisão dos dados depende exclusivamente da alimentação e disponibilidade da API do Portal Nacional de Contratações Públicas (PNCP).


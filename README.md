# Rubber Guardian - Controller EPI 🛡️

![Status](https://img.shields.io/badge/Status-Em%20Desenvolvimento-green)
![Python](https://img.shields.io/badge/Python-3.10+-blue.svg)
![License](https://img.shields.io/badge/License-MIT-yellow.svg)

Software de desktop para gestão e controle de Equipamentos de Proteção Individual (EPIs), desenvolvido para simplificar o rastreamento de entregas, a gestão de estoque e a análise de consumo em ambientes corporativos.

O sistema utiliza uma interface gráfica moderna e intuitiva construída com **CustomTkinter** e armazena todos os dados em uma **planilha Excel**, facilitando o acesso e a portabilidade para pequenas e médias operações.


## ✨ Funcionalidades Principais

- **📝 Lançamento de Entregas:** Registre de forma rápida e segura a entrega de EPIs aos funcionários, com autocomplete inteligente de nomes e equipamentos.
- **📦 Gestão de Estoque:** Realize entradas de novos lotes de EPIs ou faça ajustes manuais no inventário com facilidade.
- **📋 Cadastro Centralizado de EPIs:** Gerencie os tipos de EPIs, seus Certificados de Aprovação (C.A.) e preços em um único local.
- **📊 Dashboard Analítico e Interativo:**
  - Visualize o consumo de EPIs por funcionário, frequência de retirada ou por tipo de EPI.
  - Filtre os dados por período, tipo de visualização (Quantidade, Valor ou Ambos) e por EPI específico.
  - Funcionalidade de **Drill-Down**: clique duas vezes em uma barra do gráfico para explorar dados mais detalhados (ex: consumo mensal de um funcionário).
  - Retorne à visualização anterior com a tecla `ESC`.
- **🧾 Inventário Completo:** Visualize todos os registros de movimentação (entradas e saídas) com filtros poderosos e cálculo de saldo em tempo real.
- **✒️ Edição e Remoção de Registros:** Edite ou remova qualquer registro diretamente pela interface, seja um lançamento de entrega ou um EPI cadastrado.
- **║█║ Barcode Generator:** Gere e salve códigos de barras (Code128) para os nomes dos EPIs, facilitando a integração com sistemas de leitura óptica.

## 🛠️ Tecnologias Utilizadas

- **Linguagem:** Python 3
- **Interface Gráfica (GUI):** [CustomTkinter](https://github.com/TomSchimansky/CustomTkinter)
- **Análise de Dados e Gráficos:**
  - [Pandas](https://pandas.pydata.org/)
  - [Matplotlib](https://matplotlib.org/)
- **Interação com Excel:** [OpenPyXL](https://openpyxl.readthedocs.io/en/stable/)
- **Geração de Código de Barras:** [python-barcode](https://pypi.org/project/python-barcode/)
- **Calendário:** [tkcalendar](https://pypi.org/project/tkcalendar/)

## 🗄️ Banco de Dados

Este software utiliza uma única planilha Excel (`.xlsx`) como banco de dados, localizada em um diretório de sua escolha (preferencialmente em um serviço de nuvem como Google Drive ou OneDrive para backup automático). A planilha é estruturada em duas abas:

1.  `CONTROLE EPI`: Armazena todos os registros de movimentação (entradas e saídas).
2.  `CADASTRO EPI`: Funciona como uma tabela de referência para os tipos de EPIs, seus C.A. e preços unitários.

## 🚀 Instalação e Configuração

Siga os passos abaixo para executar o software em sua máquina.

**1. Clone o Repositório**
```bash
git clone https://github.com/seu-usuario/seu-repositorio.git
cd seu-repositorio
Use code with caution.
Markdown
2. Crie um Ambiente Virtual (Recomendado)
Generated bash
python -m venv venv
source venv/bin/activate  # No Windows: venv\Scripts\activate
Use code with caution.
Bash
3. Instale as Dependências
Crie um arquivo requirements.txt com o seguinte conteúdo:
Generated txt
customtkinter
pandas
openpyxl
matplotlib
tkcalendar
python-barcode
Pillow
Use code with caution.
Txt
Em seguida, instale as dependências:
Generated bash
pip install -r requirements.txt
Use code with caution.
Bash
4. Configure os Caminhos
Esta é a etapa mais importante. Abra o arquivo .pyw e altere as constantes BASE_DIR e LOGO_PATH para os caminhos corretos em sua máquina.
Generated python
# --- Caminhos ---
# IMPORTANTE: Altere este caminho para o diretório onde sua planilha será armazenada.
BASE_DIR = r"G:\Meu Drive\CONTROLLER\DATA CENTER" 
DB_PATH = os.path.join(BASE_DIR, "BANCO_DE_DADOS_EPI.xlsx")
# Opcional: Altere o caminho para o logo da sua empresa.
LOGO_PATH = os.path.join(BASE_DIR, "LOGO_RUBBERGATTI.png")
Use code with caution.
Python
5. Execute a Aplicação
Execute o arquivo principal para iniciar o software. A planilha BANCO_DE_DADOS_EPI.xlsx e suas abas serão criadas automaticamente no primeiro uso, se não existirem.
Generated bash
python "CONTROLE DE EPIS.pyw"
Use code with caution.
Bash
🗺️ Roadmap (Melhorias Futuras)
Migrar o banco de dados de Excel para SQLite para maior performance e segurança de dados.
Implementar um sistema de logging para registrar erros e eventos importantes em um arquivo de texto.
Adicionar uma funcionalidade de exportação de relatórios (PDF ou CSV) a partir do Dashboard e do Inventário.
Criar um executável (.exe) com PyInstaller para facilitar a distribuição.

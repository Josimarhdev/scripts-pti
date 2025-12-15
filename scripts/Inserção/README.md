# Ferramentas de Automação e Gestão de Dados (UVR)

Este repositório contém ferramentas desenvolvidas em Python para automatizar o fluxo de dados de Unidades de Valorização de Recicláveis (UVRs). 

O projeto visa eliminar lacunas, garantindo padronização na limpeza de dados, geração de relatórios em Excel e integração automática com APIs externas.

## Visão Geral

O sistema é dividido em duas ferramentas principais que atuam em etapas distintas do ciclo de vida dos dados:

### 1. Processamento e Padronização de Planilhas
Responsável por transformar dados brutos e desestruturados em relatórios padronizados.
- **Foco:** Limpeza de dados (Data Cleaning) e preenchimento de templates.
- **Destaque:** Utiliza lógica *Fuzzy* (aproximação de texto) para categorizar automaticamente despesas e materiais que possuem descrições variadas ou erros de digitação.
- **Saída:** Planilhas Excel (`.xlsx`) formatadas e prontas para validação.

### 2. Automação de Envio (API)
Responsável por converter os dados validados em formato digital e enviá-los para o banco de dados.
- **Foco:** Interoperabilidade e transmissão de dados.
- **Funcionalidade:** Lê planilhas Excel e tabelas CSV de referência, converte as informações para payloads JSON estruturados e realiza envios em lote via requisições HTTP (`POST`) com autenticação segura.

---

## Tecnologias Principais

O projeto é construído inteiramente em **Python 3.x**, utilizando as seguintes bibliotecas principais:

- **Manipulação de Dados:** `pandas`, `openpyxl`
- **Lógica e Processamento de Texto:** `fuzzywuzzy`
- **Conectividade e Web:** `requests`
- **Configuração:** `python-dotenv`

---

### Pré-requisitos Gerais

Certifique-se de ter o Python instalado. Recomenda-se o uso de um ambiente virtual (`venv`).

1. **Clone o repositório:**
   ```bash
   git clone 
   cd seu-repositorio
   ```

2. **Instale as dependências unificadas:**
   ```bash
   pip install pandas openpyxl fuzzywuzzy python-dotenv
   ```

### Fluxo de Trabalho Sugerido

1. Utilize o **planilha_lacunas** (gerador de planilhas) para sanear os dados brutos recebidos das UVRs.
2. Valide as planilhas geradas na pasta de saída.
4. Gere o JSON com **json_script**.
5. Utilize o **enviar_payloads** para processar essas planilhas validadas e submeter os dados à API.


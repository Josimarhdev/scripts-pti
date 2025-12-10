# Scripts

Este repositório centraliza um conjunto de ferramentas desenvolvidas em Python para a automação de processos, análise de dados e monitoramento de Unidades de Valorização de Recicláveis (UVRs) e gestão municipal.

O projeto atua como um hub de integração, cobrindo desde a extração de dados brutos, até a geração de relatórios gerenciais e envio de informações via API.

## 🚀 Visão Geral dos Módulos

O sistema é dividido em quatro frentes principais de atuação:

### 1. Automação de Monitoramento 
Responsável pela extração primária de dados. Este módulo conecta-se ao banco de dados, executa a extração e processa as regras de negócio para gerar as planilhas de controle oficial (Forms 1 a 4). Inclui validação cruzada de dados, gestão de abas mensais e atualização automática de status de envio.

### 2. Análise de Engajamento GRS
Um módulo focado em Business Intelligence que consome os dados do monitoramento para calcular KPIs de desempenho. O script cruza informações de múltiplos formulários para gerar uma "Nota de Engajamento" para cada município, produzindo relatórios visuais (Heatmaps) que classificam o envio de dados entre níveis Alto, Médio e Baixo.

### 3. Gerador de Indicadores Financeiros 
Ferramenta de normalização de dados desestruturados. Utiliza algoritmos de *Fuzzy Matching* (aproximação de texto) para ler descrições de despesas e receitas, mapeando-as automaticamente para templates padronizados. Essencial para transformar dados brutos de entrada em indicadores macro e microeconômicos consistentes.

### 4. Integração via API
Módulo final de ponte de dados. Automatiza a conversão de planilhas processadas (`.xlsx` e `.csv`) em payloads JSON estruturados e realiza o envio em lote para sistemas externos via requisições HTTP (POST) autenticadas.

## 🛠️ Stack Tecnológica

O ecossistema é construído inteiramente em **Python 3**, utilizando as seguintes tecnologias principais:

- **Manipulação de Dados:** `pandas`, `openpyxl`, `numpy`
- **Processamento de Texto:** `fuzzywuzzy`, `python-Levenshtein`
- **Conectividade:** `psycopg2` (PostgreSQL), `requests` (API REST)
- **Infraestrutura:** Gestão de variáveis de ambiente (`dotenv`)
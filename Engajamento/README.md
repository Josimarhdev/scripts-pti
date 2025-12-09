# Script de Análise de Engajamento GRS

Este projeto automatiza a verificação do engajamento de Municípios e Unidades de Valorização de Recicláveis (UVRs) no envio dos formulários de monitoramento. O script cruza dados de diferentes planilhas de controle, calcula percentuais de envio e gera um relatório visual com indicadores de desempenho.

## 📋 Funcionalidades

- **Leitura de Dados:** Importa dados de arquivos Excel de monitoramento (Forms 1, 2, 3 e 4).
- **Verificação de Status:** Checa se os formulários fixos (1, 2 e 3) foram enviados ou duplicados.
- **Contagem Mensal (Form 4):** Contabiliza envios mensais para 2024 e 2025, calculando a expectativa de envios baseada no mês atual (lógica dinâmica).
- **Cálculo de Engajamento:** Gera uma nota de engajamento baseada na razão entre *Envios Realizados* vs *Envios Esperados*.
- **Relatório Visual:** Gera uma planilha Excel (`analise_engajamento.xlsx`) formatada com cores condicionais para facilitar a leitura.

## 📂 Estrutura de Diretórios Necessária

O script utiliza caminhos relativos para localizar os arquivos de entrada e o utilitário de cores. A estrutura de pastas deve seguir o padrão abaixo:

    Projeto/
    ├── Monitoramento/
    │   ├── scripts/
    │   │   └── utils.py          # Contém o dicionário 'cores_regionais'
    │   └── outputs/
    │       └── GRS/
    │           ├── 0 - Monitoramento Form 1, 2 e 3.xlsx
    │           └── 0 - Monitoramento Form 4.xlsx
    │
    └── engajamento/              # Pasta onde este script reside
        ├── script_engajamento.py # (Seu arquivo atual)
        └── outputs/              # Onde o relatório final será salvo automaticamente
            └── analise_engajamento.xlsx

## 🛠️ Pré-requisitos

O script requer **Python 3** e as bibliotecas `pandas` e `openpyxl`.

Instalação das dependências via pip:

    pip install pandas openpyxl

## 🚀 Como Executar

1. Certifique-se de que os arquivos de entrada (`0 - Monitoramento...`) estejam na pasta correta (`../Monitoramento/outputs/GRS/`).
2. Execute o script via terminal dentro da pasta onde o arquivo `.py` está salvo:

    python nome_do_seu_script.py

3. O resultado será gerado na subpasta `outputs/` (criada automaticamente se não existir) dentro do diretório do script.

## 📊 Lógica de Cálculo do Engajamento

O nível de engajamento é definido pela porcentagem de formulários entregues em relação ao total esperado até a data atual.

### Definição de Níveis
- **Alto (Verde Escuro):** > 90% de envio.
- **Médio (Amarelo):** Entre 60% e 90% de envio.
- **Baixo (Vermelho):** < 60% de envio.

### Critérios de Contagem
- **Form 1, 2 e 3:** Conta 1 ponto se o status na planilha for "Enviado" ou "Duplicado".
- **Form 4:** Conta envios mensais acumulados nas abas correspondentes (ex: `01.24`, `05.25`).
  - **2024:** Expectativa fixa (considera meses específicos definidos no código).
  - **2025:** Expectativa dinâmica (aumenta conforme o mês atual avança).

## 🎨 Legenda de Cores na Saída

O arquivo Excel gerado aplica as seguintes formatações condicionais:

- **Regional (Coluna A):** Colore a célula com a cor oficial da regional (importada de `utils.py`).
- **Status de Envio (Forms 1, 2, 3):**
  - 🟢 **Verde Claro (C6EFCE):** Enviado.
  - 🔴 **Vermelho Claro (FFC7CE):** Ausente/Não Enviado.
- **Nível de Engajamento:**
  - 🟢 **Verde Escuro:** Alto.
  - 🟡 **Amarelo:** Médio.
  - 🔴 **Vermelho:** Baixo.

## ⚠️ Tratamento de Erros

O script possui validações automáticas para:
- Arquivos de entrada inexistentes ou caminhos incorretos (exibe alerta no console).
- Abas do Excel com nomes alterados.
- Leitura de abas mensais do Form 4 (ignora abas que não seguem o padrão de nome `MM.YY`).
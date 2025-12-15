# Automação de Monitoramento de Formulários

Este repositório contém o ecossistema de scripts para automação, validação e monitoramento de envios de formulários. O sistema atua desde a extração de dados diretamente do banco de dados até a geração de relatórios em Excel com formatação condicional avançada.

## Fluxo de Execução

O processo é orquestrado pelo script `EXECUTAR_TODOS.py` e opera em três etapas principais:

1.  **Extração (ETL)**: Conexão segura ao banco de dados (PostgreSQL) para execução de *queries* SQL. Os resultados são convertidos automaticamente para arquivos `.csv` na pasta `inputs`.
2.  **Processamento (Forms)**: Leitura dos CSVs e atualização das planilhas de monitoramento (Forms 1 a 4), aplicando regras de negócio, cálculo de atrasos, gestão de abas mensais e formatação visual.
3.  **Validação Cruzada**: Execução de scripts de auditoria que comparam os dados gerados com regras de validação visual (pintura de células, checagem de regionais e duplicatas).

## Estrutura do Repositório

- **`inputs/`**: Centraliza os dados de entrada.
    - Pastas das regionais (`0 - Belém`, `0 - Expansão`, `0 - GRS II`, `0 - Expansão MS`).
    - Arquivos `.sql` para consultas ao banco.
    - Arquivos `.csv` gerados automaticamente pela extração.
- **`outputs/`**: Destino dos arquivos processados e relatórios finais organizados por pasta.
- **`scripts/`**:
    - `EXECUTAR_TODOS.py`: Orquestrador geral.
    - `script_form[1-4].py`: Lógica individual de cada formulário.
    - `script_validacao.py`: Gera relatórios de auditoria visual.
    - `lib_validacao.py` e `utils.py`: Bibliotecas auxiliares de estilo, conexão e normalização.

## Configuração e Pré-requisitos

O projeto utiliza variáveis de ambiente para conexão com o banco de dados.

**1. Instale as dependências:**

```bash
pip install -r requirements.txt
```

**2. Configure o ambiente:**

Crie um arquivo `.env` na raiz do projeto (ou na pasta `scripts/`) com as seguintes credenciais:

```env
DB_NAME=nome_do_banco
DB_USER=usuario
DB_PASSWORD=senha
DB_HOST=host
DB_PORT=porta
```

## Regras de Negócio por Formulário

### Form 1
- **Escopo**: Município.
- **Lógica**: Verifica se houve envio contabilizando apenas o município. Atualiza as planilhas auxiliares (GRS, Expansão, Belém e Expansão MS).

### Form 2
- **Escopo**: Município + UVR.
- **Lógica**: Aprofunda a verificação cruzando o Município com o número da UVR, garantindo que cada unidade específica tenha um envio reportado.

### Form 3
- **Escopo**: Município + UVR + Empreendimento.
- **Lógica**: Cruza município com o número da UVR, também verificando se cada unidade específica teve um envio reportado.

### Form 4 
- **Escopo**: Município + UVR (Visualização Mensal).
- **Abas Dinâmicas**: O script cria ou atualiza automaticamente abas no formato `MM.AA` (ex: `01.25`).
- **Status Calculados**:
    - `Enviado`: Dados constam na base.
    - `Duplicado`: Mais de um envio para a mesma referência (o script concatena e exibe todas as datas).
    - `Atrasado`: Sem envio no mês anterior.
    - `Atrasado >= 2`: Sem envio há dois meses ou mais.
- **Gestão Inteligente de Irregulares**:
    - A aba "Irregulares" é recriada a cada execução, porém **preserva** validações manuais anteriores (colunas de resposta da TI ou Regional) cruzando chaves de identificação.
    - Novos envios irregulares são inseridos automaticamente no fim da lista.
- **Expansão MS**: Suporte total à regional de Mato Grosso do Sul.

## Script de Validação Visual

Após a geração dos formulários, o módulo de validação (`script_validacao.py`) gera relatórios de auditoria:

- **Verificação Visual**: Células inconsistentes são pintadas de vermelho automaticamente.
- **Mapeamento de Regionais**: Utiliza um dicionário interno para garantir que cidades estejam vinculadas à regional correta (ex: "Cascavel" -> "Valquiria"), corrigindo desvios na fonte de dados.

## Como Executar

1. Certifique-se de que o arquivo `.env` está configurado e as planilhas base estão na pasta `inputs/`.
2. Navegue até a pasta de scripts e execute o orquestrador:

```bash
cd scripts
python EXECUTAR_TODOS.py
```

O script irá conectar ao banco, baixar os dados, processar todos os formulários sequencialmente e salvar os resultados na pasta `outputs/`.
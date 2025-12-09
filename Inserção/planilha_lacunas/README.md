# Gerador de Planilhas de Indicadores (UVR)

Script em Python para automação e geração de planilhas de indicadores macro e microeconômicos para Unidades de Valorização de Recicláveis (UVRs). O sistema processa dados brutos, aplica limpeza de texto e utiliza lógica *fuzzy* (aproximação de texto) para preencher templates Excel padronizados.

## 📋 Funcionalidades

- **Geração em Lote:** Cria arquivos individuais por Município, UVR e Mês de Referência.
- **Tratamento de Strings:** Normalização de texto e remoção de caracteres especiais.
- **Fuzzy Matching:** Mapeamento inteligente de descrições de despesas, receitas e materiais que não correspondem exatamente ao template.
- **Mapa de Sinônimos:** Dicionário customizável para forçar correspondências de termos específicos (ex: "EPI", "PET Colorido").
- **Estrutura de Dados:** Separação automática entre dados operacionais, de manutenção e fundo de caixa.

## 🛠️ Dependências

O projeto utiliza Python 3.x e as seguintes bibliotecas:

- `pandas` (Manipulação de dados)
- `openpyxl` (Leitura e escrita de arquivos Excel)
- `fuzzywuzzy` (Comparação de strings e fuzzy logic)
- `python-Levenshtein` (Opcional, mas recomendado para performance do fuzzywuzzy)

Para instalar:

```bash
pip install pandas openpyxl fuzzywuzzy python-Levenshtein
```

## 📂 Estrutura de Arquivos

O script exige uma estrutura de diretórios específica para localizar os inputs. Certifique-se de organizar as pastas da seguinte maneira:

```text
/raiz_do_projeto
│
├── script.py                  # Este script principal
├── planilhas_geradas/         # (Criada automaticamente) Destino dos outputs
│
└── Inserção/
    └── planilha_lacunas/
        └── inputs/
            ├── dados.xlsx     # Fonte dos dados
            │                  # Abas necessárias: 'Micro Dados - Filtrados', 'Macro Dados - Filtrados'
            │
            └── template.xlsx  # Modelo Excel para preenchimento
                               # Abas necessárias: 'Receitas', 'Despesas', 'Materiais', 'Macro Dados'
```

## ⚙️ Configuração

No início do código (`script.py`), você pode ajustar as seguintes constantes globais:

- **`DEBUG_MODE`**: Defina como `True` para visualizar no console os detalhes de cada correspondência feita pelo algoritmo *fuzzy*. Útil para entender por que um item foi classificado de determinada forma.
- **`MATCH_THRESHOLD`**: Limiar de similaridade (inteiro de 0 a 100). O padrão é `70`. Itens com similaridade abaixo deste valor serão categorizados como "Outros".
- **`SYNONYM_MAP`**: Dicionário para correções manuais. Adicione termos aqui quando a correspondência automática falhar consistentemente.

## 🚀 Como Executar

1. Verifique se os arquivos `dados.xlsx` e `template.xlsx` estão na pasta `Inserção/planilha_lacunas/inputs`.
2. Execute o script via terminal:
   ```bash
   python script.py
   ```
3. Acompanhe o log no terminal para ver quais arquivos estão sendo gerados.
4. As planilhas preenchidas estarão na pasta `planilhas_geradas`.

## 🧠 Lógica de Preenchimento

### 1. Macro Dados
Preenche dados cadastrais (Município, UVR, Data), número de catadores, renda média e totais calculados.
- **Cálculo:** Separa o total de despesas em três categorias: Operação (Geral), Manutenção (termos como "conserto", "reforma") e Fundo de Caixa.

### 2. Receitas
Agrupa receitas por subtipo e utiliza *Fuzzy Matching* para encontrar a coluna correspondente no template `Receitas`.
- **Regra:** Se a confiança do match for > 80, aloca na coluna específica. Caso contrário, ou se não houver match, soma em "Outro Tipo".

### 3. Despesas
Analisa registros marcados como "Despesa".
- **Regra:** Compara a descrição da despesa com os nomes das linhas do template `Despesas`. Se o *score* for maior que o `MATCH_THRESHOLD` (70), preenche a linha correspondente. Caso contrário, soma em "Outras Despesas de Operação".

### 4. Materiais
Analisa registros de "Receita da venda de recicláveis".
- **Regra:** Tenta mapear o nome do material (ex: "Papelão Ondulado") para o subtipo no template `Materiais`.
- **Fallback:** Se não houver correspondência alta, tenta identificar a categoria do material e alocá-lo em uma categoria genérica (ex: "Outro Plástico", "Outro Papel").


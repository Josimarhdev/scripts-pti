# Projeto de Automação de Envio de Formulários

Este projeto automatiza o processo de geração e envio de formulários para uma API, com base em dados de planilhas Excel (.xlsx) e arquivos CSV de referência.

O fluxo de trabalho é dividido em duas etapas principais:

1. Geração de Payload – Leitura de arquivos `.xlsx` e `.csv`, processamento dos dados e criação de arquivos `.json` formatados.  
2. Envio de Payload – Leitura dos `.json` gerados e envio (`POST`) para uma API com autenticação via Bearer Token.

---

## Requisitos

- Python 3.7+
- Bibliotecas listadas em `requirements.txt`

---

## Instalação

1. Clone este repositório (ou tenha todos os arquivos na mesma pasta).  
2. Instale as dependências necessárias executando o comando no terminal:

   ```bash
   pip install -r requirements.txt
   ```

---

## Estrutura de Pastas

A estrutura de diretórios deve seguir o padrão abaixo:

```
/projeto_automacao/
│
├── .env                 
├── json_script.py    
├── enviar_payloads.py   
├── requirements.txt     
│
├── inputs/              
│   │
│   │ 
│   ├── Cafelandia_UVR-1_12-2024.xlsx
│   ├── Foz_do_Iguacu_UVR-6_01-2025.xlsx
│   ├── Marechal_Candido_Rondon_UVR-2_12-2024.xlsx
│   │
│   │ 
│   ├── frd_categoria_despesa.csv
│   ├── frd_categoria_servico.csv
│   ├── frd_tipo_despesa.csv
│   ├── geral_subtipo_reciclavel.csv
│   └── geral_uvrs.csv
│
└── outputs/         
    │
    ├── payload_Cafelandia_UVR-1_12-2024.json
    ├── payload_Foz_do_Iguacu_UVR-6_01-2025.json
    └── ...
```

---

## Configuração (.env)

Antes de enviar os dados (Passo 2), é necessário criar um arquivo chamado `.env` na pasta principal do projeto.  
Este arquivo armazena suas credenciais da API de forma segura.

### Exemplo (produção):

```
url=http://localhost:8000/form4/create
token=SEU_TOKEN_SECRETO_VEM_AQUI
```

### Exemplo (teste com httpbin.org):

```
url=https://httpbin.org/post
token=TOKEN_DE_TESTE_12345
```

---

## Como Usar

O processo é dividido em dois passos simples:

---

### Passo 1: Gerar os Payloads JSON

Execute o script `json_script.py`.

Este script irá:

- Ler todos os arquivos `.xlsx` da pasta `inputs/dados`.
- Ler todos os arquivos `.csv` de referência da pasta `inputs/tabelas_referencia`.
- Processar os dados e gerar os arquivos `payload_*.json` na pasta `outputs/`.

No terminal, execute:

```bash
python3 json_script.py
```

---

### Passo 2: Enviar os Payloads para a API

Depois de gerar os JSONs e configurar o arquivo `.env`, execute o script `enviar_payloads.py`.

Este script irá:

- Ler a `url` e o `token` do arquivo `.env`.
- Ler cada arquivo `.json` da pasta `outputs/`.
- Fazer um POST para a API com o conteúdo do JSON e o Bearer Token de autenticação.

No terminal, execute:

```bash
python3 enviar_payloads.py
```

---

## Exemplo de Execução

```bash
$ python3 json_script.py
> Gerando payloads a partir das planilhas...

$ python3 enviar_payloads.py
> Enviando payloads para API...
> Payload payload_Cafelandia_UVR-1_12-2024.json enviado com sucesso!
```

---

## requirements.txt (exemplo)

O arquivo `requirements.txt` deve conter as bibliotecas abaixo (ajuste conforme necessário):

```
pandas
openpyxl
python-dotenv
requests
```

---

## Tecnologias Utilizadas

- Python 3.7+
- pandas – Manipulação de dados Excel/CSV  
- openpyxl – Leitura de arquivos `.xlsx`  
- requests – Envio de requisições HTTP  
- dotenv – Leitura de variáveis de ambiente  

---



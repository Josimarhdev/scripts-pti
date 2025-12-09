from openpyxl import Workbook
from pathlib import Path 
import os
import sys
import pandas as pd
from dotenv import load_dotenv 

# Tenta importar a função de exportação da nossa lib
try:
    from Monitoramento.scripts.lib_validacao import export_query_to_csv
except ImportError:
    # Fallback caso a estrutura de pastas esteja diferente
    from Monitoramento.scripts.lib_validacao import export_query_to_csv

# Carrega variáveis de ambiente
load_dotenv()

# --- CONFIGURAÇÃO INICIAL ---
pasta_scripts = Path(__file__).parent
pasta_inputs = pasta_scripts.parent / "inputs"
pasta_saida = pasta_scripts.parent / "outputs"

# Inicializa os workbooks globais
belem_wb = Workbook()
expansao_wb = Workbook()
grs_wb = Workbook()
expansao_ms_wb = Workbook()

# Remove aba padrão
for wb in [belem_wb, expansao_wb, grs_wb, expansao_ms_wb]:
    wb.remove(wb.active)

# Disponibiliza globais
globals().update({
    "belem_wb": belem_wb,
    "expansao_wb": expansao_wb,
    "grs_wb": grs_wb,
    "expansao_ms_wb": expansao_ms_wb
})

#ATUALIZAÇÃO DOS DADOS (BANCO -> CSV)

print("\n=== ETAPA 1: Atualizando Bases de Dados (Forms 1 a 4) ===")

db_vars = ["DB_NAME", "DB_USER", "DB_PASSWORD", "DB_HOST"]
tem_credenciais = all(os.getenv(var) for var in db_vars)

# Nome do arquivo SQL -> Nome do CSV que será gerado
mapa_queries = [
    ("form1.sql", "form1.csv"),
    ("form2.sql", "form2.csv"),
    ("form3.sql", "form3.csv"),
    ("form4.sql", "form4.csv")
]

if tem_credenciais:
    print("Credenciais encontradas. Iniciando extração dos dados...")
    try:
        # Loop para gerar os 4 CSVs baseados nos SQLs
        for sql_file, csv_file in mapa_queries:
            print(f"Processando: {sql_file} -> {csv_file}...")
            export_query_to_csv(sql_file, csv_file, pasta_inputs)
        
        print(">>> Bases atualizadas com sucesso.")

    except Exception as e:
        print(f"[ERRO CRÍTICO] Falha na conexão ou extração dos Forms: {e}")
        print("Tentando continuar com os CSVs existentes (se houver)...")
else:
    print("[AVISO] Sem credenciais de banco. Usando arquivos CSV locais antigos.")

# EXECUÇÃO DOS SCRIPTS DE FORMATAÇÃO

print("\n=== ETAPA 2: Gerando Planilhas de Monitoramento (Forms) ===")

# Verifica se os CSVs existem antes de rodar
for _, csv_file in mapa_queries:
    caminho_csv = pasta_inputs / csv_file
    if not caminho_csv.exists():
        print(f"[ALERTA] O arquivo {csv_file} não existe na pasta inputs!")

try:
    exec(open("scripts/script_form1.py").read())
    print("Form 1: OK")
    exec(open("scripts/script_form2.py").read())
    print("Form 2: OK")
    exec(open("scripts/script_form3.py").read())
    print("Form 3: OK")
    exec(open("scripts/script_form4.py").read())
    print("Form 4: OK")
except Exception as e:
    print(f"[ERRO] Falha na execução dos scripts de formulário: {e}")



# Validação

print("\n=== Executando Script de Validação ===")

if tem_credenciais:
    try:
        # Este script já tem sua própria lógica de buscar no banco (consulta_grs, etc)
        exec(open("scripts/script_validacao.py").read())
        print(">>> Validação finalizado.")
    except Exception as e:
        print(f"[ERRO] Falha no script de validação: {e}")
else:
    print("[PULADO] Sem credenciais para script de validação.")


# SALVAMENTO FINAL DOS WORKBOOKS

print("\n=== ETAPA 4: Salvando Arquivos Monitoramento ===")

mapa_saida = {
    belem_wb: "Belém",
    expansao_wb: "Expansão",
    grs_wb: "GRS",
    expansao_ms_wb: "Expansão MS"
}

for wb, nome_pasta in mapa_saida.items():
    caminho_final_pasta = pasta_saida / nome_pasta
    caminho_final_pasta.mkdir(parents=True, exist_ok=True)
    
    caminho_arquivo = caminho_final_pasta / "0 - Monitoramento Form 1, 2 e 3.xlsx"
    
    # Só salva se tiver algo na planilha 
    if len(wb.sheetnames) > 0:
        wb.save(caminho_arquivo)
        print(f"Salvo: {caminho_arquivo}")
    else:
        print(f"[AVISO] Workbook vazio para {nome_pasta}, não salvo.")

print("\nProcesso Geral Finalizado.")
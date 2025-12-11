import sys
from pathlib import Path
from datetime import datetime
import pandas as pd

try:
    from lib_validacao import export_query_to_csv, processar_e_salvar_excel
except ImportError:
    from Monitoramento.scripts.lib_validacao import export_query_to_csv, processar_e_salvar_excel

print(">>> Iniciando Validacao: Banco de Dados -> Excel")

pasta_scripts = Path(__file__).resolve().parent
pasta_root = pasta_scripts.parent
pasta_inputs = pasta_root / "inputs"
pasta_outputs = pasta_root / "outputs"

# Extração do Banco de Dados para a pasta INPUTS 

print(f"Buscando dados no banco e salvando em: {pasta_inputs}")

# GRS
export_query_to_csv("consulta_grs.sql", "data_grs.csv", pasta_inputs)

# Expansão
export_query_to_csv("consulta_expansao.sql", "data_expansao.csv", pasta_inputs)


# Geração dos Relatórios Excel na pasta OUTPUTS 
now = datetime.now()
timestamp = now.strftime("%d-%m-%Y %H-%M")

# Processamento GRS
csv_grs = pasta_inputs / "data_grs.csv"
if csv_grs.exists():
    df_grs = pd.read_csv(csv_grs)
    
    # Define pasta de saída do GRS (Cria se não existir)
    output_grs = pasta_outputs / "GRS"
    output_grs.mkdir(parents=True, exist_ok=True)
    
    arquivo_final = output_grs / f"1 - Formulários - GRS - {timestamp}.xlsx"
    processar_e_salvar_excel(df_grs, arquivo_final)
else:
    print(f"[AVISO] {csv_grs.name} não encontrado. Pulando geração do relatório GRS.")

# Processamento Expansão
csv_exp = pasta_inputs / "data_expansao.csv"
if csv_exp.exists():
    df_exp = pd.read_csv(csv_exp)
    
    # Define pasta de saída da Expansão
    output_exp = pasta_outputs / "Expansão"
    output_exp.mkdir(parents=True, exist_ok=True)
    
    arquivo_final = output_exp / f"1 - Formulários - Expansão - {timestamp}.xlsx"
    processar_e_salvar_excel(df_exp, arquivo_final)
else:
    print(f"[AVISO] {csv_exp.name} não encontrado. Pulando geração do relatório Expansão.")

print(">>> Validação finalizado.")
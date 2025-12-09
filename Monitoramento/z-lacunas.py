# -*- coding: utf-8 -*-
import pandas as pd
from pathlib import Path
from datetime import datetime
from unidecode import unidecode
import re 



pasta_script = Path(__file__).resolve().parent

pasta_base = pasta_script
caminho_planilha_original = pasta_base / "outputs" / "grs_atualizado_form4.xlsx"
caminho_csv_lacunas = pasta_base / "inputs" / "lacunas.csv"
caminho_relatorio_final = pasta_base / "outputs" / "relatorio_lacunas_encontradas.xlsx"
meses_referencia_analise = ['11.24', '12.24', '01.25', '02.25', '03.25', '04.25']



def normalizar_municipio(texto):

    if not isinstance(texto, str):
        return ""
    # Converte para minúsculo e remove acentos
    texto = unidecode(texto.lower())
    # Remove "municipio de " se existir no início
    if texto.startswith('municipio de '):
        texto = texto[13:]
    # Remove caracteres especiais, mantendo apenas letras e números
    texto = re.sub(r'[^a-z0-9]', '', texto)
    return texto

def normalizar_uvr(uvr):
    """Extrai apenas os dígitos de um número de UVR para padronização."""
    if uvr is None:
        return ""
    # Converte para string e extrai apenas os dígitos
    uvr_str = str(uvr)
    numeros = re.sub(r'\D', '', uvr_str)
    # Remove zeros à esquerda para padronizar 
    if numeros:
        return str(int(numeros))
    return ""

def converter_data_para_mes_ano(data):
    """Converte datas para o formato MM.YY."""
    if pd.isna(data):
        return None
    try:
        return pd.to_datetime(data, dayfirst=True).strftime('%m.%y')
    except (ValueError, TypeError):
        return None



print(f"Analisando o arquivo: {caminho_planilha_original}")
lista_de_lacunas = []
try:
    excel_file = pd.ExcelFile(caminho_planilha_original)
except FileNotFoundError:
    print(f"ERRO: O arquivo '{caminho_planilha_original}' não foi encontrado.")
    exit()

for mes_ano in meses_referencia_analise:
    print(f"  -> Verificando a aba '{mes_ano}'...")
    if mes_ano in excel_file.sheet_names:
        df_mes = pd.read_excel(excel_file, sheet_name=mes_ano)
        lacunas_mes = df_mes[
            df_mes['Data de Envio'].isna() &
            ~df_mes['Situação'].isin(['Sem Técnico', 'Outras Ocorrências'])
        ]
        if not lacunas_mes.empty:
            for _, row in lacunas_mes.iterrows():
                lista_de_lacunas.append({
                    'Regional': row.get('Regional'),
                    'Município': row.get('Município'),
                    'UVR': row.get('UVR'),
                    'Mês da Lacuna': mes_ano
                })
    else:
        print(f"  -> AVISO: A aba '{mes_ano}' não foi encontrada no arquivo. Pulando...")

if not lista_de_lacunas:
    print("Nenhuma lacuna de envio foi encontrada nos meses especificados. Encerrando o script.")
    exit()

df_lacunas_identificadas = pd.DataFrame(lista_de_lacunas)
print(f"\nTotal de {len(df_lacunas_identificadas)} lacunas identificadas na planilha original.")



print(f"\nLendo o arquivo de envios: {caminho_csv_lacunas}")
try:
    df_envios_csv = pd.read_csv(caminho_csv_lacunas, dtype=str)
except FileNotFoundError:
    print(f"ERRO: O arquivo '{caminho_csv_lacunas}' não foi encontrado.")
    exit()

df_envios_csv.rename(columns={
    'Município': 'municipio_csv',
    'UVR Número': 'uvr_csv',
    'Data de referência': 'data_ref_csv' # Ajuste o nome da coluna conforme seu arquivo
}, inplace=True)

df_envios_csv['mes_ano_csv'] = df_envios_csv['data_ref_csv'].apply(converter_data_para_mes_ano)
print("Processamento do arquivo CSV concluído.")

# --- CRUZAMENTO DE DADOS ---

print("\nCruzando dados com normalização avançada...")

# Cria a chave usando as novas funções de normalização
df_lacunas_identificadas['chave'] = (
    df_lacunas_identificadas['Município'].apply(normalizar_municipio) + "_" +
    df_lacunas_identificadas['UVR'].apply(normalizar_uvr) + "_" +
    df_lacunas_identificadas['Mês da Lacuna']
)
df_envios_csv['chave'] = (
    df_envios_csv['municipio_csv'].apply(normalizar_municipio) + "_" +
    df_envios_csv['uvr_csv'].apply(normalizar_uvr) + "_" +
    df_envios_csv['mes_ano_csv']
)

# Realiza a junção dos DataFrames
relatorio_final_df = pd.merge(
    df_lacunas_identificadas,
    df_envios_csv[['chave', 'data_ref_csv']],
    on='chave',
    how='left'
)

relatorio_final_df = relatorio_final_df.dropna(subset=['data_ref_csv'])

# --- GERAÇÃO DO RELATÓRIO FINAL ---

if relatorio_final_df.empty:
    print("\nNenhuma das lacunas identificadas foi encontrada no arquivo CSV. Nenhum relatório será gerado.")
else:
    print(f"\n{len(relatorio_final_df)} lacunas encontradas no arquivo CSV. Gerando relatório...")
    relatorio_final_df = relatorio_final_df[[
        'Regional', 'Município', 'UVR', 'Mês da Lacuna', 'data_ref_csv'
    ]].rename(columns={
        'data_ref_csv': 'Data de Referência CSV'
    })
    try:
        with pd.ExcelWriter(caminho_relatorio_final, engine='openpyxl') as writer:
            relatorio_final_df.to_excel(writer, index=False, sheet_name='Lacunas Encontradas')
            worksheet = writer.sheets['Lacunas Encontradas']
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                adjusted_width = (max_length + 2)
                worksheet.column_dimensions[column_letter].width = adjusted_width
        print(f"\nRelatório gerado com sucesso em: {caminho_relatorio_final}")
    except Exception as e:
        print(f"\nOcorreu um erro ao salvar o relatório: {e}")
import pandas as pd
import json
import re
import os
from datetime import datetime
import sys
import glob 

def safe_float(val):
  
    try:
        return float(str(val).replace(',', '.'))
    except (ValueError, TypeError):
        return 0.0

def format_date(date_str):
    
    if isinstance(date_str, datetime):
        return date_str.strftime('%Y-%m-01')
    
    match = re.match(r'(\d{1,2})/(\d{4})', str(date_str))
    if match:
        mes, ano = match.groups()

        return f"{ano}-{int(mes):02d}-01"
    
    try:
        return pd.to_datetime(date_str).strftime('%Y-%m-01')
    except:
        return datetime.now().strftime('%Y-01-01')

def create_payload(excel_file_path, ref_dir):

    # Extrair informações do nome do arquivo 
    base_filename = os.path.basename(excel_file_path)
    # Tenta extrair (Nome do Município)_UVR-(Número)_(Mês)-(Ano).xlsx
    match = re.match(r'^(.*?)_UVR-(\d+)_(\d{1,2})-(\d{4})\.xlsx$', base_filename, re.IGNORECASE)
    
    if not match:
        print(f"  [AVISO] Não foi possível analisar o nome do arquivo: {base_filename}. Tentando continuar...", file=sys.stderr)

        municipio = base_filename.split('_')[0].replace('.xlsx', '')
        uvr_num = "1" # Suposição padrão
    else:
        municipio, uvr_num, mes, ano = match.groups()

    try:
        uvrs_path = os.path.join(ref_dir, "geral_uvrs.csv")
        subtipo_path = os.path.join(ref_dir, "geral_subtipo_reciclavel.csv")
        tipo_despesa_path = os.path.join(ref_dir, "frd_tipo_despesa.csv")
        cat_despesa_path = os.path.join(ref_dir, "frd_categoria_despesa.csv")
        cat_servico_path = os.path.join(ref_dir, "frd_categoria_servico.csv")

        df_uvrs = pd.read_csv(uvrs_path, on_bad_lines='skip')
        
        df_subtipo = pd.read_csv(subtipo_path, header=0, on_bad_lines='skip') 
        
        df_tipo_despesa = pd.read_csv(tipo_despesa_path, on_bad_lines='skip')
        df_cat_despesa = pd.read_csv(cat_despesa_path, on_bad_lines='skip')
        df_cat_servico = pd.read_csv(cat_servico_path, on_bad_lines='skip')
        
    except Exception as e:
        print(f"  [ERRO] Erro ao carregar arquivos de referência CSV: {e}", file=sys.stderr)
        print(f"  Verifique se os arquivos de referência (geral_uvrs, etc.) estão na pasta 'inputs'.", file=sys.stderr)
        return None

    try:
      
        df_macro = pd.read_excel(excel_file_path, sheet_name='Macro Dados')
        df_materiais = pd.read_excel(excel_file_path, sheet_name='Materiais', header=1) # O cabeçalho começa na linha 2 (índice 1)
        df_receitas = pd.read_excel(excel_file_path, sheet_name='Receitas')
        df_despesas = pd.read_excel(excel_file_path, sheet_name='Despesas', header=1) # O cabeçalho começa na linha 2 (índice 1)

    except Exception as e:
        print(f"  [ERRO] Erro ao carregar abas do arquivo Excel '{base_filename}': {e}", file=sys.stderr)
        print("  Verifique se o arquivo não está corrompido e se as abas 'Macro Dados', 'Materiais', 'Receitas', e 'Despesas' existem.", file=sys.stderr)
        return None

    if df_macro.empty:
        print(f"  [ERRO] A aba 'Macro Dados' está vazia.", file=sys.stderr)
        return None

    macro_data = df_macro.iloc[0]

    fk_guvr_id = None
    try:
     
        uvr_num_formatado = f"UVR {int(uvr_num):02d}"

        municipio_formatado = str(municipio).replace('_', ' ')
        municipio_upper = municipio_formatado.upper() # Ex: "FOZ DO IGUAÇU"
        
        df_uvrs['guvr_nome_upper'] = df_uvrs['guvr_nome'].str.upper()
        
        # Tenta encontrar a correspondência exata primeiro
        
        resultado_uvr = df_uvrs[
            (df_uvrs['guvr_nome_upper'] == f"{municipio_upper} - {uvr_num_formatado}")
        ]
        
        # Se não achar, tenta uma busca mais flexível 
        if resultado_uvr.empty:
             resultado_uvr = df_uvrs[
                (df_uvrs['guvr_nome_upper'].str.contains(municipio_upper)) &
                (df_uvrs['guvr_nome_upper'].str.contains(uvr_num_formatado))
            ]

        if not resultado_uvr.empty:
            fk_guvr_id = int(resultado_uvr.iloc[0]['guvr_id'])
        else:
            print(f"  [AVISO] Não foi possível encontrar o 'fk_guvr_id' para {municipio} UVR {uvr_num}. Usando 'None'.", file=sys.stderr)

    except Exception as e:
        print(f"  [AVISO] Erro ao procurar 'fk_guvr_id': {e}. Usando 'None'.", file=sys.stderr)
        fk_guvr_id = None

    # Pega o texto original da observação
    obs_original = str(macro_data.get("Observações", ""))
        
    # Adiciona o prefixo
    obs_final = "formulário migrado do Reciclômetro v1: " + obs_original
    step1 = {
        "mcmr_observacoes": obs_final,
        "mcmr_renda": safe_float(macro_data.get("Renda Média", 0.0)),
        "mcmr_quantidade": int(safe_float(macro_data.get("Número de Catadores", 0))),
        "rd_mes_referencia": format_date(macro_data.get("Mês Referência", "")),
        "fk_guvr_id": fk_guvr_id
    }

    # Step 2: Informações Financeiras 
    step2 = {
        "if_receita_total_venda_reciclaveis": safe_float(macro_data.get("Receita Venda Recicláveis", 0.0)),
        "if_fundo_caixa_total": safe_float(macro_data.get("Fundo de Caixa", 0.0)),
        "if_despesa_total_operacao_uvr": safe_float(macro_data.get("Despesa Operação", 0.0)),
        "if_despesa_total_manutencao_uvr": safe_float(macro_data.get("Despesa Manutenção", 0.0))
    }

    #  Step 3: Fixo 
    step3 = {
        "vr_emicao_notas_ficais": False,
        "anf_nota_fiscal": ""
    }

    #  Step 4: Venda de Recicláveis 
    step4 = {"tipo_formulario": True, "frd_venda_reciclaveis": []}
    
    # Criar mapa de subtipos para consulta rápida
    # Limpa os nomes para correspondência: minúsculas e sem espaços
    df_subtipo['subtipo_clean'] = df_subtipo['gsr_subtipo'].str.strip().str.lower()
    mapa_subtipo = df_subtipo.set_index('subtipo_clean')['gsr_id'].to_dict()

    for _, row in df_materiais.iterrows():
        quantidade = safe_float(row.get('Quantidade'))
        valor = safe_float(row.get('Valor'))
        
        # Só adiciona se houver quantidade OU valor (qualquer um deles > 0)
        if quantidade > 0 or valor > 0:
            subtipo_nome = str(row.get('Subtipo', '')).strip().lower()
            
            # Encontra o ID do subtipo
            fk_gsr_id = mapa_subtipo.get(subtipo_nome)
            
            if fk_gsr_id:
                # Calcular preço médio
                preco_medio = (valor / quantidade) if quantidade > 0 else 0.0
                
                item_material = {
                    "fk_gsr_id": int(fk_gsr_id),
                    "fk_gtr_id": None,
                    "vr_quantidade": quantidade,
                    "vr_valor": valor,
                    "vr_emicao_notas_ficais": False,
                    "vr_preco_medio": preco_medio
                }
                step4["frd_venda_reciclaveis"].append(item_material)
            else:
                # Não reporta aviso se o subtipo estiver vazio (linhas de Categoria/Tipo)
                if subtipo_nome: 
                   print(f"  [AVISO] Subtipo de material não encontrado no mapeamento: '{row.get('Subtipo')}'", file=sys.stderr)

# Step 5: Receitas 
    step5 = []
    
    # Criar mapa de serviços
    # Garante que a coluna de comparação esteja limpa (minúsculas, sem espaços)
    df_cat_servico['categoria_clean'] = df_cat_servico['cs_categoria'].str.strip().str.lower()
    
    # Mapeamento manual dos nomes da aba 'Receitas' (colunas do Excel)
    # para os nomes do CSV 'frd_categoria_servico' (valores do banco)
    mapa_servicos_nomes = {
        'Prestação de Serviço': 'contratos de serviço de reciclagem (triagem)',
        'Logística Reversa': 'logistica reversa',
        'Convênios': 'convenios',
        'Termo de Cooperação': 'termo de cooperação',
        'Outro Tipo': 'outro'
    }
    
    mapa_servico_final = {}
    mapa_cat_servico = df_cat_servico.set_index('categoria_clean')['cs_id'].to_dict()

    # Mapeia os nomes das colunas do Excel para os IDs de serviço reais
    for col_excel, nome_csv in mapa_servicos_nomes.items():
        id_servico = mapa_cat_servico.get(nome_csv)
        if id_servico:
            mapa_servico_final[col_excel] = id_servico
        else:
             print(f"  [AVISO] Categoria de serviço não encontrada no mapeamento: '{nome_csv}'", file=sys.stderr)

    if df_receitas.empty:
        print("  [AVISO] Aba 'Receitas' está vazia.", file=sys.stderr)
    else:
        # Pega a primeira linha de dados da aba de receitas
        receitas_data = df_receitas.iloc[0]
        
        for col_nome, fk_cs_id in mapa_servico_final.items():
            # Verifica se a coluna existe na aba de receitas
            if col_nome in receitas_data:
                valor_receita = safe_float(receitas_data.get(col_nome, 0.0))
                
                # Adiciona ao step5 se o valor for positivo
                if valor_receita > 0:
                    fk_cs_id_int = int(fk_cs_id)
                    
                    # Se for Logística Reversa (ID 2), preenche os campos específicos
                    if fk_cs_id_int == 2:
                        ps_periodo_repasse = 1
                        ps_repasse_aquisicao_equipamentos = False
                    else:
                        ps_periodo_repasse = None
                        ps_repasse_aquisicao_equipamentos = None
                    
                    item_servico = {
                        "fk_cs_id": fk_cs_id_int,
                        "ps_periodo_repasse": ps_periodo_repasse,
                        "ps_repasse_aquisicao_equipamentos": ps_repasse_aquisicao_equipamentos,
                        "ps_valor": valor_receita 
                    }
                    step5.append(item_servico)
            else:
                 print(f"  [AVISO] Coluna de receita esperada não encontrada na aba 'Receitas': '{col_nome}'", file=sys.stderr)
            
# Step 6: Despesas 
    step6 = []
    
    # Criar mapa de tipos de despesa 
    df_tipo_despesa['tipo_clean'] = df_tipo_despesa['td_tipo'].str.strip().str.lower()
    mapa_tipo_despesa = df_tipo_despesa.set_index('tipo_clean')
    
    for _, row in df_despesas.iterrows():
        
        nome_despesa = str(row.get('Nome', '')).strip().lower()
        valor_despesa = safe_float(row.get('Valor'))
        
        if nome_despesa and valor_despesa > 0:
            # Encontra o ID do tipo de despesa
            if nome_despesa in mapa_tipo_despesa.index:
                tipo_despesa_info = mapa_tipo_despesa.loc[nome_despesa]
                fk_td_id = tipo_despesa_info['td_id']
                
                item_despesa = {
                    "de_valor": valor_despesa,
                    "fk_cd_id": 3, # Receita da UVR
                    "fk_td_id": int(fk_td_id)
                }
                step6.append(item_despesa)
            else:
                 print(f"  [AVISO] Tipo de despesa não mapeado: '{row.get('Nome')}'", file=sys.stderr)

    # Step 7: Rejeitos 
    step7 = {
        "bg_total_rejeitos": safe_float(macro_data.get("Rejeito", 0.0))
    }

    # Montagem Final 
    payload = {
        "step1": step1,
        "step2": step2,
        "step3": step3,
        "step4": step4,
        "step5": step5,
        "step6": step6,
        "step7": step7
    }
    
    return payload

#  EXECUÇÃO PRINCIPAL
if __name__ == "__main__":
    
    try:
        SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
    except NameError:
         SCRIPT_DIR = os.getcwd() 
    
    INPUT_DIR = os.path.join(SCRIPT_DIR, "inputs")

    OUTPUT_DIR = os.path.join(SCRIPT_DIR, "outputs")
    
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    

    # Procura por arquivos .xlsx na pasta de inputs
    search_pattern = os.path.join(INPUT_DIR, "*.xlsx")
    excel_files = glob.glob(search_pattern)
    

    excel_files = [f for f in excel_files if not os.path.basename(f).startswith('~$')]
    
    if not excel_files:
        print(f"Erro: Nenhum arquivo '.xlsx' encontrado em {INPUT_DIR}", file=sys.stderr)
        print("Por favor, verifique se os seus arquivos Excel (ex: Cafelândia_UVR-1_12-2024.xlsx) estão na pasta 'inputs'.", file=sys.stderr)
        sys.exit(1) # Sai do script se nenhum arquivo for encontrado

    print(f"Encontrados {len(excel_files)} arquivo(s) Excel para processar...")

    for excel_file_path in excel_files:
        
        base_filename = os.path.basename(excel_file_path)
        
        print(f"\n--- Processando arquivo: {base_filename} ---")
        
        final_json_data = create_payload(excel_file_path, INPUT_DIR) 
        
        if final_json_data:

            json_output = json.dumps(final_json_data, indent=4, ensure_ascii=False)
            
            print(f"  Payload gerado com sucesso para {base_filename}.")
            
            base_json_name = f"payload_{base_filename.replace('.xlsx', '')}.json"
            output_filename = os.path.join(OUTPUT_DIR, base_json_name)
            
            try:
                # Salva o arquivo JSON
                with open(output_filename, 'w', encoding='utf-8') as f:
                    f.write(json_output)
                print(f"  Salvo em: {output_filename}")
            except Exception as e:
                print(f"  [ERRO] Erro ao salvar arquivo JSON: {e}", file=sys.stderr)
        else:
            print(f"  [FALHA] Falha ao gerar o payload para {base_filename}. Verifique os erros acima.", file=sys.stderr)
    
    print("\nProcessamento concluído.")
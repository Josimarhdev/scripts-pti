import os
import glob
import json
import sys
import requests
from dotenv import load_dotenv

def enviar_payloads():
   
    #Carrega todos os arquivos .json da pasta outputs e faz um POST
    #para a API definida no arquivo .env
   
    
    # 1. Carregar Configurações do .env 
    try:
        SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
        
        env_path = os.path.join(SCRIPT_DIR, '.env')
        load_dotenv(dotenv_path=env_path)

        # Pega a URL e o Token
        API_URL = os.getenv("url")
        API_TOKEN = os.getenv("token")

        # Verifica se as variáveis foram carregadas
        if not API_URL or not API_TOKEN:
            print("[ERRO] As variáveis 'url' e 'token' não foram encontradas no arquivo .env.", file=sys.stderr)
            print("Verifique se o arquivo .env está na mesma pasta do script.", file=sys.stderr)
            return

    except Exception as e:
        print(f"[ERRO] Falha ao carregar o arquivo .env: {e}", file=sys.stderr)
        return

    # 2. Definir Cabeçalhos da Requisição 
    headers = {
        "Authorization": f"Bearer {API_TOKEN}",
        "Content-Type": "application/json"
    }

    # 3. Encontrar os Arquivos JSON para Enviar 
    # O script vai ler da pasta 'outputs', onde o 'gerar_payload.py' salva os arquivos.
    json_dir = os.path.join(SCRIPT_DIR, "outputs")
    search_pattern = os.path.join(json_dir, "payload_*.json")
    
    # Lista todos os arquivos JSON na pasta
    json_files = glob.glob(search_pattern)

    if not json_files:
        print(f"[AVISO] Nenhum arquivo .json encontrado em: {json_dir}", file=sys.stderr)
        print("Certifique-se de que o script 'gerar_payload.py' já foi executado.", file=sys.stderr)
        return

    print(f"Encontrados {len(json_files)} arquivos JSON para enviar para {API_URL}...")

    # 4. Loop de Envio (POST para cada arquivo) 
    sucessos = 0
    falhas = 0

    for json_path in json_files:
        filename = os.path.basename(json_path)
        print(f"\n--- Processando: {filename} ---")

        try:
            # Carrega o conteúdo do arquivo JSON
            with open(json_path, 'r', encoding='utf-8') as f:
                payload_data = json.load(f)

        except json.JSONDecodeError:
            print(f"  [FALHA] O arquivo '{filename}' está mal formatado (não é um JSON válido). Pulando.", file=sys.stderr)
            falhas += 1
            continue
        except Exception as e:
            print(f"  [FALHA] Erro ao ler o arquivo '{filename}': {e}. Pulando.", file=sys.stderr)
            falhas += 1
            continue
            
        # Tenta fazer o POST
        try:
            
            response = requests.post(API_URL, headers=headers, json=payload_data, timeout=15)

            # Verifica se o POST foi bem-sucedido 
            if 200 <= response.status_code < 300:
                print(f"  [SUCESSO] Status: {response.status_code}. Formulário enviado.")
                sucessos += 1
            else:
                # Se a API der um erro (4xx ou 5xx)
                print(f"  [FALHA] Status: {response.status_code}. A API retornou um erro.", file=sys.stderr)
                print(f"  Resposta da API: {response.text}", file=sys.stderr)
                falhas += 1

        except requests.exceptions.ConnectionError:
            print(f"  [FALHA] Erro de conexão. Não foi possível conectar a {API_URL}.", file=sys.stderr)
            print("  Verifique se o seu servidor (localhost:8000) está rodando.", file=sys.stderr)
            falhas += 1
        except requests.exceptions.Timeout:
            print(f"  [FALHA] A requisição demorou muito (timeout). Pulando.", file=sys.stderr)
            falhas += 1
        except Exception as e:
            print(f"  [FALHA] Ocorreu um erro inesperado durante o POST: {e}", file=sys.stderr)
            falhas += 1

    # 5. Resumo Final 
    print("\n--- Processamento Concluído ---")
    print(f"Sucessos: {sucessos}")
    print(f"Falhas:   {falhas}")


# Executa a função principal
if __name__ == "__main__":
    enviar_payloads()
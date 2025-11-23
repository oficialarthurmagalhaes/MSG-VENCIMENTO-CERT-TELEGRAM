import pandas as pd
import openpyxl
import requests
import os
from dotenv import load_dotenv
import time # Mantido, embora não mais usado no loop, pode ser útil em caso de múltiplos envios futuros

# --- 1. CONFIGURAÇÕES ESSENCIAIS (AGORA LENDO DO .ENV) ---
# Carrega as variáveis de ambiente do arquivo .env
load_dotenv()

# Tenta ler as credenciais. Se não encontrar, o script para.
TOKEN = os.getenv('TELEGRAM_TOKEN')
CHAT_ID = os.getenv('TELEGRAM_CHAT_ID')

ARQUIVO_EXCEL = 'dados.xlsx' # O nome da sua planilha

# Verifica se as credenciais foram carregadas
if not TOKEN or not CHAT_ID:
    print("❌ ERRO: O TOKEN ou CHAT_ID não foram carregados.")
    print("Verifique se o arquivo '.env' existe e se as chaves 'TELEGRAM_TOKEN' e 'TELEGRAM_CHAT_ID' estão corretas.")
    # Forçando a saída segura
    exit(1)


# --- 2. FUNÇÃO DE ENVIO (Com suporte a HTML) ---
def enviar_telegram(mensagem):
#Realiza a chamada à API para enviar a mensagem (suporta HTML).
    url = f"https://api.telegram.org/bot{TOKEN}/sendMessage"
    dados = {
        "chat_id": CHAT_ID,
        "text": mensagem,
        "parse_mode": "HTML" # Essencial para interpretar o <b>negrito</b>
    }
    try:
        response = requests.post(url, data=dados)
        if response.status_code == 200:
            print("✅ Relatório enviado ao Telegram com sucesso!")
            return True
        else:
            print(f"❌ Erro ao enviar. Código: {response.status_code}")
            print(f"Detalhes do erro: {response.text}")
            return False
    except Exception as e:
        print(f"❌ Erro de conexão com a API do Telegram: {e}")
        return False

# --- 3. LÓGICA PRINCIPAL: Processar Excel e Consolidar Mensagem ---
def processar_e_enviar_alertas():
    
    if not os.path.exists(ARQUIVO_EXCEL):
        print(f"❌ Arquivo '{ARQUIVO_EXCEL}' não encontrado. Coloque a planilha na mesma pasta do script.")
        return

    try:
        # Lê a planilha inteira
        df = pd.read_excel(ARQUIVO_EXCEL)
        
        # Variáveis para acumular o relatório e contar os alertas
        relatorio_detalhado = ""
        certificados_em_alerta = 0
        
        # Loop principal: Percorre cada linha
        for index, linha in df.iterrows():
            
            # Extração dos dados
            # Garanta que os nomes das colunas aqui (ex: 'Código') batem com os nomes no seu Excel.
            codigo = linha['Código']
            empresa = linha['Empresa']
            dias = linha['Dias']
            
            # Formatação da Data para padrão BR (dd/mm/yyyy)
            validade_formatada = ""
            try:
                # Tenta usar strftime no objeto datetime do Pandas
                validade_formatada = linha['Validade'].strftime("%d/%m/%Y")
            except AttributeError:
                # Caso a célula do Excel não esteja como data, trata como string
                validade_formatada = str(linha['Validade'])
            
            
            # --- CONDIÇÃO DE FILTRO: DIAS RESTANTES (1 a 7) ---
            # Verifica se 'dias' é um número e está dentro do limite de 7 dias
            if isinstance(dias, (int, float)) and 0 < dias <= 7:
                
                # Monta a linha de alerta com negrito (<b>)
                aviso_linha = f"<b>- {empresa}</b> ({codigo}) vence em {dias} dias! [{validade_formatada}]\n"
                
                # Acumula a linha no relatório detalhado
                relatorio_detalhado += aviso_linha
                
                # Incrementa o contador
                certificados_em_alerta += 1

        
        # --- ENVIO FINAL (FORA DO LOOP) ---
        
        if certificados_em_alerta > 0:
            
            # Monta o cabeçalho final da mensagem
            texto_final = (f"⚠️ Bom dia! Há <b>{certificados_em_alerta}</b> certificado(s) próximo(s) do vencimento:\n\n"
                           f"{relatorio_detalhado}")
            
            enviar_telegram(texto_final)
            
        else:
            print("✅ Não há certificados próximos do vencimento. Nenhuma mensagem enviada ao Telegram.")
            
    except Exception as e:
        # Erro genérico (muitas vezes causado por nomes de colunas errados ou dados mal formatados)
        print(f"❌ Ocorreu um erro inesperado durante o processamento: {e}")
        print("Dica: Verifique se as colunas 'Código', 'Empresa', 'Dias' e 'Validade' existem no Excel.")

# --- EXECUÇÃO DO SCRIPT ---
if __name__ == "__main__":
    processar_e_enviar_alertas()
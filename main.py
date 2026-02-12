import gspread
from oauth2client.service_account import ServiceAccountCredentials
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import urllib.parse
import os

scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name('credenciais.json', scope)
client = gspread.authorize(creds)


planilha = client.open("Automatizacao").sheet1 



chrome_options = Options()

dir_path = os.getcwd()
chrome_options.add_argument(f"user-data-dir={dir_path}\\sessao_whatsapp")


chrome_options.add_argument("--headless") 
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
driver.get("https://web.whatsapp.com")

print("Aguardando carregamento do WhatsApp...")
time.sleep(20) 

def enviar_avisos():
    try:
        dados = planilha.get_all_records()
    except Exception as e:
        print(f"Erro ao ler planilha: {e}")
        return

    for i, linha in enumerate(dados):
      
        status = str(linha.get('Status', '')).strip().lower()
        avisado = str(linha.get('Avisados', '')).strip()

        if status == 'pronto' and avisado != 'Sim':
            nome_completo = str(linha.get('Nome do Cliente', 'Cliente'))
            primeiro_nome = nome_completo.split(" ")[0]
            data_venda = str(linha.get('Data da Venda', ''))
            
            
            telefone = str(linha.get('Telefone', '')).split('.')[0].strip().replace(" ", "").replace("-", "")
            
            if not telefone or len(telefone) < 10:
                print(f"⚠️ Telefone inválido para {nome_completo}: {telefone}")
                continue

          
            mensagem = (
                f"Olá {primeiro_nome}! 👋\n\n"
                f"Boas notícias! Seus óculos da compra de {data_venda} já estão prontos aqui na Qoculos. 👓\n\n"
                f"Pode vir buscar quando quiser! Estamos te esperando."
            )
            
            texto_url = urllib.parse.quote(mensagem)
            link = f"https://web.whatsapp.com/send?phone={telefone}&text={texto_url}"
            
            try:
                print(f"Processando: {nome_completo}...")
                driver.get(link)
                
               
                campo_texto = WebDriverWait(driver, 50).until(
                    EC.presence_of_element_located((By.XPATH, '//div[@contenteditable="true"][@data-tab="10"]'))
                )
                
                time.sleep(3) 
                
              
                campo_texto.send_keys(Keys.ENTER)
                
            
                time.sleep(8) 
                
                
                planilha.update_cell(i + 2, 5, "Sim")
                print(f"✅ Mensagem enviada para: {nome_completo}")
                
            except Exception as e:
                print(f"❌ Falha ao interagir com o WhatsApp para {nome_completo}. Verifique a conexão.")
            
           
            time.sleep(10)


enviar_avisos()
print("\n--- Processo Finalizado ---")
driver.quit()
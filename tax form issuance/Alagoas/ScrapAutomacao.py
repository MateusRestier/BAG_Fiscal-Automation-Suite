import subprocess
import os
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

"""Config dotenv"""
from dotenv import load_dotenv
from pathlib import Path
def localizar_env(diretorio_raiz="PRIVATE_BAG.ENV"):
    path = Path(__file__).resolve()
    for parent in path.parents:
        possible = parent / diretorio_raiz / ".env"
        if possible.exists():
            return possible
    raise FileNotFoundError(f"Arquivo .env não encontrado dentro de '{diretorio_raiz}'.")
env_path = localizar_env()
load_dotenv(dotenv_path=env_path)

def configure_browser():
    options = Options()
    options.add_argument("--start-maximized")
    #options.add_argument("--headless")
    options.add_argument("--disable-gpu")

    # Obtem a versão do Chrome instalada (Windows)
    try:
        output = subprocess.check_output(
            r'reg query "HKEY_CURRENT_USER\Software\Google\Chrome\BLBeacon" /v version',
            shell=True
        ).decode()
        chrome_version = "Desconhecida"
        for line in output.splitlines():
            if "version" in line.lower():
                chrome_version = line.split()[-1]
                break
    except Exception:
        chrome_version = "Desconhecida"

    print(f"🌐 Versão do Chrome instalada: {chrome_version}")

    # Instala e usa o ChromeDriver compatível
    driver_path = ChromeDriverManager().install()
    chromedriver_version = os.path.basename(os.path.dirname(driver_path))
    print(f"🧩 Versão do ChromeDriver utilizada: {chromedriver_version}")

    driver = webdriver.Chrome(service=Service(driver_path), options=options)
    return driver

def fazer_login(driver, cnpj, loja, login, senha):
    # Acessa o site
    url = "https://contribuinte.sefaz.al.gov.br/"
    driver.get(url)
    print(f"🔗 Acessando o site: {url}")

    # Aguardar até que o botão 'Cobrança de Documentos Fiscais Eletrônicos' esteja clicável
    WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.ID, "link-cobranca-dfe"))
    )
    
    # Encontrar o botão de cobrança e clicar
    botao_cobranca = driver.find_element(By.ID, "link-cobranca-dfe")
    botao_cobranca.click()
    print("✅ Clique realizado no botão 'Cobrança de Documentos Fiscais Eletrônicos'.")

    # Aguardar até que o botão 'login' esteja clicável
    WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, "a.alert-link[jhitranslate='global.messages.info.authenticated.link']"))
    )
    
    # Encontrar o botão de login e clicar
    botao_login = driver.find_element(By.CSS_SELECTOR, "a.alert-link[jhitranslate='global.messages.info.authenticated.link']")
    botao_login.click()
    print("✅ Clique realizado no botão 'login'.")

    # Aguardar até que os campos de login estejam visíveis
    WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.ID, "username"))
    )
    
    # Preencher o campo de login
    driver.find_element(By.ID, "username").send_keys(login)
    print(f"➡️ Usuário {login} preenchido.")
    
    # Preencher o campo de senha
    driver.find_element(By.ID, "password").send_keys(senha)
    print("➡️ Senha preenchida.")
    
    # Clicar no botão de entrar
    driver.find_element(By.CSS_SELECTOR, "button[type='submit']").click()
    print("Esperando 30 segundos para carregar a pagina depois do login...")
    time.sleep(30)  # Espera um pouco para garantir que a página carregue

    # Aguardar até que a mensagem com o "X" para fechar apareça
    try:
        # Verifica se o "X" aparece na tela (clica após 15 segundos, se aparecer)
        WebDriverWait(driver, 15).until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, "span[aria-hidden='true']"))
        )
        
        # Se encontrado, clica no "X"
        fechar_mensagem = driver.find_element(By.CSS_SELECTOR, "span[aria-hidden='true']")
        fechar_mensagem.click()
        print("✅ Mensagem fechada com sucesso.")
        
    except Exception as e:
        # Se não encontrar o "X" após 15 segundos, segue para o próximo passo
        print(f"⚠️ Nenhuma mensagem com 'X' foi encontrada ou clicada. Erro: {e}")
    
    # Esperar para garantir que a página do Google foi carregada
    time.sleep(2)

def emitir_guias():
    # Lista de lojas com CNPJ, Loja, Login e Senha

    lojas = [
        (os.getenv("CNPJLOJA75"), 75, os.getenv("LOGINLOJA75"), os.getenv("SENHALOJA75")),
        (os.getenv("CNPJLOJA76"), 76, os.getenv("LOGINLOJA76"), os.getenv("SENHALOJA76")),
        (os.getenv("CNPJLOJA86"), 86, os.getenv("LOGINLOJA86"), os.getenv("SENHALOJA86")),
        (os.getenv("CNPJLOJA89"), 89, os.getenv("LOGINLOJA89"), os.getenv("SENHALOJA89")),
        (os.getenv("CNPJLOJA151"), 151, os.getenv("LOGINLOJA151"), os.getenv("SENHALOJA151")),
    ]

    # Configura o navegador
    driver = configure_browser()
    
    # Loop para realizar login em todas as lojas
    for cnpj, loja, login, senha in lojas:
        print(f"\nIniciando login para a loja {loja} ({cnpj})...")

        # Faz login para cada loja
        fazer_login(driver, cnpj, loja, login, senha)

        # Aguardar até que o botão 'Minhas Cobranças' esteja clicável
        WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.ID, "link-acesso-obrigacoes-acessorias"))
        )

        # Encontrar o botão 'Minhas Cobranças' e clicar
        botao_minhas_cobrancas = driver.find_element(By.ID, "link-acesso-obrigacoes-acessorias")
        botao_minhas_cobrancas.click()
        print("✅ Clique realizado no botão 'Minhas Cobranças'.")
        input("Pressione Enter para continuar...")

        # Após completar as ações para uma loja, reinicia o navegador para garantir que a sessão seja limpa
        print("🔄 Reiniciando o navegador para garantir que a sessão anterior seja limpa.")
        driver.quit()

        # Reinicia o navegador para o próximo login
        driver = configure_browser()

    # Fechar o navegador após a execução de todos os logins
    driver.quit()



def main():
    # Chama a função emitir_guias
    emitir_guias()

if __name__ == "__main__":
    main()

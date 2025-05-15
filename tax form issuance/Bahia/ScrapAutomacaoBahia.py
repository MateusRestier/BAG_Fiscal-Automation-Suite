try:
    import os
    import time
    import pyautogui
    import pythoncom
    import pyodbc
    import subprocess
    import pyautogui
    from datetime import datetime, timedelta
    from collections import defaultdict
    from datetime import datetime
    from win32com.client import Dispatch
    from selenium import webdriver
    from selenium.webdriver.chrome.service import Service
    from webdriver_manager.chrome import ChromeDriverManager
    from selenium.common.exceptions import StaleElementReferenceException, ElementNotInteractableException
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.common.action_chains import ActionChains
except Exception as e:
    print(f"❌ Erro ao importar bibliotecas: {e}")

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

LOGIN = os.getenv("LOGIN_ECONET")
SENHA = os.getenv("SENHA_ECONET")

dir_down = os.getenv("DIR_DOWN_FICAL_BAHIA")

URL1 = "https://www.econeteditora.com.br/links_pagina_inicial/calculos/icmsba/diferencial_aliquotas/index.php?form[regimeOrigem]=C&form[destinatario]=N&form[regimeDestinatario]=C&form[beneficio_origem]=N&form[beneficio_destino]=N&form[acao]=formulario"
URL2 = "https://www.econeteditora.com.br///links_pagina_inicial/calculos/icmsba/Simulador_BA/calculo4.php?select_operacao=est"

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

def enviar_email_alerta():
    try:
        pythoncom.CoInitialize()
        outlook = Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)  # 0 = Email
        mail.To = "mateus.restier@bagaggio.com.br; rafaella.camacho@bagaggio.com.br; jessica.rodrigues@bagaggio.com.br"
        mail.Subject = "AUTOMÁTICO - FALHA EM RESOLVER O CAPTCHA (ANTECIPADOS BAHIA)"
        mail.Body = (
            "Olá,\n"
            "O robô não conseguiu resolver o CAPTCHA e avançar no login automático.\n"
            "Faça a conexão remotamente e resolva o CAPTCHA para dar continuidade à automação.\n\n"
            "Atenciosamente,\n"
            "Automação"
        )
        mail.Send()
        print("📧 E-mail de alerta enviado com sucesso.")
    except Exception as e:
        print(f"❌ Falha ao enviar e-mail de alerta: {e}")
    finally:
        pythoncom.CoUninitialize()

def enviar_email_encerramento():
    try:
        pythoncom.CoInitialize()
        outlook = Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.To = "mateus.restier@bagaggio.com.br; rafaella.camacho@bagaggio.com.br; jessica.rodrigues@bagaggio.com.br"
        mail.Subject = "AUTOMÁTICO - SCRIPT ENCERRADO POR NÃO RESOLVER CAPTCHA (ANTECIPADOS BAHIA)"
        mail.Body = (
            "Olá,\n\n"
            "O robô foi encerrado automaticamente após 12 horas de espera.\n"
            "O CAPTCHA não foi resolvido e o login não foi realizado.\n\n"
            "Atenciosamente,\n"
            "Automação"
        )
        mail.Send()
        print("📧 E-mail de encerramento enviado com sucesso.")
    except Exception as e:
        print(f"❌ Falha ao enviar e-mail de encerramento: {e}")
    finally:
        pythoncom.CoUninitialize()

def fazer_login(driver):
    driver.get(URL1)

    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.NAME, "Log"))).send_keys(LOGIN)
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.NAME, "Sen"))).send_keys(SENHA)

    print("⚠️ Tentando clicar automaticamente no CAPTCHA...")
    time.sleep(0.5)
    pyautogui.moveTo(559, 388, duration=3) # coordenadas do captcha num pc com tela principal em 1366x768
    pyautogui.click()
    print("🖱️ Clique no reCAPTCHA realizado. Aguardando liberação do botão de login...")

    captcha_timer = time.time()
    alerta_enviado = False
    timeout_segundos = 12 * 60 * 60  # 12 horas

    while True:
        try:
            botao = driver.find_element(By.ID, "login_ver")
            if botao.is_enabled():
                print("✅ CAPTCHA validado. Clicando no botão de login...")
                botao.click()
                print("✅ Login realizado com sucesso!")
                break
        except:
            pass

        tempo_decorrido = time.time() - captcha_timer

        # Envia alerta após 1 minuto
        if not alerta_enviado and tempo_decorrido > 60:
            print("⏰ 1 minuto se passou e o CAPTCHA ainda não foi resolvido. Enviando alerta...")
            enviar_email_alerta()
            alerta_enviado = True

        # Encerra após 12h se o botão ainda não estiver habilitado
        if tempo_decorrido > timeout_segundos:
            print("⏳ 12 horas se passaram sem login. Enviando e-mail e encerrando script...")
            enviar_email_encerramento()
            driver.quit()
            exit()

        time.sleep(1)


def fc_antecipadobahia(driver):

    class ZeroValueException(Exception):
        pass

    def tentar_voltar(driver_local):
        try:
            link_voltar = driver_local.find_element(By.XPATH, "//a[contains(@href,'javascript:history.back()')]")
            link_voltar.click()
            print("   ↩️ Botão 'voltar' clicado.")
        except:
            print("   ↩️ Botão 'voltar' não encontrado. Usando driver.back().")
            driver_local.back()

    def preencher_calcular(driver_local, baseicms, valipi, alqicms):
        tentativas = 0
        num_tentativas = 30

        while tentativas < num_tentativas:
            tentativas += 1
            print(f"   🛠️ Tentativa #{tentativas} de preencher e calcular...")

            try:
                driver_local.get(URL1)
                time.sleep(2)

                # Valor da operação (BASEICMS)
                WebDriverWait(driver_local, 20).until(
                    EC.element_to_be_clickable((By.NAME, "form[vlr_operacao]"))
                ).send_keys(str(baseicms).replace('.', ','))

                # Valor do IPI (VALIPI)
                driver_local.find_element(By.NAME, "form[vlr_ipi]").send_keys(str(valipi).replace('.', ','))

                # Alíquota interestadual (ALQICMS)
                select = WebDriverWait(driver_local, 10).until(
                    EC.element_to_be_clickable((By.NAME, "form[aliq_interestadual]"))
                )
                select.click()
                driver_local.find_element(
                    By.XPATH, f"//select[@name='form[aliq_interestadual]']/option[@value='{int(alqicms)}']"
                ).click()

                # Alíquota Interna (BA) = 20,5
                driver_local.find_element(By.NAME, "form[aliq_interna]").send_keys("20,5")

                # Clicar em 'Calcular'
                WebDriverWait(driver_local, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//input[@type='submit' and @value='Calcular']"))
                ).click()

                # Capturar valor final
                valor_elemento = WebDriverWait(driver_local, 20).until(
                    EC.presence_of_element_located((
                        By.XPATH, "//tr[td[contains(text(),'Valor - Antecipação parcial')]]/td[2]"
                    ))
                )
                valor = valor_elemento.text.strip()

                if valor == "R$ 0,00":
                    raise ZeroValueException("Retorno foi R$ 0,00.")

                print(f"🔢 Valor obtido da Antecipação parcial: {valor}")
                return valor

            except (StaleElementReferenceException, ElementNotInteractableException):
                print("   ⚠️ Elemento temporariamente inacessível. Recarregando e tentando novamente...")
                tentar_voltar(driver_local)

            except ZeroValueException as e:
                print(f"   ⚠️ {e} Tentando novamente após voltar...")
                tentar_voltar(driver_local)

            except Exception as e:
                print(f"   ⚠️ Erro inesperado: {type(e).__name__}. Tentando novamente...")
                tentar_voltar(driver_local)

            time.sleep(2)

        print(f"❌ Falha após {num_tentativas} tentativas. Registro ignorado.")
        return None

    conn = pyodbc.connect(
        f"DRIVER={{ODBC Driver 17 for SQL Server}};"
        f"SERVER={os.getenv('DB_SERVER_EXCEL')},{os.getenv('DB_PORT_EXCEL')};"
        f"DATABASE={os.getenv('DB_DATABASE_EXCEL')};"
        f"UID={os.getenv('DB_USER_EXCEL')};"
        f"PWD={os.getenv('DB_PASSWORD_EXCEL')}"
    )
    cursor = conn.cursor()

    hoje = datetime.now().strftime('%Y%m%d')

    cursor.execute(f"""
        SELECT ID, BASEICMS, VALIPI, ALQICMS, EMISSÃO
        FROM dbo.FC_AntecipadoBahia
        WHERE AntecipacaoParcial IS NULL
          AND EMISSÃO <= '{hoje}'
    """)
    rows = cursor.fetchall()

    print(f"🔎 Encontrados {len(rows)} registros para cálculo...")

    for row in rows:
        _id, baseicms, valipi, alqicms, emissao = row

        # Verifica se algum valor essencial é None
        if baseicms is None or valipi is None or alqicms is None:
            print(f"❌ Registro ID={_id} possui valor nulo (BASEICMS, VALIPI ou ALQICMS). Pulando registro.")
            continue

        print(f"\n➡️ Processando ID={_id}, BASEICMS={baseicms}, VALIPI={valipi}, ALQICMS={alqicms}")
        valor = preencher_calcular(driver, baseicms, valipi, alqicms)

        if valor is None:
            print(f"   ⛔ Ignorando ID={_id} (sem valor obtido).")
            continue

        valor_numerico = valor.replace("R$ ", "").replace(".", "").replace(",", ".")
        cursor.execute("""
            UPDATE dbo.FC_AntecipadoBahia
            SET AntecipacaoParcial = ?
            WHERE ID = ?
        """, (valor_numerico, _id))
        conn.commit()

        print(f"✅ Atualizado com sucesso. ID={_id}, Valor={valor_numerico}")

    cursor.close()
    conn.close()
    print("🏁 Finalizado processamento de todos os registros.")


def fc_antecipadobahiast(driver):
    """
    Preenche o cálculo de ST na URL2 usando dados de FC_AntecipadoBahiaST.
    Captura o valor do ICMS ST corretamente da linha que contém o texto:
    'Valor do ICMS Substituição Tributária'.
    """

    class EmptyValueException(Exception):
        pass

    def preencher_calcular_st(driver_local, baseicms, alqicms, alqipi, mva):
        max_tentativas = 30
        tentativas = 0

        while tentativas < max_tentativas:
            tentativas += 1
            print(f"   🛠️ Tentativa #{tentativas} de preencher e calcular ST...")

            try:
                driver_local.get(URL2)
                time.sleep(2)

                # Preencher os campos necessários
                WebDriverWait(driver_local, 20).until(EC.presence_of_element_located((By.NAME, "campo1")))

                driver_local.find_element(By.NAME, "campo1").clear()
                driver_local.find_element(By.NAME, "campo1").send_keys(str(baseicms).replace('.', ','))

                driver_local.find_element(By.NAME, "campo9").click()
                driver_local.find_element(
                    By.XPATH, f"//select[@name='campo9']/option[@value='{int(alqicms)}']"
                ).click()

                driver_local.find_element(By.NAME, "campo12").clear()
                driver_local.find_element(By.NAME, "campo12").send_keys(str(alqipi).replace('.', ','))

                driver_local.find_element(By.NAME, "campo15").clear()
                driver_local.find_element(By.NAME, "campo15").send_keys(str(mva))

                driver_local.find_element(By.NAME, "campo18").clear()
                driver_local.find_element(By.NAME, "campo18").send_keys("20,5")

                # Rola para o fim da página antes de clicar
                driver_local.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                time.sleep(1)

                # Clicar em Calcular
                driver_local.find_element(By.NAME, "botao").click()

                # Rola novamente para garantir renderização
                driver_local.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                time.sleep(1)

                # Captura o valor no td da Substituição Tributária
                valor_st_element = WebDriverWait(driver_local, 20).until(
                    EC.presence_of_element_located((
                        By.XPATH, "//tr[td[2][contains(text(),'Valor do ICMS Substituição Tributária')]]/td[4]"
                    ))
                )
                valor_st = valor_st_element.text.strip()

                if valor_st in ["R$", ""]:
                    raise EmptyValueException("Valor vazio ou apenas 'R$'.")

                print(f"   💰 Valor ICMS ST capturado: {valor_st}")
                return valor_st

            except (StaleElementReferenceException, ElementNotInteractableException) as e:
                print(f"   ⚠️ Erro de interação (tentativa {tentativas}): {type(e).__name__}. Repetindo...")

            except EmptyValueException as e:
                print(f"   ⚠️ {e} Repetindo...")

            except Exception as e:
                print(f"   ⚠️ Erro inesperado (tentativa {tentativas}): {type(e).__name__}. Repetindo...")

            time.sleep(2)

        print(f"❌ Falha após {max_tentativas} tentativas. Registro ignorado.")
        return None

    print("➡️ Iniciando cálculo de ST em URL2...")

    conn = pyodbc.connect(
        f"DRIVER={{ODBC Driver 17 for SQL Server}};"
        f"SERVER={os.getenv('DB_SERVER_EXCEL')},{os.getenv('DB_PORT_EXCEL')};"
        f"DATABASE={os.getenv('DB_DATABASE_EXCEL')};"
        f"UID={os.getenv('DB_USER_EXCEL')};"
        f"PWD={os.getenv('DB_PASSWORD_EXCEL')}"
    )
    cursor = conn.cursor()

    cursor.execute("""
        SELECT ID, BASEICMS, ALQICMS, ALQIPI, MVA
        FROM dbo.FC_AntecipadoBahiaST
        WHERE SubstituicaoTributaria IS NULL
    """)
    rows = cursor.fetchall()

    print(f"🔎 Encontrados {len(rows)} registros para Substituição Tributária...")

    for row in rows:
        _id, baseicms, alqicms, alqipi, mva = row

        # Se qualquer valor essencial for None, pula este registro
        if baseicms is None or alqicms is None or alqipi is None or mva is None:
            print(f"❌ ID={_id} contém valor nulo (BASEICMS, ALQICMS, ALQIPI ou MVA). Pulando registro.")
            continue

        print(f"\n📝 Processando ST do ID={_id}, BASEICMS={baseicms}, ALQICMS={alqicms}, ALQIPI={alqipi}, MVA={mva}")

        valor_st = preencher_calcular_st(driver, baseicms, alqicms, alqipi, mva)
        if not valor_st:
            print(f"   ❌ Valor ST não obtido para ID={_id}. Pulando registro.")
            continue

        valor_st_numerico = (
            valor_st.replace("R$", "").replace(".", "").replace(",", ".").strip()
        )

        cursor.execute("""
            UPDATE dbo.FC_AntecipadoBahiaST
            SET SubstituicaoTributaria = ?
            WHERE ID = ?
        """, (valor_st_numerico, _id))
        conn.commit()

        print(f"   ✅ SubstituicaoTributaria atualizada para ID={_id}: {valor_st_numerico}")

    cursor.close()
    conn.close()
    print("🏁 Fim do processo de ST.")


def verificar_pendencia_financeira(driver):
    """
    Acessa a página inicial da Econet e verifica se há aviso de pendência financeira.
    Caso haja, envia e-mail automático para o responsável.
    """
    print("🔍 Verificando pendência financeira na conta...")

    try:
        driver.get("https://www.econeteditora.com.br/index.asp?url=/")
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "body")))

        aviso = driver.find_elements(By.XPATH, "//td[contains(text(),'pendência financeira')]")
        if aviso:
            print("🚨 Pendência financeira detectada. Enviando e-mail...")

            pythoncom.CoInitialize()
            outlook = Dispatch("Outlook.Application")
            mail = outlook.CreateItem(0)
            mail.To = "mateus.restier@bagaggio.com.br; rafaella.camacho@bagaggio.com.br; jessica.rodrigues@bagaggio.com.br"
            mail.Subject = "AUTOMÁTICO: 🚨 ECONET Pendência Financeira Detectada"
            mail.Body = (
                "Olá,\n\n"
                "A automação identificou que sua assinatura na Econet apresenta pendência financeira.\n"
                "Evite o bloqueio da conta acessando a área do cliente o quanto antes.\n\n"
                "Link: https://www.econeteditora.com.br/index.asp?url=/\n\n"
                "Atenciosamente,\nAutomação"
            )
            mail.Send()
            print("📧 E-mail enviado com sucesso.")
        else:
            print("✅ Nenhuma pendência financeira detectada.")
    except Exception as e:
        print(f"❌ Erro ao verificar pendência ou enviar e-mail: {e}")
    finally:
        pythoncom.CoUninitialize()
        

def emissaoantecipado(driver):
    """
    Acessa a página de emissão do DAE na SEFAZ BA e preenche os dados com base nos registros
    agrupados por LOJA e data de emissão da tabela FC_AntecipadoBahia.
    Agora os campos de data (vencimento e pagamento) são preenchidos sem barras (ex: "03042025").
    """
    print("➡️ Acessando página de emissão do DAE (Antecipado) na SEFAZ BA...")

    try:

        def navegar_emissao (driver):
            url = "https://servicos.sefaz.ba.gov.br/sistemas/arasp/pagamento/modulos/dae/pagamento/dae_pagamento.aspx"
            driver.get(url)

            WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.ID, "PHConteudo_ddl_antecipacao_tributaria"))
            )
            print("✅ Página carregada. Selecionando código de receita...")

            driver.find_element(By.ID, "PHConteudo_ddl_antecipacao_tributaria").click()
            driver.find_element(By.XPATH, "//option[@value='2175|formulario']").click()
            time.sleep(0.5)

            WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "PHConteudo_rb_dae_normal_1"))
            )
            radio = driver.find_element(By.ID, "PHConteudo_rb_dae_normal_1")
            driver.execute_script("arguments[0].scrollIntoView(true);", radio)
            radio.click()
            time.sleep(0.5)

        navegar_emissao (driver)

        # Conectar ao banco
        conn = pyodbc.connect(
            f"DRIVER={{ODBC Driver 17 for SQL Server}};"
            f"SERVER={os.getenv('DB_SERVER_EXCEL')},{os.getenv('DB_PORT_EXCEL')};"
            f"DATABASE={os.getenv('DB_DATABASE_EXCEL')};"
            f"UID={os.getenv('DB_USER_EXCEL')};"
            f"PWD={os.getenv('DB_PASSWORD_EXCEL')}"
        )
        cursor = conn.cursor()

        cursor.execute("""
            SELECT ID, EMISSÃO, LOJA, IE, NF, AntecipacaoParcial
            FROM dbo.FC_AntecipadoBahia
            WHERE 1 = 1 
            AND AntecipacaoParcial IS NOT NULL
            AND IE IS NOT NULL
            AND GUIAEMITIDA < 1;
        """) # se quiser emitir uma guia específica, só mudar essa consulta filtrando a emissao e a loja desejada
        rows = cursor.fetchall()

        if not rows:
            print("ℹ️ Nenhuma guia de Antecipado encontrada para emitir. Processo encerrado.")
            cursor.close()
            conn.close()
            return

        grupos = defaultdict(list)
        for row in rows:
            chave = (row[1], row[2])  # Agrupa por EMISSÃO (row[1]) e LOJA (row[2])
            grupos[chave].append({
                'ID': row[0],
                'IE': row[3],
                'NF': row[4],
                'AntecipacaoParcial': row[5]
            })

        def increment_guiaemitida(cursor, conn, tabela, id_list):
            """
            Incrementa a coluna GUIAEMITIDA para os registros na tabela especificada
            cujos IDs estão em id_list, somando +1 ao valor atual (considerando NULL como 0).
            """
            if not id_list:
                print("Nenhum ID para atualizar GUIAEMITIDA.")
                return
            placeholders = ','.join('?' for _ in id_list)
            sql = f"UPDATE dbo.FC_AntecipadoBahia SET GUIAEMITIDA = COALESCE(GUIAEMITIDA, 0) + 1 WHERE ID IN ({placeholders})"
            cursor.execute(sql, id_list)
            conn.commit()
            print(f"Incrementado GUIAEMITIDA para {cursor.rowcount} registros na tabela {tabela}.")

        # Definir datas
        hoje = datetime.now()
        amanha = hoje + timedelta(days=1)
        depois_de_amanha = hoje + timedelta(days=2)

        # Formatos desejados
        hoje_str = hoje.strftime('%d%m%Y')                # exemplo: 11042025
        mes_ano = hoje.strftime('%m/%Y')                  # exemplo: 04/2025
        amanha_str = amanha.strftime('%d%m%Y')            # exemplo: 12042025
        depois_de_amanha_str = depois_de_amanha.strftime('%d%m%Y')  # exemplo: 13042025

        for (emissao, loja), registros in grupos.items():
            print(f"\n🧾 Preenchendo grupo: LOJA={loja}, EMISSÃO={emissao}")

            # Preencher IE
            WebDriverWait(driver, 60).until(
                EC.presence_of_element_located((By.ID, "PHconteudoSemAjax_txt_num_inscricao_estad"))
            )
            campo_ie = driver.find_element(By.ID, "PHconteudoSemAjax_txt_num_inscricao_estad")
            campo_ie.send_keys(registros[0]['IE'])
            print(f"➡️ IE preenchido: {registros[0]['IE']}")
            time.sleep(0.5)

            # Converter data de emissão para ddmmyyyy (sem barras)
            emissao_formatada = f"{emissao[6:8]}{emissao[4:6]}{emissao[0:4]}"
            campo_venc = driver.find_element(By.ID, "PHconteudoSemAjax_txt_dtc_vencimento")
            ActionChains(driver).move_to_element(campo_venc).click().pause(1).perform()
            #campo_venc.clear() # com isso aq ta bugando nessa opção
            campo_venc.send_keys(emissao_formatada)
            print(f"📆 Emissão preenchida (vencimento): {emissao_formatada}")
            time.sleep(0.5)
            driver.find_element(By.ID, "PHconteudoSemAjax_txt_dtc_max_pagamento").click()
            time.sleep(0.5)

            # Preencher data de pagamento (amanhã) no formato ddmmyyyy
            campo_pag = driver.find_element(By.ID, "PHconteudoSemAjax_txt_dtc_max_pagamento")
            #ActionChains(driver).move_to_element(campo_pag).click().pause(1).perform() # com isso aq ta bugando nessa opção
            #campo_pag.clear() # com isso aq ta bugando nessa opção
            campo_pag.send_keys(amanha_str)
            print(f"📆 Pagamento preenchido: {amanha_str}")
            time.sleep(0.5)
            driver.find_element(By.ID, "PHconteudoSemAjax_txt_val_principal").click()
            time.sleep(0.5)

            # Scroll até metade da página
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight / 4);")
            time.sleep(0.5)

            # Calcular valor total
            total_float = sum([
                float(r['AntecipacaoParcial'].replace("R$", "").strip().replace(",", "."))
                for r in registros if r['AntecipacaoParcial']
            ])
            # Formata para duas casas decimais e remove o ponto decimal
            total_formatado = f"{total_float:.2f}"  # ex: "284.16"
            valor_site = total_formatado.replace(".", "")  # ex: "28416"
            print(f"💰 Valor total calculado: {total_formatado} -> Valor digitado: {valor_site}")
            campo_valor = driver.find_element(By.ID, "PHconteudoSemAjax_txt_val_principal")
            campo_valor.send_keys(valor_site)
            time.sleep(0.5)

            # Preencher referência (mês/ano) - mantemos o formato mm/aaaa
            campo_ref = driver.find_element(By.ID, "PHconteudoSemAjax_txt_mes_ano_referencia_6anos")
            ActionChains(driver).move_to_element(campo_ref).click().pause(1).perform()
            #campo_ref.clear()
            campo_ref.send_keys(mes_ano)
            print(f"📆 Referência preenchida: {mes_ano}")
            time.sleep(0.5)

            # Scroll até o final da página
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(0.5)

            # Preencher notas fiscais (não repetir)
            nfs_unicas = list(dict.fromkeys([r['NF'] for r in registros]))[:15]
            print(f"🧾 Notas fiscais a preencher: {nfs_unicas}")
            for idx, nf in enumerate(nfs_unicas):
                id_input = f"PHconteudoSemAjax_txt_num_nota_fiscal{'' if idx == 0 else str(idx+1)}"
                campo_nf = driver.find_element(By.ID, id_input)
                campo_nf.send_keys(nf)
                print(f"   NF inserida no campo {id_input}: {nf}")
                time.sleep(0.5)

            # Preencher quantidade de NFs
            campo_qtd = driver.find_element(By.ID, "PHconteudoSemAjax_txt_qtd_nota_fiscal")
            campo_qtd.send_keys(str(len(nfs_unicas)))
            print(f"🧮 Quantidade de NFs preenchida: {len(nfs_unicas)}")
            time.sleep(0.5)

            # Preencher descrição
            descricao = f"Antecipado - {loja} - {emissao[0:4]}/{emissao[4:6]}/{emissao[6:8]}"
            campo_desc = driver.find_element(By.ID, "PHconteudoSemAjax_txt_des_informacoes_complementares")
            campo_desc.send_keys(descricao)
            print(f"📝 Descrição preenchida: {descricao}")
            time.sleep(0.5)

            # Clicar no botão "Visualizar o DAE"
            print("➡️ Clicando em 'Visualizar o DAE'...")
            WebDriverWait(driver, 60).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "label[for='PHconteudoSemAjax_btn_visualizar']"))
            )
            driver.find_element(By.CSS_SELECTOR, "label[for='PHconteudoSemAjax_btn_visualizar']").click()
            time.sleep(3)

            # Espera a nova página carregar e clica em "Imprimir o DAE"
            print("➡️ Clicando em 'Imprimir o DAE'...")
            WebDriverWait(driver, 60).until(
                EC.presence_of_element_located((By.ID, "PHConteudo_rep_dae_receita_btn_imprimir_0"))
            )
            driver.find_element(By.ID, "PHConteudo_rep_dae_receita_btn_imprimir_0").click()
            time.sleep(3)

            # Trocar para a janela que contém a página do boleto
            try:
                main_handle = driver.current_window_handle
                boleto_handle = None
                timeout = 30  # segundos
                start_time = time.time()
                while time.time() - start_time < timeout:
                    handles = driver.window_handles
                    print(f"Janelas abertas: {handles}")
                    for handle in handles:
                        if handle != main_handle:
                            try:
                                driver.switch_to.window(handle)
                                current_url = driver.current_url
                                print(f"Verificando janela {handle} com URL: {current_url}")
                                if "BoletoDae.aspx" in current_url:
                                    boleto_handle = handle
                                    print(f"➡️ Janela do boleto encontrada: {handle}")
                                    break
                            except Exception as ex:
                                print(f"Erro ao tentar acessar a janela {handle}: {ex}")
                    if boleto_handle:
                        break
                    time.sleep(1)
                if boleto_handle is None:
                    print("❌ Janela do boleto não encontrada, permanecendo na janela atual.")
                time.sleep(3)
            except Exception as e:
                print(f"❌ Erro ao tentar trocar de janela: {e}")
                boleto_handle = None
                # Prossegue com a janela atual

            # Salvar a página do boleto como PDF usando CDP do Chrome
            print("➡️ Salvando a página do boleto como PDF usando CDP do Chrome...")
            try:
                pdf = driver.execute_cdp_cmd("Page.printToPDF", {"printBackground": True})
                print("✅ Comando Page.printToPDF executado com sucesso.")
            except Exception as ex:
                print(f"❌ Erro ao executar Page.printToPDF: {ex}")
                # Se ocorrer erro, continuar para a próxima etapa
                pdf = None
            if pdf:
                try:
                    import base64
                    pdf_data = base64.b64decode(pdf["data"])
                    print("✅ PDF data decodificada com sucesso.")
                except Exception as ex:
                    print(f"❌ Erro ao decodificar PDF data: {ex}")
                    pdf_data = None
            if pdf_data:
                now = datetime.now()
                year = now.strftime("%Y")
                month = now.strftime("%m")
                day = now.strftime("%d")
                download_dir = f"{dir_down}Contabilidade\\Fiscal\\{year}\\LUCRO REAL\\SHEHRAZADE\\{month}.{year}\\ICMS\\ICMS ANTECIPADO E ST\\BAHIA\\Antecipado\\{year}\\{month}\\{day}"
                print(f"➡️ Diretório de download: {download_dir}")
                if not os.path.exists(download_dir):
                    print("➡️ Diretório não existe, criando...")
                    os.makedirs(download_dir)
                    print("✅ Diretório criado.")
                else:
                    print("✅ Diretório já existe.")
                # Substituir barras por hífen na descrição para criar um nome de arquivo válido
                safe_descricao = descricao.replace("/", "-")
                file_path = os.path.join(download_dir, f"{safe_descricao}.pdf")
                print(f"➡️ Caminho completo para salvar o PDF: {file_path}")
                try:
                    with open(file_path, "wb") as f:
                        f.write(pdf_data)
                    print(f"✅ PDF salvo em: {file_path}")
                except Exception as ex:
                    print(f"❌ Erro ao salvar PDF: {ex}")
            time.sleep(3)

            # Incrementar a coluna GUIAEMITIDA para os registros deste grupo
            id_list = [r['ID'] for r in registros]
            try:
                increment_guiaemitida(cursor, conn, "FC_AntecipadoBahia", id_list)
            except Exception as ex:
                print(f"❌ Erro ao incrementar GUIAEMITIDA: {ex}")
            time.sleep(3)

            # Fechar a janela do boleto e voltar para a janela principal
            try:
                print("➡️ Fechando a janela do boleto e voltando para a janela principal...")
                driver.close()
                driver.switch_to.window(main_handle)
                print("✅ Janela do boleto fechada. Voltando para a janela principal.")
            except Exception as ex:
                print(f"❌ Erro ao fechar a janela do boleto ou voltar para a janela principal: {ex}")
            time.sleep(3)

            #input("Pressione ENTER para continuar...")  # Pausa para revisão antes de enviar

            # No final do grupo, volta até a página da emissão
            navegar_emissao (driver)

        cursor.close()
        conn.close()

    except Exception as e:
        print(f"❌ Erro durante emissão do DAE: {type(e).__name__} - {e}")


def emissaoantecipadost(driver):
    """
    Acessa a página de emissão do DAE na SEFAZ BA e preenche os dados com base nos registros
    agrupados por LOJA e data de emissão da tabela FC_AntecipadoBahiaST.
    Agora os campos de data (vencimento e pagamento) são preenchidos sem barras (ex: "03042025").
    """
    print("➡️ Acessando página de emissão do DAE (AntecipadoST) na SEFAZ BA...")

    try:

        def navegar_emissao (driver):
            url = "https://servicos.sefaz.ba.gov.br/sistemas/arasp/pagamento/modulos/dae/pagamento/dae_pagamento.aspx"
            driver.get(url)

            WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.ID, "PHConteudo_ddl_antecipacao_tributaria"))
            )
            print("✅ Página carregada. Selecionando código de receita...")

            driver.find_element(By.ID, "PHConteudo_ddl_antecipacao_tributaria").click()
            driver.find_element(By.XPATH, "//option[@value='1145|campanha']").click()
            time.sleep(0.5)

            WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "PHConteudo_rb_dae_normal_1"))
            )
            radio = driver.find_element(By.ID, "PHConteudo_rb_dae_normal_1")
            driver.execute_script("arguments[0].scrollIntoView(true);", radio)
            radio.click()
            time.sleep(0.5)

        navegar_emissao (driver)

        # Conectar ao banco
        conn = pyodbc.connect(
            f"DRIVER={{ODBC Driver 17 for SQL Server}};"
            f"SERVER={os.getenv('DB_SERVER_EXCEL')},{os.getenv('DB_PORT_EXCEL')};"
            f"DATABASE={os.getenv('DB_DATABASE_EXCEL')};"
            f"UID={os.getenv('DB_USER_EXCEL')};"
            f"PWD={os.getenv('DB_PASSWORD_EXCEL')}"
        )
        cursor = conn.cursor()

        cursor.execute("""
            SELECT ID, EMISSÃO, LOJA, IE, NF, SubstituicaoTributaria
            FROM dbo.FC_AntecipadoBahiaST
            WHERE 1 = 1 
            AND SubstituicaoTributaria IS NOT NULL
            AND IE IS NOT NULL
            AND MVA IS NOT NULL
            AND GUIAEMITIDA < 1;
        """) # se quiser emitir uma guia específica, só mudar essa consulta filtrando a emissao e a loja desejada
        rows = cursor.fetchall()

        if not rows:
            print("ℹ️ Nenhuma guia de Antecipado encontrada para emitir. Processo encerrado.")
            cursor.close()
            conn.close()
            return

        grupos = defaultdict(list)
        for row in rows:
            chave = (row[1], row[2])  # Agrupa por EMISSÃO (row[1]) e LOJA (row[2])
            grupos[chave].append({
                'ID': row[0],
                'IE': row[3],
                'NF': row[4],
                'SubstituicaoTributaria': row[5]
            })

        def increment_guiaemitida(cursor, conn, tabela, id_list):
            """
            Incrementa a coluna GUIAEMITIDA para os registros na tabela especificada
            cujos IDs estão em id_list, somando +1 ao valor atual (considerando NULL como 0).
            """
            if not id_list:
                print("Nenhum ID para atualizar GUIAEMITIDA.")
                return
            placeholders = ','.join('?' for _ in id_list)
            sql = f"UPDATE dbo.FC_AntecipadoBahiaST SET GUIAEMITIDA = COALESCE(GUIAEMITIDA, 0) + 1 WHERE ID IN ({placeholders})"
            cursor.execute(sql, id_list)
            conn.commit()
            print(f"Incrementado GUIAEMITIDA para {cursor.rowcount} registros na tabela {tabela}.")

        # Definir datas
        hoje = datetime.now()
        amanha = hoje + timedelta(days=1)
        depois_de_amanha = hoje + timedelta(days=2)

        # Formatos desejados
        hoje_str = hoje.strftime('%d%m%Y')                # exemplo: 11042025
        mes_ano = hoje.strftime('%m/%Y')                  # exemplo: 04/2025
        amanha_str = amanha.strftime('%d%m%Y')            # exemplo: 12042025
        depois_de_amanha_str = depois_de_amanha.strftime('%d%m%Y')  # exemplo: 13042025

        for (emissao, loja), registros in grupos.items():
            print(f"\n🧾 Preenchendo grupo: LOJA={loja}, EMISSÃO={emissao}")

            # Preencher IE
            WebDriverWait(driver, 60).until(
                EC.presence_of_element_located((By.ID, "PHconteudoSemAjax_txt_num_inscricao_estad"))
            )
            campo_ie = driver.find_element(By.ID, "PHconteudoSemAjax_txt_num_inscricao_estad")
            campo_ie.send_keys(registros[0]['IE'])
            print(f"➡️ IE preenchido: {registros[0]['IE']}")
            time.sleep(0.5)

            # Converter data de emissão para ddmmyyyy (sem barras)
            emissao_formatada = f"{emissao[6:8]}{emissao[4:6]}{emissao[0:4]}"
            campo_venc = driver.find_element(By.ID, "PHconteudoSemAjax_txt_dtc_vencimento")
            ActionChains(driver).move_to_element(campo_venc).click().pause(1).perform()
            #campo_venc.clear() # com isso aq ta bugando nessa opção
            campo_venc.send_keys(emissao_formatada)
            print(f"📆 Emissão preenchida (vencimento): {emissao_formatada}")
            time.sleep(0.5)
            driver.find_element(By.ID, "PHconteudoSemAjax_txt_dtc_max_pagamento").click()
            time.sleep(0.5)

            # Preencher data de pagamento (amanhã) no formato ddmmyyyy
            campo_pag = driver.find_element(By.ID, "PHconteudoSemAjax_txt_dtc_max_pagamento")
            #ActionChains(driver).move_to_element(campo_pag).click().pause(1).perform() # com isso aq ta bugando nessa opção
            #campo_pag.clear() # com isso aq ta bugando nessa opção
            campo_pag.send_keys(amanha_str)
            print(f"📆 Pagamento preenchido: {amanha_str}")
            time.sleep(0.5)
            driver.find_element(By.ID, "PHconteudoSemAjax_txt_val_principal").click()
            time.sleep(0.5)

            # Scroll até metade da página
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight / 4);")
            time.sleep(0.5)

            # Calcular valor total
            total_float = sum([
                float(r['SubstituicaoTributaria'].replace("R$", "").strip().replace(",", "."))
                for r in registros if r['SubstituicaoTributaria']
            ])
            # Formata para duas casas decimais e remove o ponto decimal
            total_formatado = f"{total_float:.2f}"  # ex: "284.16"
            valor_site = total_formatado.replace(".", "")  # ex: "28416"
            print(f"💰 Valor total calculado: {total_formatado} -> Valor digitado: {valor_site}")
            campo_valor = driver.find_element(By.ID, "PHconteudoSemAjax_txt_val_principal")
            campo_valor.send_keys(valor_site)
            time.sleep(0.5)

            # Preencher referência (mês/ano) - mantemos o formato mm/aaaa
            campo_ref = driver.find_element(By.ID, "PHconteudoSemAjax_txt_mes_ano_referencia_6anos")
            ActionChains(driver).move_to_element(campo_ref).click().pause(1).perform()
            #campo_ref.clear()
            campo_ref.send_keys(mes_ano)
            print(f"📆 Referência preenchida: {mes_ano}")
            time.sleep(0.5)

            # Scroll até o final da página
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(0.5)

            # Preencher notas fiscais (não repetir)
            nfs_unicas = list(dict.fromkeys([r['NF'] for r in registros]))[:15]
            print(f"🧾 Notas fiscais a preencher: {nfs_unicas}")
            for idx, nf in enumerate(nfs_unicas):
                id_input = f"PHconteudoSemAjax_txt_num_nota_fiscal{'' if idx == 0 else str(idx+1)}"
                campo_nf = driver.find_element(By.ID, id_input)
                campo_nf.send_keys(nf)
                print(f"   NF inserida no campo {id_input}: {nf}")
                time.sleep(0.5)

            # Preencher quantidade de NFs
            campo_qtd = driver.find_element(By.ID, "PHconteudoSemAjax_txt_qtd_nota_fiscal")
            campo_qtd.send_keys(str(len(nfs_unicas)))
            print(f"🧮 Quantidade de NFs preenchida: {len(nfs_unicas)}")
            time.sleep(0.5)

            # Preencher descrição
            descricao = f"AntecipadoST - {loja} - {emissao[0:4]}/{emissao[4:6]}/{emissao[6:8]}"
            campo_desc = driver.find_element(By.ID, "PHconteudoSemAjax_txt_des_informacoes_complementares")
            campo_desc.send_keys(descricao)
            print(f"📝 Descrição preenchida: {descricao}")
            time.sleep(0.5)

            # Clicar no botão "Visualizar o DAE"
            print("➡️ Clicando em 'Visualizar o DAE'...")
            WebDriverWait(driver, 60).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "label[for='PHconteudoSemAjax_btn_visualizar']"))
            )
            driver.find_element(By.CSS_SELECTOR, "label[for='PHconteudoSemAjax_btn_visualizar']").click()
            time.sleep(3)

            # Espera a nova página carregar e clica em "Imprimir o DAE"
            print("➡️ Clicando em 'Imprimir o DAE'...")
            WebDriverWait(driver, 60).until(
                EC.presence_of_element_located((By.ID, "PHConteudo_rep_dae_receita_btn_imprimir_0"))
            )
            driver.find_element(By.ID, "PHConteudo_rep_dae_receita_btn_imprimir_0").click()
            time.sleep(3)

            # Trocar para a janela que contém a página do boleto
            try:
                main_handle = driver.current_window_handle
                boleto_handle = None
                timeout = 30  # segundos
                start_time = time.time()
                while time.time() - start_time < timeout:
                    handles = driver.window_handles
                    print(f"Janelas abertas: {handles}")
                    for handle in handles:
                        if handle != main_handle:
                            try:
                                driver.switch_to.window(handle)
                                current_url = driver.current_url
                                print(f"Verificando janela {handle} com URL: {current_url}")
                                if "BoletoDae.aspx" in current_url:
                                    boleto_handle = handle
                                    print(f"➡️ Janela do boleto encontrada: {handle}")
                                    break
                            except Exception as ex:
                                print(f"Erro ao tentar acessar a janela {handle}: {ex}")
                    if boleto_handle:
                        break
                    time.sleep(1)
                if boleto_handle is None:
                    print("❌ Janela do boleto não encontrada, permanecendo na janela atual.")
                time.sleep(3)
            except Exception as e:
                print(f"❌ Erro ao tentar trocar de janela: {e}")
                boleto_handle = None
                # Prossegue com a janela atual

            # Salvar a página do boleto como PDF usando CDP do Chrome
            print("➡️ Salvando a página do boleto como PDF usando CDP do Chrome...")
            try:
                pdf = driver.execute_cdp_cmd("Page.printToPDF", {"printBackground": True})
                print("✅ Comando Page.printToPDF executado com sucesso.")
            except Exception as ex:
                print(f"❌ Erro ao executar Page.printToPDF: {ex}")
                # Se ocorrer erro, continuar para a próxima etapa
                pdf = None
            if pdf:
                try:
                    import base64
                    pdf_data = base64.b64decode(pdf["data"])
                    print("✅ PDF data decodificada com sucesso.")
                except Exception as ex:
                    print(f"❌ Erro ao decodificar PDF data: {ex}")
                    pdf_data = None
            if pdf_data:
                now = datetime.now()
                year = now.strftime("%Y")
                month = now.strftime("%m")
                day = now.strftime("%d")
                download_dir = f"{dir_down}Contabilidade\\Fiscal\\{year}\\LUCRO REAL\\SHEHRAZADE\\{month}.{year}\\ICMS\\ICMS ANTECIPADO E ST\\BAHIA\\AntecipadoST\\{year}\\{month}\\{day}"
                print(f"➡️ Diretório de download: {download_dir}")
                if not os.path.exists(download_dir):
                    print("➡️ Diretório não existe, criando...")
                    os.makedirs(download_dir)
                    print("✅ Diretório criado.")
                else:
                    print("✅ Diretório já existe.")
                # Substituir barras por hífen na descrição para criar um nome de arquivo válido
                safe_descricao = descricao.replace("/", "-")
                file_path = os.path.join(download_dir, f"{safe_descricao}.pdf")
                print(f"➡️ Caminho completo para salvar o PDF: {file_path}")
                try:
                    with open(file_path, "wb") as f:
                        f.write(pdf_data)
                    print(f"✅ PDF salvo em: {file_path}")
                except Exception as ex:
                    print(f"❌ Erro ao salvar PDF: {ex}")
            time.sleep(3)

            # Incrementar a coluna GUIAEMITIDA para os registros deste grupo
            id_list = [r['ID'] for r in registros]
            try:
                increment_guiaemitida(cursor, conn, "FC_AntecipadoBahiaST", id_list)
            except Exception as ex:
                print(f"❌ Erro ao incrementar GUIAEMITIDA: {ex}")
            time.sleep(3)

            # Fechar a janela do boleto e voltar para a janela principal
            try:
                print("➡️ Fechando a janela do boleto e voltando para a janela principal...")
                driver.close()
                driver.switch_to.window(main_handle)
                print("✅ Janela do boleto fechada. Voltando para a janela principal.")
            except Exception as ex:
                print(f"❌ Erro ao fechar a janela do boleto ou voltar para a janela principal: {ex}")
            time.sleep(3)

            #input("Pressione ENTER para continuar...")  # Pausa para revisão antes de enviar

            # No final do grupo, volta até a página da emissão
            navegar_emissao (driver)

        cursor.close()
        conn.close()

    except Exception as e:
        print(f"❌ Erro durante emissão do DAE: {type(e).__name__} - {e}")

def main():

    """ALIMENTAR TABELAS DADOS_EXCEL"""
    import AntecipadosBanco
    AntecipadosBanco.main()

    """CONFIGURAR BROWSER"""
    driver = configure_browser()

    """ECONET"""
    fazer_login(driver)
    fc_antecipadobahia(driver)   
    fc_antecipadobahiast(driver)
    verificar_pendencia_financeira(driver)

    """EMISSÃO GUIA"""
    emissaoantecipado(driver)
    emissaoantecipadost(driver)
    driver.quit()

    """INSERIR DADOS NA TABELA DE GUIAS EMITIDAS"""
    import ExtracaopdfEnviaremail
    ExtracaopdfEnviaremail.main()

if __name__ == "__main__":
    main()
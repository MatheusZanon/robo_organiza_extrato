# =========================IMPORTAÇÕES DE BIBLIOTECAS E COMPONENTES========================
from components.importacao_diretorios_windows import listagem_pastas, listagem_arquivos, pega_nome
from components.extract_text_pdf import extract_text_pdf
from components.importacao_caixa_dialogo import DialogBox
from components.checar_ativacao_google_drive import checa_google_drive
from components.configuracao_db import configura_db
from components.configuracao_selenium_drive import configura_selenium_driver
import tkinter as tk
import mysql.connector
import re
from pathlib import Path
import shutil 
from openpyxl import load_workbook
import win32com.client as win32
from dotenv import load_dotenv
import os
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import  NoSuchElementException
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ================= CARREGANDO VARIÁVEIS DE AMBIENTE======================
load_dotenv()

# =====================CONFIGURAÇÂO DO BANCO DE DADOS======================
db_conf, conn, cursor = configura_db()

# =============CHECANDO SE O GOOGLE FILE STREAM ESTÁ INICIADO NO SISTEMA==============
checa_google_drive()

# ================CONFIGURAÇÃO DO SELENIUM CHROME DRIVER=====================
chrome_options, servico = configura_selenium_driver()
# Configurando Conta de Automação do Nibo
automacao_email = os.getenv('SELENIUM_USER')
automacao_senha = os.getenv('SELENIUM_PASSWORD')

# ==================CAIXA DE DIALOGO INICIAL============================
def main():
    root = tk.Tk()
    app = DialogBox(root)
    root.mainloop()
    return app.particao, app.rotina, app.mes, app.ano

if __name__ == "__main__":
    particao, rotina, mes, ano = main()


# ========================PARAMETROS INICIAS==============================
dir_clientes_itaperuna = f"{particao}:\\Meu Drive\\Cobranca_Clientes_terceirizacao\\Clientes Itaperuna"
dir_clientes_manaus = f"{particao}:\\Meu Drive\\Cobranca_Clientes_terceirizacao\\Clientes Manaus"
lista_dir_clientes = [dir_clientes_itaperuna, dir_clientes_manaus]
dir_extratos = f"{particao}:\\Meu Drive\\Robo_Emissao_Relatorios_do_Mes\\faturas_human_{mes}_{ano}"


# ==================== MÉTODOS DE AUXÍLIO====================================
def procura_cliente(nome_cliente):
    try:
        # PROCURA CLIENTE AO QUAL O EXTRATO PERTENCE
        query_procura_cliente = "SELECT * FROM clientes_financeiro WHERE nome_razao_social = %s"
        values_procura_cliente = (nome_cliente,)
        with mysql.connector.connect(**db_conf) as conn, conn.cursor() as cursor:
            cursor.execute(query_procura_cliente, values_procura_cliente)
            cliente = cursor.fetchone()
            conn.commit()
        if cliente:
            return cliente
        else:
            cliente_novo = procura_cliente(str(nome_cliente).replace("S S", "S/S"))
            if cliente_novo:
                return cliente_novo
    except Exception as error:
        print(error)

def procura_valores(cliente_id):
    try:
        # SE ACHAR O CLIENTE VERIFICA SE OS VALORES DO EXTRATO JÁ NÃO FORAM REGISTRADOS                  
        query_procura_valores = """
                                SELECT * FROM clientes_financeiro_valores WHERE 
                                cliente_id = %s AND mes = %s AND ano = %s 
                                """
        values_procura_valores = (cliente_id, mes, ano)
        with mysql.connector.connect(**db_conf) as conn, conn.cursor() as cursor:
            cursor.execute(query_procura_valores, values_procura_valores)
            valores = cursor.fetchone()
            conn.commit()
        if valores:
            return valores
    except Exception as error:
        print(error)

def procura_pasta_cliente(nome):
    try:
        nome = nome.replace("S/S", "S S")
        caminho_pasta_cliente = ""
        for diretorio in lista_dir_clientes:
            if not caminho_pasta_cliente == "":
                break 
            else:
                pastas_cliente = listagem_pastas(diretorio)
                for pasta in pastas_cliente:
                    if not caminho_pasta_cliente == "":
                        break 
                    else:
                        nome_pasta_cliente = pega_nome(pasta)
                        if nome_pasta_cliente == nome:
                            sub_pastas_cliente = listagem_pastas(pasta)
                            for sub_pasta in sub_pastas_cliente:
                                if sub_pasta.__contains__(f"{mes}-{ano}"):
                                    caminho_pasta_cliente = sub_pasta
                                    break
        return caminho_pasta_cliente
    except Exception as error:
        print(error)

def procura_elemento(driver, xpath, tempo_espera):
    try:
        WebDriverWait(driver, float(tempo_espera)).until(EC.presence_of_element_located((By.XPATH, xpath)))
        time.sleep(0.1)
        elemento = WebDriverWait(driver, float(tempo_espera)).until(EC.visibility_of_element_located((By.XPATH, xpath)))
        if elemento.is_displayed() and elemento.is_enabled():
            return elemento
    except TimeoutException:
        return None

def procura_todos_elementos(driver, xpath, tempo_espera):
    try:
        WebDriverWait(driver, float(tempo_espera)).until(EC.presence_of_all_elements_located((By.XPATH, xpath)))
        time.sleep(0.1)
        elementos = WebDriverWait(driver, float(tempo_espera)).until(EC.visibility_of_all_elements_located((By.XPATH, xpath)))
        return elementos
    except TimeoutException:
        return None

def encontrar_elemento_shadow_root(driver, host, elemento, timeout):
    """Espera por um elemento dentro de um shadow-root até que o elemento esteja presente ou o tempo limite seja atingido."""
    end_time = time.time() + float(timeout)
    while True:
        try:
            # Tenta encontrar o elemento usando JavaScript
            js_script = f"""
            return document.querySelector('{host}').shadowRoot.querySelector('{elemento}');
            """
            element = driver.execute_script(js_script)
            if element:
                return element
        except Exception as e:
            pass  # Ignora erros e tenta novamente até que o tempo limite seja atingido
        time.sleep(0.1)  # Espera 1 segundo antes de tentar novamente
        if time.time() > end_time:
            break  # Sai do loop se o tempo limite for atingido
    return None

def agendar_lancamento(driver, valor_fatura):
    try:
        elemento_agenda_lancamento = procura_elemento(driver, """//*[@id="EntityDetailsContainer"]/h4[1]/a""", 15)
        if elemento_agenda_lancamento:
            elemento_agenda_lancamento.click()
            if int(mes) == 12:
                mes_agenda = "01"
                ano_agenda = str(int(ano) + 1)
            else:
                if int(mes) > 0 and int(mes) < 10: 
                    mes_agenda = "0" + str(int(mes) + 1)
                elif int(mes) > 9:
                    mes_agenda = str(int(mes) + 1)
                ano_agenda = ano
            texto_data_lancamento = f"02/{mes_agenda}/{ano_agenda}"

            # Vencimento
            elemento_vencimento = encontrar_elemento_shadow_root(driver, "#app", """div > ngb-modal-window > div > div > app-receivement-schedule-create >"""+
                                                                 """ div > div.modal-body > form > div:nth-child(2) > div:nth-child(2) > form-helper:nth-child(1)"""+
                                                                 """ > div:nth-child(1) > div > calendar > div > div > div > input""", 2)
            driver.execute_script(f"""arguments[0].value='{texto_data_lancamento}'""", elemento_vencimento)
            time.sleep(0.1)

            # Previsão
            elemento_previsao = encontrar_elemento_shadow_root(driver, "#app", """div > ngb-modal-window > div > div > app-receivement-schedule-create > """+
                                                               """div > div.modal-body > form > div:nth-child(2) > div:nth-child(2) > form-helper:nth-child(2)"""+
                                                               """ > div:nth-child(1) > div > calendar > div > div > div > input""", 2)
            driver.execute_script(f"""arguments[0].value='{texto_data_lancamento}'""", elemento_previsao)
            time.sleep(0.1)

            # Descrição
            elemento_descricao = encontrar_elemento_shadow_root(driver, "#app", "#description", 2)
            driver.execute_script("""arguments[0].value='Salários a pagar, FGTS, GPS, provisão direitos trabalhistas, """+
                                  f"""vale transporte e taxa de administração de pessoas {mes}/{ano}'""", elemento_descricao)
            time.sleep(0.1)

            # Categoria
            elemento_categoria = encontrar_elemento_shadow_root(driver, "#app", """div > ngb-modal-window > div > div > app-receivement-schedule-create """+
                                                                """> div > div.modal-body > form > div:nth-child(2) > div.row.mt-3 > app-schedule-category > """+
                                                                """fieldset > div.ng-untouched.ng-valid.ng-dirty > div > app-schedule-category-item > div > """+
                                                                """form-helper.col-4 > div:nth-child(1) > div > app-category-select > ng-select > div > div > """+
                                                                """div.ng-input > input[type=text]""", 2)
            driver.execute_script(f"""arguments[0].value='Gestão de Mão de Obra Terceirizada'""", elemento_categoria)
            time.sleep(0.1)

            # Valor
            elemento_valor = encontrar_elemento_shadow_root(driver, "#app", """div > ngb-modal-window > div > div > app-receivement-schedule-create """+
                                                            """> div > div.modal-body > form > div:nth-child(2) > div.row.mt-3 > app-schedule-category """+
                                                            """> fieldset > div.ng-untouched.ng-valid.ng-dirty > div > app-schedule-category-item > div > """+
                                                            """div > form-helper > div:nth-child(1) > div > input""", 2)
            driver.execute_script(f"""arguments[0].value='{valor_fatura}'""", elemento_valor)
            time.sleep(0.1)

            # Automatizar Cobrança
            elemento_botao_automatizar = encontrar_elemento_shadow_root(driver, "#app", """div > ngb-modal-window > div > div > app-receivement-schedule-create """+
                                                                        """> div > div.modal-body > form > div:nth-child(2) > div:nth-child(7) > panel-toggle >"""+
                                                                        """ div > div > ui-switch > button""", 2)
            driver.execute_script(f"""arguments[0].click();""", elemento_botao_automatizar)
            time.sleep(0.1)

            # Botao Enviar Imediatamente
            elemento_envio_imediato = encontrar_elemento_shadow_root(driver, "#app", """#entryToday""", 2)
            driver.execute_script(f"""arguments[0].click();""", elemento_envio_imediato)
            time.sleep(0.1)

            # TODO CLICAR NO BOTAO DE AGENDAR LA BOLETA

            # Botão Fechar
            elemento_fechar = encontrar_elemento_shadow_root(driver, "#app", """div > ngb-modal-window > div > div > app-receivement-schedule-create > div > """+
                                                             """modal-header > div > button""", 10)
            driver.execute_script("""return arguments[0].click();""", elemento_fechar)
            time.sleep(0.2)
    except Exception as error:
        print(error)

def baixar_boleto_lancamento(driver, elemento_search):
    try:
        elemento_lista_lancamentos = procura_todos_elementos(driver, """//*[@id="openScheduleList"]/tbody/tr[*]/td[2]/a""", 8)
    except TypeError:
        pass
    try:
        if elemento_lista_lancamentos is not None:
            for elemento in elemento_lista_lancamentos:
                if ("Salários a pagar, FGTS, GPS, provisão direitos trabalhistas, "+
                    f"vale transporte e taxa de administração de pessoas {mes}/{ano}") in elemento.text:
                    elemento.click()
                    time.sleep(0.5)
                    indice = 1
                    achouElemento = False
                    while achouElemento == False:
                        elemento_cobrar_boleto = encontrar_elemento_shadow_root(driver, "#app", f"#ngb-nav-{str(indice)}", 1)
                        if elemento_cobrar_boleto and "Cobrar via boleto" in elemento_cobrar_boleto.text:
                            achouElemento = True
                            driver.execute_script("""return arguments[0].click();""", elemento_cobrar_boleto)
                            time.sleep(0.2)
                            elemento_download = encontrar_elemento_shadow_root(driver, "#app", f"""#ngb-nav-{str(indice)}-panel > settings > div > app-schedule-entry-promise > """+
                                                                            """div > app-entry-promise-details-emitted > div > section > table > tbody > tr > """+
                                                                            """td:nth-child(6) > div > a""", 10)
                            time.sleep(0.2)
                            driver.execute_script("""return arguments[0].click();""", elemento_download)
                        else:
                            indice += 1
                    elemento_fechar = encontrar_elemento_shadow_root(driver, "#app", """div > ngb-modal-window > div > div > app-receivement-schedule-details """+
                                                                    """> div > modal-header > div > button""", 10)
                    driver.execute_script("""return arguments[0].click();""", elemento_fechar)
                    time.sleep(0.2)
                    elemento_search.clear()
                    break
        else:
            elemento_search.clear()
            print("Nenhum lançamento encontrado!")
    except Exception as error:
        print(f"Deu algum erro ao tentar baixar o boleto: {error}")

def start_chrome():
    try:
        # Iniciar o Chrome com as opções configuradas e o serviço
        print("Iniciando navegador chrome...")
        driver = webdriver.Chrome(service=servico, options=chrome_options) 
        actions = ActionChains(driver)
        driver.maximize_window()
        driver.get("https://passport.nibo.com.br/account/login?email=&returnUrl=%2fauthorize%3fresponse_type"+
                   "%3dtoken%26client_id%3d103416FE-A280-466A-9D28-642ACEE21C3B%26lu%3d1%26redirect_uri%3dhttps"+
                   "%253a%252f%252fempresa.nibo.com.br%252fUser%252fLogonWithToken%253freturnUrl%253d%252fOrganization")
        elemento_email = procura_elemento(driver, """//*[@id="Username"]""", 30)
        elemento_email.send_keys(automacao_email)
        elemento_btn_continue = procura_elemento(driver, """//*[@id="continue-button"]""", 30)
        elemento_btn_continue.click()
        elemento_senha = procura_elemento(driver, """//*[@id="Password"]""", 30)
        elemento_senha.send_keys(automacao_senha)
        elemento_btn_entrar = procura_elemento(driver, """//*[@id="password"]/div[3]/input""", 30) 
        elemento_btn_entrar.click()
        return actions, driver
    except Exception as error:
        print(f"Chrome Driver retornou um erro: {error}")
        driver.quit()
        start_chrome()

# ==================== MÉTODOS DE CADA ETAPA DO PROCESSO=======================
def organiza_extratos():
    try:
        pasta_faturas = listagem_pastas(dir_extratos)
        for pasta in pasta_faturas:
            extratos = listagem_arquivos(pasta)
            for extrato in extratos:
                if extrato.__contains__(".pdf"):
                    nome_extrato = pega_nome(extrato)
                    texto_pdf = extract_text_pdf(extrato)

                    # VARREDURA DE DADOS DO EXTRATO PDF
                    # Exemplo de extração para Texto
                    match_centro_custo = re.search(r"C\.Custo:\s*(.*)", texto_pdf)
                    if match_centro_custo:
                        nome_centro_custo = match_centro_custo.group(1).replace("í", "i").replace("ó", "o")
                        #print(f"Centro de Custo: {nome_centro_custo}")
                        partes = nome_centro_custo.split(" - ", 1)
                        if len(partes) > 1:
                            nome_centro_custo_mod = partes[1].strip()
                    
                    # PROCURA CLIENTE AO QUAL O EXTRATO PERTENCE
                    print(nome_centro_custo_mod)
                    cliente = procura_cliente(nome_centro_custo_mod)
                    if cliente:
                        cliente_id = cliente[0]
                        caminho_pasta_cliente = procura_pasta_cliente(nome_centro_custo_mod)
                        valores_extrato = procura_valores(cliente_id)
                        if valores_extrato:
                            print("Esses valores de extrato ja foram registrados!\n")
                            print(f"Cliente: {cliente}\n Caminho da pasta: {caminho_pasta_cliente}\n Extrato: {extrato}\n")
                        else:
                            # CONVÊNIO FÁRMACIA
                            match_convenio_farm = re.search(r"244CONVÊNIO FARMÁCIA\s*([\d.,]+)", texto_pdf)
                            if match_convenio_farm:
                                convenio_farmacia = float(match_convenio_farm.group(1).replace(".", "").replace(",", "."))
                            else:
                                convenio_farmacia = 0
                            #print(f"Convênio Farmácia: {convenio_farmacia}")

                            # DESCONTO ADIANTAMENTO SALARIAL
                            match_adiant_salarial = re.search(r"981DESCONTO ADIANTAMENTO SALARIAL\s*([\d.,]+)", texto_pdf)
                            if match_adiant_salarial:
                                adiant_salarial = float(match_adiant_salarial.group(1).replace(".", "").replace(",", "."))
                            else: 
                                adiant_salarial = 0
                            #print(f"Desconto Adiantamento Salarial: {adiant_salarial}")

                            # NUMERO DE EMPREGADOS
                            # Exemplo de extração para Números
                            match_demitido = re.search(r"No. Empregados: Demitido:\s*(\d+)", texto_pdf)
                            if match_demitido:
                                demitido = match_demitido.group(1)
                                match_num_empregados = re.search(r"No. Empregados: Demitido:\s+" + demitido + 
                                                                r"\s*(\d+)", texto_pdf)
                                if match_num_empregados:
                                    num_empregados = match_num_empregados.group(1)
                                else: 
                                    num_empregados = 0 
                            else:
                                num_empregados = 0
                            #print(f"Número de empregados: {num_empregados}")

                            # NUMERO DE ESTAGIARIOS
                            match_transferido = re.search(r"No. Estagiários: Transferido:\s*(\d+)", texto_pdf)
                            if match_transferido:
                                transferido = match_transferido.group(1)
                                match_num_estagiarios = re.search(r"No. Estagiários: Transferido:\s+" + transferido + 
                                                                r"\s*(\d+)", texto_pdf)
                                if match_num_estagiarios:
                                    num_estagiarios = match_num_estagiarios.group(1)
                                else: 
                                    num_estagiarios = 0
                            else:
                                num_estagiarios = 0
                            #print(f"Número de estagiários: {num_estagiarios}")

                            # TRABALHANDO
                            match_ferias = re.search(r"Trabalhando: Férias:\s*(\d+)", texto_pdf)
                            if match_ferias:
                                ferias = match_ferias.group(1)
                                match_trabalhando = re.search(r"Trabalhando: Férias:\s+" + ferias + r"\s*(\d+)", texto_pdf)
                                if match_trabalhando:
                                    trabalhando = match_trabalhando.group(1)
                                else:
                                    trabalhando = 0
                            else:
                                trabalhando = 0
                            #print(f"Trabalhando: {trabalhando}")

                            # SALARIO CONTRIBUIÇÃO EMPREGADOS
                            match_salario_contri_empregados = re.search(r"Salário contribuição empregados:\s*([\d.,]+)", texto_pdf)
                            if  match_salario_contri_empregados:
                                salario_contri_empregados = float(match_salario_contri_empregados
                                                                .group(1).replace(".", "").replace(",", "."))
                            else: 
                                salario_contri_empregados = 0
                            #print(f"Salário contribuição Empregados: {salario_contri_empregados}")

                            # SALARIO CONTRIBUIÇÃO CONTRIBUINTES
                            match_salario_contri_contribuintes = re.search(r"Salário contribuição contribuintes:\s*([\d.,]+)", 
                                                                        texto_pdf)
                            if  match_salario_contri_contribuintes:
                                salario_contri_contribuintes = float(match_salario_contri_contribuintes
                                                                    .group(1).replace(".", "").replace(",", "."))
                            else:
                                salario_contri_contribuintes = 0
                            #print(f"Salário contribuição Contribuintes: {salario_contri_contribuintes}")
                            
                            # SOMA DOS SALARIOS
                            soma_salarios_provdt = salario_contri_empregados + salario_contri_contribuintes
                            #print(f"Soma dos salários: {soma_salarios_provdt}")

                            # VALOR DO INSS
                            # A expressão regular procura por um ou mais números seguidos por qualquer coisa (não capturada)
                            # e então "Total INSS:"
                            match_inss = re.search(r"Total INSS:\s*([\d.,]+)", texto_pdf)
                            if match_inss:
                                inss = float(match_inss.group(1).replace(".", "").replace(",", "."))
                            else:
                                inss = 0
                            #print(f"Total INSS: {inss}")

                            # VALOR DO FGTS
                            match_fgts = re.search(r"Valor do FGTS:\s*([\d.,]+)", texto_pdf)
                            if  match_fgts:
                                fgts = float(match_fgts.group(1).replace(".", "").replace(",", "."))
                            else:
                                fgts = 0
                            #print(f"Valor do FGTS: {fgts}")

                            # VALOR DO IRRF
                            match_base_iss = re.search(r"([\d.,]+)\s+Valor Total do IRRF: Base ISS:", texto_pdf)
                            if match_base_iss:
                                base_iss = match_base_iss.group(1)
                                match_irrf = re.search(r"([\d.,]+)\s+" + base_iss + r"\s+Valor Total do IRRF: Base ISS:", texto_pdf)
                                if match_irrf:
                                    irrf = float(match_irrf.group(1).replace(".", "").replace(",", "."))
                                else:
                                    irrf = 0
                            else:
                                irrf = 0
                            #print(f"Valor Total do IRRF: {irrf}")

                            # LÍQUIDO CENTRO DE CUSTO
                            match_liquido = re.search(r"Líquido Centro de Custo:\s*([\d.,]+)", texto_pdf)
                            if  match_liquido:
                                liquido_centro_custo = float(match_liquido.group(1).replace(".", "").replace(",", "."))
                            else:
                                liquido_centro_custo = 0
                            #print(f"Líquido Centro de Custo: {liquido_centro_custo}")
                            # LIQUIDO CENTRO DE CUSTO ENTRA NA COLUNA SALARIOS_PAGAR DO BANCO


                            # INSERÇÃO DE DADOS NO BANCO
                            query_insert_valores = """INSERT INTO clientes_financeiro_valores 
                                                    (cliente_id, convenio_farmacia, adiant_salarial, numero_empregados, 
                                                    numero_estagiarios, trabalhando, salario_contri_empregados, 
                                                    salario_contri_contribuintes, soma_salarios_provdt, inss, fgts, irrf, 
                                                    salarios_pagar, mes, ano)
                                                    VALUES (%s, %s,  %s,  %s,  %s,  %s,  %s,  %s,  %s,  %s,  %s,  %s, %s, %s, %s)
                                                    """
                            values_insert_valores = (cliente_id, convenio_farmacia, adiant_salarial, num_empregados, 
                                                        num_estagiarios, trabalhando, salario_contri_empregados, 
                                                        salario_contri_contribuintes, soma_salarios_provdt, inss, fgts, 
                                                        irrf, liquido_centro_custo, mes, ano
                                                        )
                            with mysql.connector.connect(**db_conf) as conn, conn.cursor() as cursor:
                                cursor.execute(query_insert_valores, values_insert_valores)
                                conn.commit()

                            # Caminho do arquivo PDF
                            caminho_pdf = Path(extrato)
                            if not nome_extrato.__contains__(f"Extrato_Mensal_{nome_centro_custo.replace("S/S", "S S")}_{mes}.{ano}"):
                                novo_nome_extrato = caminho_pdf.with_name(f"Extrato_Mensal_{nome_centro_custo.replace("S/S", "S S")}_{mes}.{ano}.pdf")
                                caminho_pdf_mod = caminho_pdf.rename(novo_nome_extrato)
                            else:
                                caminho_pdf_mod = caminho_pdf
                            # Caminho da pasta de destino (o caminho que vem da sua variável)
                            caminho_destino = Path(caminho_pasta_cliente)
                            # Verifica se a pasta de destino existe; se não, cria a pasta
                            caminho_destino.mkdir(parents=True, exist_ok=True)
                            # Copiar o arquivo PDF para a pasta de destino
                            shutil.copy(caminho_pdf_mod, caminho_destino / caminho_pdf_mod.name)
                    else:
                        print("Cliente não encontrado!\n")

                    
    except Exception as error:
        if error.args == ("'NoneType' object is not iterable",):
            print("O diretório informado não foi especificado!")
        else:
            print(f"O sistema retornou um erro: {error}")

def gera_fatura():
    try:
        modelo_fatura = Path(f"{particao}:\\Meu Drive\\Arquivos_Automacao\\Fatura_Detalhada_Modelo_00.0000_python.xlsx")
        fatura_pronta = False

        for diretorio in lista_dir_clientes:
            pastas_regioes = listagem_pastas(diretorio)
            for pasta_cliente in pastas_regioes:
                nome_pasta_cliente = pega_nome(pasta_cliente)
                sub_pastas_cliente = listagem_pastas(pasta_cliente)
                for sub_pasta in sub_pastas_cliente:
                    if sub_pasta.__contains__(f"{mes}-{ano}"):
                        arquivos_cliente = listagem_arquivos(sub_pasta)
                        for arquivo in arquivos_cliente:
                            if (arquivo.__contains__("Fatura_Detalhada_") 
                            and arquivo.__contains__(nome_pasta_cliente)
                            and arquivo.__contains__(".pdf")):
                                fatura_pronta = True
                                break
                            else:
                                fatura_pronta = False
                        if fatura_pronta == True:
                            fatura_pronta = False
                            break
                        elif fatura_pronta == False:
                            print(pasta_cliente, fatura_pronta)
                            cliente = procura_cliente(nome_pasta_cliente)
                            if cliente:
                                cliente_id = cliente[0]
                                valores_financeiro = procura_valores(cliente_id)
                                if valores_financeiro:
                                    caminho_sub_pasta = Path(sub_pasta)

                                    # Variáveis para planilha
                                    nome_fatura = f"Fatura_Detalhada_{nome_pasta_cliente}_{mes}.{ano}.xlsx"
                                    caminho_fatura = f"{caminho_sub_pasta}\\{nome_fatura}"
                                    
                                    # COPIANDO A FATURA MODELO PARA A PASTA DO CLIENTE
                                    shutil.copy(modelo_fatura, caminho_sub_pasta / nome_fatura)
                                   
                                    # FORMATANDO A FATURA
                                    try:
                                        workbook = load_workbook(caminho_fatura)
                                        sheet = workbook.active
                                        # nome da planilha (em baixo)
                                        sheet.title = f"{mes}.{ano}"
                                        # titulo da fatura
                                        sheet['D2'] = f"Fatura Detalhada - {nome_pasta_cliente}"
                                        # numero de funcionarios
                                        if valores_financeiro[4] == 1:
                                            sheet['J6'] = 1
                                            sheet['K6'] = 'funcionário'
                                        else:
                                            sheet['J6'] = valores_financeiro[4] + valores_financeiro[5]
                                        # salários a pagar
                                        sheet['A7'] = f"Salários a pagar {mes}.{ano}"
                                        sheet['J7'] = valores_financeiro[13]
                                        salarios_pagar = valores_financeiro[13]
                                        # inss
                                        sheet['A8'] = f"GPS (Guia da Previdência Social) {mes}.{ano}"
                                        sheet['J8'] = valores_financeiro[10]
                                        inss = valores_financeiro[10]
                                        # fgts
                                        sheet['A9'] = f"FGTS (Fundo de Garantia por Tempo de Serviço) {mes}.{ano}"
                                        sheet['J9'] = valores_financeiro[11]
                                        fgts = valores_financeiro[11]
                                        # adiantamento salarial
                                        if not valores_financeiro[3] == None:
                                            sheet['J10'] = valores_financeiro[3]
                                            adiant_salarial = valores_financeiro[3]
                                        else:
                                            adiant_salarial = 0
                                        # provisão de direitos trabalhistas
                                        sheet['A11'] = f"Provisão de Direitos Trabalhistas {mes}.{ano}"
                                        sheet['E11'] = valores_financeiro[9]
                                        soma_salarios_provdt = valores_financeiro[9]
                                        # irrf (folha de pagamento)
                                        if not valores_financeiro[12] == None:
                                            sheet['J12'] = valores_financeiro[12]
                                            irrf = valores_financeiro[12]
                                        else:
                                            irrf = 0
                                        # mensalidade do ponto eletrônico
                                        if not valores_financeiro[14] == None:
                                            sheet['J13'] = valores_financeiro[14]
                                            mensal_ponto = valores_financeiro[14]
                                        else:
                                            mensal_ponto = 0
                                        # assinatura eletrônica
                                        if not valores_financeiro[15] == None:
                                            sheet['J14'] = valores_financeiro[15]
                                            assinatura_elet = valores_financeiro[15]
                                        else:
                                            assinatura_elet = 0
                                        # vale transporte
                                        sheet['A15'] = f"Vale Transporte {mes}/{ano}"
                                        if not valores_financeiro[16] == None:
                                            sheet['J15'] = valores_financeiro[16]
                                            vale_transp = valores_financeiro[16]
                                        else:
                                            vale_transp = 0
                                        # vale refeição
                                        sheet['A16'] = f"Vale Refeição {mes}/{ano}"
                                        if not valores_financeiro[17] == None:
                                            sheet['J16'] = valores_financeiro[17]
                                            vale_refeic = valores_financeiro[17]
                                        else:
                                            vale_refeic = 0
                                        # saúde e segurança do trabalho
                                        if not valores_financeiro[18] == None:
                                            sheet['J17'] = valores_financeiro[18] 
                                            sst = valores_financeiro[18]
                                        else:
                                            sst = 0
                                        # convênio farmácia
                                        if not valores_financeiro[2] == None:
                                            sheet['J18'] = valores_financeiro[2]
                                            conv_farmacia = valores_financeiro[2]
                                        else:
                                            conv_farmacia = 0
                                        # provisao de direitos trabalhistas
                                        prov_direitos = soma_salarios_provdt * 0.3487
                                        # percentual human
                                        percent_human = soma_salarios_provdt * 0.12
                                        # economia mensal
                                        valor1 = round(soma_salarios_provdt * 0.8027, 2)
                                        valor2 = round(soma_salarios_provdt * 0.4287, 2)
                                        eco_mensal = valor1 - valor2
                                        workbook.save(caminho_fatura)
                                        workbook.close()

                                        # valor total da fatura
                                        fatura = (salarios_pagar + inss + fgts + adiant_salarial + prov_direitos
                                                  + irrf + mensal_ponto + assinatura_elet + vale_transp + vale_refeic
                                                  + sst + conv_farmacia + percent_human
                                                 )
                                        total_fatura = round(fatura, 2)
                                        print(f"Total da Fatura: {total_fatura}")
                                        
                                        # GERANDO PDF DA FATURA
                                        try:
                                            print("Terminada a formatação da fatura, gerando pdf...")
                                            excel = win32.gencache.EnsureDispatch('Excel.Application')
                                            excel.Visible = True
                                            wb = excel.Workbooks.Open(caminho_fatura)
                                            ws = wb.Worksheets[f"{mes}.{ano}"]
                                            time.sleep(3)
                                            ws.ExportAsFixedFormat(0, sub_pasta + f"\\Fatura_Detalhada_{nome_pasta_cliente}_{mes}.{ano}")
                                            wb.Close()
                                            excel.Quit()
                                            print("Pdf gerado com sucesso.")

                                            # INSERINDO VALOR DA FATURA NO BANCO
                                            print("Inserindo valores no banco.")
                                            query_fatura = """UPDATE clientes_financeiro_valores SET percentual_human = %s,
                                                            economia_mensal = %s, total_fatura = %s WHERE 
                                                            cliente_id = %s AND mes = %s AND ano = %s
                                                        """
                                            with mysql.connector.connect(**db_conf) as conn, conn.cursor() as cursor:
                                                cursor.execute(query_fatura, (percent_human, eco_mensal, total_fatura, cliente_id, mes, ano))
                                                conn.commit()
                                            print("Valores inseridos com sucesso!")
                                        except Exception as error:
                                            wb.Close()
                                            wb.Quit()
                                            print(error)
                                    except Exception as error:
                                        print(error)
                                else: 
                                    print("Cliente não possue valores para gerar fatura!")
                            else:
                                print("Cliente não encontrado!")
    except Exception as error:
        return (error)

def gera_boleto(): 
    try:
        boleto = False
        actions, driver = start_chrome()
        time.sleep(1)
        elemento_cards = procura_todos_elementos(driver, """//*[@id="myorganizations-container"]"""+
                                                """/div/div[3]/ng-include[2]/div[*]/a/h3/span""", 10)
        for card in elemento_cards:
            if card.text == "HUMAN SOLUCOES E DESENVOLVIMENTOS EM RECURSOS HUMANOS LTDA":
                card.click()
                break
        time.sleep(1.5)
        elemento_contatos = procura_elemento(driver, """//*[@id="page-organization-details"]"""+
                                                """/div[5]/div/div[1]/div/div/ul[2]/li[3]/a""", 10)
        if elemento_contatos:
            actions.move_to_element(elemento_contatos).perform()
            elemento_clientes = procura_elemento(driver, """//*[@id="page-organization-details"]"""+
                                                    """/div[5]/div/div[1]/div/div/ul[2]"""+
                                                    """/li[3]/ul/li[1]/a""", 10)
            elemento_clientes.click()
            time.sleep(1)
    except Exception as web_error:
        print (web_error)
    try:
        for diretorio in lista_dir_clientes:
            pastas_regioes = listagem_pastas(diretorio)
            for pasta_cliente in pastas_regioes:
                nome_pasta_cliente = pega_nome(pasta_cliente)
                sub_pastas_cliente = listagem_pastas(pasta_cliente)
                for sub_pasta in sub_pastas_cliente:
                    if sub_pasta.__contains__(f"{mes}-{ano}"):
                        arquivos_cliente = listagem_arquivos(sub_pasta)
                        for arquivo in arquivos_cliente:
                            if arquivo.__contains__("Boleto_Recebimento_") and arquivo.__contains__(nome_pasta_cliente):
                                boleto = True
                                break
                            else:
                                boleto = False
                        if boleto == True:
                            boleto = False
                            break
                        elif boleto == False:
                            cliente = procura_cliente(nome_pasta_cliente)
                            if cliente:
                                cliente_id = cliente[0]
                                cliente_cnpj = cliente[2]
                                cliente_cpf = cliente[3]
                                valores = procura_valores(cliente_id)
                                if valores:
                                    valor_fatura = valores[21]
                                    print(f"{nome_pasta_cliente} vai precisar de um boleto. Valor da fatura é: {valor_fatura}")
                                    elemento_search = procura_elemento(driver, """//*[@id="entityList_filter"]"""+
                                                                      """/label/input""", 10)   
                                    if elemento_search:                           
                                        if not cliente_cnpj == '' or not cliente_cnpj == None:
                                            elemento_search.send_keys(cliente_cnpj)
                                        elif not cliente_cpf == '' or not cliente_cpf == None:
                                            elemento_search.send_keys(cliente_cpf)                       
                                    time.sleep(2)
                                    try:
                                        elemento_lista_clientes = procura_todos_elementos(driver,"""//*[@id="entityList"]"""+
                                                                    """/tbody/tr/td[1]/a""" , 15)
                                    except NoSuchElementException:
                                        elemento_lista_clientes = procura_todos_elementos(driver,"""//*[@id="entityList"]"""+
                                                                    """/tbody/tr[*]/td[1]/a""" , 15)
                                    for cliente_lista in elemento_lista_clientes:
                                        if cliente_lista.text.__contains__(cliente_cnpj) or cliente_lista.text.__contains__(cliente_cpf):
                                            cliente_lista.click()
                                            try:
                                                elemento_sem_lancamento = procura_elemento(driver, """/html/body/div[*]/div/div[2]/div[3]"""+
                                                                                        """/div/div/div[2]/div[2]/div/div[3]/div/div/p""", 5)

                                                if elemento_sem_lancamento:
                                                    time.sleep(0.7)
                                                    agendar_lancamento(driver, valor_fatura)
                                                    time.sleep(1.5)
                                                    baixar_boleto_lancamento(driver, elemento_search)
                                                elif elemento_sem_lancamento == None:
                                                    time.sleep(0.7)
                                                    baixar_boleto_lancamento(driver, elemento_search)
                                            except NoSuchElementException:
                                                print("Algum objeto nao foi encontrado!")
                                else:
                                    print(f"Valores de financeiro não encontrados para {nome_pasta_cliente}")
                            else:
                                print(f"Cliente {nome_pasta_cliente} não encontrado!")                    
    except Exception as error:
        print(error)
    print("PROCESSO DE BOLETO ENCERRADO!")
    input()   
    driver.quit() 

def envia_arquivos():
    try:    
        print("Processo de enviar os arquivos para cada cliente (extrato, fatura, boleto)")
    except Exception as error:
        print (error)

# ========================LÓGICA DE EXECUÇÃO DO ROBÔ===========================
if rotina == "1. Organizar Extratos":
    organiza_extratos()
    gera_fatura()
    gera_boleto()
    envia_arquivos()
elif rotina == "2. Gerar Fatura Detalhada":
    gera_fatura()
    gera_boleto()
    envia_arquivos()
elif rotina == "3. Gerar Boletos":
    gera_boleto()
    envia_arquivos()
elif rotina == "4. Enviar Arquivos":
    envia_arquivos()
else:
    print("Nenhuma rotina selecionada, encerrando o robô.")
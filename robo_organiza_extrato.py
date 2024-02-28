# =========================IMPORTAÇÕES DE BIBLIOTECAS E COMPONENTES========================
from components.importacao_diretorios_windows import listagem_pastas, listagem_arquivos,listagem_arquivos_downloads, pega_nome
from components.extract_text_pdf import extract_text_pdf
from components.importacao_automacao_excel_pandas import carrega_arquivo
from components.importacao_caixa_dialogo import DialogBox
from components.checar_ativacao_google_drive import checa_google_drive
from components.configuracao_db import configura_db, ler_sql
from components.configuracao_selenium_drive import configura_selenium_driver
from components.enviar_emails import enviar_email_com_anexos
import tkinter as tk
import mysql.connector
from re import search
from pathlib import Path
from shutil import copy
from openpyxl import load_workbook
from openpyxl.styles import Border, Side, NamedStyle
import win32com.client as win32
from dotenv import load_dotenv
import os
from time import sleep, time
import pandas as pd
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
modelo_fatura = Path(f"{particao}:\\Meu Drive\\Arquivos_Automacao\\Fatura_Detalhada_Modelo_00.0000_python.xlsx")
planilha_vales_sst = Path(f"{particao}:\\Meu Drive\\Relatorio_Vales_Saude_Seguranca\\{mes}-{ano}\\Relatorio_Vales_Saude_Seguranca_{mes}.{ano}.xlsx")
planilha_reembolsos = Path(f"{particao}:\\Meu Drive\\Relatorio_Boletos_Salario_Reembolso\\{mes}-{ano}\\Relatorio_Boletos_Salario_Reembolso.xlsx")

# ==================== MÉTODOS DE AUXÍLIO====================================
def pega_valores_vales_reembolsos(cliente_id, centro_custo):
    try:
        print(centro_custo)
        df_vales_sst = pd.read_excel(planilha_vales_sst, usecols='C:H', skiprows=1)
        vales = df_vales_sst.loc[df_vales_sst['CLIENTE'] == centro_custo, ['Vale Transporte', 'Assinatura Eletronica', 'Vale Refeição', 'Ponto Eletrônico', 'Saúde/Segurança do Trabalho']]
        if not vales.empty:
            vale_transporte = str(vales['Vale Transporte'].values[0]).replace("R$", "").replace(",", ".")
            assinat_eletronica = str(vales['Assinatura Eletronica'].values[0]).replace("R$", "").replace(",", ".")
            vale_refeicao = str(vales['Vale Refeição'].values[0]).replace("R$", "").replace(",", ".")
            ponto_eletronico = str(vales['Ponto Eletrônico'].values[0]).replace("R$", "").replace(",", ".")
            sst = str(vales['Saúde/Segurança do Trabalho'].values[0]).replace("R$", "").replace(",", ".")
            print(vale_transporte, assinat_eletronica, vale_refeicao, ponto_eletronico, sst)
        else:
            vale_transporte = 0
            assinat_eletronica = 0
            vale_refeicao = 0
            ponto_eletronico = 0
            sst = 0

        df_reembolsos = pd.read_excel(planilha_reembolsos, usecols='B:D', skiprows=1)
        reembolsos = df_reembolsos[(df_reembolsos['CLIENTE'] == centro_custo) & (df_reembolsos['Descrição'].notnull()) & (df_reembolsos['Valor'].notnull())]
        descricao_reembolsos = reembolsos['Descrição'].tolist()
        valores_reembolsos = reembolsos['Valor'].tolist()
        print(descricao_reembolsos, valores_reembolsos)
        input()
        if not descricao_reembolsos == [] and not valores_reembolsos == []:
            for i in range(len(valores_reembolsos)):
                query_insert_reembolsos = ler_sql('sql/registrar_valores_reembolsos.sql')
                values = (cliente_id, descricao_reembolsos[i], valores_reembolsos[i], mes, ano)
                with mysql.connector.connect(**db_conf) as conn, conn.cursor() as cursor:
                    cursor.execute(query_insert_reembolsos, values)
                    conn.commit()
        return vale_transporte, assinat_eletronica, vale_refeicao, ponto_eletronico, sst
    except Exception as error:
        print(error)

def procura_cliente(nome_cliente):
    try:
        query_procura_cliente = ler_sql('sql/procura_cliente.sql')
        values_procura_cliente = (nome_cliente,)
        with mysql.connector.connect(**db_conf) as conn, conn.cursor() as cursor:
            cursor.execute(query_procura_cliente, values_procura_cliente)
            cliente = cursor.fetchone()
            conn.commit()
        if cliente:
            return cliente
        else:
            cliente_mod = procura_cliente_mod(str(nome_cliente).replace("S S", "S/S"))
            return cliente_mod
    except Exception as error:
        print(error)

def procura_cliente_mod(nome_cliente):
    try:
        query_procura_cliente = ler_sql('sql/procura_cliente.sql')
        values_procura_cliente = (nome_cliente,)
        with mysql.connector.connect(**db_conf) as conn, conn.cursor() as cursor:
            cursor.execute(query_procura_cliente, values_procura_cliente)
            cliente = cursor.fetchone()
            conn.commit()
        if cliente:
            return cliente
    except Exception as error:
        print(error)

def procura_valores(cliente_id):
    try:                
        query_procura_valores = ler_sql('sql/procura_valores_financeiro.sql')
        values_procura_valores = (cliente_id, mes, ano)
        with mysql.connector.connect(**db_conf) as conn, conn.cursor() as cursor:
            cursor.execute(query_procura_valores, values_procura_valores)
            valores = cursor.fetchall()
            conn.commit()
        if valores and len(valores) == 1:
            valores_tupla = valores[0]
            return valores_tupla
        elif valores and len(valores) >= 1:
            query_valores = ler_sql('sql/soma_valores_multiplos.sql')
            values = (cliente_id, mes, ano)
            with mysql.connector.connect(**db_conf) as conn, conn.cursor() as cursor:
                cursor.execute(query_valores, values)
                valores = cursor.fetchone()
                conn.commit()
            return valores
    except Exception as error:
        print(error)

def procura_valores_com_codigo(cliente_id, cod_centro_custo):
    try:                
        query_procura_valores = ler_sql('sql/procura_valor_com_codigo_empresa.sql')
        values_procura_valores = (cliente_id, cod_centro_custo, mes, ano)
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

def procura_elemento(driver, tipo_seletor:str, elemento, tempo_espera):
    """
        Function to search for an element using the specified selector type, element, and wait time.
        driver: The WebDriver instance to use for element search.
        tipo_seletor: The type of selector to use (e.g., 'ID', 'CLASS_NAME', 'XPATH', 'TAG_NAME').
        elemento: The element to search for.
        tempo_espera: The maximum time to wait for the element to be located.
        :return: The located element if found, otherwise None.
    """
    try:
        seletor = getattr(By, tipo_seletor.upper())
        WebDriverWait(driver, float(tempo_espera)).until(EC.presence_of_element_located((seletor, elemento)))
        sleep(0.1)
        elemento = WebDriverWait(driver, float(tempo_espera)).until(EC.visibility_of_element_located((seletor, elemento)))
        if elemento.is_displayed() and elemento.is_enabled():
            return elemento
    except TimeoutException:
        return None

def procura_todos_elementos(driver, tipo_seletor:str, elemento, tempo_espera):
    """
    A function that searches for all elements based on the given selector type and element, within a specified waiting time.
    
    Args:
        driver: The WebDriver instance to use for locating the elements.
        tipo_seletor: A string representing the type of selector to use (e.g., 'ID', 'CLASS_NAME', 'XPATH', 'TAG_NAME').
        elemento: The element to search for.
        tempo_espera: The maximum time to wait for the elements to be present before throwing a TimeoutException.
        
    Returns:
        A list of WebElement objects representing the found elements, or None if the elements are not found within the specified waiting time.
    """
    try:
        seletor = getattr(By, tipo_seletor.upper())
        WebDriverWait(driver, float(tempo_espera)).until(EC.presence_of_all_elements_located((seletor, elemento)))
        sleep(0.1)
        elementos = WebDriverWait(driver, float(tempo_espera)).until(EC.visibility_of_all_elements_located((seletor, elemento)))
        if elementos:
            return elementos
    except TimeoutException:
        return None

def encontrar_elemento_shadow_root(driver, host, elemento, timeout):
    """Espera por um elemento dentro de um shadow-root até que o elemento esteja presente ou o tempo limite seja atingido."""
    end_time = time() + float(timeout)
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
        sleep(0.1)  # Espera 1 segundo antes de tentar novamente
        if time() > end_time:
            break  # Sai do loop se o tempo limite for atingido
    return None

def agendar_lancamento(driver, valor_fatura):
    print("AGENDANDO LANÇAMENTO")
    try:
        elemento_agenda_lancamento = procura_elemento(driver, "xpath", """//*[@id="EntityDetailsContainer"]/h4[1]/a""", 15)
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
            sleep(0.1)

            # Previsão
            elemento_previsao = encontrar_elemento_shadow_root(driver, "#app", """div > ngb-modal-window > div > div > app-receivement-schedule-create > """+
                                                               """div > div.modal-body > form > div:nth-child(2) > div:nth-child(2) > form-helper:nth-child(2)"""+
                                                               """ > div:nth-child(1) > div > calendar > div > div > div > input""", 2)
            driver.execute_script(f"""arguments[0].value='{texto_data_lancamento}'""", elemento_previsao)
            sleep(0.1)

            # Descrição
            elemento_descricao = encontrar_elemento_shadow_root(driver, "#app", "#description", 2)
            driver.execute_script("""arguments[0].value='Salários a pagar, FGTS, GPS, provisão direitos trabalhistas, """+
                                  f"""vale transporte e taxa de administração de pessoas {mes}/{ano}'""", elemento_descricao)
            sleep(0.1)

            # Categoria
            elemento_categoria = encontrar_elemento_shadow_root(driver, "#app", """div > ngb-modal-window > div > div > app-receivement-schedule-create """+
                                                                """> div > div.modal-body > form > div:nth-child(2) > div.row.mt-3 > app-schedule-category > """+
                                                                """fieldset > div.ng-untouched.ng-valid.ng-dirty > div > app-schedule-category-item > div > """+
                                                                """form-helper.col-4 > div:nth-child(1) > div > app-category-select > ng-select > div > div > """+
                                                                """div.ng-input > input[type=text]""", 2)
            driver.execute_script(f"""arguments[0].value='Gestão de Mão de Obra Terceirizada'""", elemento_categoria)
            sleep(0.1)

            # Valor
            elemento_valor = encontrar_elemento_shadow_root(driver, "#app", """div > ngb-modal-window > div > div > app-receivement-schedule-create """+
                                                            """> div > div.modal-body > form > div:nth-child(2) > div.row.mt-3 > app-schedule-category """+
                                                            """> fieldset > div.ng-untouched.ng-valid.ng-dirty > div > app-schedule-category-item > div > """+
                                                            """div > form-helper > div:nth-child(1) > div > input""", 2)
            driver.execute_script(f"""arguments[0].value='{valor_fatura}'""", elemento_valor)
            sleep(0.1)

            # Automatizar Cobrança
            elemento_botao_automatizar = encontrar_elemento_shadow_root(driver, "#app", """div > ngb-modal-window > div > div > app-receivement-schedule-create """+
                                                                        """> div > div.modal-body > form > div:nth-child(2) > div:nth-child(7) > panel-toggle >"""+
                                                                        """ div > div > ui-switch > button""", 2)
            driver.execute_script(f"""arguments[0].click();""", elemento_botao_automatizar)
            sleep(0.1)

            # Botao Enviar Imediatamente
            elemento_envio_imediato = encontrar_elemento_shadow_root(driver, "#app", """#entryToday""", 2)
            driver.execute_script(f"""arguments[0].click();""", elemento_envio_imediato)
            sleep(0.1)

            # TODO CLICAR NO BOTAO DE AGENDAR LA BOLETA

            # Botão Fechar
            elemento_fechar = encontrar_elemento_shadow_root(driver, "#app", """div > ngb-modal-window > div > div > app-receivement-schedule-create > div > """+
                                                             """modal-header > div > button""", 10)
            driver.execute_script("""return arguments[0].click();""", elemento_fechar)
            sleep(0.2)
    except Exception as error:
        print(error)

def baixar_boleto_lancamento(driver, elemento_search):
    print("BAIXANDO BOLETO")
    try:
        elemento_lista_lancamentos = procura_todos_elementos(driver, "xpath", """//*[@id="openScheduleList"]/tbody/tr[*]/td[2]/a""", 8)
    except TypeError:
        pass
    try:
        if elemento_lista_lancamentos is not None:
            for elemento in elemento_lista_lancamentos:
                if ("Salários a pagar, FGTS, GPS, provisão direitos trabalhistas, "+
                    f"vale transporte e taxa de administração de pessoas {mes}/{ano}") in elemento.text:
                    elemento.click()
                    sleep(0.5)
                    indice = 1
                    achouElemento = False
                    while achouElemento == False:
                        elemento_cobrar_boleto = encontrar_elemento_shadow_root(driver, "#app", f"#ngb-nav-{str(indice)}", 1)
                        if elemento_cobrar_boleto and "Cobrar via boleto" in elemento_cobrar_boleto.text:
                            achouElemento = True
                            driver.execute_script("""return arguments[0].click();""", elemento_cobrar_boleto)
                            sleep(0.2)
                            elemento_download = encontrar_elemento_shadow_root(driver, "#app", f"""#ngb-nav-{str(indice)}-panel > settings > div > app-schedule-entry-promise > """+
                                                                            """div > app-entry-promise-details-emitted > div > section > table > tbody > tr > """+
                                                                            """td:nth-child(6) > div > a""", 10)
                            sleep(0.2)
                            driver.execute_script("""return arguments[0].click();""", elemento_download)
                        else:
                            indice += 1
                    elemento_fechar = encontrar_elemento_shadow_root(driver, "#app", """div > ngb-modal-window > div > div > app-receivement-schedule-details """+
                                                                    """> div > modal-header > div > button""", 10)
                    driver.execute_script("""return arguments[0].click();""", elemento_fechar)
                    sleep(0.2)
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
        elemento_email = procura_elemento(driver, "xpath", """//*[@id="Username"]""", 30)
        elemento_email.send_keys(automacao_email)
        elemento_btn_continue = procura_elemento(driver, "xpath","""//*[@id="continue-button"]""", 30)
        elemento_btn_continue.click()
        elemento_senha = procura_elemento(driver, "xpath", """//*[@id="Password"]""", 30)
        elemento_senha.send_keys(automacao_senha)
        elemento_btn_entrar = procura_elemento(driver, "xpath", """//*[@id="password"]/div[3]/input""", 30) 
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

                    # Nome do Centro de Custo
                    match_centro_custo = search(r"C\.Custo:\s*(.*)", texto_pdf)
                    if match_centro_custo:
                        nome_centro_custo = match_centro_custo.group(1).replace("í", "i").replace("ó", "o")
                        partes = nome_centro_custo.split(" - ", 1)
                        if len(partes) > 1:
                            nome_centro_custo_mod = partes[1].strip()
                            cod_centro_custo = partes[0].strip()
                    
                    print(nome_centro_custo)

                    cliente = procura_cliente(nome_centro_custo_mod)
                    if cliente:
                        cliente_id = cliente[0]
                        caminho_pasta_cliente = procura_pasta_cliente(nome_centro_custo_mod)
                        valores_extrato = procura_valores_com_codigo(cliente_id, cod_centro_custo)
                        if valores_extrato:
                            print("Esses valores de extrato ja foram registrados!\n")
                        else:
                            # CONVÊNIO FÁRMACIA
                            match_convenio_farm = search(r"244CONVÊNIO FARMÁCIA\s*([\d.,]+)", texto_pdf)
                            if match_convenio_farm:
                                convenio_farmacia = float(match_convenio_farm.group(1).replace(".", "").replace(",", "."))
                            else:
                                convenio_farmacia = 0

                            # DESCONTO ADIANTAMENTO SALARIAL
                            match_adiant_salarial = search(r"981DESCONTO ADIANTAMENTO SALARIAL\s*([\d.,]+)", texto_pdf)
                            if match_adiant_salarial:
                                adiant_salarial = float(match_adiant_salarial.group(1).replace(".", "").replace(",", "."))
                            else: 
                                adiant_salarial = 0

                            # NUMERO DE EMPREGADOS
                            match_demitido = search(r"No. Empregados: Demitido:\s*(\d+)", texto_pdf)
                            if match_demitido:
                                demitido = match_demitido.group(1)
                                match_num_empregados = search(r"No. Empregados: Demitido:\s+" + demitido + 
                                                                r"\s*(\d+)", texto_pdf)
                                if match_num_empregados:
                                    num_empregados = match_num_empregados.group(1)
                                else: 
                                    num_empregados = 0 
                            else:
                                num_empregados = 0

                            # NUMERO DE ESTAGIARIOS
                            match_transferido = search(r"No. Estagiários: Transferido:\s*(\d+)", texto_pdf)
                            if match_transferido:
                                transferido = match_transferido.group(1)
                                match_num_estagiarios = search(r"No. Estagiários: Transferido:\s+" + transferido + 
                                                                r"\s*(\d+)", texto_pdf)
                                if match_num_estagiarios:
                                    num_estagiarios = match_num_estagiarios.group(1)
                                else: 
                                    num_estagiarios = 0
                            else:
                                num_estagiarios = 0

                            # TRABALHANDO
                            match_ferias = search(r"Trabalhando: Férias:\s*(\d+)", texto_pdf)
                            if match_ferias:
                                ferias = match_ferias.group(1)
                                match_trabalhando = search(r"Trabalhando: Férias:\s+" + ferias + r"\s*(\d+)", texto_pdf)
                                if match_trabalhando:
                                    trabalhando = match_trabalhando.group(1)
                                else:
                                    trabalhando = 0
                            else:
                                trabalhando = 0

                            # SALARIO CONTRIBUIÇÃO EMPREGADOS
                            match_salario_contri_empregados = search(r"Salário contribuição empregados:\s*([\d.,]+)", texto_pdf)
                            if  match_salario_contri_empregados:
                                salario_contri_empregados = float(match_salario_contri_empregados
                                                                .group(1).replace(".", "").replace(",", "."))
                            else: 
                                salario_contri_empregados = 0

                            # SALARIO CONTRIBUIÇÃO CONTRIBUINTES
                            match_salario_contri_contribuintes = search(r"Salário contribuição contribuintes:\s*([\d.,]+)", 
                                                                        texto_pdf)
                            if  match_salario_contri_contribuintes:
                                salario_contri_contribuintes = float(match_salario_contri_contribuintes
                                                                    .group(1).replace(".", "").replace(",", "."))
                            else:
                                salario_contri_contribuintes = 0
                            
                            # SOMA DOS SALARIOS
                            soma_salarios_provdt = salario_contri_empregados + salario_contri_contribuintes

                            # VALOR DO INSS
                            match_inss = search(r"Total INSS:\s*([\d.,]+)", texto_pdf)
                            if match_inss:
                                inss = float(match_inss.group(1).replace(".", "").replace(",", "."))
                            else:
                                inss = 0

                            # VALOR DO FGTS
                            match_fgts = search(r"Valor do FGTS:\s*([\d.,]+)", texto_pdf)
                            if  match_fgts:
                                fgts = float(match_fgts.group(1).replace(".", "").replace(",", "."))
                            else:
                                fgts = 0

                            # VALOR DO IRRF
                            match_base_iss = search(r"([\d.,]+)\s+Valor Total do IRRF: Base ISS:", texto_pdf)
                            if match_base_iss:
                                base_iss = match_base_iss.group(1)
                                match_irrf = search(r"([\d.,]+)\s+" + base_iss + r"\s+Valor Total do IRRF: Base ISS:", texto_pdf)
                                if match_irrf:
                                    irrf = float(match_irrf.group(1).replace(".", "").replace(",", "."))
                                else:
                                    irrf = 0
                            else:
                                irrf = 0

                            # LÍQUIDO CENTRO DE CUSTO - entra na coluna salarios a pagar
                            match_liquido = search(r"Líquido Centro de Custo:\s*([\d.,]+)", texto_pdf)
                            if  match_liquido:
                                liquido_centro_custo = float(match_liquido.group(1).replace(".", "").replace(",", "."))
                            else:
                                liquido_centro_custo = 0

                            vale_transporte, assinat_eletronica, vale_refeicao, ponto_eletronico, sst = pega_valores_vales_reembolsos(cliente_id, nome_centro_custo_mod.replace("S/S", "S S"))

                            # INSERÇÃO DE DADOS NO BANCO
                            query_insert_valores = ler_sql('sql/registrar_valores_extrato.sql')
                            values_insert_valores = (cliente_id, cod_centro_custo, convenio_farmacia, adiant_salarial, num_empregados, 
                                                        num_estagiarios, trabalhando, salario_contri_empregados, 
                                                        salario_contri_contribuintes, soma_salarios_provdt, inss, fgts, 
                                                        irrf, liquido_centro_custo, vale_transporte, assinat_eletronica, 
                                                        vale_refeicao, ponto_eletronico, sst, mes, ano
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
                            copy(caminho_pdf_mod, caminho_destino / caminho_pdf_mod.name)
                    else:
                        print("Cliente não encontrado!\n")

    except Exception as error:
        if error.args == ("'NoneType' object is not iterable",):
            print("O diretório informado não foi especificado!")
        else:
            print(f"O sistema retornou um erro: {error}")

def gera_fatura():
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
                                    print(valores_financeiro)
                                    input()
                                    caminho_sub_pasta = Path(sub_pasta)

                                    # Variáveis para planilha
                                    nome_fatura = f"Fatura_Detalhada_{nome_pasta_cliente}_{mes}.{ano}.xlsx"
                                    caminho_fatura = f"{caminho_sub_pasta}\\{nome_fatura}"
                                    
                                    # COPIANDO A FATURA MODELO PARA A PASTA DO CLIENTE
                                    copy(modelo_fatura, caminho_sub_pasta / nome_fatura)
                                   
                                    try:
                                        # FORMATANDO A FATURA                                       
                                        workbook = load_workbook(caminho_fatura)
                                        sheet = workbook.active
                                        # Criar um estilo de número personalizado para moeda
                                        style_moeda = NamedStyle(name="estilo_moeda", number_format='_-R$ * #,##0.00_-;-R$ * #,##0.00_-;_-R$ * "-"??_-;_-@_-')
                                        # Adicionar o estilo ao workbook (necessário apenas uma vez)
                                        workbook.add_named_style(style_moeda)
                                        # nome da planilha (em baixo)
                                        sheet.title = f"{mes}.{ano}"
                                        # titulo da fatura
                                        sheet['D2'] = f"Fatura Detalhada - {nome_pasta_cliente}"
                                        # convênio farmácia
                                        if not valores_financeiro[1] == None:
                                            sheet['J18'] = valores_financeiro[1]
                                            conv_farmacia = valores_financeiro[1]
                                        else:
                                            conv_farmacia = 0
                                        # adiantamento salarial
                                        if not valores_financeiro[2] == None:
                                            sheet['J10'] = valores_financeiro[2]
                                            adiant_salarial = valores_financeiro[2]
                                        else:
                                            adiant_salarial = 0
                                        # numero de funcionarios
                                        if valores_financeiro[3] + valores_financeiro[4] == 1:
                                            sheet['J6'] = 1
                                            sheet['K6'] = 'funcionário'
                                        else:
                                            sheet['J6'] = valores_financeiro[3] + valores_financeiro[4]
                                        # salários a pagar
                                        sheet['A7'] = f"Salários a pagar {mes}.{ano}"
                                        sheet['J7'] = valores_financeiro[12]
                                        salarios_pagar = valores_financeiro[12]
                                        # inss
                                        sheet['A8'] = f"GPS (Guia da Previdência Social) {mes}.{ano}"
                                        sheet['J8'] = valores_financeiro[9]
                                        inss = valores_financeiro[9]
                                        # fgts
                                        sheet['A9'] = f"FGTS (Fundo de Garantia por Tempo de Serviço) {mes}.{ano}"
                                        sheet['J9'] = valores_financeiro[10]
                                        fgts = valores_financeiro[10]
                                        # provisão de direitos trabalhistas
                                        sheet['A11'] = f"Provisão de Direitos Trabalhistas {mes}.{ano}"
                                        sheet['E11'] = valores_financeiro[8]
                                        soma_salarios_provdt = valores_financeiro[8]
                                        # irrf (folha de pagamento)
                                        if not valores_financeiro[11] == None:
                                            sheet['J12'] = valores_financeiro[11]
                                            irrf = valores_financeiro[11]
                                        else:
                                            irrf = 0
                                        # vale transporte
                                        sheet['A15'] = f"Vale Transporte {mes}/{ano}"
                                        if not valores_financeiro[13] == None:
                                            sheet['J15'] = valores_financeiro[13]
                                            vale_transp = valores_financeiro[13]
                                        else:
                                            vale_transp = 0
                                        # assinatura eletrônica
                                        if not valores_financeiro[14] == None:
                                            sheet['J14'] = valores_financeiro[14]
                                            assinatura_elet = valores_financeiro[14]
                                        else:
                                            assinatura_elet = 0
                                        # vale refeição
                                        sheet['A16'] = f"Vale Refeição {mes}/{ano}"
                                        if not valores_financeiro[15] == None:
                                            sheet['J16'] = valores_financeiro[15]
                                            vale_refeic = valores_financeiro[15]
                                        else:
                                            vale_refeic = 0
                                        # mensalidade do ponto eletrônico
                                        if not valores_financeiro[16] == None:
                                            sheet['J13'] = valores_financeiro[16]
                                            mensal_ponto = valores_financeiro[16]
                                        else:
                                            mensal_ponto = 0
                                        # saúde e segurança do trabalho
                                        if not valores_financeiro[17] == None:
                                            sheet['J17'] = valores_financeiro[17] 
                                            sst = valores_financeiro[17]
                                        else:
                                            sst = 0

                                        #reembolsos
                                        query_procura_reembolsos = ler_sql('sql/procura_valores_reembolsos.sql')
                                        values_procura_reembolsos = (cliente_id, mes, ano)
                                        with mysql.connector.connect(**db_conf) as conn, conn.cursor() as cursor:
                                            cursor.execute(query_procura_reembolsos, values_procura_reembolsos)
                                            reembolsos = cursor.fetchall()
                                            conn.commit()
                                        if not reembolsos == []:
                                            reembolso_total = 0
                                            cel_1 = 23
                                            cel_2 = 24
                                            for reembolso in reembolsos:
                                                cel_1 += 1
                                                cel_2 += 1
                                                sheet.insert_rows(19)
                                                sheet['A19'] = reembolso[0]
                                                sheet['J19'] = reembolso[1]
                                                sheet['J19'].style = style_moeda
                                                sheet['G19'].border = Border(left=Side(style='thin'))
                                                sheet['I19'].border = Border(right=Side(style='thin'))
                                                sheet['L19'].border = Border(right=Side(style='thin'))
                                                sheet[f'J{cel_1 - 4}'] = f'=E11*H{cel_1 - 4}'
                                                sheet[f'J{cel_1 - 3}'] = f'=SUM(J7:L{cel_1 - 4})'
                                                sheet[f'H{cel_1 + 2}'] = f'=H{cel_1}-H{cel_2}'
                                                sheet[f'J{cel_1}'] = f'=E11*H{cel_1}'
                                                sheet[f'J{cel_2}'] = f'=E11*H{cel_2}'
                                                sheet[f'J{cel_1 + 2}'] = f'=J{cel_1}-J{cel_2}'
                                                reembolso_total = reembolso_total + reembolso[1]
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
                                                  + sst + conv_farmacia + percent_human + reembolso_total
                                                 )
                                        total_fatura = round(fatura, 2)
                                        print(f"Total da Fatura: {total_fatura}")
                                        input()
                                        # GERANDO PDF DA FATURA
                                        try:
                                            print("Terminada a formatação da fatura, gerando pdf...")
                                            excel = win32.gencache.EnsureDispatch('Excel.Application')
                                            excel.Visible = True
                                            wb = excel.Workbooks.Open(caminho_fatura)
                                            ws = wb.Worksheets[f"{mes}.{ano}"]
                                            sleep(3)
                                            ws.ExportAsFixedFormat(0, sub_pasta + f"\\Fatura_Detalhada_{nome_pasta_cliente}_{mes}.{ano}")
                                            wb.Close()
                                            excel.Quit()
                                            print("Pdf gerado com sucesso.")

                                            print("Inserindo valores no banco.")
                                            query_fatura = ler_sql('sql/registrar_valores_fatura.sql')
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
        actions, driver = start_chrome()
        sleep(1)
        elemento_cards = procura_todos_elementos(driver, "xpath", """//*[@id="myorganizations-container"]"""+
                                                """/div/div[3]/ng-include[2]/div[*]/a/h3/span""", 20)
        for card in elemento_cards:
            if card.text == "HUMAN SOLUCOES E DESENVOLVIMENTOS EM RECURSOS HUMANOS LTDA":
                card.click()
                break
        sleep(1.5)
        elemento_contatos = procura_elemento(driver, "xpath", """//*[@id="page-organization-details"]/div[5]/div/div[2]"""+
                                             """/div[2]/div/div/ul[2]/li[3]/a/span""", 20)
        if elemento_contatos:
            actions.move_to_element(elemento_contatos).perform()
            elemento_clientes = procura_elemento(driver, "xpath", """//*[@id="page-organization-details"]/div[5]/div/div[2]"""+
                                                 """/div[2]/div/div/ul[2]/li[3]/ul/li[1]/a/span""", 20)
            elemento_clientes.click()
            sleep(1)
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
                        print(pasta_cliente, boleto)
                        if boleto == True:
                            boleto = False
                            break
                        elif boleto == False:
                            cliente = procura_cliente(nome_pasta_cliente)
                            print(cliente)
                            if cliente:
                                cliente_id = cliente[0]
                                cliente_cnpj = cliente[2]
                                cliente_cpf = cliente[3]
                                valores = procura_valores(cliente_id)
                                if valores:
                                    print(valores)
                                    valor_fatura = valores[20]
                                    print(f"{nome_pasta_cliente} vai precisar de um boleto. Valor da fatura é: {valor_fatura}")
                                    elemento_search = procura_elemento(driver, "xpath", """//*[@id="entityList_filter"]"""+
                                                                      """/label/input""", 15)   
                                    if elemento_search:                           
                                        if not cliente_cnpj == '' or not cliente_cnpj == None:
                                            elemento_search.send_keys(cliente_cnpj)
                                        elif not cliente_cpf == '' or not cliente_cpf == None:
                                            elemento_search.send_keys(cliente_cpf)                       
                                    sleep(2)
                                    try:
                                        elemento_lista_clientes = procura_todos_elementos(driver, "xpath", """//*[@id="entityList"]"""+
                                                                    """/tbody/tr/td[1]/a""" , 15)
                                    except NoSuchElementException:
                                        elemento_lista_clientes = procura_todos_elementos(driver, "xpath", """//*[@id="entityList"]"""+
                                                                    """/tbody/tr[*]/td[1]/a""" , 15)
                                    for cliente_lista in elemento_lista_clientes:
                                        if cliente_lista.text.__contains__(cliente_cnpj) or cliente_lista.text.__contains__(cliente_cpf):
                                            cliente_lista.click()
                                            try:
                                                sleep(0.7)
                                                elemento_sem_lancamento = procura_elemento(driver, "class_name", """generic-list-no-content""", 15)
                                                if elemento_sem_lancamento:
                                                    agendar_lancamento(driver, valor_fatura)
                                                    sleep(1.5)
                                                    baixar_boleto_lancamento(driver, elemento_search)
                                                elif elemento_sem_lancamento == None:
                                                    baixar_boleto_lancamento(driver, elemento_search)                                                   
                                                sleep(2)
                                                arquivos_downloads = listagem_arquivos_downloads()
                                                arquivo_mais_recente = max(arquivos_downloads, key=os.path.getmtime)
                                                if (arquivo_mais_recente.__contains__(".pdf") and not 
                                                    arquivo_mais_recente.__contains__(f"Boleto_Recebimento_{nome_pasta_cliente.replace("S/S", "S S")}_{mes}.{ano}")):
                                                    caminho_pdf = Path(arquivo_mais_recente)
                                                    novo_nome_boleto = caminho_pdf.with_name(f"Boleto_Recebimento_{nome_pasta_cliente.replace("S/S", "S S")}_{mes}.{ano}.pdf")
                                                    caminho_pdf_mod = caminho_pdf.rename(novo_nome_boleto)
                                                    caminho_destino = Path(pasta_cliente)
                                                    caminho_destino.mkdir(parents=True, exist_ok=True)
                                                    copy(caminho_pdf_mod, caminho_destino / caminho_pdf_mod.name)
                                                    if os.path.exists(caminho_pdf_mod):
                                                        os.remove(caminho_pdf_mod)
                                            except NoSuchElementException:
                                                print("Algum objeto nao foi encontrado!")
                                else:
                                    print(f"Valores de financeiro não encontrados para {nome_pasta_cliente}")
                            else:
                                print(f"Cliente {nome_pasta_cliente} não encontrado!")                    
    except Exception as error:
        print(error)
    print("PROCESSO DE BOLETO ENCERRADO!")
    driver.quit() 

def envia_arquivos():
    try:  
        anexos = []  
        for diretorio in lista_dir_clientes:
            pastas_regioes = listagem_pastas(diretorio)
            for pasta_cliente in pastas_regioes:
                nome_pasta_cliente = pega_nome(pasta_cliente)
                sub_pastas_cliente = listagem_pastas(pasta_cliente)
                for sub_pasta in sub_pastas_cliente:
                    if sub_pasta.__contains__(f"{mes}-{ano}"):
                        arquivos_cliente = listagem_arquivos(sub_pasta)
                        for arquivo in arquivos_cliente:
                            if arquivo.__contains__("Extrato_Mensal_") and arquivo.__contains__(f"{nome_pasta_cliente}_{mes}.{ano}.pdf"):
                                extrato = True
                                anexos.append(arquivo)
                            elif arquivo.__contains__("Fatura_Detalhada_") and arquivo.__contains__(f"{nome_pasta_cliente}_{mes}.{ano}.pdf"):
                                fatura = True
                                anexos.append(arquivo)
                            elif arquivo.__contains__("Boleto_Recebimento_") and arquivo.__contains__(f"{nome_pasta_cliente}_{mes}.{ano}.pdf"):
                                boleto = True
                                anexos.append(arquivo)
                        print(nome_pasta_cliente)
                        print(f"Extrato: {extrato},\n Fatura: {fatura},\n Boleto: {boleto}")
                        print(f"Arquivos:\n {anexos}")
                        if extrato == True and fatura == True and boleto == True:
                            try:
                                cliente = procura_cliente(nome_pasta_cliente)
                                if cliente:
                                    cliente_id = cliente[0]
                                    valores_extrato = procura_valores(cliente_id)
                                    if valores_extrato:
                                        print(valores_extrato)
                                        print(f"Fará o envio para o cliente {nome_pasta_cliente}")
                                        input()
                                        enviar_email_com_anexos("mzblannes@outlook.com", "Teste de Automacao", "testando componente de envio de emails com python", anexos)
                                    else:
                                        print("Valores de financeiro não encontrados!")
                            except Exception as error:
                                print (error)
                            finally:
                                extrato = False
                                anexos = []
                                fatura = False
                                boleto = False
                        else:
                            print("Cliente não possui todos os arquivos necessários para o envio!")
                            input()
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
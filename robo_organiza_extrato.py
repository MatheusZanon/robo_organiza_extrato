from components.importacao_diretorios_windows import listagem_pastas, listagem_arquivos, pega_nome
from components.extract_text_pdf import extract_text_pdf
from components.db_config import db_config
from components.importacao_caixa_dialogo import DialogBox
import tkinter as tk
import mysql.connector
from mysql.connector import errorcode
import re
from pathlib import Path
import shutil 
from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
import win32com.client as win32
import time
import subprocess
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains

# =====================CONFIGURAÇÂO DO BANCO DE DADOS======================
db_conf = db_config()
try: 
    conn = mysql.connector.connect(**db_conf)
    cursor = conn.cursor()
    print(" * Conexão bem sucedida!")

except mysql.connector.Error as err:
    if err.errno == errorcode.ER_ACCESS_DENIED_ERROR:
        print("Tem algo de erro com seu nome ou senha.")
    elif err.errno == errorcode.ER_BAD_DB_ERROR:
        print("Esse banco não existe!")
    else:
        print(err) 


# =============CHECANDO SE O GOOGLE FILE STREAM ESTÁ INICIADO NO SISTEMA==============
# Nome do processo do Google Drive File Stream
nome_processo_drive = "GoogleDriveFS.exe"

# Listar processos em execução e verificar se o Google Drive File Stream está entre eles
processo_ativo = False

try:
    processos = subprocess.check_output(['tasklist']).decode('cp1252').split('\r\n')
except UnicodeDecodeError:
    processos = subprocess.check_output(['tasklist']).decode('utf-16').split('\r\n')

for proc in processos:
    if nome_processo_drive in proc:
        processo_ativo = True
        break

# Se o Google Drive File Stream não estiver em execução, iniciá-lo
if not processo_ativo:
    caminho_executavel_drive = r"C:\Program Files\Google\Drive File Stream\launch.bat"
    subprocess.Popen(caminho_executavel_drive, shell=True)
    time.sleep(3)


# ================CONFIGURAÇÃO DO SELENIUM CHROME DRIVER=====================
# Opções de inicialização
chrome_options = Options()
chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
chrome_options.add_experimental_option('useAutomationExtension', False)
# Caminho para o chrome driver
caminho_drive = r'documents\\chromedriver-win64\\chromedriver.exe'
# Configurar o serviço do ChromeDriver
servico = Service(caminho_drive)


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


# ==================== MÉTODOS DE CADA ETAPA DO PROCESSO=======================
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
        element_email = driver.find_element(By.XPATH, """//*[@id="Username"]""")
        element_email.send_keys("automacao@exponential-co.com")
        element_btn_continue = driver.find_element(By.XPATH, """//*[@id="continue-button"]""")
        element_btn_continue.click()
        time.sleep(0.5)
        element_senha = driver.find_element(By.XPATH, """//*[@id="Password"]""")
        element_senha.send_keys("""NRbTK*Agd#T10{""")
        element_btn_entrar = driver.find_element(By.XPATH, """//*[@id="password"]/div[3]/input""")
        element_btn_entrar.click()
        return actions, driver
    except Exception as error:
        print(f"Chrome Driver retornou um erro: {error}")

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
                        nome_centro_custo = match_centro_custo.group(1)
                        print(f"Centro de Custo: {nome_centro_custo}")
                        partes = nome_centro_custo.split(" - ", 1)
                        if len(partes) > 1:
                            nome_centro_custo_mod = partes[1]
                    
                    # CONVÊNIO FÁRMACIA
                    match_convenio_farm = re.search(r"244CONVÊNIO FARMÁCIA\s*([\d.,]+)", texto_pdf)
                    if match_convenio_farm:
                        convenio_farmacia = float(match_convenio_farm.group(1).replace(".", "").replace(",", "."))
                    else:
                        convenio_farmacia = 0
                    print(f"Convênio Farmácia: {convenio_farmacia}")

                    # DESCONTO ADIANTAMENTO SALARIAL
                    match_adiant_salarial = re.search(r"981DESCONTO ADIANTAMENTO SALARIAL\s*([\d.,]+)", texto_pdf)
                    if match_adiant_salarial:
                        adiant_salarial = float(match_adiant_salarial.group(1).replace(".", "").replace(",", "."))
                    else: 
                        adiant_salarial = 0
                    print(f"Desconto Adiantamento Salarial: {adiant_salarial}")

                    # NUMERO DE EMPREGADOS
                    # Exemplo de extração para Números
                    match_demitido = re.search(r"No. Empregados: Demitido:\s*(\d+)", texto_pdf)
                    if match_demitido:
                        demitido = match_demitido.group(1)
                        match_num_empregados = re.search(r"No. Empregados: Demitido:\s+" + demitido + r"\s*(\d+)", texto_pdf)
                        if match_num_empregados:
                            num_empregados = match_num_empregados.group(1)
                        else: 
                            num_empregados = 0 
                    else:
                        num_empregados = 0
                    print(f"Número de empregados: {num_empregados}")

                    # NUMERO DE ESTAGIARIOS
                    match_transferido = re.search(r"No. Estagiários: Transferido:\s*(\d+)", texto_pdf)
                    if match_transferido:
                        transferido = match_transferido.group(1)
                        match_num_estagiarios = re.search(r"No. Estagiários: Transferido:\s+" + transferido + r"\s*(\d+)", texto_pdf)
                        if match_num_estagiarios:
                            num_estagiarios = match_num_estagiarios.group(1)
                        else: 
                            num_estagiarios = 0
                    else:
                        num_estagiarios = 0
                    print(f"Número de estagiários: {num_estagiarios}")

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
                    print(f"Trabalhando: {trabalhando}")

                    # SALARIO CONTRIBUIÇÃO EMPREGADOS
                    match_salario_contri_empregados = re.search(r"Salário contribuição empregados:\s*([\d.,]+)", texto_pdf)
                    if  match_salario_contri_empregados:
                        salario_contri_empregados = float(match_salario_contri_empregados.group(1).replace(".", "").replace(",", "."))
                    else: 
                        salario_contri_empregados = 0
                    print(f"Salário contribuição Empregados: {salario_contri_empregados}")

                    # SALARIO CONTRIBUIÇÃO CONTRIBUINTES
                    match_salario_contri_contribuintes = re.search(r"Salário contribuição contribuintes:\s*([\d.,]+)", texto_pdf)
                    if  match_salario_contri_contribuintes:
                        salario_contri_contribuintes = float(match_salario_contri_contribuintes.group(1).replace(".", "").replace(",", "."))
                    else:
                        salario_contri_contribuintes = 0
                    print(f"Salário contribuição Contribuintes: {salario_contri_contribuintes}")
                    
                    # SOMA DOS SALARIOS
                    soma_salarios_provdt = salario_contri_empregados + salario_contri_contribuintes
                    print(f"Soma dos salários: {soma_salarios_provdt}")

                    # VALOR DO INSS
                    # A expressão regular procura por um ou mais números seguidos por qualquer coisa (não capturada)
                    # e então "Total INSS:"
                    match_inss = re.search(r"Total INSS:\s*([\d.,]+)", texto_pdf)
                    if match_inss:
                        inss = float(match_inss.group(1).replace(".", "").replace(",", "."))
                    else:
                        inss = 0
                    print(f"Total INSS: {inss}")

                    # VALOR DO FGTS
                    match_fgts = re.search(r"Valor do FGTS:\s*([\d.,]+)", texto_pdf)
                    if  match_fgts:
                        fgts = float(match_fgts.group(1).replace(".", "").replace(",", "."))
                    else:
                        fgts = 0
                    print(f"Valor do FGTS: {fgts}")

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
                    print(f"Valor Total do IRRF: {irrf}")

                    # LÍQUIDO CENTRO DE CUSTO
                    match_liquido = re.search(r"Líquido Centro de Custo:\s*([\d.,]+)", texto_pdf)
                    if  match_liquido:
                        liquido_centro_custo = float(match_liquido.group(1).replace(".", "").replace(",", "."))
                    else:
                        liquido_centro_custo = 0
                    print(f"Líquido Centro de Custo: {liquido_centro_custo}")
                    # LIQUIDO CENTRO DE CUSTO ENTRA NA COLUNA SALARIOS_PAGAR DO BANCO
                    

                    # PROCURA CLIENTE AO QUAL O EXTRATO PERTENCE
                    cliente = procura_cliente(nome_centro_custo_mod)
                    if cliente:
                        cliente_id = cliente[0]
                        caminho_pasta_cliente = procura_pasta_cliente(nome_centro_custo_mod)
                        print(f"Cliente: {cliente}\n Caminho da pasta: {caminho_pasta_cliente}\n Extrato: {extrato}\n")
                        input()
                        valores_extrato = procura_valores(cliente_id)
                        print(valores_extrato)
                        input()
                        if valores_extrato:
                            print("Esses valores de extrato ja foram registrados!")
                        else:
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
                            if not nome_extrato.__contains__(f"Extrato_Mensal_{nome_centro_custo}_{mes}.{ano}"):
                                novo_nome_extrato = caminho_pdf.with_name(f"Extrato_Mensal_{nome_centro_custo}_{mes}.{ano}")
                                caminho_pdf.rename(novo_nome_extrato)
                            # Caminho da pasta de destino (o caminho que vem da sua variável)
                            caminho_destino = Path(caminho_pasta_cliente)
                            # Verifica se a pasta de destino existe; se não, cria a pasta
                            caminho_destino.mkdir(parents=True, exist_ok=True)
                            # Copiar o arquivo PDF para a pasta de destino
                            shutil.copy(caminho_pdf, caminho_destino / caminho_pdf.name)
                    else:
                        print("Cliente não encontrado!")

                    
    except Exception as error:
        if error.args == ("'NoneType' object is not iterable",):
            print("O diretório informado não foi especificado!")
        else:
            print(f"O sistema retornou um erro: {error}")

def gera_fatura():
    try:
        modelo_fatura = Path(f"{particao}:\\Meu Drive\\Arquivos_Automacao\\Fatura_Detalhada_Modelo_00.0000_python.xlsx")
        caminho_final = ""
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
                            if arquivo.__contains__("Fatura_Detalhada_") and arquivo.__contains__(nome_pasta_cliente):
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
                                print(f"Cliente: {cliente}")
                                valores_financeiro = procura_valores(cliente_id)
                                print(f"Valores: {valores_financeiro}")
                                input()
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
                                        sheet['E2'] = f"Fatura Detalhada - {nome_pasta_cliente}"
                                        # numero de funcionarios
                                        if valores_financeiro[4] == 1:
                                            sheet['J6'] = 1
                                            sheet['K6'] = 'funcionário'
                                        else:
                                            sheet['J6'] = valores_financeiro[4]
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
                    else:
                        continue

    except Exception as error:
        return (error)

def gera_boleto(): 
    try:  
        actions, driver = start_chrome()
        time.sleep(1)
        element_cards = driver.find_elements(By.XPATH, """//*[@id="myorganizations-container"]/div/div[3]/ng-include[2]/div[*]/a/h3/span""")
        for card in element_cards:
            if card.text == "HUMAN SOLUCOES E DESENVOLVIMENTOS EM RECURSOS HUMANOS LTDA":
                card.click()
                break
        element_contatos = driver.find_element(By.XPATH, """//*[@id="page-organization-details"]/div[5]/div/div[1]/div/div/ul[2]/li[3]""")
        actions.move_to_element(element_contatos).perform()
        time.sleep(0.5)
        element_contatos_clientes = driver.find_element(By.XPATH, """//*[@id="page-organization-details"]/div[5]/div/div[1]/div/div/ul[2]/li[3]/ul/li[1]""")
        element_contatos_clientes.click()
        element_search = driver.find_element(By.XPATH, """//*[@id="entityList_filter"]/label/input""")
    except Exception as web_error:
        print (web_error.args)

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
                            print(f"{nome_pasta_cliente} vai precisar de um boleto.")
                            cliente = procura_cliente(nome_pasta_cliente.replace("S S", "S/S"))
                            if cliente:
                                print(cliente)
                                cliente_id = cliente[0]
                                cliente_cnpj = cliente[2]
                                cliente_cpf = cliente[3]
        input()
        driver.quit()
    except Exception as error:
        print (error)

def envia_arquivos():
    try:    
        print("Processo de enviar os arquivos para cada cliente (extrato, fatura, boleto)")
    except Exception as error:
        print (error)

# ========================CÓDIGO PRINCIPAL DO ROBÔ===========================
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
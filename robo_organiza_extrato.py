# =========================IMPORTAÇÕES DE BIBLIOTECAS E COMPONENTES========================
from components.importacao_diretorios_windows import *
from components.extract_text_pdf import extract_text_pdf
from components.importacao_caixa_dialogo import DialogBox
from components.checar_ativacao_google_drive import checa_google_drive
from components.configuracao_db import configura_db, ler_sql
from components.procura_cliente import procura_cliente, procura_cliente_por_id, procura_clientes
from components.procura_valores import procura_valores, procura_valores_com_codigo, procura_salarios_com_codigo
from components.procura_elementos_web import procura_elemento, procura_todos_elementos, encontrar_elemento_shadow_root
from components.configuracao_selenium_drive import configura_selenium_driver
from components.enviar_emails import enviar_email_com_anexos
from components.integracao_nibo import pegar_empresa_por_id, pegar_agendamento_de_pagamento_cliente_por_data, agendar_recebimento, cancelar_agendamento_de_recebimento
import tkinter as tk
import mysql.connector
from re import search
from pathlib import Path
from shutil import copy, move
from openpyxl import load_workbook
from openpyxl.styles import Border, Side, NamedStyle
import win32com.client as win32
from dotenv import load_dotenv
import os
import pythoncom
from time import sleep, time
from datetime import date
import pandas as pd
from flask import Flask, request
from flask_restful import Resource, Api, reqparse


# ================= CARREGANDO VARIÁVEIS DE AMBIENTE======================
load_dotenv()

# =====================CONFIGURAÇÂO DO BANCO DE DADOS======================
db_conf = configura_db()

# =============CHECANDO SE O GOOGLE FILE STREAM ESTÁ INICIADO NO SISTEMA==============
checa_google_drive()

chrome_options, servico = configura_selenium_driver()

# ================CONFIGURAÇÃO DAS VARIAVEIS DE AMBIENTE=====================
automacao_email = os.getenv('SELENIUM_USER')
automacao_senha = os.getenv('SELENIUM_PASSWORD')
email_gestor = os.getenv('EMAIL_GESTOR')
corpo_email = os.getenv('CORPO_EMAIL')


# ==================== MÉTODOS DE AUXÍLIO====================================
def cria_fatura(cliente_id, nome_cliente, caminho_sub_pasta_cliente, valores_financeiro, mes, ano, modelo_fatura):
    caminho_sub_pasta = Path(caminho_sub_pasta_cliente)
    nome_fatura = f"Fatura_Detalhada_{nome_cliente}_{ano}.{mes}.xlsx"
    caminho_fatura = f"{caminho_sub_pasta}\\{nome_fatura}"     
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
        sheet['D2'] = f"Fatura Detalhada - {nome_cliente}"
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
        reembolso_total = 0
        LINHA = 19 
        if not reembolsos == []:
            cel_1 = 23
            cel_2 = 24
            for reembolso in reembolsos:
                cel_1 += 1
                cel_2 += 1
                sheet.insert_rows(19)
                sheet[f'J{LINHA}'].style = style_moeda
                sheet[f'A{LINHA}'].border = Border(bottom=Side(style='thin'), left=Side(style='thin'))
                sheet[f'B{LINHA}'].border = Border(bottom=Side(style='thin'))
                sheet[f'C{LINHA}'].border = Border(bottom=Side(style='thin'))
                sheet[f'D{LINHA}'].border = Border(bottom=Side(style='thin'))
                sheet[f'E{LINHA}'].border = Border(bottom=Side(style='thin'))
                sheet[f'F{LINHA}'].border = Border(bottom=Side(style='thin'))
                sheet[f'G{LINHA}'].border = Border(bottom=Side(style='thin'), left=Side(style='thin'))
                sheet[f'H{LINHA}'].border = Border(bottom=Side(style='thin'))
                sheet[f'I{LINHA}'].border = Border(bottom=Side(style='thin'), right=Side(style='thin'))
                sheet[f'J{LINHA}'].border = Border(bottom=Side(style='thin'))
                sheet[f'K{LINHA}'].border = Border(bottom=Side(style='thin'))
                sheet[f'L{LINHA}'].border = Border(bottom=Side(style='thin'), right=Side(style='thin'))
                sheet[f'A{LINHA}'] = reembolso[0]
                sheet[f'J{LINHA}'] = reembolso[1]
                sheet[f'J{cel_1 - 4}'] = f'=E11*H{cel_1 - 4}'
                sheet[f'J{cel_1 - 3}'] = f'=SUM(J7:L{cel_1 - 4})'
                sheet[f'H{cel_1 + 2}'] = f'=H{cel_1}-H{cel_2}'
                sheet[f'J{cel_1}'] = f'=E11*H{cel_1}'
                sheet[f'J{cel_2}'] = f'=E11*H{cel_2}'
                sheet[f'J{cel_1 + 2}'] = f'=J{cel_1}-J{cel_2}'
                reembolso_total = reembolso_total + reembolso[1]
        # provisao de direitos trabalhistas
        prov_direitos = round(soma_salarios_provdt * 0.3487, 2)
        # percentual human
        percent_human = round(soma_salarios_provdt * 0.12, 2)
        # economia mensal
        valor1 = round(soma_salarios_provdt * 0.8027, 2)
        valor2 = round(soma_salarios_provdt * 0.4287, 2)
        eco_mensal = round(valor1 - valor2, 2)
        eco_liquida = round(eco_mensal - percent_human, 2)
        workbook.save(caminho_fatura)
        workbook.close()
        # valor total da fatura
        fatura = (salarios_pagar + inss + fgts + adiant_salarial + prov_direitos
                    + irrf + mensal_ponto + assinatura_elet + vale_transp + vale_refeic
                    + sst + conv_farmacia + percent_human + reembolso_total
                    )
        total_fatura = round(fatura, 2)

        # GERANDO PDF DA FATURA
        try:
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            excel.Visible = True
            wb = excel.Workbooks.Open(caminho_fatura)
            ws = wb.Worksheets[f"{mes}.{ano}"]
            sleep(3)

            ws.ExportAsFixedFormat(0, caminho_sub_pasta_cliente + f"\\Fatura_Detalhada_{nome_cliente}_{ano}.{mes}")
            wb.Close()
            excel.Quit()
            query_fatura = ler_sql('sql/registra_valores_fatura.sql')
            with mysql.connector.connect(**db_conf) as conn, conn.cursor() as cursor:
                cursor.execute(query_fatura, (percent_human, eco_mensal, eco_liquida, total_fatura, cliente_id, mes, ano))
                conn.commit()
        except Exception as error:
            print(error)
    except Exception as error:
        print(error)

def copia_boleto_baixado(nome_cliente, mes, ano, pasta_cliente):
    try:
        arquivos_downloads = listagem_arquivos_downloads()
        arquivo_mais_recente = max(arquivos_downloads, key=os.path.getmtime)
        if (arquivo_mais_recente.__contains__(".pdf") 
            and not arquivo_mais_recente.__contains__(f"Boleto_Recebimento_{nome_cliente.replace("S/S", "S S")}_{ano}.{mes}")):
            caminho_pdf = Path(arquivo_mais_recente)
            novo_nome_boleto = caminho_pdf.with_name(f"Boleto_Recebimento_{nome_cliente.replace("S/S", "S S")}_{ano}.{mes}.pdf")
            caminho_pdf_mod = caminho_pdf.rename(novo_nome_boleto)
            sleep(0.5)
            copy(caminho_pdf_mod, pasta_cliente / caminho_pdf_mod.name)
            if os.path.exists(caminho_pdf_mod):
                os.remove(caminho_pdf_mod)
            else:
                print("Arquivo nao encontrado no caminho para remocão!")
        else:
            print("Arquivo de boleto não encontrado!")
    except Exception as error:
        print(f"Erro ao copiar o arquivo: {error}")

def valida_clientes(clientes, dir_extratos) -> list[int]:
    clientes_validos: list[int] = []
    try:
        pasta_faturas = listagem_pastas(dir_extratos)
        if pasta_faturas:
            for pasta in pasta_faturas:
                pasta_novos_extratos = Path(pasta) if Path(pasta).is_dir() and Path(pasta).name.find(f"novos_extratos") == 0 else None
            if pasta_novos_extratos and pasta_novos_extratos.is_dir():
                extratos = listagem_arquivos(pasta_novos_extratos)
                for cliente in clientes:
                    if extratos:
                        for extrato in extratos:
                            if extrato.__contains__(".pdf"):
                                texto_pdf = extract_text_pdf(extrato)

                                # Nome do Centro de Custo
                                match_centro_custo = search(r"C\.Custo:\s*(.*)", texto_pdf)
                                if match_centro_custo:
                                    nome_centro_custo = match_centro_custo.group(1).replace("í", "i").replace("ó", "o")
                                    partes = nome_centro_custo.split(" - ", 1)
                                    if len(partes) > 1:
                                        nome_centro_custo_mod = partes[1].strip()
                                
                                cliente_db_extrato = procura_cliente(nome_centro_custo_mod, db_conf)
                                cliente_id = int(cliente_db_extrato[0])
                                cliente_is_active = bool(cliente_db_extrato[7])

                                if cliente_id == cliente and cliente_is_active == True:
                                    if clientes_validos.count(cliente_id) == 0:
                                        clientes_validos.append(cliente_id)
                                    break
                            else:
                                print(f"O arquivo {extrato} não é um PDF.")
                    else:
                        print(f"O cliente {cliente} não possui extratos no diretório {pasta_novos_extratos}.")
            else:
                print(f"Não há extratos no diretório {pasta_novos_extratos}.")
        else:
            print(f"Pasta de faturas não encontrada.")
    except Exception as error:
        print(error)
    return clientes_validos

# ==================== MÉTODOS DE CADA ETAPA DO PROCESSO=======================
def organiza_extratos(mes, ano, dir_extratos, lista_dir_clientes):
    try:
        pasta_faturas = listagem_pastas(dir_extratos)
        for pasta in pasta_faturas:
            if Path(pasta).name.find(f"novos_extratos") == 0:
                continue

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

                    cliente = procura_cliente(nome_centro_custo_mod, db_conf)
                    if cliente and cliente[7] == True:
                        cliente_id = cliente[0]
                        caminho_pasta_cliente = Path(procura_pasta_cliente(nome_centro_custo_mod, lista_dir_clientes))
                        caminho_sub_pasta_cliente = Path(f"{caminho_pasta_cliente}\\{ano}-{mes}")
                        caminho_sub_pasta_cliente.mkdir(parents=True, exist_ok=True)
                        
                        # CONVÊNIO FÁRMACIA
                        match_convenio_farm = search(r"\d{3}\s*CONV[EÊ]NIO\s+FARM[AÁ]CIA\s*([\d.,]+)", texto_pdf)
                        if match_convenio_farm:
                            convenio_farmacia = float(match_convenio_farm.group(1).replace(".", "").replace(",", "."))
                        else:
                            convenio_farmacia = 0

                        # DESCONTO ADIANTAMENTO SALARIAL
                        match_adiant_salarial = search(r"\d{3}\s*DESCONTO ADIANTAMENTO SALARIAL\s*([\d.,]+)", texto_pdf)
                        if match_adiant_salarial:
                            adiant_salarial = float(match_adiant_salarial.group(1).replace(".", "").replace(",", "."))
                        else: 
                            adiant_salarial = 0
                        if adiant_salarial == 0:
                            match_adiant_salarial = search(r"\d{3}\s*DESC.ADIANT.SALARIAL\s*([\d.,]+)", texto_pdf)
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

                        valores_extrato = procura_valores_com_codigo(cliente_id, cod_centro_custo, db_conf, mes, ano)
                        if valores_extrato: # INSERÇÃO DE DADOS NO BANCO (ATUALIZA REGISTRO)
                            query_update_valores = ler_sql('sql/atualiza_valores_extrato.sql')
                            values_update_valores = (convenio_farmacia, adiant_salarial, num_empregados, 
                                                        num_estagiarios, trabalhando, salario_contri_empregados, 
                                                        salario_contri_contribuintes, soma_salarios_provdt, inss, fgts, 
                                                        irrf, liquido_centro_custo, cliente_id, cod_centro_custo, int(mes), ano
                                                        )
                            with mysql.connector.connect(**db_conf) as conn, conn.cursor() as cursor:
                                cursor.execute(query_update_valores, values_update_valores)
                                conn.commit()
                        else: # INSERÇÃO DE DADOS NO BANCO (CRIA REGISTRO)
                            try:
                                query_insert_valores = ler_sql('sql/registra_valores_extrato.sql')
                                values_insert_valores = (cliente_id, cod_centro_custo, convenio_farmacia, adiant_salarial, num_empregados, 
                                                            num_estagiarios, trabalhando, salario_contri_empregados, 
                                                            salario_contri_contribuintes, soma_salarios_provdt, inss, fgts, 
                                                            irrf, liquido_centro_custo, mes, ano, 0, 0
                                                            )
                                with mysql.connector.connect(**db_conf) as conn, conn.cursor() as cursor:
                                    cursor.execute(query_insert_valores, values_insert_valores)
                                    conn.commit()
                            except Exception as error:
                                print(f"Erro ao registrar os valores de extrato: {error}")
                        
                        caminho_pdf = Path(extrato)
                        if not nome_extrato.__contains__(f"Extrato_Mensal_{nome_centro_custo.replace("S/S", "S S")}_{ano}.{mes}"):
                            novo_nome_extrato = caminho_pdf.with_name(f"Extrato_Mensal_{nome_centro_custo.replace("S/S", "S S").strip()}_{ano}.{mes}.pdf")
                            caminho_pdf_mod = caminho_pdf.rename(novo_nome_extrato)
                        else:
                            caminho_pdf_mod = caminho_pdf
                        caminho_destino = Path(caminho_sub_pasta_cliente)
                        copy(caminho_pdf_mod, caminho_destino / caminho_pdf_mod.name)
                    else:
                        print(f"Cliente não encontrado ou inativo: {nome_centro_custo}\n")
    except Exception as error:
        if error.args == ("'NoneType' object is not iterable",):
            print("O diretório informado não foi especificado!")
        else:
            print(f"O sistema retornou um erro: {error}")

def gera_fatura(mes, ano, lista_dir_clientes, modelo_fatura):
    try:
        pythoncom.CoInitialize()
        for diretorio in lista_dir_clientes:
            pastas_regioes = listagem_pastas(diretorio)
            for pasta_cliente in pastas_regioes:
                nome_pasta_cliente = pega_nome(pasta_cliente)
                sub_pastas_cliente = listagem_pastas(pasta_cliente)
                for sub_pasta in sub_pastas_cliente:
                    if sub_pasta.__contains__(f"{ano}-{mes}"):
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
                            cliente = procura_cliente(nome_pasta_cliente, db_conf)
                            if cliente and cliente[7] == True:
                                cliente_id = cliente[0]
                                valores_financeiro = procura_valores(cliente_id, db_conf, mes, ano)
                                if valores_financeiro != None:
                                    cria_fatura(cliente_id, nome_pasta_cliente, sub_pasta, valores_financeiro, mes, ano, modelo_fatura)
                                else: 
                                    print("Cliente não possui valores para gerar fatura!")
                            else:
                                print("Cliente não encontrado ou inativo!")
    except Exception as error:
        return (error)
    finally:
        pythoncom.CoUninitialize()

def gera_boleto(mes, ano, lista_dir_clientes):
    try:
        for diretorio in lista_dir_clientes:
            pastas_regioes = listagem_pastas(diretorio)
            for pasta_cliente in pastas_regioes:
                nome_pasta_cliente = pega_nome(pasta_cliente)
                sub_pastas_cliente = listagem_pastas(pasta_cliente)
                for sub_pasta in sub_pastas_cliente:
                    if sub_pasta.__contains__(f"{ano}-{mes}"):
                        caminho_destino = Path(sub_pasta)
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
                            cliente = procura_cliente(nome_pasta_cliente, db_conf)
                            if cliente and cliente[7] == True:
                                cliente_id = cliente[0]
                                valores = procura_valores(cliente_id, db_conf, mes, ano)
                                valor_fatura = valores[20]
                                empresa = pegar_empresa_por_id(cliente_id)
                                if valor_fatura:
                                    print(f"Agendando boleto para {nome_pasta_cliente} no valor de R${valor_fatura}...")
                                    recebimento = agendar_recebimento(empresa, valor_fatura, mes, ano)
                                    if recebimento:
                                        copia_boleto_baixado(nome_pasta_cliente, mes, ano, caminho_destino)
                                else:
                                    print(f"Valor da fatura não encontrado para {nome_pasta_cliente}")
                            else:
                                print(f"Cliente {nome_pasta_cliente} não encontrado ou inativo!")                    
    except Exception as error:
        print(error)
    print("PROCESSO DE BOLETO ENCERRADO!")

def envia_arquivos(mes, ano, lista_dir_clientes):
    try:  
        input("APERTE QUALQUER TECLA PARA ENVIAR OS ARQUIVOS")
        for diretorio in lista_dir_clientes:
            pastas_regioes = listagem_pastas(diretorio)
            for pasta_cliente in pastas_regioes:
                anexos = []
                extrato = False
                fatura = False
                boleto = False
                nome_pasta_cliente = pega_nome(pasta_cliente)
                sub_pastas_cliente = listagem_pastas(pasta_cliente)
                for sub_pasta in sub_pastas_cliente:
                    if sub_pasta.__contains__(f"{ano}-{mes}"):
                        arquivos_cliente = listagem_arquivos(sub_pasta)
                        for arquivo in arquivos_cliente:
                            if arquivo.__contains__("Extrato_Mensal_") and arquivo.__contains__(f"{nome_pasta_cliente}_{ano}.{mes}.pdf"):
                                extrato = True
                                anexos.append(arquivo)
                            elif arquivo.__contains__("Fatura_Detalhada_") and arquivo.__contains__(f"{nome_pasta_cliente}_{ano}.{mes}.pdf"):
                                fatura = True
                                anexos.append(arquivo)
                            elif arquivo.__contains__("Boleto_Recebimento_") and arquivo.__contains__(f"{nome_pasta_cliente}_{ano}.{mes}.pdf"):
                                boleto = True
                                anexos.append(arquivo)
                        if extrato == True and fatura == True and boleto == True:
                            try:
                                cliente = procura_cliente(nome_pasta_cliente, db_conf)
                                if cliente and cliente[7] == True:
                                    cliente_id = cliente[0]
                                    cliente_email = cliente[4]
                                    valores_extrato = procura_valores(cliente_id, db_conf, mes, ano)
                                    if valores_extrato and valores_extrato[21] == 0 and not cliente_email == None:
                                        enviar_email_com_anexos(f"{cliente_email}, {email_gestor}", f"Documentos de Terceirização - {nome_pasta_cliente}", 
                                                               f"{corpo_email}", anexos)
                                        query_atualiza_anexos = ler_sql("sql/atualiza_anexos_cliente.sql")
                                        values_anexos = (cliente_id, mes, ano)
                                        with mysql.connector.connect(**db_conf) as conn, conn.cursor() as cursor:
                                            cursor.execute(query_atualiza_anexos, values_anexos)
                                            conn.commit()
                                        break
                                    elif valores_extrato == None:
                                        print("Valores de financeiro não encontrados!")
                                    elif valores_extrato[21] == 1:
                                        print("Anexos já enviados para o cliente!")
                                    elif cliente_email == None:
                                        print("Cliente sem email!")
                                else:
                                    print("Cliente não encontrado ou inativo!")
                            except Exception as error:
                                print (error)
                        else:
                            print("Cliente não possui todos os arquivos necessários para o envio!")
    except Exception as error:
        print (error)


# ================== MÉTODOS PARA REFAZER O PROCESSO ==================
def reorganiza_extratos(mes, ano, dir_extratos, lista_dir_clientes, clientes):
    try:
        pasta_faturas = listagem_pastas(dir_extratos)
        pasta_novos_extratos = None
        for pasta in pasta_faturas:
            pasta_novos_extratos = Path(pasta) if Path(pasta).is_dir() and Path(pasta).name.find(f"novos_extratos") == 0 else None
        
        if pasta_novos_extratos:
            extratos = listagem_arquivos(pasta_novos_extratos)
            for cliente in clientes:
                print(cliente)
                if extratos:
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

                            cliente_db = procura_cliente(nome_centro_custo_mod, db_conf)
                            cliente_id = cliente_db[0]
                            cliente_is_active = bool(cliente_db[7])

                            if cliente_id == cliente and cliente_is_active == True:
                                caminho_pasta_cliente = Path(procura_pasta_cliente(nome_centro_custo_mod, lista_dir_clientes))
                                caminho_sub_pasta_cliente = Path(f"{caminho_pasta_cliente}\\{ano}-{mes}")
                                caminho_sub_pasta_cliente.mkdir(parents=True, exist_ok=True)
                                salarios_extrato = procura_salarios_com_codigo(cliente_id, cod_centro_custo, db_conf, mes, ano)
                                if salarios_extrato:
                                    print(f"{nome_centro_custo} ja possui valores registrados!\n")
                                else:
                                    # CONVÊNIO FÁRMACIA
                                    match_convenio_farm = search(r"\d{3}\s*CONV[EÊ]NIO\s+FARM[AÁ]CIA\s*([\d.,]+)", texto_pdf)
                                    if match_convenio_farm:
                                        convenio_farmacia = float(match_convenio_farm.group(1).replace(".", "").replace(",", "."))
                                    else:
                                        convenio_farmacia = 0

                                    # DESCONTO ADIANTAMENTO SALARIAL
                                    match_adiant_salarial = search(r"\d{3}\s*DESCONTO ADIANTAMENTO SALARIAL\s*([\d.,]+)", texto_pdf)
                                    if match_adiant_salarial:
                                        adiant_salarial = float(match_adiant_salarial.group(1).replace(".", "").replace(",", "."))
                                    else: 
                                        adiant_salarial = 0
                                    if adiant_salarial == 0:
                                        match_adiant_salarial = search(r"\d{3}\s*DESC.ADIANT.SALARIAL\s*([\d.,]+)", texto_pdf)
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

                                    # INSERÇÃO DE DADOS NO BANCO
                                    query_update_valores = ler_sql('sql/atualiza_valores_extrato.sql')
                                    values_update_valores = (convenio_farmacia, adiant_salarial, num_empregados, 
                                                                num_estagiarios, trabalhando, salario_contri_empregados, 
                                                                salario_contri_contribuintes, soma_salarios_provdt, inss, fgts, 
                                                                irrf, liquido_centro_custo, cliente_id, int(mes), ano
                                                                )
                                    with mysql.connector.connect(**db_conf) as conn, conn.cursor() as cursor:
                                        cursor.execute(query_update_valores, values_update_valores)
                                        conn.commit()

                                    caminho_pdf = Path(extrato)
                                    if not nome_extrato.__contains__(f"Extrato_Mensal_{nome_centro_custo.replace("S/S", "S S")}_{ano}.{mes}"):
                                        novo_nome_extrato = caminho_pdf.with_name(f"Extrato_Mensal_{nome_centro_custo.replace("S/S", "S S").strip()}_{ano}.{mes}.pdf")
                                        caminho_pdf_mod = caminho_pdf.rename(novo_nome_extrato)
                                    else:
                                        caminho_pdf_mod = caminho_pdf
                                    caminho_destino = Path(caminho_sub_pasta_cliente)

                                    regiao_cliente = str(cliente_db[6]).strip().lower()
                                    caminho_destino_relatorios = os.path.join(*[part for part in caminho_pdf_mod.parts if caminho_pdf_mod.parts.index(part) < len(caminho_pdf_mod.parts) - 2])

                                    if regiao_cliente == 'itaperuna':
                                        caminho_destino_relatorios = Path([pasta for pasta in pasta_faturas if 'itaperuna' in str(pasta)][0])
                                    elif regiao_cliente == 'manaus':
                                        caminho_destino_relatorios = Path([pasta for pasta in pasta_faturas if 'manaus' in str(pasta)][0])
                                    
                                    caminho_destino_relatorios = Path(caminho_destino_relatorios / caminho_pdf_mod.name)
                                    copy(caminho_pdf_mod, caminho_destino / caminho_pdf_mod.name)
                                    move(caminho_pdf_mod, caminho_destino_relatorios)
                                    break
                            else:
                                print(f"{nome_centro_custo} não é o extrato do cliente atual, indo para o próximo!\n")
                else:
                    print(f"Não existem extratos para o mes {mes} e ano {ano}!\n")
        else:
            print(f"Não existe pasta novos_extratos!\n")
    except Exception as error:
        if error.args == ("'NoneType' object is not iterable",):
            print("O diretório informado não foi especificado!")
        else:
            print(f"O sistema retornou um erro: {error}")

def refazer_fatura(mes, ano, lista_dir_clientes, modelo_fatura, lista_clientes_refazer):
    try:
        pythoncom.CoInitialize()
        for cliente_id in lista_clientes_refazer:
            cliente = procura_cliente_por_id(cliente_id, db_conf)
            if cliente and cliente[7] == True:
                caminho_pasta_cliente = procura_pasta_cliente(cliente[1], lista_dir_clientes)
                nome_pasta_cliente = pega_nome(caminho_pasta_cliente)
                if nome_pasta_cliente:
                    sub_pastas_clientes = listagem_pastas(caminho_pasta_cliente)
                    sub_pasta = None
                    for sub_pasta_cliente in sub_pastas_clientes:
                        if f"{ano}-{mes}" == Path(sub_pasta_cliente).name:
                            sub_pasta = sub_pasta_cliente
                    if sub_pasta:
                        valores_financeiro = procura_valores(cliente_id, db_conf, mes, ano)
                        if valores_financeiro:
                            cria_fatura(cliente_id, nome_pasta_cliente, sub_pasta, valores_financeiro, mes, ano, modelo_fatura)
                        else: 
                            print("Cliente não possui valores para gerar fatura!")
                    else:
                        print(f"A pasta {ano}-{mes} não existe para o cliente {cliente[1]}!")
                else:
                    print(f"Pasta do cliente {cliente[1]} não encontrada!")
            else:
                print(f"Cliente não encontrado ou inativo: {cliente[1]}\n")
    except Exception as error:
        print(error)
    finally:
        pythoncom.CoUninitialize()

def refazer_boleto(mes, ano, lista_dir_clientes, lista_clientes_refazer):
    for cliente_id in lista_clientes_refazer:
        empresa = pegar_empresa_por_id(cliente_id)
        if empresa:
            agendamento_pagamento = pegar_agendamento_de_pagamento_cliente_por_data(empresa['id'], mes, ano)
            if agendamento_pagamento != False:
                deletado = cancelar_agendamento_de_recebimento(agendamento_pagamento['scheduleId'])
                if deletado:
                    print(f"Recebimento {agendamento_pagamento['description']} deletado!")
                    valores_financeiro = procura_valores(cliente_id, db_conf, mes, ano)

                    if not valores_financeiro:
                        print(f"Valores de financeiro não encontrados para o cliente {empresa['name']}!")
                        input("Pressione Enter para prosseguir para o proximo cliente...")
                        continue

                    recebimento = agendar_recebimento(empresa, valores_financeiro[20], mes, ano)
                    if recebimento:
                        print(f"Recebimento {recebimento['idAgendamento']} agendado!")
                        try:
                            cliente_db = procura_cliente_por_id(cliente_id, db_conf)
                            if cliente_db and cliente_db[7] == True:
                                for diretorio in lista_dir_clientes:
                                    pastas_regioes = listagem_pastas(diretorio)
                                    for pasta_cliente in pastas_regioes:
                                        nome_pasta_cliente = pega_nome(pasta_cliente)
                                        if nome_pasta_cliente.__contains__(str(cliente_db[1])):
                                            # PASTA CLIENTE ENCONTRADA
                                            sub_pastas_cliente = listagem_pastas(pasta_cliente)
                                            for sub_pasta in sub_pastas_cliente:
                                                if sub_pasta.__contains__(f"{ano}-{mes}"):
                                                    # PASTA ANO-MES ENCONTRADA
                                                    caminho_destino = Path(sub_pasta)
                                                    copia_boleto_baixado(nome_pasta_cliente, mes, ano, caminho_destino)
                                        else:
                                            print(f"Pasta do cliente {cliente_db[1]} não encontrada!")
                            else:
                                print(f"Cliente {cliente_db[1]} não encontrado ou inativo!")
                        except Exception as error:
                            print(f"Erro ao salvar o boleto na pasta do cliente: {error}")
                            input("Pressione Enter para prosseguir para o proximo cliente...")
                    else:
                        print(f"Recebimento {recebimento['idAgendamento']} não pode ser agendado!")
                else:
                    print(f"Recebimento {agendamento_pagamento['description']} não pode ser deletado!")
            else:
                print(f"Nenhum agendamento encontrado para o cliente {empresa['name']}!")
        else:
            print(f"Nenhuma empresa encontrada para o ID {cliente_id}!")
 
def zerar_valores(mes, ano, lista_clientes):
    try:
        for cliente in lista_clientes:
            query_zera_valores = ler_sql("sql/zerar_valores.sql")
            values_zera_valores = (mes, ano, cliente)
            with mysql.connector.connect(**db_conf) as conn, conn.cursor() as cursor:
                cursor.execute(query_zera_valores, values_zera_valores)
                conn.commit()
    except Exception as error:
        print(f"Erro ao zerar os valores: {error}")
        input("Pressione Enter para prosseguir para o proximo cliente...")

# ==================CAIXA DE DIALOGO INICIAL============================
app = Flask(__name__)
api = Api(app)

class execute(Resource):
    def post(self):
        print("Requisição Recebida!")
        parser = reqparse.RequestParser()
        parser.add_argument('mes', type=int, required=True)
        parser.add_argument('ano', type=int, required=True)
        parser.add_argument('particao', required=True)
        parser.add_argument('rotina', required=True)
        parser.add_argument('clientes', type=list, location='json', required=False)
        json = parser.parse_args()

        mes = json['mes']
        ano = json['ano']
        particao = json['particao']
        rotina = json['rotina']
        clientes = json['clientes'] if json['clientes'] is not None else []

        mes = int(mes)
        if mes < 10:
            mes = f"0{mes}"

        # ========================PARAMETROS INICIAS==============================
        dir_clientes_itaperuna = f"{particao}:\\Meu Drive\\Cobranca_Clientes_terceirizacao\\Clientes Itaperuna"
        dir_clientes_manaus = f"{particao}:\\Meu Drive\\Cobranca_Clientes_terceirizacao\\Clientes Manaus"
        lista_dir_clientes = [dir_clientes_itaperuna, dir_clientes_manaus]
        dir_extratos = f"{particao}:\\Meu Drive\\Robo_Emissao_Relatorios_do_Mes\\faturas_human_{ano}_{mes}"
        modelo_fatura = Path(f"templates\\Fatura_Detalhada_Modelo_0000.00_python.xlsx")
        sucesso = False

        # ========================LÓGICA DE EXECUÇÃO DO ROBÔ===========================
        if rotina == "1. Organizar Extratos":
            organiza_extratos(mes, ano, dir_extratos, lista_dir_clientes)
            gera_fatura(mes, ano, lista_dir_clientes, modelo_fatura)
            gera_boleto(mes, ano, lista_dir_clientes)
            envia_arquivos(mes, ano, lista_dir_clientes)
            sucesso = True
        elif rotina == "2. Gerar Fatura Detalhada":
            gera_fatura(mes, ano, lista_dir_clientes, modelo_fatura)
            gera_boleto(mes, ano, lista_dir_clientes)
            envia_arquivos(mes, ano, lista_dir_clientes)
            sucesso = True
        elif rotina == "3. Gerar Boletos":
            gera_boleto(mes, ano, lista_dir_clientes)
            envia_arquivos(mes, ano, lista_dir_clientes)
            sucesso = True
        elif rotina == "4. Enviar Arquivos":
            envia_arquivos(mes, ano)
            sucesso = True
        elif rotina == "5. Refazer Processo":
            if len(clientes) > 0:
                clientes = [int(id) for id in clientes]
                clientes_validos = valida_clientes(clientes, dir_extratos)
                clientes_invalidos = list(set(clientes) - set(clientes_validos))
                print(f"Valores a serem zerados: {clientes_validos}")
                if len(clientes_validos) > 0:
                    zerar_valores(mes, ano, clientes_validos)
                    print(f"Valores Zerados de {clientes_validos}")
                    reorganiza_extratos(mes, ano, dir_extratos, lista_dir_clientes, clientes_validos)
                    print(f"Extratos Reorganizados de {clientes_validos}")
                    refazer_fatura(mes, ano, lista_dir_clientes, modelo_fatura, clientes_validos)
                    refazer_boleto(mes, ano, lista_dir_clientes, clientes_validos)
                    envia_arquivos(mes, ano, lista_dir_clientes)
                    sucesso = True
                    print("Processo finalizado com sucesso!")
                    if len(clientes_invalidos) > 0:
                        print(f"Os seguintes clientes não continham extratos à refazer: {clientes_invalidos}")
                    print("Processo de refazer concluído com sucesso! Encerrando o robo...")
                else:
                    print("Nenhum cliente valido, encerrando o robô...")
                    sucesso = True
            else:
                print("Nenhum cliente solicitado, encerrando o robô...")
                sucesso = False
        else:
            print("Nenhuma rotina selecionada, encerrando o robô...")
            sucesso = False
        
        if sucesso:
            return {'message': 'Arquivos de Terceirização gerados com sucesso'}, 200
        else:
            return {'message': 'Erro ao gerar arquivos'}, 500

class shutdown(Resource):
    def post(self):
        try:
            os._exit(0)
        except Exception as e:
            print(f'Erro ao executar o comando de shutdown: {e}')   
    
api.add_resource(execute, '/')
api.add_resource(shutdown, '/shutdown')

if __name__ == "__main__":
    app.run(debug=True, port=5000)
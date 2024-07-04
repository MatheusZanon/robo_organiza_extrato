from components.configuracao_db import configura_db
from components.procura_cliente import procura_cliente_por_id
from dotenv import load_dotenv
from datetime import datetime
from time import sleep
import os, requests, json, platform
from pathlib import Path
from winreg import OpenKey, QueryValueEx, HKEY_CURRENT_USER

# ================= CARREGANDO VARIÁVEIS DE AMBIENTE======================
load_dotenv()

# =====================CONFIGURAÇÂO DO BANCO DE DADOS======================
db_conf = configura_db()

# ================= VARIAVEIS DE AMBIENTE DO NIBO =================
NIBO_API_BASE_URL = os.getenv('NIBO_API_BASE_URL')
NIBO_API_TOKEN = os.getenv('NIBO_API_TOKEN')
NIBO_ORGANIZATION = os.getenv('NIBO_ORGANIZATION')
NIBO_CATEGORY_ID = os.getenv('NIBO_CATEGORY_ID')
NIBO_AUTOMACAO_ID = os.getenv('NIBO_AUTOMACAO_ID')

def listar_empresas_clientes():
    empresas = []
    response = requests.get(f"{NIBO_API_BASE_URL}/empresas/v1/customers?organization={NIBO_ORGANIZATION}&ApiToken={NIBO_API_TOKEN}")
    if response.status_code == 200:
        for empresa in response.json()['items']:
            empresas.append(empresa)
    return empresas

def pegar_empresa_por_id(id):
    cliente = procura_cliente_por_id(id, db_conf)
    if cliente:
        cnpj = cliente[2]
        cpf = cliente[3]
        if cnpj:
            cnpj_formatado = str(cnpj.replace('.', '').replace('/', '').replace('-', ''))
        else:
            cnpj_formatado = ''
        if cpf:
            cpf_formatado = str(cpf.replace('.', '').replace('/', '').replace('-', ''))
        else:
            cpf_formatado = ''
        try:
            response = requests.get(f"{NIBO_API_BASE_URL}/empresas/v1/customers/?organization={NIBO_ORGANIZATION}&$filter=document/number eq '{cnpj_formatado or cpf_formatado}'&ApiToken={NIBO_API_TOKEN}")
        except Exception as e:
            print(f"Erro ao buscar empresa: {e}")

        if response.status_code == 200:
            data = response.json()
            return data['items'][0]

def agendar_recebimento(empresa, valor, mes, ano):
    # Verificar se o valor é um float, int ou string representando um número
    try:
        valor = float(valor)
        # Truncar para no máximo duas casas decimais
        valor = round(valor, 2)
    except Exception as error:
        print(f"Erro ao tratar o valor: {error}")
        return False

    # Incrementar o mês
    mes = int(mes)
    ano = int(ano)
    
    # Verificar se o mês é o de dezembro
    if mes == 12:
        mes_vencimento = 1
        ano_vencimento = ano + 1
    else:
        mes_vencimento = mes + 1
        ano_vencimento = ano
    
    # Ajustar o formato do mês para dois dígitos
    mes_vencimento = f"{mes_vencimento:02d}"

    now = datetime.now()
    today_datetime = now.strftime("%Y-%m-%dT%H:%M:%S")

    # Definir o dia da data de vencimento
    if 2 < now.day <= 5:
        dia = f"{now.day:02d}"
    else:
        dia = "02"

    # Criar a data de lançamento no formato ISO 8601
    data_vencimento = f"{ano_vencimento}-{mes_vencimento}-{dia}T00:00:00"

    # Verificar se a data de vencimento é retroativa, se for retroativa, retornar uma mensagem de erro
    if not (now < datetime.fromisoformat(data_vencimento)):
        print(f"Data de vencimento retroativa, data de vencimento: {data_vencimento} | data atual: {today_datetime}")
        return False
    
    description_agendamento = f"Salários a pagar, FGTS, GPS, provisão direitos trabalhistas, vale transporte e taxa de administração de pessoas {mes:02d}/{ano}"

    json_agendamento = {
        "stakeholderId": str(empresa['id']),
        "description": description_agendamento,
        "value": valor,
        "scheduleDate": today_datetime,
        "dueDate": data_vencimento,
        "categoryId": NIBO_CATEGORY_ID,
        "isFlagged": False,
    }

    try:
        response_agendamento = requests.post(f"{NIBO_API_BASE_URL}/empresas/v1/schedules/credit/?organization={NIBO_ORGANIZATION}&ApiToken={NIBO_API_TOKEN}", json=json_agendamento)
        if response_agendamento.status_code == 200:
            response_data_agendamento = response_agendamento.json()
            json_boleto = {
                "accountId": NIBO_AUTOMACAO_ID,
                "scheduleId": response_data_agendamento,
                "dueDate": data_vencimento,
                "items": [{
                    "description": "Sem detalhamento",
                    "quantity": 1,
                    "value": valor
                }]
            }

            try:
                response_boleto = requests.post(f"{NIBO_API_BASE_URL}/empresas/v1/schedules/credit/{response_data_agendamento}/promise?organization={NIBO_ORGANIZATION}&ApiToken={NIBO_API_TOKEN}", json=json_boleto)
                response_data_boleto = response_boleto.json()
                if response_boleto.status_code == 200:
                    response_data_boleto_agendado = consultar_boleto_recebimento_agendado(response_data_agendamento)
                    boleto = download_boleto_recebimento(response_data_boleto_agendado)
                    if boleto:
                        response = {
                            "idAgendamento": response_data_agendamento,
                            "idBoleto": response_data_boleto
                        }
                        return response
                    else:
                        print(f"Erro ao baixar boleto: {boleto}")
                        return False
                else:
                    print(f"Erro ao gerar boleto: {response_boleto.status_code} - {response_boleto.text}")
                    return False
            except Exception as e:
                print(f"Erro ao gerar boleto: {e}")
                input()
                return False
        else:
            print(f"Erro ao agendar Recebimento: {response_agendamento.status_code} - {response_agendamento.text}")
            return False
    except Exception as e:
        print(f"Erro ao agendar Recebimento: {e}")
        input()
        return False

def cancelar_agendamento_de_recebimento(id_agendamento):
    try:
        response = requests.delete(f"{NIBO_API_BASE_URL}/empresas/v1/schedules/debit/{id_agendamento}?organization={NIBO_ORGANIZATION}&ApiToken={NIBO_API_TOKEN}")
        if response.status_code == 204:
            return True
        else:
            return False
    except Exception as e:
        print(f"Erro ao cancelar agendamento: {e}")
        return False

def pegar_agendamento_de_pagamento_cliente_por_data(id_cliente, mes, ano):
    # Incrementar o mês
    mes = int(mes)
    ano = int(ano)
    
    # Verificar se o mês é o de dezembro
    if mes == 12:
        mes_vencimento = 1
        ano_vencimento = ano + 1
    else:
        mes_vencimento = mes + 1
        ano_vencimento = ano

    try:
        response = requests.get(f"{NIBO_API_BASE_URL}/empresas/v1/customers/{id_cliente}/schedules/?organization={NIBO_ORGANIZATION}&$filter=year(dueDate) eq {ano_vencimento} and month(dueDate) eq {mes_vencimento} and category/id eq {NIBO_CATEGORY_ID}&ApiToken={NIBO_API_TOKEN}")
        if response.status_code == 200:
            data = response.json()
            return data['items'][0]
        else:
            False
    except Exception as e:
        print(f"Erro ao buscar Agendamento: {e}")
        return False

def consultar_boleto_recebimento_agendado(id_agendamento):
    try:
        response = requests.get(f"{NIBO_API_BASE_URL}/empresas/v1/schedules/credit/{id_agendamento}/promise?organization={NIBO_ORGANIZATION}&ApiToken={NIBO_API_TOKEN}")

        if response.status_code == 200:
            data = response.json()
            return data['items'][0]
    except Exception as e:
        print(f"Erro ao buscar Agendamento: {e}")
        return False
    return False

def get_download_path():
    """Returns the default downloads path for Linux, MacOS, or Windows."""
    if platform.system() == 'Windows':
        sub_key = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders'
        downloads_guid = '{374DE290-123F-4565-9164-39C4925E467B}'
        
        with OpenKey(HKEY_CURRENT_USER, sub_key) as key:
            downloads_path = QueryValueEx(key, downloads_guid)[0]
        
        return str(downloads_path)
    
    else:
        return str(Path.home() / 'Downloads')

def download_boleto_recebimento(boleto):
    try:
        retry = 0
        while retry <= 2:
            response_data_boleto_agendado_pdf = requests.get(boleto['url'], stream=True)

            if response_data_boleto_agendado_pdf.status_code == 200 and response_data_boleto_agendado_pdf.headers.get('Content-Type') == 'application/pdf':
                arquivo = f"{get_download_path()}\\boleto.pdf"
                with open(arquivo, 'wb') as file:
                    file.write(response_data_boleto_agendado_pdf.content)
                    print("Boleto gerado com sucesso!")
                    return True
            else:
                retry += 1
                sleep(1)

        print(f"Erro ao baixar o boleto! {response_data_boleto_agendado_pdf.status_code}")
        return False
    except Exception as error:
        print(f"Erro ao baixar o boleto: {error}")
        input()
        return False
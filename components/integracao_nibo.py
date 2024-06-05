from components.configuracao_db import configura_db
from components.procura_cliente import procura_cliente_por_id
from dotenv import load_dotenv
from datetime import datetime
from dateutil.relativedelta import relativedelta
import os, requests, json

# ================= CARREGANDO VARIÁVEIS DE AMBIENTE======================
load_dotenv()

# =====================CONFIGURAÇÂO DO BANCO DE DADOS======================
db_conf = configura_db()

def listar_empresas_clientes():
    NIBO_API_BASE_URL = os.getenv('NIBO_API_BASE_URL')
    NIBO_API_TOKEN = os.getenv('NIBO_API_TOKEN')
    NIBO_ORGANIZATION = os.getenv('NIBO_ORGANIZATION')
    empresas = []
    response = requests.get(f"{NIBO_API_BASE_URL}/empresas/v1/customers?organization={NIBO_ORGANIZATION}&ApiToken={NIBO_API_TOKEN}")
    if response.status_code == 200:
        for empresa in response.json()['items']:
            empresas.append(empresa)
    return empresas

def pegar_empresa_por_id(id):
    NIBO_API_BASE_URL = os.getenv('NIBO_API_BASE_URL')
    NIBO_API_TOKEN = os.getenv('NIBO_API_TOKEN')
    NIBO_ORGANIZATION = os.getenv('NIBO_ORGANIZATION')

    cliente = procura_cliente_por_id(id, db_conf)
    if cliente:
        cnpj = cliente[2]
        cpf = cliente[3]

        try:
            response = requests.get(f"{NIBO_API_BASE_URL}/empresas/v1/customers/?organization={NIBO_ORGANIZATION}&$filter=document/number eq '{cnpj or cpf}'&ApiToken={NIBO_API_TOKEN}")
        except Exception as e:
            print(f"Erro ao buscar empresa: {e}")

        if response.status_code == 200:
            data = response.json()
            return data['items'][0]

def pegar_agendamento_de_pagamento_cliente_por_data(id_cliente, mes, ano):
    NIBO_API_BASE_URL = os.getenv('NIBO_API_BASE_URL')
    NIBO_API_TOKEN = os.getenv('NIBO_API_TOKEN')
    NIBO_ORGANIZATION = os.getenv('NIBO_ORGANIZATION')
    CATEGORY_ID = os.getenv('CATEGORY_ID')

    if int(mes) == 12:
        mes = "01"
    else:
        if int(mes) > 0 and int(mes) < 10: 
            mes = "0" + str(int(mes) + 1)
        elif int(mes) > 9:
            mes = str(int(mes) + 1)

    try:
        response = requests.get(f"{NIBO_API_BASE_URL}/empresas/v1/customers/{id_cliente}/schedules/?organization={NIBO_ORGANIZATION}&$filter=year(dueDate) eq {ano} and month(dueDate) eq {mes} and category/id eq {CATEGORY_ID}&ApiToken={NIBO_API_TOKEN}")
    except Exception as e:
        print(f"Erro ao buscar Agendamento: {e}")
    
    if response.status_code == 200:
        data = response.json()
        return data['items'][0]

def agendar_recebimento(cliente_id, valor, mes, ano):
    NIBO_API_BASE_URL = os.getenv('NIBO_API_BASE_URL')
    NIBO_API_TOKEN = os.getenv('NIBO_API_TOKEN')
    NIBO_ORGANIZATION = os.getenv('NIBO_ORGANIZATION')
    CATEGORY_ID = os.getenv('CATEGORY_ID')

    if int(mes) == 12:
        mes = "01"
        ano = str(int(ano) + 1)
    else:
        if int(mes) > 0 and int(mes) < 10: 
            mes = "0" + str(int(mes) + 1)
        elif int(mes) > 9:
            mes = str(int(mes) + 1)
    now = datetime.now()
    today_datetime = now.strftime("%Y-%m-%dT%H:%M:%S")
    if now.day > 2 and now.day < 5:
        data_lancamento = f"0{str(now.day)}{mes}{ano}"
    else:
        data_lancamento = f"02{mes}{ano}"
    print(f"Data lançamento boleto: {data_lancamento} | Data de hoje: {today_datetime}")

    """try:
        response = requests.post(f"{NIBO_API_BASE_URL}/empresas/v1/schedules/credit/?organization={NIBO_ORGANIZATION}&ApiToken={NIBO_API_TOKEN}", json={
            stakeholderId: cliente_id,
            value: valor,

        })"""

def cancelar_agendamento_de_recebimento(id_agendamento):
    NIBO_API_BASE_URL = os.getenv('NIBO_API_BASE_URL')
    NIBO_API_TOKEN = os.getenv('NIBO_API_TOKEN')
    NIBO_ORGANIZATION = os.getenv('NIBO_ORGANIZATION')
    CATEGORY_ID = os.getenv('CATEGORY_ID')
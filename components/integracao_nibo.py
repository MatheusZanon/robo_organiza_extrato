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

def agendar_recebimento(cliente, valor, mes, ano):
    NIBO_API_BASE_URL = os.getenv('NIBO_API_BASE_URL')
    NIBO_API_TOKEN = os.getenv('NIBO_API_TOKEN')
    NIBO_ORGANIZATION = os.getenv('NIBO_ORGANIZATION')
    NIBO_CATEGORY_ID = os.getenv('NIBO_CATEGORY_ID')
    NIBO_AUTOMACAO_ID = os.getenv('NIBO_AUTOMACAO_ID')

    # Verificar se o valor é um float, int ou string representando um número
    try:
        valor = float(valor)
        # Truncar para no máximo duas casas decimais
        valor = round(valor, 2)
    except ValueError:
        raise ValueError("O valor deve ser um número válido.")

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

    json_agendamento = {
        "stakeholderId": str(cliente['id']),
        "description": f"Salários a pagar, FGTS, GPS, provisão direitos trabalhistas, vale transporte e taxa de administração de pessoas {mes:02d}/{ano}",
        "value": valor,
        "scheduleDate": today_datetime,
        "dueDate": data_vencimento,
        "categoryId": NIBO_CATEGORY_ID,
        "isFlagged": False,
    }

    print(f"agendamento: {json.dumps(json_agendamento, indent=4)}")
    input("Pressione Enter para criar o boleto...")

    try:
        response_agendamento = requests.post(f"{NIBO_API_BASE_URL}/empresas/v1/schedules/credit/?organization={NIBO_ORGANIZATION}&ApiToken={NIBO_API_TOKEN}", json=json_agendamento)

        if response_agendamento.status_code == 200:
            response_data_agendamento = response_agendamento.json()
            print(f"response: {json.dumps(response_data_agendamento, indent=4)}")
            input()

            json_boleto = {
                "accountId": NIBO_AUTOMACAO_ID,
                "scheduleId": response_data_agendamento,
                "value": valor,
                "dueDate": data_vencimento,
                "bankSlipInstructions": "Salários a pagar, FGTS, GPS, provisão direitos trabalhistas, vale transporte e taxa de administração de pessoas {mes:02d}/{ano}",
                "stakeholderInfo": {
                    "document": cliente['document']['number'],
                    "name": cliente['name'],
                    "email": cliente['email'],
                    "street": cliente['address']['line1'],
                    "number": cliente['address']['number'],
                    "district": cliente['address']['district'],
                    "complement": cliente['address']['line2'],
                    "state": cliente['address']['state'],
                    "city": cliente['address']['city'],
                    "zipCode": str(cliente['address']['zipCode']).strip()
                },
                "items": [{
                    "description": "Sem detalhamento",
                    "quantity": 1,
                    "value": valor
                }]
            }

            print(f"boleto: {json.dumps(json_boleto, indent=4)}")
            print(f"url: {NIBO_API_BASE_URL}/empresas/v1/schedules/credit/{response_data_agendamento}/promise?organization={NIBO_ORGANIZATION}&ApiToken={NIBO_API_TOKEN}")
            input("Pressione Enter para criar o boleto...")

            try:
                response_boleto = requests.post(f"{NIBO_API_BASE_URL}/empresas/v1/schedules/credit/{response_data_agendamento}/promise?organization={NIBO_ORGANIZATION}&ApiToken={NIBO_API_TOKEN}", json=json_boleto)

                response_data_boleto = response_boleto.json()
                print(f"response: {json.dumps(response_data_boleto, indent=4)}")
                if response_boleto.status_code == 200:
                    return {
                        "idAgendamento": response_data_agendamento,
                        "idBoleto": response_data_boleto
                    }
            except Exception as e:
                print(f"Erro ao gerar boleto: {e}")
    except Exception as e:
        print(f"Erro ao agendar Recebimento: {e}")

def cancelar_agendamento_de_recebimento(id_agendamento):
    NIBO_API_BASE_URL = os.getenv('NIBO_API_BASE_URL')
    NIBO_API_TOKEN = os.getenv('NIBO_API_TOKEN')
    NIBO_ORGANIZATION = os.getenv('NIBO_ORGANIZATION')

    try:
        response = requests.delete(f"{NIBO_API_BASE_URL}/empresas/v1/schedules/debit/{id_agendamento}?organization={NIBO_ORGANIZATION}&ApiToken={NIBO_API_TOKEN}")
        print(f"response: {json.dumps(response.json(), indent=4)}")
        if response.status_code == 204:
            return True
        else:
            return False
    except Exception as e:
        print(f"Erro ao cancelar agendamento: {e}")
        return False
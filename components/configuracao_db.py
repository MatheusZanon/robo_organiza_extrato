import os
from aws_parameters import get_ssm_parameter

def configura_db():    
    db_conf = {
        "host": get_ssm_parameter('/human/DB_HOST'),
        "user": get_ssm_parameter('/human/DB_USER'),
        "password": get_ssm_parameter('/human/DB_PASS'),
        "database": get_ssm_parameter('/human/DB_NAME')
    }
    return db_conf

def ler_sql(arquivo_sql):
    with open(arquivo_sql, 'r', encoding='utf-8') as arquivo:
        return arquivo.read()
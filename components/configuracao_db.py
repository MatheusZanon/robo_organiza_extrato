import os

def configura_db():    
    db_conf = {
        "host": os.getenv('DB_HOST'),
        "user": os.getenv('DB_USER'),
        "password": os.getenv('DB_PASS'),
        "database": os.getenv('DB_NAME')
    }
    return db_conf

def ler_sql(arquivo_sql):
    with open(arquivo_sql, 'r', encoding='utf-8') as arquivo:
        return arquivo.read()
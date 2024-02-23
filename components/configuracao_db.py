import os
import mysql.connector
from mysql.connector import errorcode

def configura_db():    
    db_conf = {
        "host": os.getenv('DB_HOST'),
        "user": os.getenv('DB_USER'),
        "password": os.getenv('DB_PASS'),
        "database": os.getenv('DB_NAME')
    }

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
    
    return db_conf, conn, cursor

def ler_sql(arquivo_sql):
    with open(arquivo_sql, 'r', encoding='utf-8') as arquivo:
        return arquivo.read()
from components.configuracao_db import ler_sql
import mysql.connector

def procura_cliente(nome_cliente, db_conf):
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
            cliente_mod = procura_cliente_mod(str(nome_cliente).replace("S S", "S/S"), db_conf)
            return cliente_mod
    except Exception as error:
        print(error)

def procura_cliente_mod(nome_cliente, db_conf):
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

def procura_cliente_por_regiao(nome_cliente, db_conf):
    try:
        query_procura_cliente = ler_sql('sql/procura_cliente_por_regiao.sql')
        values_procura_cliente = (nome_cliente,)
        with mysql.connector.connect(**db_conf) as conn, conn.cursor() as cursor:
            cursor.execute(query_procura_cliente, values_procura_cliente)
            cliente = cursor.fetchone()
            conn.commit()
        if cliente:
            return cliente
        else:
            return None
    except Exception as error:
        print(error)

def procura_cliente_por_id(cliente_id, db_conf):
    try:
        query_procura_cliente = ler_sql('sql/procura_cliente_por_id.sql')
        values_procura_cliente = (cliente_id,)
        with mysql.connector.connect(**db_conf) as conn, conn.cursor() as cursor:
            cursor.execute(query_procura_cliente, values_procura_cliente)
            cliente = cursor.fetchone()
            conn.commit()
            conn.close()
        if cliente:
            return cliente
        else:
            return None
    except Exception as error:
        print(error)

def procura_clientes(db_conf):
    try:
        query_procura_clientes = ler_sql('sql/procura_clientes.sql')
        with mysql.connector.connect(**db_conf) as conn, conn.cursor() as cursor:
            cursor.execute(query_procura_clientes)
            clientes = cursor.fetchall()
            conn.commit()
            conn.close()
        return clientes
    except Exception as error:
        print(error)
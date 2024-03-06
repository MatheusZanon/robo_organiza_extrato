from components.configuracao_db import ler_sql
import mysql.connector

def procura_valores(cliente_id, db_conf, mes, ano):
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


def procura_valores_com_codigo(cliente_id, cod_centro_custo, db_conf, mes, ano):
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
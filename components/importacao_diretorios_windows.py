"""IMPORTAÇÃO DE BIBLIOTECA PARA MEXER COM DIRETÓRIOS E ARQUIVOS NO WINDOWS"""
import os
from pathlib import Path

def listagem_pastas(diretorio):
    try:
        lista_pastas = []
        pastas = Path(diretorio)
        for pasta in pastas.iterdir():
            if os.path.isdir(pasta):
                lista_pastas.append(f"{pasta}")
        return sorted(lista_pastas) 
    except FileNotFoundError as notFoundError:
        print(notFoundError)
    except Exception as exc:
        print(f"Ocorreu alguma falha no processo: {exc}")

def listagem_arquivos(diretorio):
    try:
        lista_arquivos = []
        arquivos = Path(diretorio)
        for arquivo in arquivos.iterdir():
            if os.path.isfile(arquivo):
                lista_arquivos.append(f"{arquivo}")
        return sorted(lista_arquivos)
    except FileNotFoundError as notFoundError:
        print(notFoundError)
    except Exception as exc:
        print(f"Ocorreu alguma falha no processo: {exc}")

def listagem_arquivos_downloads():
    try:
        downloads = os.path.expanduser("~") + "\\Downloads"
        lista_arquivos = []
        for arquivo in Path(downloads).iterdir():
            if arquivo.is_file():
                lista_arquivos.append(str(arquivo))
        return lista_arquivos
    except FileNotFoundError as notFoundError:
        print(notFoundError)
    except Exception as exc:
        print(f"Ocorreu alguma falha no processo: {exc}")

def procura_pasta_cliente(nome, lista_dir_clientes):
    try:
        nome = nome.replace("S/S", "S S")
        caminho_pasta_cliente = None
        for diretorio in lista_dir_clientes:
            if not caminho_pasta_cliente == None:
                break 
            else:
                pastas_cliente = listagem_pastas(diretorio)
                for pasta in pastas_cliente:
                    if not caminho_pasta_cliente == None:
                        break 
                    else:
                        nome_pasta_cliente = pega_nome(pasta)
                        if nome_pasta_cliente == nome:
                            caminho_pasta_cliente = pasta
                            return caminho_pasta_cliente
        return caminho_pasta_cliente
    except Exception as error:
        print(error)

def pega_nome(path):
    try:
        nome_objeto = os.path.basename(path)
        return nome_objeto
    except Exception as exc:
        print(f"Ocorreu alguma falha no processo: {exc}")

import pandas as pd

def carrega_arquivo(arquivo):
    try:
        df = pd.read_excel(arquivo)
        return df
    except Exception as error:
        print(error)
from components.importacao_diretorios_windows import listagem_pastas, listagem_arquivos, pega_nome
from components.extract_text_pdf import extract_text_pdf
from components.db_config import db_config
from components.importacao_caixa_dialogo import DialogBox
import tkinter as tk
import mysql.connector
from mysql.connector import errorcode
import re

# CONFIGURAÇÂO DO BANCO DE DADOS
db_conf = db_config()
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

# CAIXA DE DIALOGO INICIAL
def main():
    root = tk.Tk()
    app = DialogBox(root)
    root.mainloop()
    return app.particao, app.mes, app.ano

if __name__ == "__main__":
    particao, mes, ano = main()

# PARAMETROS INICIAS
diretorio_extratos = f"{particao}:\\Meu Drive\\Robo_Emissao_Relatorios_do_Mes\\faturas_human_{mes}_{ano}"

# MÉTODOS DE CADA ETAPA DO PROCESSO
def organiza_extratos():
    try:
        pasta_faturas = listagem_pastas(diretorio_extratos)
        for pasta in pasta_faturas:
            extratos = listagem_arquivos(pasta)
            for extrato in extratos:
                if extrato.__contains__(".pdf"):
                    nome_extrato = pega_nome(extrato)
                    texto_pdf = extract_text_pdf(extrato)

                    # VARREDURA DE DADOS DO EXTRATO PDF
                    # Exemplo de extração para Texto
                    match_centro_custo = re.search(r"C\.Custo:\s*(.*)", texto_pdf)
                    if match_centro_custo:
                        nome_centro_custo = match_centro_custo.group(1)
                        print(f"Centro de Custo: {nome_centro_custo}")

                    # NUMERO DE EMPREGADOS
                    # Exemplo de extração para Números
                    match_demitido = re.search(r"No. Empregados: Demitido:\s*(\d+)", texto_pdf)
                    if match_demitido:
                        demitido = match_demitido.group(1)
                        match_num_empregados = re.search(r"No. Empregados: Demitido:\s+" + demitido + r"\s*(\d+)", texto_pdf)
                        if match_num_empregados:
                            num_empregados = match_num_empregados.group(1)
                            print(f"Número de empregados: {num_empregados}")

                    # NUMERO DE ESTAGIARIOS
                    match_transferido = re.search(r"No. Estagiários: Transferido:\s*(\d+)", texto_pdf)
                    if match_transferido:
                        transferido = match_transferido.group(1)
                        match_num_estagiarios = re.search(r"No. Estagiários: Transferido:\s+" + transferido + r"\s*(\d+)", texto_pdf)
                        if match_num_estagiarios:
                            num_estagiarios = match_num_estagiarios.group(1)
                            print(f"Número de estagiários: {num_estagiarios}")

                    # TRABALHANDO
                    match_ferias = re.search(r"Trabalhando: Férias:\s*(\d+)", texto_pdf)
                    if match_ferias:
                        ferias = match_ferias.group(1)
                        match_trabalhando = re.search(r"Trabalhando: Férias:\s+" + ferias + r"\s*(\d+)", texto_pdf)
                        if match_trabalhando:
                            trabalhando = match_trabalhando.group(1)
                            print(f"Trabalhando: {trabalhando}")

                    # SALARIO CONTRIBUIÇÃO EMPREGADOS
                    match_salario_contri_empregados = re.search(r"Salário contribuição empregados:\s*([\d.,]+)", texto_pdf)
                    if  match_salario_contri_empregados:
                        salario_contri_empregados = match_salario_contri_empregados.group(1).replace(".", "").replace(",", ".")
                        print(f"Salário contribuição Empregados: {salario_contri_empregados}")

                    # SALARIO CONTRIBUIÇÃO CONTRIBUINTES
                    match_salario_contri_contribuintes = re.search(r"Salário contribuição contribuintes:\s*([\d.,]+)", texto_pdf)
                    if  match_salario_contri_contribuintes:
                        salario_contri_contribuintes = match_salario_contri_contribuintes.group(1).replace(".", "").replace(",", ".")
                        print(f"Salário contribuição Contribuintes: {salario_contri_contribuintes}")

                    # VALOR DO INSS
                    # A expressão regular procura por um ou mais números seguidos por qualquer coisa (não capturada)
                    # e então "Total INSS:"
                    match_inss = re.search(r"Total INSS:\s*([\d.,]+)", texto_pdf)
                    if match_inss:
                        inss = match_inss.group(1).replace(".", "").replace(",", ".")
                        print(f"Total INSS: {inss}")

                    # VALOR DO FGTS
                    match_fgts = re.search(r"Valor do FGTS:\s*([\d.,]+)", texto_pdf)
                    if  match_fgts:
                        fgts = match_fgts.group(1).replace(".", "").replace(",", ".")
                        print(f"Valor do FGTS: {fgts}")

                    # VALOR DO IRRF
                    match_base_iss = re.search(r"([\d.,]+)\s+Valor Total do IRRF: Base ISS:", texto_pdf)
                    if match_base_iss:
                        base_iss = match_base_iss.group(1)
                        match_irrf = re.search(r"([\d.,]+)\s+" + base_iss + r"\s+Valor Total do IRRF: Base ISS:", texto_pdf)
                        if match_irrf:
                            irrf = match_irrf.group(1)
                            print(f"Valor Total do IRRF: {irrf}")

                    # LÍQUIDO CENTRO DE CUSTO
                    match_liquido = re.search(r"Líquido Centro de Custo:\s*([\d.,]+)", texto_pdf)
                    if  match_liquido:
                        liquido_centro_custo = match_liquido.group(1).replace(".", "").replace(",", ".")
                        print(f"Líquido Centro de Custo: {liquido_centro_custo}")

                    input()
    except Exception as error:
        print(error)

def gera_fatura():
    return "Processo de gerar fatura_detalhada"

def gera_boleto():
    return "Processo de agendar e baixar boletos do Nibo"


# CÓDIGO PRINCIPAL DO ROBÔ
organiza_extratos()

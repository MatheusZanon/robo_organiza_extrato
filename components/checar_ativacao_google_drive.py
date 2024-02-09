import time
import subprocess

def checa_google_drive():
    # Nome do processo do Google Drive File Stream
    nome_processo_drive = "GoogleDriveFS.exe"

    # Listar processos em execução e verificar se o Google Drive File Stream está entre eles
    processo_ativo = False
    try:
        processos = subprocess.check_output(['tasklist']).decode('cp1252').split('\r\n')
    except UnicodeDecodeError:
        processos = subprocess.check_output(['tasklist']).decode('utf-16').split('\r\n')

    for proc in processos:
        if nome_processo_drive in proc:
            processo_ativo = True
            break

    # Se o Google Drive File Stream não estiver em execução, iniciá-lo
    if not processo_ativo:
        caminho_executavel_drive = r"C:\Program Files\Google\Drive File Stream\launch.bat"
        subprocess.Popen(caminho_executavel_drive, shell=True)
        time.sleep(3)
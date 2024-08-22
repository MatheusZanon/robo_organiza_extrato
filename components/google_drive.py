from aws_parameters import get_ssm_parameter
import json
import os
from pathlib import Path
import io
import boto3
from botocore.exceptions import ClientError
from google.auth.transport.requests import Request
from google.auth import identity_pool
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload

def get_secret():
    secret_name = "GoogleFederationConfig"
    region_name = "sa-east-1"

    # Create a Secrets Manager client
    session = boto3.session.Session()
    client = session.client(
        service_name='secretsmanager',
        region_name=region_name
    )
    try:
        get_secret_value_response = client.get_secret_value(
            SecretId=secret_name
        )
    except ClientError as e:
        # For a list of exceptions thrown, see
        # https://docs.aws.amazon.com/secretsmanager/latest/apireference/API_GetSecretValue.html
        raise e

    secret = get_secret_value_response['SecretString']
    return json.loads(secret)

def carregar_credenciais():
    try:
        secret_name = "GoogleFederationConfig"
        secret_json = get_secret(secret_name)
        
        credentials = identity_pool.Credentials.from_info(secret_json)

        SCOPES = [get_ssm_parameter('/human/API_SCOPES')]
        credentials = credentials.with_scopes(SCOPES)

        credentials.refresh(Request())
        return credentials
    except Exception as error:
        print(error)

def autenticacao_google_drive():
    try:
        service_name = get_ssm_parameter('/human/API_NAME')
        service_version = get_ssm_parameter('/human/API_VERSION')
        credentials = carregar_credenciais()
        drive_service = build(service_name, service_version, credentials=credentials)
        return drive_service
    except Exception as error:
        print(error)


def lista_pastas_em_diretorio(driver_service, folder_id):
    try:
        query = f"'{folder_id}' in parents and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
        results = driver_service.files().list(q=query, pageSize=80, fields="files(id, name)").execute()
        items = results.get('files', [])
        return items
    except Exception as error:
        print(error)

def pegar_arquivo(files: list[dict[str, any]], nome_arquivo: str):
    try:
        for file in files:
            if file['name'] == nome_arquivo:
                return file
        return None
    except Exception as error:
        print(error)

def lista_pastas_subpastas_em_diretorio(folder_id):
    try:
        all_files = []
        folders_to_process = [folder_id]
    except Exception as error:
        print(error)

def procura_pasta_drive_por_nome(driver_service, folder_name: str) -> list[dict] | None:
    """
    .. summary::
    Procura pastas no Google Drive pelo nome fornecido.

    Args:
        folder_name (str): Nome da pasta a ser pesquisada. Caracteres '/' são tratados como separadores de diretórios e 
                           "S/S" é substituído por "S S" para evitar problemas na pesquisa.

    Returns:
        list[dict]: Lista de dicionários contendo os dados das pastas encontradas. Cada dicionário inclui pelo menos as chaves 'id' e 'name'.
        None: Caso nenhuma pasta seja encontrada.

    Note:
        - A função realiza uma busca pela pasta com o nome exato fornecido e retorna todas as pastas que correspondem ao nome especificado.
        - Se várias pastas com o mesmo nome forem encontradas, todas serão retornadas na lista.

    Examples:
        1. Busca de uma pasta pelo nome:
            >>> procura_pasta_drive_por_nome("Documentos")
        # A pasta buscada será "Documentos".
        
        2. Busca de uma pasta com o nome incluindo separadores de diretórios:
            >>> procura_pasta_drive_por_nome("Relatórios/2023")
        # A pasta buscada será "Relatórios/2023".
        
        3. Caso especial onde "S/S" é convertido para "S S":
            >>> procura_pasta_drive_por_nome("Empresa S/S/relatorios/2024")
        # A pasta buscada será "Empresa S S/relatorios/2024".
    """
    try:
        # Substitui "S/S" por "S S" e mantém a estrutura de diretórios com "/"
        folder_name = folder_name.replace("S/S", "S S")
        query = f"name = '{folder_name}' and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
        found_items = []
        page_token = None
        while True:
            # Executa a busca e coleta os resultados
            results = driver_service.files().list(
                q=query, 
                pageSize=80, 
                fields="files(id, name), nextPageToken", 
                pageToken=page_token
            ).execute()
            items = results.get('files', [])
            found_items.extend(items)
            page_token = results.get('nextPageToken')
            if not page_token:
                break

        if not found_items:
            print("Nenhuma pasta encontrada.")
            return None

        print("Pastas encontradas:")
        for item in found_items:
            print(f'{item["name"]} ({item["id"]})')
        return found_items

    except Exception as error:
        print(f"Erro ao procurar pasta no Google Drive: {error}")
        return None

def procura_subpasta_drive_por_nome(driver_service, parent_folder_name: str, subfolder_name: str, intermediate_folders: list[str] = None, recursive=False) -> list[dict] | None:
    """
    Procura uma subpasta no Google Drive pelo nome da pasta pai e nome da subpasta.
    
    Args:
        parent_folder_name (str): Nome da pasta pai onde a busca começará. Caracteres '/' são tratados como separadores de diretórios.
        subfolder_name (str): Nome da subpasta a ser pesquisada. Caracteres '/' são tratados como separadores de diretórios.
        intermediate_folders (list of str, optional): Lista de nomes de diretórios intermediários a serem pesquisados sequencialmente antes da subpasta final. Cada nome deve corresponder a uma pasta que deve ser encontrada na ordem especificada. Caracteres '/' são tratados como separadores de diretórios.
        recursive (bool, optional): Se verdadeiro, a busca será feita recursivamente em todas as subpastas, não apenas no nível atual.

    Returns:
        list[dict]: Lista de dicionários contendo os dados das subpastas encontradas. Cada dicionário inclui pelo menos as chaves 'id' e 'name'.
        None: Caso nenhuma subpasta seja encontrada.

    Note:
        - Se `recursive` for verdadeiro, a função irá buscar a subpasta em todas as subpastas do diretório pai especificado ou da última pasta intermediária encontrada.
        - Se `intermediate_folders` for fornecido, a função procurará cada pasta intermediária na ordem fornecida antes de buscar pela subpasta final. Caso algum diretório intermediário não seja encontrado, a função retorna `None`.

    Examples:
        1. Busca simples (diretório pai é diretamente pesquisado):
            >>> procura_subpasta_drive_por_nome("Cliente", "2023")
        # A função irá buscar na pasta "Cliente/2023".
        
        2. Busca com diretórios intermediários (procura na pasta 'Cliente', em seguida 'Economia Mensal', e finalmente 'Relatórios' para encontrar '2023'):
            >>> procura_subpasta_drive_por_nome("Cliente", "2023", intermediate_folders=["Economia Mensal", "Relatórios"])
        # A função irá buscar na estrutura "Cliente/Economia Mensal/Relatórios/2023".
        
        3. Busca recursiva (procura '2023' em todas as subpastas a partir da pasta 'Cliente'):
            >>> procura_subpasta_drive_por_nome("Cliente", "2023", recursive=True)
        # A função irá buscar recursivamente na estrutura "Cliente/*/2023".
        
        4. Busca com diretórios intermediários e recursiva (procura '2023' em 'Cliente', 'Economia Mensal', e subpastas):
            >>> procura_subpasta_drive_por_nome("Cliente", "2023", intermediate_folders=["Economia Mensal"], recursive=True)
        # A função irá buscar na estrutura "Cliente/Economia Mensal/*/2023".
        
        5. Busca em pastas com nome contendo "S/S" (o caractere '/' é tratado como separador de diretórios e "S/S" é substituído por "S S"):
            >>> procura_subpasta_drive_por_nome("EMPRESA TESTE S/S", "2024/06/relatorios", intermediate_folders=["financeiro"])
        # A função irá buscar na estrutura "EMPRESA TESTE S S/financeiro/2024/06/relatorios".
    """
    try:
        # Substituir "S/S" por "S S" e tratar '/' como separadores de diretórios
        parent_folder_name = parent_folder_name.replace("S/S", "S S")
        subfolder_name = subfolder_name.replace("S/S", "S S")
        if intermediate_folders:
            intermediate_folders = [folder.replace("S/S", "S S") for folder in intermediate_folders]

        # Função auxiliar para buscar a pasta pai
        def find_parent_folder(parent_name):
            parent_name_parts = parent_name.split("/")
            parent_folder_id = None
            for part in parent_name_parts:
                query = f"name = '{part}' and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
                if parent_folder_id:
                    query = f"'{parent_folder_id}' in parents and {query}"
                results = driver_service.files().list(q=query, pageSize=80, fields="files(id, name)").execute()
                items = results.get('files', [])
                if not items:
                    return None
                parent_folder_id = items[0]['id']
            return items[0]

        # Primeiro, encontre a pasta pai pelo nome
        parent_folder = find_parent_folder(parent_folder_name)
        if not parent_folder:
            print("Pasta pai não encontrada.")
            return None

        parent_folder_id = parent_folder['id']

        # Se houver diretórios intermediários, procure dentro deles na sequência
        if intermediate_folders:
            for intermediate_folder_name in intermediate_folders:
                query = f"'{parent_folder_id}' in parents and name = '{intermediate_folder_name}' and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
                results = driver_service.files().list(q=query, pageSize=80, fields="files(id, name)").execute()
                items = results.get('files', [])
                if not items:
                    print(f"Diretório intermediário '{intermediate_folder_name}' não encontrado.")
                    return None

                intermediate_folder = items[0]
                parent_folder_id = intermediate_folder['id']

        # Se a busca for recursiva
        if recursive:
            # Função auxiliar para busca recursiva
            def busca_recursiva(folder_id, subfolder_name):
                found_items = []
                page_token = None
                while True:
                    query = f"'{folder_id}' in parents and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
                    results = driver_service.files().list(
                        q=query, pageSize=80, fields="files(id, name, parents)", pageToken=page_token
                    ).execute()
                    items = results.get('files', [])
                    for item in items:
                        if item['name'] == subfolder_name:
                            found_items.append(item)
                        # Busca recursiva nas subpastas
                        found_items.extend(busca_recursiva(item['id'], subfolder_name))
                    
                    page_token = results.get('nextPageToken')
                    if not page_token:
                        break
                return found_items

            return busca_recursiva(parent_folder_id, subfolder_name)

        # Caso não seja recursiva, procure a subpasta diretamente
        query = f"'{parent_folder_id}' in parents and name = '{subfolder_name}' and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
        results = driver_service.files().list(q=query, pageSize=80, fields="files(id, name)").execute()
        items = results.get('files', [])
        if not items:
            print("Nenhuma subpasta encontrada.")
            return None

        print("Subpastas encontradas:")
        for item in items:
            print(f'{item["name"]} ({item["id"]})')
        return items  # Retorna todos os resultados encontrados

    except Exception as error:
        print(f"Erro ao procurar subpasta no Google Drive: {error}")
        return None

def listar_arquivos_drive(driver_service, folder_id):
    try:
        driver_service = autenticacao_google_drive()
        query = f"parents in '{folder_id}' and trashed = false"

        results = driver_service.files().list(q=query, pageSize=80, fields="files(id, name, mimeType, parents, modifiedTime), nextPageToken").execute()
        arquivos = results.get('files', [])

        return arquivos
    except Exception as error:
        print(f"Erro ao procurar arquivos no Google Drive: {error}")
        return None

def encontrar_pasta_por_nome(driver_service, folder_id, nome_pasta_desejada):
    try:
        # Listar todos os arquivos e pastas no diretório atual
        arquivos = listar_arquivos_drive(driver_service, folder_id)

        if arquivos is None:
            return None

        # Verificar se há uma pasta com o nome desejado
        for arquivo in arquivos:
            if arquivo['mimeType'] == 'application/vnd.google-apps.folder' and arquivo['name'] == nome_pasta_desejada:
                return arquivo  # Retornar a pasta encontrada
        
        return None

        # Se não encontrado, repetir o processo para todas as subpastas || RECURSIVO
        # for arquivo in arquivos:
        #    if arquivo['mimeType'] == 'application/vnd.google-apps.folder':
        #        subpasta = encontrar_pasta_por_nome(arquivo['id'], nome_pasta_desejada)
        #        if subpasta:
        #            return subpasta  # Retornar a subpasta encontrada

        return None  # Se nenhuma pasta for encontrada
    except Exception as error:
        print(f"Erro ao procurar pasta no Google Drive: {error}")
        return None

def encontrar_arquivo_por_nome(driver_service, folder_id, nome_arquivo_desejado):
    try:
        # Listar todos os arquivos e pastas no diretório atual
        arquivos = listar_arquivos_drive(driver_service, folder_id)

        if arquivos is None:
            return None

        # Verificar se há uma pasta com o nome desejado
        for arquivo in arquivos:
            if arquivo['name'] == nome_arquivo_desejado:
                return arquivo  # Retornar a pasta encontrada
        
        return None
    except Exception as error:
        print(f"Erro ao procurar arquivo no Google Drive: {error}")
        return None

def criar_pasta_drive(driver_service, folder_name, parent_folder_id):
    try:
        if not folder_name or not parent_folder_id:
            raise ValueError("Nome da pasta ou ID do diretório pai não podem estar vazio.")
        
        driver_service = autenticacao_google_drive()
        folder_metadata = {
            'name': folder_name,
            'mimeType': 'application/vnd.google-apps.folder',
            'parents': [parent_folder_id]
        }

        query = f"name = '{folder_name}' and mimeType = 'application/vnd.google-apps.folder' and trashed = false"

        results = driver_service.files().list(q=query, pageSize=80, fields="files(id, name)").execute()
        items = results.get('files', [])

        if items:
            raise ValueError("Pasta ja existe! Por favor escolha outro nome.")

        created_folder = driver_service.files().create(body=folder_metadata, fields='id').execute()
        return created_folder
    except Exception as error:
        print(f"Erro ao criar pasta no Google Drive: {error}")
        return None

def upload_arquivo_drive(driver_service, file_path, folder_id):
    try:
        if not file_path or not folder_id:
            raise ValueError("Caminho do arquivo ou ID da pasta não podem estar vazio.")

        driver_service = autenticacao_google_drive()
        file_metadata = {
            'name': os.path.basename(file_path),
            'parents': [folder_id]
        }

        media = MediaFileUpload(file_path, resumable=True)
        created_file = driver_service.files().create(body=file_metadata, media_body=media, fields='id').execute()
        return created_file
    except Exception as error:
        print(f"Erro ao criar arquivo no Google Drive: {error}")
        return None

def download_arquivo_drive(drive_service, file_id, file_path):
    try:
        request = drive_service.files().get_media(fileId=file_id)

        fh = io.FileIO(file_path, 'wb')
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while done is False:
            status, done = downloader.next_chunk()
            print(f"Download {int(status.progress() * 100)}%.")
        
        if done:
            return Path(file_path)
        
    except Exception as error:
        print(f"Erro ao baixar arquivo do Google Drive: {error}")
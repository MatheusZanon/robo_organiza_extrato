from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options

def configura_selenium_driver():
    # Opções de inicialização
    chrome_options = Options()
    # Para retirar a mensagem do Chrome de que "uma automação está controlando esse navegador" 
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option('useAutomationExtension', False)
    prefs = {"credentials_enable_service": False, 
            "profile.password_manager_enabled": False, # Desativar o prompt de salvamento de senha
            "download.prompt_for_download": False, # Para desativar a caixa de diálogo de download
            "plugins.always_open_pdf_externally": True # Para desativar a caixa de diálogo de download
            }
    chrome_options.add_experimental_option("prefs", prefs)

    # Caminho para o chrome driver
    caminho_drive = r'documents\\chromedriver-win64\\chromedriver.exe'
    # Configurar o serviço do ChromeDriver
    servico = Service(caminho_drive)

    return chrome_options, servico

from time import sleep, time
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

def procura_elemento(driver, tipo_seletor:str, elemento, tempo_espera):
    """
        Function to search for an element using the specified selector type, element, and wait time.
        driver: The WebDriver instance to use for element search.
        tipo_seletor: The type of selector to use (e.g., 'ID', 'CLASS_NAME', 'XPATH', 'TAG_NAME').
        elemento: The element to search for.
        tempo_espera: The maximum time to wait for the element to be located.
        :return: The located element if found, otherwise None.
    """
    try:
        seletor = getattr(By, tipo_seletor.upper())
        WebDriverWait(driver, float(tempo_espera)).until(EC.presence_of_element_located((seletor, elemento)))
        sleep(0.1)
        elemento = WebDriverWait(driver, float(tempo_espera)).until(EC.visibility_of_element_located((seletor, elemento)))
        if elemento.is_displayed() and elemento.is_enabled():
            return elemento
    except TimeoutException:
        return None

def procura_todos_elementos(driver, tipo_seletor:str, elemento, tempo_espera):
    """
    A function that searches for all elements based on the given selector type and element, within a specified waiting time.
    
    Args:
        driver: The WebDriver instance to use for locating the elements.
        tipo_seletor: A string representing the type of selector to use (e.g., 'ID', 'CLASS_NAME', 'XPATH', 'TAG_NAME').
        elemento: The element to search for.
        tempo_espera: The maximum time to wait for the elements to be present before throwing a TimeoutException.
        
    Returns:
        A list of WebElement objects representing the found elements, or None if the elements are not found within the specified waiting time.
    """
    try:
        seletor = getattr(By, tipo_seletor.upper())
        WebDriverWait(driver, float(tempo_espera)).until(EC.presence_of_all_elements_located((seletor, elemento)))
        sleep(0.1)
        elementos = WebDriverWait(driver, float(tempo_espera)).until(EC.visibility_of_all_elements_located((seletor, elemento)))
        if elementos:
            return elementos
    except TimeoutException:
        return None
    
def encontrar_elemento_shadow_root(driver, host, elemento, timeout):
    """Espera por um elemento dentro de um shadow-root até que o elemento esteja presente ou o tempo limite seja atingido."""
    end_time = time() + float(timeout)
    while True:
        try:
            # Tenta encontrar o elemento usando JavaScript
            js_script = f"""
            return document.querySelector('{host}').shadowRoot.querySelector('{elemento}');
            """
            element = driver.execute_script(js_script)
            if element:
                return element
        except Exception as e:
            pass  # Ignora erros e tenta novamente até que o tempo limite seja atingido
        sleep(0.1)  # Espera 1 segundo antes de tentar novamente
        if time() > end_time:
            break  # Sai do loop se o tempo limite for atingido
    return None
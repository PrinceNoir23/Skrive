import sys
import json
import subprocess
import socket
import re
import os
import pyautogui
import time
import urllib.request

import argparse
import pyperclip
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from pathlib import Path


CHROME_PATH = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
BASE_USER_DATA_DIR = r"C:\Skrive_Chrome\ChromeDebugProfiles"
START_PORT = 9222
MAX_PORT = 9300



def is_chrome_debugger_alive(port):
    """Verifica si Chrome está corriendo y responde al protocolo DevTools."""
    try:
        with urllib.request.urlopen(f"http://127.0.0.1:{port}/json/version", timeout=1) as response:
            return response.status == 200
    except:
        return False


def launch_chrome_debug(port, user_data_dir):
    """Lanza Chrome en modo debugging con el puerto dado."""
    os.makedirs(user_data_dir, exist_ok=True)
    args = [
        "powershell", "-Command",
        f'Start-Process "{CHROME_PATH}" -ArgumentList "--remote-debugging-port={port}", "--user-data-dir={user_data_dir}"'
        # f'Start-Process "{CHROME_PATH}" -ArgumentList "--remote-debugging-port={port}", "--user-data-dir={user_data_dir}" -Verb RunAs'
    ]
    subprocess.run(args, shell=True)
    print(f"[✔] Chrome lanzado en modo debugging - puerto {port}")


def ensure_chrome_debug_running(reuse_existing=True):
    """Busca un puerto libre o activo. Si reuse_existing=True, reutiliza Chrome si ya está abierto."""
    for port in range(START_PORT, MAX_PORT):
        user_data_dir = os.path.join(BASE_USER_DATA_DIR, f"profile_{port}")

        if is_chrome_debugger_alive(port):
            if reuse_existing:
                print(f"[🔁] Reutilizando Chrome ya activo en el puerto {port}")
                return port
            else:
                continue  # Saltar al siguiente para abrir uno nuevo

        elif not is_port_in_use(port):
            launch_chrome_debug(port, user_data_dir)

            for _ in range(10):
                if is_chrome_debugger_alive(port):
                    return port
                time.sleep(0.5)

            raise RuntimeError(f"Chrome no respondió en el puerto {port}")
        else:
            print(f"[i] Puerto {port} ocupado pero no responde a DevTools. Probando siguiente...")

    raise RuntimeError("No se encontró un puerto libre ni activo entre 9222 y 9300.")


def is_port_in_use(port):
    """Verifica si un puerto está en uso (sin importar si es Chrome válido o no)."""
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        return s.connect_ex(('127.0.0.1', port)) == 0


# ----------------------
# EJEMPLO DE USO:
# ----------------------

Reusar = True  
# Cambia a False si quieres forzar abrir nueva ventana en cada test
# Klokken

if __name__ == "__main__":
    # Cambia esto a False si quieres forzar abrir nueva ventana en cada test
    debug_port = ensure_chrome_debug_running(reuse_existing=Reusar)

    print(f"[🚀] Chrome listo para usar en puerto {debug_port}")


# Ejecutar el comando de PowerShell para reactivar PSReadLine
subprocess.run(['powershell', '-Command', 'Import-Module PSReadLine'])


script_dir = os.path.dirname(os.path.abspath(__file__))
json_path = os.path.join(script_dir, "NewCase.json")



# Verificar si el archivo JSON existe
if not os.path.exists(json_path):
    print(f"Error: No existe el archivo en la ruta: {json_path}")
    sys.exit(1)

# Leer el contenido del JSON
try:
    with open(json_path, 'r', encoding='utf-8-sig') as f:
        data = json.load(f)
except Exception as e:
    print(f"Error al leer el JSON: {e}")
    sys.exit(1)

# ----------------------------------------------------------------------------------------------------------
# Variables

issue = data["Issue"]
SoftwareVersion = data["Software Version"]
CompName= data["&Company Name"]
CompSID = data["S&ID"]
RootCause = data["RC"]
Solution = data["Solution"]

A_description = issue + " on " + SoftwareVersion
A_CaseTitle = "{} / SID: {} / {}".format(CompName, CompSID, A_description)
A_Product = data["Modulo"]
A_Dongle= data["&Dongle"]
A_Version = data["Version"]
A_ScannerSN = data["Scanne&r S/N"]
A_AddInfo = data["Probing Questions (Add Info)"]
A_Conclusion = "RC: " + RootCause + "\n" + "S: " + Solution
A_email = data["&Email"]
A_Categ = data["Categ"]
A_CategArea = data["CategoryArea"]
A_CasegType = data["CaseType"]
A_Phone_Tit =  data["PhT"]  
A_Phone_Note =  data["PhN"] 
A_Intl_Tit =  data["InT"] 
A_Intl_Note =  data["InN"] 
A_Remote_Tit =  "Remote Session Desktop" 
A_Remote_Note =  data["RMTSS"] 

# ----------------------------------------------------------------------------------------------------------


# Inicializar navegador en modo oculto (posición fuera de pantalla)
options = webdriver.ChromeOptions()
# options.add_argument("--window-position=-32000,-32000")  # Mover ventana fuera de pantalla
# options.add_argument("--window-size=1920,1080")
options.debugger_address = f"127.0.0.1:{debug_port}"  # Usa el puerto que te devolvió la función 


script_dir = os.path.dirname(os.path.abspath(__file__))  # ej: ...\Skrive\Templates
root_dir = os.path.abspath(os.path.join(script_dir, ".."))  # sube una carpeta
executable_path = os.path.join(root_dir, "chromedriver.exe")

service = Service(executable_path=executable_path)
driver = webdriver.Chrome(service=service, options=options)

time.sleep(1)
driver.maximize_window()


parser = argparse.ArgumentParser()
parser.add_argument('--seccion1', action='store_true')
parser.add_argument('--seccion2', action='store_true')
parser.add_argument('--seccion3', action='store_true')
parser.add_argument('--seccion4', action='store_true')
parser.add_argument('--seccion5', action='store_true')
args = parser.parse_args()


def get_tabs(port):
    with urllib.request.urlopen(f"http://127.0.0.1:{port}/json") as response:
        return json.load(response)

def focus_tab(driver, port, url_fragment):
    tabs = get_tabs(port)
    for tab in tabs:
        if url_fragment in tab.get("url", ""):
            driver.switch_to.window(tab["id"])
            print(f"[✔] Cambiado a la pestaña con URL que contiene: {url_fragment}")
            return True
    print(f"[✘] No se encontró una pestaña con URL que contenga: {url_fragment}")
    return False



actions = ActionChains(driver)

wait = WebDriverWait(driver, 25)
# Abrir página de login

if data["Case Link"] == "":
    found = focus_tab(driver, debug_port, "&pagetype=entityrecord&etn=incident")

    if not found:
        # Cambiar el foco a la nueva pestaña (la última en la lista)
        if focus_tab(driver, debug_port, "about:blank"):
            time.sleep(0.5)  # Espera para que se abra bien

            driver.get("https://3shape.crm4.dynamics.com/main.aspx?appid=366b8060-2eea-e811-a959-000d3aba0c96&pagetype=entityrecord&etn=incident")
            time.sleep(0.5)
        else:
            driver.get("https://3shape.crm4.dynamics.com/main.aspx?appid=366b8060-2eea-e811-a959-000d3aba0c96&pagetype=entityrecord&etn=incident")
            time.sleep(10)
    if not found:    
        site_map_button = wait.until(
            EC.element_to_be_clickable((By.XPATH, '//*[@aria-label="Site Map"]'))
        )
        site_map_button.click()      
    driver.maximize_window()


else:
    A_link = data["Case Link"]
    if not focus_tab(driver, debug_port, "about:blank"):
        driver.get(A_link)

    time.sleep(0.5)
 
    driver.get(A_link)
    time.sleep(0.5)
    driver.maximize_window()
    
def readJson():
    try:
        with open(json_path, 'r', encoding='utf-8-sig') as f:
            data = json.load(f)
    except Exception as e:
        print(f"Error al leer el JSON: {e}")
        sys.exit(1)








time.sleep(15)

# ----------------------------------------------------------------------------------------------------------


# *************** SECTION 1 ***************  
# Llenar el formulario de SUMMARY
def seccion1 ():

    
    # # driver.find_element(By.CSS_SELECTOR, '[aria-label="Description"]').send_keys(A_description)

    # # dropdown_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@aria-label='Special Attention']")))
    # # dropdown_button.click()
    # # time.sleep(0.5)
    # # dropdown_button.send_keys('n')
    # # time.sleep(0.5)
    # # dropdown_button.send_keys(Keys.RETURN)  # Ejemplo de enviar una tecla hacia abajo




    # # time.sleep(0.5)


    # # dropdown_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@aria-label='Priority']")))
    # # dropdown_button.click()
    # # time.sleep(0.5)
    # # dropdown_button.send_keys('S')
    # # time.sleep(0.5)
    # # dropdown_button.send_keys(Keys.RETURN)  # Ejemplo de enviar una tecla hacia abajo

    # # time.sleep(0.5)


    # # # Espera a que el botón esté presente y haga clic en él para abrir el desplegable
    # # dropdown_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@aria-label='Origin']")))
    # # dropdown_button.click()
    # # time.sleep(0.5)
    # # dropdown_button.send_keys('P')
    # # time.sleep(0.5)
    # # dropdown_button.send_keys(Keys.RETURN)  # Ejemplo de enviar una tecla hacia abajo

    # # time.sleep(0.5)

    # # editfill = driver.find_element(By.CSS_SELECTOR, '[aria-label="Dongle Number, Lookup"]')
    # # time.sleep(1)
    # # editfill.click()
    # # time.sleep(1)
    # # editfill.clear()
    # # time.sleep(1)    
    # # editfill.send_keys(A_Dongle)
    # # time.sleep(3.5)
    # # editfill.send_keys(Keys.RETURN)  # Ejemplo de enviar una tecla hacia abajo
    # # time.sleep(1)

    # # if A_Product == "Unite" or A_Product == "TRIOS Software":
    # #     editfill = driver.find_element(By.CSS_SELECTOR, '[aria-label="Product, Lookup"]')
    # #     time.sleep(1)
    # #     editfill.click()
    # #     time.sleep(1)
    # #     editfill.clear()
    # #     time.sleep(1)
    # #     editfill.send_keys(A_Product)
    # #     time.sleep(3.5)
    # #     editfill.send_keys(Keys.RETURN)  # Ejemplo de enviar una tecla hacia abajo
    # #     time.sleep(1)

    # #     editfill = driver.find_element(By.CSS_SELECTOR, '[aria-label="Software Version, Lookup"]')
    # #     time.sleep(1)
    # #     editfill.click()
    # #     time.sleep(1)
    # #     editfill.clear()
    # #     time.sleep(1)
    # #     editfill.send_keys(A_Version)
    # #     time.sleep(3.5)
    # #     editfill.send_keys(Keys.RETURN)  # Ejemplo de enviar una tecla hacia abajo
    # #     time.sleep(1)

    # # elif A_Product == "TRIOS":
    # #     editfill = driver.find_element(By.CSS_SELECTOR, '[aria-label="Scanner S/N"]')
    # #     time.sleep(1)
    # #     editfill.click()
    # #     time.sleep(1)
    # #     editfill.clear()
    # #     time.sleep(1)
    # #     editfill.send_keys(A_ScannerSN)
    # #     time.sleep(3.5)
    # #     editfill.send_keys(Keys.RETURN)  # Ejemplo de enviar una tecla hacia abajo
    # #     time.sleep(1)


    # # editfill = driver.find_element(By.CSS_SELECTOR, '[aria-label="Responsible contact, Lookup"]')
    # # time.sleep(1)
    # # editfill.click()
    # # time.sleep(1)
    # # editfill.clear()
    # # time.sleep(1)
    # # editfill.send_keys(A_email)
    # # time.sleep(3.5)
    # # editfill.send_keys(Keys.RETURN)  # Ejemplo de enviar una tecla hacia abajo
    # # time.sleep(1)


    # # if A_Categ == "Request":
    # #     editfill = driver.find_element(By.CSS_SELECTOR, '[aria-label="Category, Lookup"]')
    # #     time.sleep(1)
    # #     editfill.click()
    # #     time.sleep(1)
    # #     editfill.clear()
    # #     time.sleep(1)
    # #     editfill.send_keys(A_Categ)
    # #     time.sleep(3.5)
    # #     editfill.send_keys(Keys.RETURN)  # Ejemplo de enviar una tecla hacia abajo
    # #     time.sleep(1)

    # #     if A_CategArea=="General Product Information":
    # #         editfill = driver.find_element(By.CSS_SELECTOR, '[aria-label="Category Area, Lookup"]')
    # #         time.sleep(1)
    # #         editfill.click()
    # #         time.sleep(1)
    # #         editfill.clear()
    # #         time.sleep(1)
    # #         editfill.send_keys(A_CategArea)
    # #         time.sleep(3.5)
    # #         editfill.send_keys(Keys.RETURN)  # Ejemplo de enviar una tecla hacia abajo
    # #         time.sleep(1)

    # #         editfill = driver.find_element(By.CSS_SELECTOR, '[aria-label="Case Type, Lookup"]')
    # #         time.sleep(1)
    # #         editfill.click()
    # #         time.sleep(1)
    # #         editfill.clear()
    # #         time.sleep(1)
    # #         editfill.send_keys(A_CasegType)
    # #         time.sleep(3.5)
    # #         editfill.send_keys(Keys.RETURN)  # Ejemplo de enviar una tecla hacia abajo
    # #         time.sleep(1)
        
    # #     editfill = driver.find_element(By.CSS_SELECTOR, '[aria-label="Category Area, Lookup"]')
    # #     time.sleep(1)
    # #     editfill.click()
    # #     time.sleep(1)
    # #     editfill.clear()
    # #     time.sleep(1)
    # #     editfill.send_keys(A_CategArea)
    # #     time.sleep(3.5)
    # #     editfill.send_keys(Keys.RETURN)  # Ejemplo de enviar una tecla hacia abajo
    # #     time.sleep(1)
    
    # # if A_Categ == "Complaint":
    # #     editfill = driver.find_element(By.CSS_SELECTOR, '[aria-label="Category, Lookup"]')
    # #     time.sleep(1)
    # #     editfill.click()
    # #     time.sleep(1)
    # #     editfill.clear()
    # #     time.sleep(1)
    # #     editfill.send_keys(A_Categ)
    # #     time.sleep(3.5)
    # #     editfill.send_keys(Keys.RETURN)  # Ejemplo de enviar una tecla hacia abajo
    # #     time.sleep(1)

    # #     editfill = driver.find_element(By.CSS_SELECTOR, '[aria-label="Category Area, Lookup"]')
    # #     time.sleep(1)
    # #     editfill.click()
    # #     time.sleep(1)
    # #     editfill.clear()
    # #     time.sleep(1)
    # #     editfill.send_keys(A_CategArea)
    # #     time.sleep(3.5)
    # #     editfill.send_keys(Keys.RETURN)  # Ejemplo de enviar una tecla hacia abajo
    # #     time.sleep(1)

    # #     editfill = driver.find_element(By.CSS_SELECTOR, '[aria-label="Case Type, Lookup"]')
    # #     time.sleep(1)
    # #     editfill.click()
    # #     time.sleep(1)
    # #     editfill.clear()
    # #     time.sleep(1)
    # #     editfill.send_keys(A_CasegType)
    # #     time.sleep(3.5)
    # #     editfill.send_keys(Keys.RETURN)  # Ejemplo de enviar una tecla hacia abajo
    # #     time.sleep(1)



    # # # Usando XPATH
    # # driver.find_element(By.XPATH, '//li[@aria-label="Description & Conclusion"]').click()
    # # time.sleep(1)

    # # driver.find_element(By.CSS_SELECTOR, '[aria-label="Additional Information"]').send_keys(A_AddInfo)

    # # driver.find_element(By.CSS_SELECTOR, '[aria-label="Conclusion"]').send_keys(A_Conclusion)
    # # time.sleep(1)
    # # save_button = driver.find_element(By.XPATH, "//span[contains(text(), 'Save')]")
    # # save_button.click()

    # # time.sleep(20)


    Share_Btn = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//button[@aria-label="Share"]'))
    )
    Share_Btn.click()
    time.sleep(1)



    # Clic en el botón "Copy link"
    copy_link_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[contains(@id, 'copyLinkItem')]")))
    copy_link_button.click()
    time.sleep(2.5)


    def update_case_link(driver, json_path):
            # Esperar a que aparezca la ventana "Link Copied"

        time.sleep(1)  # pequeña espera para asegurar que el portapapeles se haya actualizado

        # Obtener link del portapapeles
        full_url = pyperclip.paste()

        # Leer el JSON
        with open(json_path, 'r', encoding='utf-8-sig') as file:
            data = json.load(file)

        # Reemplazar el valor de la clave "Case Link"
        if "Case Link" in data:
            data["Case Link"] = full_url
        else:
            print('"Case Link" no encontrado en el JSON.')
        # Guardar los cambios
        with open(json_path, 'w', encoding='utf-8-sig') as file:
            json.dump(data, file, indent=4)
        time.sleep(1)

        


    update_case_link(driver,json_path)

    close_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[contains(@id, 'dialogCloseIconButton')]")))
    close_button.click()
    time.sleep(1)  # pequeña espera para asegurar que el portapapeles se haya actualizado

    # Ruta del escritorio del usuario
    desktop_path = os.path.join(os.path.expanduser("~"), r"Desktop\CasesJSON")
    # Verificar si el archivo JSON existe
    

    # Construcción del nombre del archivo
    backup_filename = f"DNG_{A_Dongle}.json"
    backup_path = os.path.join(desktop_path, backup_filename)

    if not os.path.exists(backup_path):
        print(f"Error: No existe el archivo en la ruta: {backup_path}")
        sys.exit(1)

    update_case_link(driver, backup_path)


# *************** SECTION 2 ***************  
# Toma el case Number, la Survey y lo guarda en el JSON

def seccion2 ():
    readJson()

    if args.seccion1:
        time.sleep(20)

    # Esperar a que "Enter a note" sea clickeable
    WebDriverWait(driver, 25).until(
        EC.element_to_be_clickable((By.XPATH, '//li[@aria-label="Summary"]'))
    )

    # Esperar a que "Enter a note" sea clickeable
    WebDriverWait(driver, 25).until(
        EC.element_to_be_clickable((By.XPATH, '//span[contains(text(), "Enter a note")]'))
    )



    # # Encontrar el campo
    casewtit = driver.find_element(By.CSS_SELECTOR, '[aria-label="Case Title"]')
    A_CaseNumber = casewtit.get_attribute("value")

    pyperclip.copy(A_CaseNumber)



    def update_case_number(json_path, A_CaseNumber):
        # Extraer solo el texto entre corchetes
        match = re.search(r'\[(.*?)\]', A_CaseNumber)
        if match:
            A_CaseNumber = match.group(1)
        else:
            print("No se encontró texto entre corchetes en el valor proporcionado.")
            return

        # Leer el JSON
        with open(json_path, 'r', encoding='utf-8-sig') as file:
            data = json.load(file)

        # Reemplazar el valor de la clave "C&ase Number"
        if "C&ase Number" in data:
            data["C&ase Number"] = A_CaseNumber
        else:
            print('"C&ase Number" no encontrado en el JSON.')

        # Guardar los cambios
        with open(json_path, 'w', encoding='utf-8-sig') as file:
            json.dump(data, file, indent=4)

    update_case_number(json_path, A_CaseNumber)

    # Ruta del escritorio del usuario
    desktop_path = os.path.join(os.path.expanduser("~"), r"Desktop\CasesJSON")
    # Verificar si el archivo JSON existe
    

    # Construcción del nombre del archivo
    backup_filename = f"DNG_{A_Dongle}.json"
    backup_path = os.path.join(desktop_path, backup_filename)

    if not os.path.exists(backup_path):
        print(f"Error: No existe el archivo en la ruta: {backup_path}")
        sys.exit(1)

    update_case_number(backup_path, A_CaseNumber)

    # [CAS-1211486-Z2G2M8]

    try:
        with open(json_path, 'r', encoding='utf-8-sig') as f:
            data = json.load(f)
    except Exception as e:
        print(f"Error al leer el JSON: {e}")
        sys.exit(1)

    A_CaseNumber = data["C&ase Number"]


    time.sleep(0.5)
    # Enviar Alt + Apyautogui
    

    pyautogui.hotkey('alt', 'a')
    # Luego, enviar flecha izquierda
    time.sleep(1.5)
    # Asegurarse que el campo tenga foco
    casewtit.click()
    time.sleep(0.5)

    # Inicializa ActionChains
    time.sleep(1)

    # Ctrl+A y Backspace
    actions.key_down(Keys.CONTROL).send_keys('a').key_up(Keys.CONTROL).perform()
    time.sleep(0.5)

    actions.send_keys(Keys.BACKSPACE)


    time.sleep(0.5)

    A_CaseTitl_Plus_Casenumber = f"{A_CaseTitle} [{A_CaseNumber}] "
    time.sleep(0.2)

    casewtit.send_keys(A_CaseTitl_Plus_Casenumber)

    time.sleep(1)
    save_button = driver.find_element(By.XPATH, "//span[contains(text(), 'Save')]")
    save_button.click()

    # Esperar hasta que el indicador de carga desaparezca
    WebDriverWait(driver, 25).until(
        EC.invisibility_of_element_located((By.CLASS_NAME, 'progressDot'))
    )

    time.sleep(10)



    WebDriverWait(driver, 25).until(
        EC.element_to_be_clickable((By.XPATH, '//span[contains(text(), "Enter a note")]'))
    )

    # Luego hacer clic en el tab "Summary"
    # Esperar a que el tab esté presente y visible
    system_tab = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '//li[@aria-label="System"]'))
    )

    # Intentar clic normal
    try:
        system_tab.click()
    except:
        # Si el click normal falla, usar JavaScript
        driver.execute_script("arguments[0].click();", system_tab)


    # time.sleep(10)

    # driver.find_element(By.XPATH, '//li[@aria-label="System"]').click()
    time.sleep(10)

    # Localizar el campo de texto usando XPath

    surveyEdit = driver.find_element(By.XPATH, '//input[@aria-label="SurveyMonkey Link"]')
    time.sleep(0.5)

    A_Survey = surveyEdit.get_attribute("value")
    time.sleep(0.5)

    pyperclip.copy(A_Survey)


    time.sleep(1.5)

    def update_survey(json_path, A_Survey):

        # Leer el JSON
        with open(json_path, 'r', encoding='utf-8-sig') as file:
            data = json.load(file)

        # Reemplazar el valor de la clave "Sur&vey"
        if "Sur&vey" in data:
            data["Sur&vey"] = A_Survey
        else:
            print('"Sur&vey" no encontrado en el JSON.')

        # Guardar los cambios
        with open(json_path, 'w', encoding='utf-8-sig') as file:
            json.dump(data, file, indent=4)


    update_survey(json_path, A_Survey)
    # Ruta del escritorio del usuario
    desktop_path = os.path.join(os.path.expanduser("~"), r"Desktop\CasesJSON")
    # Verificar si el archivo JSON existe
    

    # Construcción del nombre del archivo
    backup_filename = f"DNG_{A_Dongle}.json"
    backup_path = os.path.join(desktop_path, backup_filename)

    if not os.path.exists(backup_path):
        print(f"Error: No existe el archivo en la ruta: {backup_path}")
        sys.exit(1)

    update_survey(backup_path, A_CaseNumber)

    # [CAS-1211486-Z2G2M8]
    try:
        with open(json_path, 'r', encoding='utf-8-sig') as f:
            data = json.load(f)
    except Exception as e:
        print(f"Error al leer el JSON: {e}")
        sys.exit(1)

    time.sleep(0.8)

    pyautogui.hotkey('alt', 'v')

    time.sleep(0.8)


# *************** SECTION 3 ***************  
# Agregar notas y enviar email
def seccion3 ():
    time.sleep(1.5)

    driver.find_element(By.XPATH, '//li[@aria-label="Summary"]').click()


    time.sleep(0.5)

    WebDriverWait(driver, 25).until(
        EC.element_to_be_clickable((By.XPATH, '//span[contains(text(), "Enter a note")]'))
    )


    def create_note(driver, wait, TITULO, NOTA,Notebool):
        note_area = wait.until(EC.element_to_be_clickable((By.XPATH, '//span[contains(text(), "Enter a note")]')))
        note_area.click()
        time.sleep(2)
        # Escribir el título de la nota
        # Ubicar directamente el input por aria-label
        notetitle_input = driver.find_element(By.XPATH, '//input[@aria-label="Create a Note title"]')
        # Click (opcional, si necesario)
        notetitle_input.click()
        # Enviar el texto
        notetitle_input.send_keys(TITULO)

        time.sleep(0.5)
        # Escribir el contenido de la nota
        noteBody_input = driver.find_element(By.XPATH, '//div[@aria-label="Rich Text Editor Control incident notetext"]')
        noteBody_input.send_keys(NOTA)

        

        if Notebool == True:
            attach_button = wait.until(EC.element_to_be_clickable((By.XPATH, '//button[@title="Add an attachment"]')))
            attach_button.click()
            time.sleep(1)

            # Ahora esperar el input de tipo file
            file_input = wait.until(EC.presence_of_element_located((By.XPATH, '//input[@type="file"]')))
            time.sleep(1)
            import os

            desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
            logs_file = os.path.join(desktop_path, "Cases_Final")

            # Suponiendo que A_Dongle ya está definido, por ejemplo:
            # A_Dongle = "Suba"

            # Buscar el archivo cuyo nombre contenga A_Dongle
            archivo_encontrado = None
            for archivo in os.listdir(logs_file):
                if A_Dongle in archivo and archivo.endswith(".zip"):
                    archivo_encontrado = os.path.join(logs_file, archivo)
                    break

            if archivo_encontrado:
                file_input.send_keys(archivo_encontrado)
            else:
                print(f"No se encontró un archivo .zip que contenga '{A_Dongle}' en {logs_file}")


            # TAB varias veces
            for _ in range(3):

                pyautogui.press('tab')
                time.sleep(0.5)


            pyautogui.press('enter')
            
            time.sleep(10)
            
        
        # Click en "Add note and close"
        note_add = wait.until(EC.element_to_be_clickable((By.XPATH, '//span[contains(text(), "Add note and close")]')))
        note_add.click()


    time.sleep(1)
    create_note(driver, wait, A_Phone_Tit, A_Phone_Note,False)
    time.sleep(3)
    create_note(driver, wait, A_Intl_Tit, A_Intl_Note,False)
    time.sleep(3)
    create_note(driver, wait, A_Remote_Tit, A_Remote_Note,False)
    time.sleep(3)
    create_note(driver, wait, A_Intl_Tit, "Logs and images are here",True)

    # Esperar hasta que el indicador de carga desaparezca
    WebDriverWait(driver, 25).until(
        EC.invisibility_of_element_located((By.CLASS_NAME, 'progressDot'))
    )

    time.sleep(10)


    # Hacer clic en el botón de "Agregar correo"
    add_email_Button = WebDriverWait(driver, 25).until(
        EC.element_to_be_clickable((By.XPATH, '//button[@data-id="notescontrol-action_bar_add_command"]'))
    )
    add_email_Button.click()

   
    time.sleep(1)

    pyautogui.hotkey('enter')


    # === 1. Entrar al iframe del email popup ===
    email_iframe = WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, 'iframe[id^="EmailPopupIframe_"]'))
    )
    driver.switch_to.frame(email_iframe)


    # # # # ---------------------------------------------------------------------
    # # # # # # Paso 1: Hacer clic en el botón del menú "Email"
    # # # # # email_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@data-id='form-selector' and contains(., 'Email')]")))
    # # # # # email_button.click()

    # # # # # # Paso 2: Esperar y hacer clic en la opción "Enhanced Email"
    # # # # # enhanced_email_option = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@role='menuitemradio' and @data-text='Enhanced Email']")))
    # # # # # enhanced_email_option.click()


    # # # # # discard_button = wait.until(EC.element_to_be_clickable((
    # # # # #     By.XPATH, "//button[@data-id='cancelButton' and .//div[text()='Discard changes']]"
    # # # # # )))
    # # # # # discard_button.click()

    # # # # ---------------------------------------------------------------------




    # === 2. Llenar el Subject ===
    time.sleep(0.5)
    readJson()
    time.sleep(0.5)
    subject_input = WebDriverWait(driver, 15).until(
        EC.presence_of_element_located((By.XPATH, '//input[@aria-label="Subject"]'))
    )
    A_CaseNumber = data["C&ase Number"]
    A_Subject = f"Regarding your case [{A_CaseNumber}]"
    subject_input.send_keys(A_Subject)

    time.sleep(10)

    # SHIFT + TAB
    pyautogui.hotkey('shift', 'tab')
    time.sleep(0.5)

    # TAB varias veces
    for _ in range(3):
        pyautogui.press('tab')
        time.sleep(0.5)

    time.sleep(0.5)



    # Construir el texto final
    closing_message = (
        "\n\nThank you for allowing me to assist you today. I would greatly appreciate it if you could take a moment to provide feedback on my service by completing the following survey:\n\n\n\n"
        "At the end of the survey, please leave a note about your experience during this interaction.\n\n\n\n"
        "Your patience and understanding throughout the process are greatly appreciated, and your feedback will help us continue improving our services as we strive for excellence.\n\n"
    )

    # Combinar el contenido principal + mensaje final
    full_text = data["EmailFinal"] + closing_message

    # Copiar al portapapeles
    pyperclip.copy(full_text)


    # Pequeña pausa para asegurar que el campo tiene foco
    time.sleep(0.5)

    # Pegar con pyautogui
    pyautogui.hotkey('ctrl', 'v')

    # Abrir buscador
    pyautogui.hotkey('ctrl', 'f')
    time.sleep(0.5)  # Espera a que aparezca el campo de búsqueda

    # Escribir lo que quieres buscar
    pyautogui.write('the following survey:')

    pyautogui.hotkey('enter')
    time.sleep(0.5)  # Espera a que aparezca el campo de búsqueda
    pyautogui.hotkey('esc')
    time.sleep(0.5)  # Espera a que aparezca el campo de búsqueda
    pyautogui.hotkey('right')
    time.sleep(0.5)  # Espera a que aparezca el campo de búsqueda
     # TAB varias veces
    for _ in range(3):
        pyautogui.hotkey('enter')
        time.sleep(0.5)

    A_surveyFinal = f"https://{data['Sur&vey']}"
    time.sleep(0.5)


    pyautogui.write(A_surveyFinal)

    time.sleep(0.8)  # Espera a que aparezca el campo de búsqueda
    pyautogui.hotkey('enter')

    time.sleep(0.5)  # Espera a que aparezca el campo de búsqueda

    # === 6. Clic en el botón Send Email ===
    send_button = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.XPATH, '//button[@aria-label="Send Email"]'))
    )
    send_button.click()

    time.sleep(10)

    # === 7. Salir completamente del iframe del email ===
    driver.switch_to.default_content()


# *************** SECTION 4 ***************  
# Terminar workflow y cerrar el caso
def seccion4 ():
    driver.find_element(By.XPATH, '//li[@aria-label="Summary"]').click()
    time.sleep(1)

    worflow_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//button[contains(@aria-label, "Categorization")]'))
    )
    worflow_button.click()
    time.sleep(3)

    next_stage_btn = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//button[@aria-label="Next Stage"]'))
    )
    next_stage_btn.click()
    time.sleep(3)


    worflow_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//button[contains(@aria-label, "Analysis")]'))
    )
    worflow_button.click()
    time.sleep(3)

    next_stage_btn = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//button[@aria-label="Next Stage"]'))
    )
    next_stage_btn.click()
    time.sleep(3)


    worflow_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//button[contains(@aria-label, "Conclusion")]'))
    )
    worflow_button.click()
    time.sleep(3)

    next_stage_btn = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//button[@aria-label="Finish"]'))
    )
    next_stage_btn.click()
    time.sleep(3)


# *************** SECTION 5 ***************  
# Resolver el caso
def seccion5 ():
    driver.set_window_size(1920, 1080)

    ResolveCase_Btn = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//button[@aria-label="Resolve Case"]'))
    )
    ResolveCase_Btn.click()
    time.sleep(3)

    # Espera a que el textarea esté presente y visible
    textarea = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '//textarea[@aria-label="Resolution"]'))
    )
    textarea.send_keys(A_Conclusion)


    time.sleep(3)


    # Clic en el botón "Copy link"
    SaveClose_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[contains(@title, 'Save and close this Case Resolution')]")))
    SaveClose_button.click()
    time.sleep(3)

    driver.maximize_window()
    time.sleep(0.5)

    cases_item = wait.until(EC.element_to_be_clickable((By.ID, "sitemap-entity-nav_cases")))
    cases_item.click()

readJson()
if args.seccion1:
    readJson()
    seccion1()
if args.seccion2:
    readJson()
    seccion2()
if args.seccion3:
    readJson()
    seccion3()
if args.seccion4:
    readJson()
    seccion4()
if args.seccion5:
    readJson()
    seccion5()
else:
    print("No se seleccionó ninguna opción.")
import sys
import json
import re
import os
import pyautogui
import time
import pyperclip
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

from selenium.webdriver.common.action_chains import ActionChains

from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager



import subprocess

# Ejecutar el comando de PowerShell para reactivar PSReadLine
subprocess.run(['powershell', '-Command', 'Import-Module PSReadLine'])


# Leer JSON
if len(sys.argv) > 1:
    json_path = sys.argv[1]
else:
    # json_path = "C:/Users/Joel Hurtado/Documents/GitHub/Skrive/Templates/datos.json"
    json_path = "C:/Users/Joel Hurtado/Documents/GitHub/Skrive/Templates/CallanDriscoll.json"
    # json_path = os.path.join(os.getcwd(), "Templates/datos.json")

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


# ----------------------------------------------------------------------------------------------------------




# Inicializar navegador en modo oculto (posición fuera de pantalla)
options = webdriver.ChromeOptions()
# options.add_argument("--window-position=-32000,-32000")  # Mover ventana fuera de pantalla
options.add_argument("--window-size=1920,1080")
executable_path = os.path.join(os.getcwd(), "chromedriver.exe")
service = Service(executable_path=executable_path)
driver = webdriver.Chrome(service=service, options=options)
driver.set_window_position(0, 0)
time.sleep(1)
driver.maximize_window()
# Abrir página de login
# driver.get("https://3shape.crm4.dynamics.com/main.aspx?appid=366b8060-2eea-e811-a959-000d3aba0c96&pagetype=entityrecord&etn=incident")
driver.get("https://3shape.crm4.dynamics.com/main.aspx?appid=366b8060-2eea-e811-a959-000d3aba0c96&pagetype=entityrecord&etn=incident&id=c97a9c02-8a09-4fec-969c-ff7b35d08489&lid=1745936471667")
time.sleep(20)


# ----------------------------------------------------------------------------------------------------------
actions = ActionChains(driver)

wait = WebDriverWait(driver, 10)

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
# # editfill.clear()
# # time.sleep(1)
# # editfill.send_keys(A_Dongle)
# # time.sleep(3.5)
# # editfill.send_keys(Keys.RETURN)  # Ejemplo de enviar una tecla hacia abajo
# # time.sleep(1)

# # if A_Product == "Unite" or A_Product == "TRIOS Software":
# #     editfill = driver.find_element(By.CSS_SELECTOR, '[aria-label="Product, Lookup"]')
# #     time.sleep(0.5)
# #     editfill.clear()
# #     time.sleep(1)
# #     editfill.send_keys(A_Product)
# #     time.sleep(3.5)
# #     editfill.send_keys(Keys.RETURN)  # Ejemplo de enviar una tecla hacia abajo
# #     time.sleep(1)

# #     editfill = driver.find_element(By.CSS_SELECTOR, '[aria-label="Software Version, Lookup"]')
# #     time.sleep(0.5)
# #     editfill.clear()
# #     time.sleep(1)
# #     editfill.send_keys(A_Version)
# #     time.sleep(3.5)
# #     editfill.send_keys(Keys.RETURN)  # Ejemplo de enviar una tecla hacia abajo
# #     time.sleep(1)

# # elif A_Product == "TRIOS":
# #     editfill = driver.find_element(By.CSS_SELECTOR, '[aria-label="Scanner S/N"]')
# #     time.sleep(0.2)
# #     editfill.clear()
# #     time.sleep(1)
# #     editfill.send_keys(A_ScannerSN)
# #     time.sleep(3.5)
# #     editfill.send_keys(Keys.RETURN)  # Ejemplo de enviar una tecla hacia abajo
# #     time.sleep(1)


# # editfill = driver.find_element(By.CSS_SELECTOR, '[aria-label="Responsible contact, Lookup"]')
# # editfill.send_keys(A_email)
# # time.sleep(3.5)
# # editfill.send_keys(Keys.RETURN)  # Ejemplo de enviar una tecla hacia abajo
# # time.sleep(1)


# # if A_Categ == "Request":
# #     editfill = driver.find_element(By.CSS_SELECTOR, '[aria-label="Category, Lookup"]')
# #     time.sleep(0.2)
# #     editfill.clear()
# #     time.sleep(1)
# #     editfill.send_keys(A_Categ)
# #     time.sleep(3.5)
# #     editfill.send_keys(Keys.RETURN)  # Ejemplo de enviar una tecla hacia abajo
# #     time.sleep(1)

# #     if A_CategArea=="General Product Information":
# #         editfill = driver.find_element(By.CSS_SELECTOR, '[aria-label="Category Area, Lookup"]')
# #         time.sleep(0.2)
# #         editfill.clear()
# #         time.sleep(1)
# #         editfill.send_keys(A_CategArea)
# #         time.sleep(3.5)
# #         editfill.send_keys(Keys.RETURN)  # Ejemplo de enviar una tecla hacia abajo
# #         time.sleep(1)

# #         editfill = driver.find_element(By.CSS_SELECTOR, '[aria-label="Case Type, Lookup"]')
# #         time.sleep(0.2)
# #         editfill.clear()
# #         time.sleep(1)
# #         editfill.send_keys(A_CasegType)
# #         time.sleep(3.5)
# #         editfill.send_keys(Keys.RETURN)  # Ejemplo de enviar una tecla hacia abajo
# #         time.sleep(1)
    
# #     editfill = driver.find_element(By.CSS_SELECTOR, '[aria-label="Category Area, Lookup"]')
# #     time.sleep(0.2)
# #     editfill.clear()
# #     time.sleep(1)
# #     editfill.send_keys(A_CategArea)
# #     time.sleep(3.5)
# #     editfill.send_keys(Keys.RETURN)  # Ejemplo de enviar una tecla hacia abajo
# #     time.sleep(1)


# # editfill = driver.find_element(By.CSS_SELECTOR, '[aria-label="Category, Lookup"]')
# # time.sleep(0.2)
# # editfill.clear()
# # time.sleep(1)
# # editfill.send_keys(A_Categ)
# # time.sleep(3.5)
# # editfill.send_keys(Keys.RETURN)  # Ejemplo de enviar una tecla hacia abajo
# # time.sleep(1)

# # editfill = driver.find_element(By.CSS_SELECTOR, '[aria-label="Category Area, Lookup"]')
# # time.sleep(0.2)
# # editfill.clear()
# # time.sleep(1)
# # editfill.send_keys(A_CategArea)
# # time.sleep(3.5)
# # editfill.send_keys(Keys.RETURN)  # Ejemplo de enviar una tecla hacia abajo
# # time.sleep(1)

# # editfill = driver.find_element(By.CSS_SELECTOR, '[aria-label="Case Type, Lookup"]')
# # time.sleep(0.2)
# # editfill.clear()
# # time.sleep(1)
# # editfill.send_keys(A_CasegType)
# # time.sleep(3.5)
# # editfill.send_keys(Keys.RETURN)  # Ejemplo de enviar una tecla hacia abajo
# # time.sleep(1)



# # # Usando XPATH
# # driver.find_element(By.XPATH, '//li[@aria-label="Description & Conclusion"]').click()
# # time.sleep(1)

# # driver.find_element(By.CSS_SELECTOR, '[aria-label="Additional Information"]').send_keys(A_AddInfo)

# # driver.find_element(By.CSS_SELECTOR, '[aria-label="Conclusion"]').send_keys(A_Conclusion)
# # time.sleep(1)
# # save_button = driver.find_element(By.XPATH, "//span[contains(text(), 'Save')]")
# # save_button.click()




# # # Esperar a que "Enter a note" sea clickeable
# # WebDriverWait(driver, 25).until(
# #     EC.element_to_be_clickable((By.XPATH, '//span[contains(text(), "Enter a note")]'))
# # )

# # # Luego hacer clic en el tab "Summary"
# # driver.find_element(By.XPATH, '//li[@aria-label="Summary"]').click()




# # # Encontrar el campo
casewtit = driver.find_element(By.CSS_SELECTOR, '[aria-label="Case Title"]')
A_CaseNumber = casewtit.get_attribute("value")

# # pyperclip.copy(A_CaseNumber)



# # def update_case_number(json_path, A_CaseNumber):
# #     # Extraer solo el texto entre corchetes
# #     match = re.search(r'\[(.*?)\]', A_CaseNumber)
# #     if match:
# #         A_CaseNumber = match.group(1)
# #     else:
# #         print("No se encontró texto entre corchetes en el valor proporcionado.")
# #         return

# #     # Leer el JSON
# #     with open(json_path, 'r', encoding='utf-8') as file:
# #         data = json.load(file)

# #     # Reemplazar el valor de la clave "C&ase Number"
# #     if "C&ase Number" in data:
# #         data["C&ase Number"] = A_CaseNumber
# #     else:
# #         print('"C&ase Number" no encontrado en el JSON.')

# #     # Guardar los cambios
# #     with open(json_path, 'w', encoding='utf-8') as file:
# #         json.dump(data, file, indent=4)

# # update_case_number(json_path, A_CaseNumber)
# # # [CAS-1211486-Z2G2M8]

try:
    with open(json_path, 'r', encoding='utf-8-sig') as f:
        data = json.load(f)
except Exception as e:
    print(f"Error al leer el JSON: {e}")
    sys.exit(1)

A_CaseNumber = data["C&ase Number"]


# # time.sleep(0.5)
# # # Enviar Alt + Apyautogui

# # pyautogui.hotkey('alt', 'a')
# # # Luego, enviar flecha izquierda
# # time.sleep(0.5)
# # # Asegurarse que el campo tenga foco
# # casewtit.click()
# # time.sleep(0.5)

# # # Inicializa ActionChains
# # time.sleep(1)

# # # Ctrl+A y Backspace
# # actions.key_down(Keys.CONTROL).send_keys('a').key_up(Keys.CONTROL).perform()

# # actions.send_keys(Keys.BACKSPACE)


# # time.sleep(0.5)

# # A_CaseTitl_Plus_Casenumber = f"{A_CaseTitle} {A_CaseNumber} "
# # time.sleep(0.2)

# # casewtit.send_keys(A_CaseTitl_Plus_Casenumber)

# # time.sleep(1)
# # save_button = driver.find_element(By.XPATH, "//span[contains(text(), 'Save')]")
# # save_button.click()

# # # time.sleep(100)



# # WebDriverWait(driver, 25).until(
# #     EC.element_to_be_clickable((By.XPATH, '//span[contains(text(), "Enter a note")]'))
# # )

# # # Luego hacer clic en el tab "Summary"
# # # Esperar a que el tab esté presente y visible
# # system_tab = WebDriverWait(driver, 10).until(
# #     EC.presence_of_element_located((By.XPATH, '//li[@aria-label="System"]'))
# # )

# # # Intentar clic normal
# # try:
# #     system_tab.click()
# # except:
# #     # Si el click normal falla, usar JavaScript
# #     driver.execute_script("arguments[0].click();", system_tab)


# # # time.sleep(10)

# # # driver.find_element(By.XPATH, '//li[@aria-label="System"]').click()
# # time.sleep(10)

# # # Localizar el campo de texto usando XPath

# # surveyEdit = driver.find_element(By.XPATH, '//input[@aria-label="SurveyMonkey Link"]')

# # A_Survey = surveyEdit.get_attribute("value")

# # pyperclip.copy(A_Survey)



# # def update_survey(json_path, A_Survey):

# #     # Leer el JSON
# #     with open(json_path, 'r', encoding='utf-8') as file:
# #         data = json.load(file)

# #     # Reemplazar el valor de la clave "Sur&vey"
# #     if "Sur&vey" in data:
# #         data["Sur&vey"] = A_Survey
# #     else:
# #         print('"Sur&vey" no encontrado en el JSON.')

# #     # Guardar los cambios
# #     with open(json_path, 'w', encoding='utf-8') as file:
# #         json.dump(data, file, indent=4)


# # update_survey(json_path, A_Survey)

# # time.sleep(0.5)

# # pyautogui.hotkey('alt', 'v')

time.sleep(0.5)



driver.find_element(By.XPATH, '//li[@aria-label="Summary"]').click()


time.sleep(0.5)

WebDriverWait(driver, 25).until(
    EC.element_to_be_clickable((By.XPATH, '//span[contains(text(), "Enter a note")]'))
)


def create_note(driver, wait, TITULO, NOTA,Notebool):
    note_area = wait.until(EC.element_to_be_clickable((By.XPATH, '//span[contains(text(), "Enter a note")]')))
    note_area.click()
    time.sleep(1)
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

        # Mandar el path del archivo
        file_input.send_keys(r'C:\Users\Joel Hurtado\Documents\Cases_Final\CAS_CallanDriscoll.zip')

        
        




        time.sleep(5)
        
    
    # Click en "Add note and close"
    note_add = wait.until(EC.element_to_be_clickable((By.XPATH, '//span[contains(text(), "Add note and close")]')))
    note_add.click()
# # time.sleep(1)
# # create_note(driver, wait, A_Phone_Tit, A_Phone_Note,False)
# # time.sleep(3)
# # create_note(driver, wait, A_Intl_Tit, A_Intl_Note,False)
# # time.sleep(3)


add_email_Button = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, '//button[@data-id="notescontrol-action_bar_add_command"]'))
)
add_email_Button.click()
actions.send_keys(Keys.ENTER)

emailSubject = driver.find_element(By.XPATH, '//input[@aria-label="Subject"]')
A_Subject = f"Regarding your case {A_CaseNumber} "
emailSubject.send_keys(A_Subject)
time.sleep(0.5)
emailSubject.send_keys(Keys.TAB)
time.sleep(0.5)
emailSubject.send_keys(Keys.TAB)
time.sleep(1)
A_emailbody = data["EmailFinal"]
time.sleep(0.5)
emailSubject.send_keys(A_emailbody)

# # driver.find_element(By.XPATH, '//li[@aria-label="Summary"]').click()
# # time.sleep(1)
# # create_note(driver, wait, A_Intl_Tit, "Logs and images are here",True)



worflow_button = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, '//button[contains(@aria-label, "Categorization")]'))
)
worflow_button.click()

next_stage_btn = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, '//button[@aria-label="Next Stage"]'))
)
next_stage_btn.click()

worflow_button = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, '//button[contains(@aria-label, "Analysis")]'))
)
worflow_button.click()

next_stage_btn = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, '//button[@aria-label="Next Stage"]'))
)
next_stage_btn.click()

worflow_button = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, '//button[contains(@aria-label, "Conclusion")]'))
)
worflow_button.click()

next_stage_btn = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, '//button[@aria-label="Finish"]'))
)
next_stage_btn.click()














time.sleep(10000)
time.sleep(2)
time.sleep(30)

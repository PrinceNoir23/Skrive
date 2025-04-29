import sys
import json
import os
import pyautogui
import time
import pyperclip
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager



def wait_for_element(driver, by, locator, timeout=10):
    """Espera a que un elemento esté visible y lo retorna."""
    try:
        wait = WebDriverWait(driver, timeout)
        element = wait.until(EC.visibility_of_element_located((by, locator)))
        return element
    except Exception as e:
        print(f"Error: No se encontró el elemento '{locator}' en {timeout} segundos.")
        raise e


# Leer JSON
if len(sys.argv) > 1:
    json_path = sys.argv[1]
else:
    json_path = "C:/Users/Joel Hurtado/Documents/GitHub/Skrive/Templates/datos.json"
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

# if "usuario" not in data or "contrasena" not in data:
#     print(f"Error: El JSON no contiene las claves 'usuario' y 'contrasena'. Contenido encontrado: {data}")
#     sys.exit(1)


issue = data["Issue"]
SoftwareVersion = data["Software Version"]
A_description = issue + " on " + SoftwareVersion

CompName= data["&Company Name"]
CompSID = data["S&ID"]

A_CaseTitle = "{} / SID: {} / {}".format(CompName, CompSID, A_description)

# Inicializar navegador en modo oculto (posición fuera de pantalla)
options = webdriver.ChromeOptions()
# options.add_argument("--window-position=-32000,-32000")  # Mover ventana fuera de pantalla
options.add_argument("--window-size=1920,1080")

executable_path = os.path.join(os.getcwd(), "chromedriver.exe")

service = Service(executable_path=executable_path)
driver = webdriver.Chrome(service=service, options=options)


# Abrir página de login
driver.get("https://3shape.crm4.dynamics.com/main.aspx?appid=366b8060-2eea-e811-a959-000d3aba0c96&pagetype=entityrecord&etn=incident")

# Llenar Description

# Espera y escribe en el campo Description
description_field = wait_for_element(driver, By.CSS_SELECTOR, 'textarea[aria-label="Description"]')
description_field.send_keys(A_description)

# Espera y escribe en el campo Case Title
case_title_field = wait_for_element(driver, By.CSS_SELECTOR, 'textarea[aria-label="Case Title"]')
case_title_field.send_keys(A_CaseTitle)



def write_in_field(driver, by, locator, text, timeout=10):
    field = wait_for_element(driver, by, locator, timeout)
    field.clear()  # Limpia antes de escribir (opcional)
    field.send_keys(text)
write_in_field(driver, By.CSS_SELECTOR, 'textarea[aria-label="Description"]', A_description)
write_in_field(driver, By.CSS_SELECTOR, '[aria-label="Case Title"]', A_CaseTitle)










# # Obtener los valores de los campos de texto
# usuario = driver.find_element(By.ID, "login_field").get_attribute("value")
# password = driver.find_element(By.ID, "password").get_attribute("value")

# # Copiar los valores al portapapeles
# time.sleep(0.5)
# pyperclip.copy(usuario)  # Copia el usuario
# print(f"Usuario copiado al portapapeles: {usuario}")


# # Simular Alt + A
# pyautogui.hotkey('alt', 'a')

# # Simular Alt + V

# time.sleep(0.5)

# # Si también necesitas copiar la contraseña, puedes hacerlo
# pyperclip.copy(password)  # Copia la contraseña
# print(f"Contraseña copiada al portapapeles: {password}")
# pyautogui.hotkey('alt', 'v')

# Click en botón "Sign in"
# driver.find_element(By.NAME, "commit").click()


# --- Aquí restauramos la ventana para hacerla visible ---
driver.set_window_position(0, 0)
time.sleep(1)
driver.maximize_window()

print("Navegador activado y maximizado después del login.")

# Mantener abierto unos segundos
time.sleep(30)



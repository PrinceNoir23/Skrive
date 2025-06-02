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
from datetime import datetime, timedelta
import threading
import subprocess

def resource_path(relative_path):
    """Obtiene la ruta absoluta al recurso, funciona tanto para desarrollo como para PyInstaller."""
    try:
        # Cuando se ejecuta el .exe empaquetado, PyInstaller crea esta carpeta temporal
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)



CHROME_PATH = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
BASE_USER_DATA_DIR = r"C:\Skrive_Chrome\ChromeDebugProfiles"
START_PORT = 9222
MAX_PORT = 9300
import sys
five9_Default = False


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
        # from pathlib import Path
        # import os

        # local_appdata = os.getenv("LOCALAPPDATA")
        # user_data_dir = Path(local_appdata) / "Google" / "Chrome" / "User Data"


        if is_chrome_debugger_alive(port):
            if reuse_existing:
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
# -------------------- NUEVO BLOQUE PARA RUTAS --------------------

def get_base_dir():
    """Retorna la ruta base, considerando ejecución como .exe o .py"""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

# Ruta al archivo JSON
json_path = os.path.join(get_base_dir(), "A_Info.json")

# Verificar si el archivo JSON existe
if not os.path.exists(json_path):
    print(f"❌ Error: No existe el archivo en la ruta: {json_path}")
    sys.exit(1)

# Leer el contenido del JSON
try:
    with open(json_path, 'r', encoding='utf-8-sig') as f:
        Info = json.load(f)
    print("✅ JSON cargado con éxito:", Info)
    print("Herre")

except Exception as e:
    print(f"❌ Error al leer el JSON: {e}")
    sys.exit(1)

# ----------------------------------------------------------------

# Inicializar navegador en modo oculto (posición fuera de pantalla)
options = webdriver.ChromeOptions()
# options.add_argument("--window-position=-32000,-32000")  # Mover ventana fuera de pantalla
# options.add_argument("--window-size=1920,1080")
options.debugger_address = f"127.0.0.1:{debug_port}"  # Usa el puerto que te devolvió la función 


script_dir = os.path.dirname(os.path.abspath(__file__))  # ej: ...\Skrive\Templates
root_dir = os.path.abspath(os.path.join(script_dir, ".."))  # sube una carpeta
# executable_path = os.path.join(root_dir, "chromedriver.exe")
executable_path = os.path.join(get_base_dir(), "chromedriver.exe")

service = Service(executable_path=executable_path)
driver = webdriver.Chrome(options=options)  # No pases `service`
# driver = webdriver.Chrome(service=service, options=options)

# time.sleep(1)
# driver.maximize_window()

wait = WebDriverWait(driver, 25)

parser = argparse.ArgumentParser()
parser.add_argument('--lunchtime', action='store_true')
parser.add_argument('--breaktime', action='store_true')
parser.add_argument('--klokken', action='store_true')
parser.add_argument('--klokkenout', action='store_true')
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

def inicializar():

    # Cambiar al tab que contenga la URL deseada (por ejemplo, "agent/home")
    found = focus_tab(driver, debug_port, "agent/")


    if not found:# Obtener el handle de la pestaña actual

        # Cambiar el foco a la nueva pestaña (la última en la lista)
        if focus_tab(driver, debug_port, "about:blank"):
            time.sleep(0.5)  # Espera para que se abra bien

            driver.get("https://app-scl.five9.com/clients/agent/main.html?role=Agent#agent/home")

        else:
            driver.get("https://app-scl.five9.com/clients/agent/main.html?role=Agent#agent/home")

        time.sleep(5)
    
        

    # Crea el driver (ajusta según el navegador que uses, aquí con Chrome)
    # Establece posición de la ventana
    # driver.set_window_position(75, 4)
    # Establece tamaño total de la ventana (no solo del contenido)
    # driver.set_window_size(643, 859)
    driver.maximize_window()

def ready():
    # button = WebDriverWait(driver, 10).until(
    #     EC.element_to_be_clickable((By.ID, "ReadyCodesLayout-ready-button"))
    # )
    # button.click()
    # button.send_keys(Keys.DOWN)
    # button.send_keys(Keys.ENTER)
    body = wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))

    # Enviar teclas como Ctrl + Alt + R
    body.send_keys(Keys.CONTROL, Keys.ALT, 'r')

def LoguearMorning():
    time.sleep(1)
    # Cambiar al tab que contenga la URL deseada (por ejemplo, "agent/home")
    found = focus_tab(driver, debug_port, "html?login")

    if not found:   
        # Cambiar el foco a la nueva pestaña (la última en la lista)
        if focus_tab(driver, debug_port, "about:blank"):

            time.sleep(0.5)  # Espera para que se abra bien

            # Ir a la URL deseada en la nueva pestaña
            driver.get("https://app.five9.com/index.html?login")

        else:
            driver.get("https://app.five9.com/index.html?login")
        time.sleep(5)  # Espera para cargar la página si es necesario



    # Ingresar valores en los inputs
    username_input = driver.find_element(By.ID, "Login-username-input")
    password_input = driver.find_element(By.ID, "Login-password-input")

    
    
    username_input.send_keys(Keys.CONTROL, 'a')
    username_input.send_keys(Info["User Name"])

    password_input.send_keys(Keys.CONTROL, 'a')
    password_input.send_keys(Info["PasswordFive"])
    time.sleep(1)
    # Hacer clic en el botón "Log In"
    login_button = driver.find_element(By.ID, "Login-login-button")
    login_button.click()

    
    agent_link = wait.until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, "a.home-link.agent"))
    )

    agent_link.click()


    
    station_input = wait.until(
        EC.visibility_of_element_located((By.ID, "SoftPhoneSetup-stationid-input"))
    )

    # Seleccionar todo el contenido actual y escribir uno nuevo
    station_input.send_keys(Keys.CONTROL, 'a')
    station_input.send_keys(Info["StationNumber"])  # Reemplaza por el valor que quieras ingresar

    
    # Esperar y hacer clic en el botón "Next"
    next_button = wait.until(
        EC.element_to_be_clickable((By.ID, "WizardBase-submit-button"))
    )
    next_button.click()


    dashboard_button = wait.until(
        EC.element_to_be_clickable((By.ID, "WizardBase-submit-button"))
    )

    dashboard_button.click()
    time.sleep(6)
    # inicializar()
    ready()

# def Access():
    # inicializar()
        

    
    # time.sleep(5)
    # driver.maximize_window()
    # time.sleep(1)

    #  # Abrir buscador
    # pyautogui.hotkey('ctrl', 'f')
    # time.sleep(0.5)  # Espera a que aparezca el campo de búsqueda

    # # Escribir lo que quieres buscar
    # pyautogui.write('Five9 Adapter')

    # pyautogui.hotkey('enter')
    # time.sleep(0.5)  # Espera a que aparezca el campo de búsqueda
    # pyautogui.hotkey('esc')
    # time.sleep(0.5)  # Espera a que aparezca el campo de búsqueda
    #  # TAB varias veces
    # for _ in range(3):
    #     pyautogui.hotkey('tab')
    #     time.sleep(0.5)
    
    # EmailLogin = Info["User Name"]
    # pyautogui.write(EmailLogin)
    # time.sleep(0.5)
    # pyautogui.hotkey('tab')
    # time.sleep(0.5)
    # PsswrdLogin = Info["PasswordFive"]
    # pyautogui.write(PsswrdLogin)
    # time.sleep(0.5)
    # pyautogui.hotkey('tab')
    # time.sleep(0.5)
    # pyautogui.hotkey('enter')
def LogOut():
    inicializar()


    driver.maximize_window()
    body = wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))

    # Enviar teclas como Ctrl + Alt + R
    body.send_keys(Keys.CONTROL, Keys.ALT,Keys.SHIFT, 'l')
    time.sleep(0.5)  # Espera a que aparezca el campo de búsqueda


    
    confirm_button = wait.until(
        EC.element_to_be_clickable((By.ID, "LogoutReasonDialog-confirm-button"))
    )

    confirm_button.click()

    #  # Abrir buscador
    # pyautogui.hotkey('ctrl','shift','alt', 'l')
    
    # time.sleep(0.5)  # Espera a que aparezca el campo de búsqueda

    # pyautogui.hotkey('enter')


# Suponiendo que five9_Default es una variable
if not five9_Default:
    sys.exit("El valor de Five9 Default es False. Deteniendo la ejecución.")
# Aquí seguiría el resto del código si la variable es True
print("Este código no se ejecutará si Five Default es False.")


if args.lunchtime:
    inicializar()
    # Esperar y hacer clic en el botón que despliega el menú
    button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, "ReadyCodesLayout-ready-button"))
    )
    button.click()
    time.sleep(1)


    # Esperar y hacer clic en la opción "Break"
    break_option = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, "ReadyCodesItem-327071-anchor"))
    )
    break_option.click()
    time.sleep(3598)
    inicializar()
    ready()
elif args.breaktime:
    inicializar()

    
    # Esperar y hacer clic en el botón que despliega el menú
    button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, "ReadyCodesLayout-ready-button"))
    )
    button.click()
    time.sleep(1)


    # Esperar y hacer clic en la opción "Break"
    break_option = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, "ReadyCodesItem-327070-anchor"))
    )
    break_option.click()
    time.sleep(898)
    inicializar()
    ready()

    # Esperar un momento para que aparezca el dropdown (ajusta si es necesario)


    # # Enviar 4 veces la tecla DOWN para moverse por las opciones
    # for _ in range(4):
    #     button.send_keys(Keys.DOWN)
    #     time.sleep(0.5)  # pequeña pausa entre pulsaciones

    # # Opcional: enviar ENTER si deseas seleccionar la opción
    # button.send_keys(Keys.ENTER)
elif args.klokken:
    LoguearMorning()
elif args.klokkenout:
    LogOut()


    

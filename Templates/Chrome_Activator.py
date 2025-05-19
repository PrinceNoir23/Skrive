import sys
import json
import subprocess
import socket
import re
import os
import pyautogui
import time
import urllib.request
import tkinter as tk
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
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import subprocess

try:
    subprocess.run(["taskkill", "/F", "/IM", "chrome.exe"], check=True)
    print("Chrome cerrado correctamente.")
except subprocess.CalledProcessError:
    print("No se pudo cerrar Chrome o no estaba abierto.")


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


def ensure_chrome_debug_running(reuse_existing=True):
    """Busca un puerto libre o activo. Si reuse_existing=True, reutiliza Chrome si ya está abierto."""
    for port in range(START_PORT, MAX_PORT):
        
        # user_data_dir = os.path.join(BASE_USER_DATA_DIR, f"profile_{port}")
        from pathlib import Path
        import os

        local_appdata = os.getenv("LOCALAPPDATA")
        user_data_dir = Path(local_appdata) / "Google" / "Chrome" / "User Data"


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



# Ejecutar el comando de PowerShell para reactivar PSReadLine
subprocess.run(['powershell', '-Command', 'Import-Module PSReadLine'])


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

driver.maximize_window()
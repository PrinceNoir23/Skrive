import subprocess
import sys
import importlib

# Paquetes pip -> módulos a importar
required_packages = {
    "pyautogui": "pyautogui",
    "pyperclip": "pyperclip",
    "selenium": "selenium",
    "webdriver-manager": "webdriver_manager"
}

def install_if_missing(pip_name, import_name):
    try:
        importlib.import_module(import_name)
        print(f"✅ '{import_name}' ya está instalado.")
    except ImportError:
        print(f"📦 Instalando '{pip_name}'...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", pip_name])

# Recorre todos los paquetes
for pip_pkg, import_name in required_packages.items():
    install_if_missing(pip_pkg, import_name)

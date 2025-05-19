@echo off
rem Eliminar carpetas completas
rmdir /s /q ".\Templates\Backups"
rmdir /s /q ".\Templates\build"
rmdir /s /q ".\Templates\dist"
rmdir /s /q ".\Templates\imgs"
rmdir /s /q ".\Templates\PythonBackup"
rmdir /s /q ".\git"

rem Eliminar archivos individuales
del /f /q ".\Templates\A_Info.json"
del /f /q ".\Templates\NewCase.json"
del /f /q "Chrome_Activator.py"
del /f /q "*.spec"
del /f /q "*.py"
del /f /q "*.ps1"
del /f /q "*.txt"
del /f /q "*.ico"
del /f /q ".\gitignore"
del /f /q ".\Templates\Z_CompareMaps.ahk"
del /f /q ".\Templates\Z_InstallPackages.py"
del /f /q ".\Templates\States_Codes.html"




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
del /f /q ".\Templates\*.spec"
del /f /q ".\Templates\*.py"
del /f /q ".\Templates\*.ps1"
del /f /q ".\Templates\*.txt"
del /f /q ".\Templates\*.ico"
del /f /q ".\Templates\*.ahk"
del /f /q ".\gitignore"
del /f /q ".\Templates\Z_CompareMaps.ahk"
del /f /q ".\Templates\Z_InstallPackages.py"
del /f /q ".\Templates\States_Codes.html"




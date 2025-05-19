# Ir a la carpeta del script
Set-Location -Path $PSScriptRoot

Write-Host "🧹 Eliminando archivos anteriores..."

# Eliminar carpetas build y dist si existen
if (Test-Path "build") {
    Remove-Item -Recurse -Force "build"
    Write-Host "✅ Carpeta 'build' eliminada."
}
if (Test-Path "dist") {
    Remove-Item -Recurse -Force "dist"
    Write-Host "✅ Carpeta 'dist' eliminada."
}

# Eliminar todos los archivos .spec
Get-ChildItem -Filter *.spec | ForEach-Object {
    Remove-Item -Force $_.FullName
    Write-Host "🗑️  Eliminado archivo: $($_.Name)"
}

Write-Host "`n🚀 Compilando ejecutables..."

# Chrome_Activator2.py → ChromeAuto.exe
pyinstaller --onefile --name "Chrome_Activator" --add-binary "chromedriver.exe;." Chrome_Activator2.py

# Five9_2.py → Five9Runner.exe
pyinstaller --onefile --name "Five9" --add-binary "chromedriver.exe;." Five9_2.py

# CRM2_2.py → CRMControl.exe
pyinstaller --onefile --name "CRM2" --add-binary "chromedriver.exe;." CRM2_2.py

Write-Host "`n✅ Compilación finalizada. Ejecutables listos en 'dist/'"

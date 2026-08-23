$ErrorActionPreference = "Stop"
Set-Location (Split-Path $PSScriptRoot -Parent)

Write-Host "Instalando dependencias de build..."
python -m pip install -r requirements.txt pyinstaller -q

Write-Host "Generando ejecutable (puede tardar)..."
python -m PyInstaller --noconfirm --clean build_exe.spec

$exe = Join-Path (Get-Location) "dist\GeneradorInformes.exe"
if (Test-Path $exe) {
    $size = [math]::Round((Get-Item $exe).Length / 1MB, 1)
    Write-Host "OK: $exe ($size MB)"
    Write-Host "Al ejecutarlo, crea la carpeta data\ junto al .exe (SQLite y exports)."
} else {
    throw "No se generó el ejecutable"
}

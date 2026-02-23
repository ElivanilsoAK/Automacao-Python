$ErrorActionPreference = "Stop"

# Caminho para o Python do venv
$VENV_PYTHON = ".\.venv\Scripts\python.exe"
$PYINSTALLER = ".\.venv\Scripts\pyinstaller.exe"

if (-not (Test-Path $VENV_PYTHON)) {
    Write-Host "Erro: Virtual environment não encontrado em .\.venv" -ForegroundColor Red
    exit 1
}

Write-Host "Limpando builds anteriores..."
if (Test-Path "build") { try { Remove-Item -Path "build" -Recurse -Force -ErrorAction SilentlyContinue } catch {} }
# dist folder might be locked by open files, skip aggressive deletion and let pyinstaller overwrite
if (Test-Path "dist") { try { Remove-Item -Path "dist" -Recurse -Force -ErrorAction SilentlyContinue } catch { Write-Host "Aviso: Não foi possível limpar pasta dist totalmente (arquivo em uso?), continuando..." -ForegroundColor Yellow } }

Write-Host "Iniciando Build com PyInstaller..."
# Executa PyInstaller através do módulo python para garantir paths
& $VENV_PYTHON -m PyInstaller BotEliva.spec --clean --noconfirm

if ($LASTEXITCODE -eq 0) {
    Write-Host "Build Concluído com Sucesso!" -ForegroundColor Green
    Write-Host "Executável em: dist\BotEliva\BotEliva.exe"
}
else {
    Write-Host "Falha no Build." -ForegroundColor Red
}

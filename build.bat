@echo off
REM ============================================================
REM  Build do .exe do Exclusor de Workflow iiLex
REM  Uso: dois cliques OU "build.bat" no terminal
REM ============================================================

echo.
echo === Limpando builds anteriores ===
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

echo.
echo === Verificando PyInstaller ===
py -m pip show pyinstaller >nul 2>&1
if errorlevel 1 (
    echo PyInstaller nao encontrado. Instalando...
    py -m pip install pyinstaller
)

echo.
echo === Gerando executavel ===
py -m PyInstaller iilex.spec --clean

echo.
if exist "dist\iiLex Workflow Cleaner.exe" (
    echo === Build concluido com sucesso! ===
    echo Arquivo: dist\iiLex Workflow Cleaner.exe
) else (
    echo *** ERRO: build falhou ***
)
echo.
pause

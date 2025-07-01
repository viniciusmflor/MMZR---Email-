@echo off
title MMZR Family Office - Sistema de Relatorios v1.0
color 0F
cls

echo.
echo ===============================================
echo   MMZR FAMILY OFFICE - SISTEMA DE RELATORIOS
echo ===============================================
echo   Versao: 1.0.0
echo   Plataforma: Windows (Microsoft Outlook)
echo ===============================================
echo.

echo Verificando ambiente Python...

python --version >nul 2>&1
if errorlevel 1 (
    echo ERRO: Python nao foi encontrado no sistema
    echo.
    echo Por favor:
    echo 1. Instale Python 3.8+ do site python.org
    echo 2. Marque a opcao "Add Python to PATH" durante instalacao
    echo 3. Reinicie o computador
    echo 4. Execute este arquivo novamente
    echo.
    pause
    exit /b 1
)

for /f "tokens=*" %%i in ('python --version') do set PYTHON_VERSION=%%i
echo Encontrado: %PYTHON_VERSION%

echo.
echo Verificando dependencias...

python -c "import pandas, openpyxl, tkinter" >nul 2>&1
if errorlevel 1 (
    echo AVISO: Algumas dependencias podem estar faltando
    echo Tentando instalar automaticamente...
    echo.
    pip install -r requirements.txt
    if errorlevel 1 (
        echo ERRO: Falha na instalacao das dependencias
        echo Execute manualmente: pip install -r requirements.txt
        echo.
        pause
        exit /b 1
    )
    echo Dependencias instaladas com sucesso
) else (
    echo Dependencias verificadas - OK
)

echo.
echo Verificando arquivos do sistema...

if not exist app.py (
    echo ERRO: Arquivo app.py nao encontrado
    echo Certifique-se de estar na pasta correta do projeto MMZR
    echo.
    pause
    exit /b 1
)

if not exist documentos\ (
    echo ERRO: Pasta documentos nao encontrada
    echo Certifique-se de estar na pasta correta do projeto MMZR
    echo.
    pause
    exit /b 1
)

echo Arquivos do sistema - OK

echo.
echo Inicializando interface grafica...
echo.

python app.py

echo.
echo ===============================================
echo   Sistema encerrado
echo ===============================================
echo.
pause 
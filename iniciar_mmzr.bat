@echo off
setlocal enabledelayedexpansion
title MMZR Family Office - Sistema de Relatorios v1.0

echo.
echo ===============================================
echo    MMZR Family Office - Sistema de Relatorios
echo              Versao 1.0.0
echo ===============================================
echo.

REM Verificar se estamos no diretorio correto
if not exist "documentos" (
    echo ERRO: Execute este arquivo na pasta raiz do projeto MMZR
    echo       A pasta deve conter o diretorio 'documentos'
    echo.
    pause
    exit /b 1
)

REM Verificar se Python esta instalado
echo Verificando instalacao do Python...
python --version >nul 2>&1
if errorlevel 1 (
    echo ERRO: Python nao encontrado no sistema
    echo.
    echo Para instalar o Python:
    echo 1. Acesse: https://www.python.org/downloads/
    echo 2. Baixe a versao mais recente do Python 3
    echo 3. Execute o instalador e marque "Add to PATH"
    echo 4. Reinicie o computador apos a instalacao
    echo.
    pause
    exit /b 1
)

echo Python encontrado com sucesso.

REM Verificar dependencias
echo Verificando dependencias do sistema...
python -c "import pandas, openpyxl, tkinter" >nul 2>&1
if errorlevel 1 (
    echo.
    echo AVISO: Algumas dependencias nao estao instaladas
    echo Instalando dependencias necessarias...
    echo.
    
    pip install -r requirements.txt
    if errorlevel 1 (
        echo.
        echo ERRO: Falha na instalacao das dependencias
        echo Tente executar manualmente: pip install -r requirements.txt
        echo.
        pause
        exit /b 1
    )
    
    echo Dependencias instaladas com sucesso.
)

echo Sistema pronto para execucao.
echo.
echo Iniciando interface grafica...
echo Aguarde alguns segundos para a janela aparecer...
echo.

REM Executar a interface grafica
python app.py

REM Verificar se houve erro na execucao
if errorlevel 1 (
    echo.
    echo ERRO: Falha na execucao da aplicacao
    echo.
         echo Possiveis solucoes:
     echo 1. Verificar se todas as planilhas estao na pasta documentos/dados/
     echo 2. Executar diagnostico: python diagnostico.py --diagnostico
     echo 3. Reinstalar dependencias: pip install -r requirements.txt
    echo.
    pause
) else (
    echo.
    echo Aplicacao encerrada normalmente.
)

endlocal 
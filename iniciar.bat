@echo off
chcp 65001 >nul
title Recordatorios de Cartera - Lomarosa
color 0A

cls
echo ============================================================
echo  SISTEMA DE RECORDATORIOS DE PAGO - CARTERA LOMAROSA
echo ============================================================
echo.

echo [1/3] Verificando Python...
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python no esta instalado o no esta en el PATH.
    echo.
    echo Por favor instala Python desde: https://www.python.org/downloads/
    echo.
    pause
    exit /b 1
)
python --version
echo [OK] Python encontrado.
echo.

echo [2/3] Verificando dependencias...
echo Instalando/verificando: Flask, Flask-CORS, python-dotenv, openpyxl, xlrd...
pip install Flask Flask-CORS python-dotenv openpyxl xlrd --quiet
if errorlevel 1 (
    echo [ERROR] No se pudieron instalar las dependencias.
    pause
    exit /b 1
)
echo [OK] Dependencias verificadas.
echo.

echo [3/3] Verificando configuracion...
if not exist ".env" (
    echo ============================================================
    echo [ADVERTENCIA] No se encontro el archivo .env
    echo ============================================================
    echo.
    echo Para enviar correos, necesitas configurar .env con:
    echo - EMAIL_USER=tu_correo@gmail.com
    echo - EMAIL_PASSWORD=tu_contraseña_aplicacion
    echo.
    echo Presiona cualquier tecla para continuar sin configuracion...
    echo.
    pause >nul
)
echo [OK] Configuracion lista.
echo.

echo ============================================================
echo  INICIANDO SERVIDOR...
echo ============================================================
echo.
echo  *** ACCEDE AQUI: http://localhost:5000 ***
echo.
echo El navegador se abrira automaticamente.
echo.
echo Presiona Ctrl+C para detener el servidor.
echo ============================================================
echo.

start http://localhost:5000

python app.py

echo.
echo ============================================================
echo  SERVIDOR DETENIDO
echo ============================================================
echo.
pause

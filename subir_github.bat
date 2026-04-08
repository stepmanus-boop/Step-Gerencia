@echo off
title Subir projeto para o GitHub
color 0A

cd /d "C:\Users\douglas.tabella\Downloads\Dashboard step"

echo ==========================================
echo   ATUALIZAR PROJETO NO GITHUB
echo ==========================================
echo.

git --version >nul 2>&1
if errorlevel 1 (
    echo ERRO: Git nao encontrado no sistema.
    pause
    exit /b
)

set /p msg=Digite a mensagem do commit: 
if "%msg%"=="" set msg=Atualizacao do projeto

echo.
echo Baixando atualizacoes do GitHub...
git pull origin main

echo.
echo Enviando suas alteracoes...
git add .
git commit -m "%msg%"
git push origin main

echo.
echo ==========================================
echo   PROCESSO FINALIZADO
echo ==========================================
pause
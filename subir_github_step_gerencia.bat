@echo off
setlocal EnableExtensions EnableDelayedExpansion
title Subir projeto para o GitHub - Step Gerencia
color 0A

set "PASTA=C:\Users\douglas.tabella\Downloads\Dashboard step gerencia"
set "REPO=https://github.com/stepmanus-boop/Step-Gerencia.git"

if not exist "%PASTA%" (
    echo ERRO: A pasta do projeto nao foi encontrada:
    echo %PASTA%
    pause
    exit /b 1
)

cd /d "%PASTA%"

echo ==========================================
echo   ATUALIZAR PROJETO NO GITHUB
echo ==========================================
echo Pasta: %PASTA%
echo Repositorio: %REPO%
echo.

git --version >nul 2>&1
if errorlevel 1 (
    echo ERRO: Git nao encontrado no sistema.
    pause
    exit /b 1
)

if not exist ".git" (
    echo Inicializando repositorio Git local...
    git init
    if errorlevel 1 (
        echo ERRO ao inicializar o Git.
        pause
        exit /b 1
    )
)

git branch -M main >nul 2>&1

git remote get-url origin >nul 2>&1
if not errorlevel 1 (
    echo Removendo remote origin antigo...
    git remote remove origin
)

echo Configurando remote origin...
git remote add origin "%REPO%"
if errorlevel 1 (
    echo ERRO ao configurar o remote origin.
    pause
    exit /b 1
)

set /p msg=Digite a mensagem do commit: 
if "%msg%"=="" set "msg=Atualizacao do projeto"

echo.
echo Adicionando arquivos...
git add .
if errorlevel 1 (
    echo ERRO ao adicionar arquivos.
    pause
    exit /b 1
)

git diff --cached --quiet
if errorlevel 1 (
    echo Criando commit...
    git commit -m "%msg%"
    if errorlevel 1 (
        echo ERRO ao criar o commit.
        pause
        exit /b 1
    )
) else (
    echo Nenhuma alteracao nova para commit.
)

echo.
echo Verificando se a branch main ja existe no remoto...
git ls-remote --exit-code --heads origin main >nul 2>&1
if errorlevel 1 (
    echo Branch remota main ainda nao existe. Enviando projeto...
) else (
    echo Baixando atualizacoes do GitHub antes do push...
    git config pull.rebase false >nul 2>&1
    git pull origin main --allow-unrelated-histories
)

echo.
echo Enviando para o GitHub...
git push -u origin main
if errorlevel 1 (
    echo.
    echo ERRO ao enviar para o GitHub.
    echo Verifique se voce esta autenticado no GitHub e se tem permissao no repositorio.
    pause
    exit /b 1
)

echo.
echo ==========================================
echo   PROCESSO FINALIZADO COM SUCESSO
echo ==========================================
pause

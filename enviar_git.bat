@echo off
setlocal EnableDelayedExpansion

REM === CONFIGURAÇÕES ===
set "REPO_URL=https://github.com/warlesson/INVEST-FIIs.git"
set "BRANCH=main"

REM === INÍCIO ===
echo.
echo Iniciando sincronização com GitHub...
echo.

REM Inicializa Git se necessário
if not exist ".git" (
    echo Inicializando repositório Git local...
    git init
)

REM Define a branch principal
git branch -M %BRANCH%

REM Adiciona arquivos
echo Adicionando arquivos...
git add .

REM Pede mensagem de commit
set /p commit_msg="Digite a mensagem do commit: "

REM Faz commit
git commit -m "!commit_msg!"

REM Adiciona ou atualiza remoto
git remote remove origin 2>nul
git remote add origin %REPO_URL%

REM Puxa as alterações remotas (evita conflitos)
echo Baixando atualizações do repositório remoto...
git pull origin %BRANCH% --rebase

REM Faz push
echo Enviando alterações para o GitHub...
git push -u origin %BRANCH%

echo.
echo ✅ Processo finalizado com sucesso!
pause

@echo off
setlocal
cd /d "%~dp0"

where python >nul 2>nul
if %errorlevel%==0 (
  python "%~dp0consolidador_server.py"
  goto :eof
)

where py >nul 2>nul
if %errorlevel%==0 (
  py -3 "%~dp0consolidador_server.py"
  goto :eof
)

echo Python nao encontrado no PATH.
echo Instale o Python 3 e depois execute novamente este arquivo.
pause

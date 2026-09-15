@echo off
setlocal
cd /d "%~dp0"

title TETORD 4.2.0 - Build Windows

echo ============================================
echo   TETORD 4.2.0 - BUILD PARA WINDOWS
echo ============================================
echo.

where node >nul 2>nul
if errorlevel 1 (
  echo [ERROR] Node.js no esta instalado o no esta en PATH.
  echo Instala Node.js LTS y vuelve a ejecutar este archivo.
  echo.
  pause
  exit /b 1
)

where npm >nul 2>nul
if errorlevel 1 (
  echo [ERROR] npm no esta disponible en PATH.
  echo Reinstala Node.js LTS y vuelve a intentarlo.
  echo.
  pause
  exit /b 1
)

echo [1/3] Node.js detectado:
node --version
echo.

echo [2/3] Instalando dependencias...
npm install
if errorlevel 1 (
  echo.
  echo [ERROR] npm install fallo.
  echo Revisa el mensaje anterior y mandame una captura si aparece un error.
  echo.
  pause
  exit /b 1
)

echo.
echo [3/3] Generando los ejecutables de TETORD...
npm run build:release
if errorlevel 1 (
  echo.
  echo [ERROR] El build de release fallo.
  echo Revisa el mensaje anterior y mandame una captura si aparece un error.
  echo.
  pause
  exit /b 1
)

echo.
echo ============================================
echo   BUILD TERMINADO
echo ============================================
echo.
echo Busca los ejecutables dentro de:
echo   %CD%\dist\
echo.
echo Archivos esperados:
echo   TETORD-Setup-4.2.0-x64.exe
echo   TETORD-4.2.0-x64.exe
echo.
echo Si ambos aparecen, el build se genero correctamente.
echo.
pause
endlocal

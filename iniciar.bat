@echo off
echo.
echo ============================================
echo   SISTEMA DE COMPRAS - Iniciando servidor
echo ============================================
echo.

cd /d "%~dp0"

if not exist "venv312\Scripts\activate.bat" (
    echo Creando entorno virtual...
    python -m venv venv312
)

call venv312\Scripts\activate.bat

echo Instalando dependencias...
pip install -r requirements.txt --quiet

echo.
echo Servidor corriendo en: http://localhost:8000
echo Presiona Ctrl+C para detener.
echo.

start "" http://localhost:8000
python -m uvicorn main:app --host 0.0.0.0 --port 8000 --reload

pause


                                            

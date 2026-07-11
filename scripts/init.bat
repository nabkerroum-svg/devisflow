@echo off
REM init.bat — initialise un environnement de developpement local DevisFlow (Windows)

cd /d "%~dp0\.."
set PROJECT_ROOT=%CD%

echo === DevisFlow Initialisation ===
echo Repertoire projet : %PROJECT_ROOT%
echo.

REM Verification Python
where python >nul 2>&1
if errorlevel 1 (
  echo Python introuvable. Installez Python 3.10+ depuis https://python.org
  exit /b 1
)
echo Python detecte.

REM Verification LibreOffice
where soffice >nul 2>&1
if errorlevel 1 (
  if not exist "C:\Program Files\LibreOffice\program\soffice.exe" (
    echo Attention : LibreOffice introuvable.
    echo La conversion DOCX vers PDF ne fonctionnera pas tant qu'il n'est pas installe.
    echo Telecharger : https://www.libreoffice.org/download/
    echo.
  ) else (
    echo LibreOffice detecte dans Program Files.
    set SOFFICE_BIN=C:\Program Files\LibreOffice\program\soffice.exe
  )
) else (
  echo LibreOffice detecte dans PATH.
)

REM Setup venv
cd backend
if not exist venv (
  echo.
  echo Creation du virtualenv...
  python -m venv venv
)

call venv\Scripts\activate.bat
echo Virtualenv active.

echo.
echo Installation des dependances Python...
pip install -q --upgrade pip
pip install -q -r requirements.txt
echo Dependances installees.

REM Creation des dossiers de stockage
if not exist storage\templates mkdir storage\templates
if not exist storage\generated mkdir storage\generated
if not exist storage\db mkdir storage\db
echo Dossiers de stockage prets.

echo.
echo === Initialisation terminee ===
echo.
echo Pour demarrer le serveur :
echo   cd backend
echo   venv\Scripts\activate
echo   uvicorn main:app --reload --port 8000
echo.
echo Puis ouvrir http://localhost:8000

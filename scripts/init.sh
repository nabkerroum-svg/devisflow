#!/usr/bin/env bash
#
# init.sh — initialise un environnement de développement local DevisFlow
#
# Crée un venv Python, installe les dépendances, vérifie LibreOffice, lance le serveur.

set -e

cd "$(dirname "$0")/.."
PROJECT_ROOT="$(pwd)"

echo "=== DevisFlow — Initialisation ==="
echo "Répertoire projet : $PROJECT_ROOT"
echo ""

# === Vérifications ===
if ! command -v python3 &> /dev/null; then
  echo "❌ Python 3 introuvable. Installez Python 3.10+ depuis https://python.org"
  exit 1
fi

PYV=$(python3 -c 'import sys; print(f"{sys.version_info.major}.{sys.version_info.minor}")')
echo "✓ Python détecté : $PYV"

if ! command -v soffice &> /dev/null; then
  echo "⚠️  LibreOffice (soffice) introuvable dans le PATH."
  echo "   La conversion DOCX→PDF ne fonctionnera pas tant qu'il n'est pas installé."
  echo ""
  echo "   Installation :"
  echo "     Linux  : sudo apt install libreoffice-writer"
  echo "     macOS  : brew install --cask libreoffice"
  echo "     Windows: https://www.libreoffice.org/download/"
  echo ""
else
  echo "✓ LibreOffice détecté : $(soffice --version | head -1)"
fi

# === Setup venv ===
cd backend

if [ ! -d "venv" ]; then
  echo ""
  echo "→ Création du virtualenv (backend/venv)..."
  python3 -m venv venv
fi

source venv/bin/activate
echo "✓ Virtualenv activé"

echo ""
echo "→ Installation des dépendances Python..."
pip install -q --upgrade pip
pip install -q -r requirements.txt
echo "✓ Dépendances installées"

# === Création des dossiers de stockage ===
mkdir -p storage/templates storage/generated storage/db
echo "✓ Dossiers de stockage prêts"

echo ""
echo "=== Initialisation terminée ✅ ==="
echo ""
echo "Pour démarrer le serveur :"
echo "  cd backend"
echo "  source venv/bin/activate"
echo "  uvicorn main:app --reload --port 8000"
echo ""
echo "Puis ouvrir http://localhost:8000"

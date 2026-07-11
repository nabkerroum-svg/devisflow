"""
Configuration centrale du projet DevisFlow.

Les chemins sont en dur pour simplifier le déploiement.
Pour la production, vous pouvez surcharger via des variables d'environnement
(voir os.environ.get plus bas).
"""
import os
from pathlib import Path

# === Chemins du projet ===
BACKEND_DIR = Path(__file__).resolve().parent
STORAGE_DIR = BACKEND_DIR / "storage"
TEMPLATES_DIR = STORAGE_DIR / "templates"      # templates Word maîtres
GENERATED_DIR = STORAGE_DIR / "generated"      # devis générés (.docx + .pdf)
DB_DIR = STORAGE_DIR / "db"
DB_PATH = DB_DIR / "devisflow.db"
SEED_DIR = BACKEND_DIR / "seed_data"
FRONTEND_DIR = BACKEND_DIR.parent / "frontend"

# Créer les dossiers s'ils n'existent pas
for d in (TEMPLATES_DIR, GENERATED_DIR, DB_DIR):
    d.mkdir(parents=True, exist_ok=True)

# === Conversion PDF ===
# Le projet utilise LibreOffice headless pour convertir .docx → .pdf
# Détection automatique du binaire selon l'OS
SOFFICE_BIN = os.environ.get("SOFFICE_BIN") or (
    "/usr/bin/soffice" if Path("/usr/bin/soffice").exists()
    else "soffice"  # fallback : suppose qu'il est dans le PATH
)

# === Base de données ===
SQLITE_URL = f"sqlite:///{DB_PATH}"

# === API ===
API_PREFIX = "/api"
CORS_ORIGINS = ["*"]  # En production, mettre l'origine du frontend ici

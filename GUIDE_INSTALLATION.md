# DevisFlow — Guide d'installation

Application de génération de devis Word/PDF pour Marie Eugénie.

## Option 1 — Docker (recommandé, le plus simple)

Prérequis : Docker et Docker Compose installés.

```bash
unzip devisflow.zip
cd devisflow
docker-compose up --build
```

Puis ouvrir un navigateur sur **http://localhost:8000**

Au premier démarrage, la base de données est créée et pré-remplie
automatiquement (modèles Word, prestations, équipements avec photos, clients,
équipe, paramètres). Les données sont conservées entre les redémarrages.

Pour arrêter : `Ctrl+C`, puis `docker-compose down`.

## Option 2 — Installation manuelle (développement)

Prérequis : Python 3.12+, LibreOffice (pour la conversion PDF).

```bash
unzip devisflow.zip
cd devisflow/backend

# Dépendances Python
pip install -r requirements.txt

# LibreOffice (Debian/Ubuntu)
sudo apt-get install -y libreoffice

# Lancer le serveur
uvicorn main:app --host 0.0.0.0 --port 8000
```

Puis ouvrir **http://localhost:8000**

## Vérification

- La page principale affiche l'interface DevisFlow (6 vues dans la barre latérale).
- Un badge en bas à droite indique « Connecté au backend ».
- La vue Bibliothèque liste les prestations et équipements réels.

## Dépannage

- **PDF non généré** : vérifier que LibreOffice est installé (`soffice --version`).
- **Port déjà utilisé** : changer le port (`--port 8001`) ou libérer le 8000.
- **Base à réinitialiser** : supprimer `backend/storage/db/devisflow.db` puis
  relancer (le seed se relance automatiquement).

## Arborescence

```
devisflow/
├── backend/            API FastAPI + moteur de génération
│   ├── main.py         point d'entrée
│   ├── models.py       tables (Client, Equipement, PrestationType, …)
│   ├── routes_*.py     endpoints API
│   ├── pricing.py      moteur de calcul des prix
│   ├── template_service.py  génération Word/PDF + injection photos
│   ├── seed_data/      modèles Word maîtres + photos de démonstration
│   └── storage/        base SQLite, documents générés, photos uploadées
├── frontend/
│   ├── index_rich.html interface principale
│   └── static/         pont API + styles
├── docker-compose.yml
└── docs/               documentation
```

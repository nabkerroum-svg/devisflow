# Guide d'installation

Ce document explique étape par étape comment installer DevisFlow sur un nouveau poste de travail, en local ou en serveur.

## Sommaire

1. [Pré-requis système](#pré-requis-système)
2. [Installation rapide via Docker](#installation-rapide-via-docker-recommandé)
3. [Installation en local sans Docker](#installation-en-local-sans-docker)
4. [Configuration de la base de données](#configuration-de-la-base-de-données)
5. [Premier lancement et vérification](#premier-lancement-et-vérification)
6. [Mises à jour](#mises-à-jour)
7. [Sauvegarde et restauration](#sauvegarde-et-restauration)
8. [Résolution des problèmes courants](#résolution-des-problèmes-courants)

---

## Pré-requis système

### Pour l'installation Docker (recommandée)
- **Docker** ≥ 20.10 — https://docs.docker.com/get-docker/
- **Docker Compose** ≥ 2.0 (inclus dans Docker Desktop)
- 4 Go de RAM, 2 Go d'espace disque

### Pour l'installation locale (sans Docker)
- **Python** ≥ 3.10 — https://www.python.org/downloads/
- **LibreOffice** (pour la conversion DOCX → PDF) — https://www.libreoffice.org/download/
- **Git** (optionnel, pour cloner) — https://git-scm.com/downloads

Compatibilité OS : Linux (Ubuntu/Debian/Fedora), macOS, Windows 10/11.

---

## Installation rapide via Docker (recommandée)

### Étape 1. Récupérer le projet

Décompresser l'archive ZIP du projet ou cloner le dépôt :

```bash
# Si vous avez le ZIP
unzip devisflow.zip
cd devisflow

# Si vous avez un dépôt Git
git clone <url-du-depot>
cd devisflow
```

### Étape 2. Construire et lancer

```bash
docker-compose up --build
```

Au premier lancement, Docker télécharge l'image Python, installe LibreOffice et toutes les dépendances. Compter **5 à 10 minutes**.

Lors des lancements suivants : `docker-compose up` suffit (≈ 10 secondes).

### Étape 3. Ouvrir l'application

L'application est disponible sur **http://localhost:8000**

La documentation API auto-générée est sur **http://localhost:8000/docs**

### Étape 4. Arrêter l'application

`Ctrl+C` dans le terminal, ou depuis un autre terminal :
```bash
docker-compose down
```

---

## Installation en local sans Docker

### Étape 1. Installer LibreOffice

Indispensable pour la conversion DOCX → PDF.

**Sur Ubuntu / Debian** :
```bash
sudo apt update
sudo apt install libreoffice-writer libreoffice-core
```

**Sur macOS** :
```bash
brew install --cask libreoffice
```

**Sur Windows** :
1. Télécharger depuis https://www.libreoffice.org/download/
2. Installer
3. Vérifier que `soffice.exe` est accessible (généralement dans `C:\Program Files\LibreOffice\program\`)
4. Ajouter ce chemin dans le PATH système, OU définir la variable d'environnement `SOFFICE_BIN` :
   ```cmd
   set SOFFICE_BIN=C:\Program Files\LibreOffice\program\soffice.exe
   ```

### Étape 2. Installer les dépendances Python

```bash
cd devisflow/backend
python -m venv venv

# Activer l'environnement virtuel
# Sur Linux/macOS :
source venv/bin/activate
# Sur Windows :
venv\Scripts\activate

pip install -r requirements.txt
```

### Étape 3. Lancer le serveur

```bash
# Toujours depuis devisflow/backend, avec venv activé :
uvicorn main:app --reload --port 8000
```

Ouvrir http://localhost:8000

---

## Configuration de la base de données

DevisFlow utilise **SQLite** par défaut — aucune configuration n'est requise. La base est créée automatiquement au premier lancement dans :

```
backend/storage/db/devisflow.db
```

### Migration vers PostgreSQL (pour la production)

Si vous voulez passer à PostgreSQL pour une utilisation multi-utilisateurs en production :

1. Installer PostgreSQL ≥ 13
2. Créer une base et un utilisateur :
   ```sql
   CREATE DATABASE devisflow;
   CREATE USER devisflow_user WITH PASSWORD 'votre_mot_de_passe';
   GRANT ALL PRIVILEGES ON DATABASE devisflow TO devisflow_user;
   ```
3. Installer le driver PostgreSQL :
   ```bash
   pip install psycopg2-binary
   ```
4. Modifier `backend/config.py` :
   ```python
   # Remplacer SQLITE_URL par :
   SQLITE_URL = "postgresql://devisflow_user:votre_mot_de_passe@localhost:5432/devisflow"
   ```
   Ou via variable d'environnement :
   ```bash
   export DATABASE_URL="postgresql://..."
   ```
5. Adapter `backend/config.py` pour lire la variable :
   ```python
   SQLITE_URL = os.environ.get("DATABASE_URL", f"sqlite:///{DB_PATH}")
   ```
6. Redémarrer — les tables sont créées automatiquement par SQLModel au lancement.

Pour migrer les données existantes de SQLite vers PostgreSQL, utiliser un outil comme `pgloader` ou exporter via SQL puis réimporter.

---

## Premier lancement et vérification

Au tout premier démarrage, DevisFlow charge automatiquement le **template Marie Eugénie** (`copro_petite.docx`) qui se trouve dans `backend/seed_data/`. Vous pouvez vérifier que tout fonctionne en générant un devis de test.

### Vérification rapide via l'interface

1. Ouvrir http://localhost:8000
2. Aller dans l'onglet **« Modèles PDF »**
3. Vous devez voir le template **« Copropriété — petite surface (Marie Eugénie) »** avec 8 variables détectées (DATE_EMISSION, DEST_LIGNE1, DEST_LIGNE2, DEST_LIGNE3, NUMERO_DEVIS, SITE_ADRESSE, SITE_CP_VILLE, TYPE_PRESTATION)
4. Aller dans l'onglet **« Créer un devis »**, remplir les champs, cliquer **« Générer le devis »**
5. Vous devez recevoir deux fichiers : `.docx` (3,8 Mo) et `.pdf` (≈ 900 Ko)

### Vérification via l'API

```bash
# Statut
curl http://localhost:8000/api/health
# Attendu : {"status":"ok"}

# Liste des templates
curl http://localhost:8000/api/templates
# Attendu : un template "copro_petite" avec is_default=true
```

Si tout fonctionne, l'installation est terminée.

---

## Mises à jour

### Quand vous récupérez une nouvelle version du code

```bash
# Récupérer le code (selon votre méthode)
git pull
# ou : décompresser la nouvelle archive par-dessus

# Avec Docker
docker-compose down
docker-compose up --build

# Sans Docker
cd backend
source venv/bin/activate    # ou venv\Scripts\activate sur Windows
pip install -r requirements.txt --upgrade
# Redémarrer le serveur
uvicorn main:app --reload --port 8000
```

Les données dans `backend/storage/` sont préservées : templates uploadés, devis générés, base SQLite.

### Quand vous voulez ajouter un nouveau template Word

**Solution graphique** (recommandé) :
1. Préparer le fichier `.docx` dans Word en remplaçant les zones variables par des marqueurs `{{ VARIABLE_EN_MAJUSCULES }}`
2. Onglet **« Modèles PDF »** → formulaire d'import → cliquer « Importer le modèle »

**Solution en ligne de commande** (pour automatisation) :
```bash
curl -X POST http://localhost:8000/api/templates/upload \
  -F "fichier=@mon_template.docx" \
  -F "code=mon_template" \
  -F "nom=Mon nouveau template" \
  -F "famille=contrat"
```

Voir le **Guide d'annotation des templates** (`docs/GUIDE_ANNOTATION_TEMPLATES.md`) pour savoir comment préparer un `.docx`.

---

## Sauvegarde et restauration

Toutes les données persistantes du projet sont dans **`backend/storage/`** :

| Dossier | Contenu |
|---|---|
| `storage/templates/` | Templates Word maîtres uploadés |
| `storage/generated/` | Tous les devis produits (.docx et .pdf) |
| `storage/db/devisflow.db` | Base SQLite (métadonnées, références) |

### Sauvegarde manuelle

```bash
# Avec Docker, l'app peut tourner pendant la sauvegarde
cd devisflow
tar czf devisflow-backup-$(date +%Y%m%d).tar.gz backend/storage/
```

### Sauvegarde automatique (cron Linux)

Ajouter dans `crontab -e` :
```cron
# Tous les jours à 2h du matin
0 2 * * * cd /chemin/vers/devisflow && tar czf /chemin/sauvegardes/devisflow-$(date +\%Y\%m\%d).tar.gz backend/storage/
```

### Restauration

```bash
# Arrêter l'application
docker-compose down       # avec Docker
# ou Ctrl+C pour le serveur uvicorn

# Restaurer
cd devisflow
rm -rf backend/storage/   # ATTENTION : supprime l'existant
tar xzf devisflow-backup-20260619.tar.gz

# Redémarrer
docker-compose up         # ou uvicorn main:app
```

---

## Résolution des problèmes courants

### « LibreOffice introuvable » à la conversion PDF

**Linux/macOS** : vérifier l'installation
```bash
which soffice
soffice --version
```

**Windows** : définir explicitement la variable d'environnement
```cmd
set SOFFICE_BIN=C:\Program Files\LibreOffice\program\soffice.exe
```

### Port 8000 déjà utilisé

Changer le port :
```bash
# Avec Docker, modifier docker-compose.yml :
# ports: ["8001:8000"]   au lieu de "8000:8000"

# Sans Docker
uvicorn main:app --port 8001
```

### Erreur « Duplicate name » à la génération

Le service `template_service.py` contient déjà la fonction `_nettoyer_doublons_zip()` qui corrige ce souci automatiquement. Si l'erreur persiste, vérifier que le template uploadé n'est pas corrompu (l'ouvrir dans Word et le ré-enregistrer).

### Le rendu PDF est différent du Word

Cela peut arriver si LibreOffice n'a pas certaines polices installées. Solution :
```bash
# Linux
sudo apt install fonts-liberation fonts-dejavu fonts-arial-fonts msttcorefonts-installer
fc-cache -f -v
```

### Les variables {{ XXX }} apparaissent telles quelles dans le devis généré

Cela signifie que `docxtpl` n'a pas reconnu les marqueurs. Causes possibles :
1. **Le marqueur est fragmenté en plusieurs runs Word** (Word a coupé `{{ NUMERO_DEVIS }}` en `{{`, ` NUMERO_DEVIS `, ` }}`). Solution : ouvrir le .docx dans Word, supprimer entièrement le marqueur et le retaper d'une seule traite sans correction automatique.
2. **Mauvaise syntaxe** : utiliser exactement `{{ VARIABLE }}` avec espaces autour, en majuscules sans accents.
3. **Le marqueur est dans une zone de texte (text frame)** : python-docx ne voit pas toujours ces zones. Solution : retaper la zone variable directement dans un paragraphe normal.

### Conteneur Docker qui crash au démarrage

```bash
docker-compose logs devisflow
```

Si le souci vient de LibreOffice qui ne démarre pas, recréer le conteneur :
```bash
docker-compose down -v
docker-compose up --build
```

# Guide de déploiement

Ce document présente plusieurs solutions d'hébergement pour DevisFlow, classées par adéquation au projet.

## Vue d'ensemble : compatibilité par solution

| Solution | Adapté ? | Raison |
|---|---|---|
| **VPS** (OVH, Scaleway, Hetzner) | ✅ Recommandé | Contrôle total, LibreOffice possible, données stockées localement |
| **Docker sur VPS** | ✅ Recommandé | Setup minimal, portable |
| **Railway** | ⚠️ Partiel | Pas de LibreOffice natif → PDF impossible sans buildpack custom |
| **Render** | ✅ Possible | Supporte Docker, mais limite ressources gratuites |
| **Hébergement mutualisé** | ❌ Non | Python pas toujours dispo, LibreOffice impossible |
| **Vercel** | ❌ Non | Serverless = stateless ; templates Word + DB persistante incompatibles |
| **Netlify** | ❌ Non | Hébergement statique uniquement |

DevisFlow génère des fichiers Word/PDF en utilisant LibreOffice ; il a besoin d'un environnement qui :
1. Exécute du **Python 3.10+** en process long
2. Permet d'installer **LibreOffice**
3. Offre un **stockage disque persistant** pour les templates et devis

Cela exclut les plateformes serverless (Vercel, Netlify, Lambda) et la majorité des hébergements mutualisés.

---

## Option 1 — VPS avec Docker (recommandée)

### Pré-requis
- VPS avec ≥ 2 Go RAM, ≥ 20 Go SSD
- OS : Ubuntu 22.04 LTS ou Debian 12
- Nom de domaine (optionnel mais recommandé)

### Recommandations de fournisseurs (Europe, RGPD)

| Fournisseur | Offre minimale | Prix indicatif | Localisation |
|---|---|---|---|
| **Scaleway** | DEV1-S | ~10 €/mois | France |
| **OVHcloud** | VPS Starter | ~6 €/mois | France |
| **Hetzner** | CX11 | ~5 €/mois | Allemagne |
| **Infomaniak** | Public Cloud S | ~15 €/mois | Suisse |

### Étape 1. Préparer le serveur

```bash
# Se connecter au VPS
ssh root@votre-ip

# Mise à jour
apt update && apt upgrade -y

# Installer Docker et Docker Compose
curl -fsSL https://get.docker.com | sh
apt install -y docker-compose-plugin

# Créer un utilisateur dédié (sécurité)
adduser devisflow
usermod -aG docker devisflow
su - devisflow
```

### Étape 2. Déployer l'application

```bash
# Récupérer le code
git clone <url-depot> devisflow
# ou : scp -r ./devisflow devisflow@votre-ip:~

cd devisflow

# Démarrer en arrière-plan
docker compose up -d --build
```

L'app tourne sur le port 8000 du VPS. Vérification :
```bash
curl http://localhost:8000/api/health
```

### Étape 3. Configurer Nginx (reverse proxy + HTTPS)

```bash
sudo apt install -y nginx certbot python3-certbot-nginx
```

Créer `/etc/nginx/sites-available/devisflow` :
```nginx
server {
    listen 80;
    server_name devisflow.votre-domaine.com;

    client_max_body_size 50M;  # autoriser les uploads de templates jusqu'à 50 Mo

    location / {
        proxy_pass http://127.0.0.1:8000;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
        proxy_read_timeout 120s;  # conversion PDF peut prendre du temps
    }
}
```

Activer + HTTPS :
```bash
sudo ln -s /etc/nginx/sites-available/devisflow /etc/nginx/sites-enabled/
sudo nginx -t
sudo systemctl reload nginx
sudo certbot --nginx -d devisflow.votre-domaine.com
```

### Étape 4. Sauvegardes automatiques

Créer `/home/devisflow/backup.sh` :
```bash
#!/bin/bash
DATE=$(date +%Y%m%d-%H%M%S)
BACKUP_DIR=/home/devisflow/backups
mkdir -p $BACKUP_DIR
cd /home/devisflow/devisflow
tar czf $BACKUP_DIR/devisflow-$DATE.tar.gz backend/storage/
# Garder uniquement les 14 derniers backups
ls -t $BACKUP_DIR/devisflow-*.tar.gz | tail -n +15 | xargs -r rm
```

```bash
chmod +x /home/devisflow/backup.sh
crontab -e
# Ajouter : 0 2 * * * /home/devisflow/backup.sh
```

### Étape 5. Monitoring minimal

```bash
# Voir les logs
docker compose logs -f --tail 100

# Voir l'utilisation ressources
docker stats devisflow
```

Pour un monitoring plus poussé, installer **Uptime Kuma** (Docker) ou **Healthchecks.io**.

---

## Option 2 — Render

[Render](https://render.com) supporte Docker et offre un plan gratuit limité.

### Configuration

1. **Pousser le projet sur GitHub** (privé OK)
2. Sur Render, créer un **Web Service**
3. Connecter le repo GitHub
4. Configuration :
   - Environment : **Docker**
   - Dockerfile path : `Dockerfile`
   - Plan : **Starter** (7 $/mois, requis pour LibreOffice qui consomme de la RAM)
5. Ajouter un **Persistent Disk** (Settings → Disks) :
   - Mount Path : `/app/backend/storage`
   - Size : 5 Go minimum
6. Déployer

### Limites Render
- Plan gratuit insuffisant (mémoire + pas de disque persistant)
- Délai de cold start si le service dort
- Mise à jour automatique au `git push`

---

## Option 3 — Railway

[Railway](https://railway.app) déploie automatiquement les projets depuis GitHub.

### Limite importante
**Railway ne fournit pas LibreOffice nativement.** Vous devez créer un Dockerfile (déjà fait dans le projet) et utiliser le build Docker, pas le buildpack par défaut.

### Configuration

1. Pousser le projet sur GitHub
2. Sur Railway, créer un nouveau projet → **Deploy from GitHub repo**
3. Sélectionner le repo
4. Railway détecte le `Dockerfile` automatiquement
5. Ajouter un **Volume** pour `storage/` :
   - Mount Path : `/app/backend/storage`
6. Variables d'environnement (optionnel) :
   - `PORT=8000`
7. Déployer

### Coût
- Plan Hobby : ~5 $/mois en usage continu
- Plan Pro : selon usage

---

## Option 4 — Docker sur n'importe quel serveur cloud

Le projet est entièrement containerisé. Toute infra qui exécute Docker convient :

- **AWS EC2** (t3.small ≈ 15 $/mois)
- **Google Cloud Run** (avec stockage GCS pour persistance)
- **DigitalOcean Droplets** (~6 $/mois)
- **Azure Container Instances**

Procédure type :
1. Build l'image : `docker build -t devisflow .`
2. Pousser sur un registry (Docker Hub, GHCR, ECR, GCR)
3. Lancer sur la plateforme cible avec volume monté pour `/app/backend/storage`
4. Configurer un reverse proxy ou Load Balancer pour HTTPS

---

## Solutions à éviter (avec raison)

### Vercel
Plateforme **serverless** : chaque requête peut être servie par un nouveau conteneur sans mémoire des précédentes. Incompatible avec :
- Base SQLite locale (pas persistante entre invocations)
- Stockage de templates Word sur disque
- LibreOffice (binaire trop volumineux pour un serverless)

Vercel est excellent pour des sites statiques ou API stateless, mais pas pour DevisFlow.

### Netlify
Hébergement **statique uniquement**. Aucun runtime Python. Inutile pour ce projet.

### Hébergements mutualisés (OVH Perso, Hostinger Premium, etc.)
- Python pas toujours installable
- LibreOffice **jamais** disponible
- Pas d'accès SSH suffisant pour configurer
- Stockage limité

Si vous n'avez qu'un mutualisé, deux solutions :
1. Souscrire en plus un VPS basique (5-10 €/mois)
2. Utiliser un service tiers de conversion DOCX→PDF (CloudConvert API, ~10 €/mois) et adapter `template_service.py` pour appeler cette API au lieu de LibreOffice

---

## Sécurité en production

### Authentification (à ajouter)
La V1 livrée n'a **pas d'authentification**. Pour une mise en production, ajouter :

```bash
pip install fastapi-users[sqlalchemy] python-jose[cryptography]
```

Et protéger les routes par token JWT. Documentation : https://fastapi-users.github.io/fastapi-users/

### Variables d'environnement à définir

En production, exporter (ou mettre dans un `.env`) :

```bash
# Sécurité
SECRET_KEY=<générer avec : openssl rand -hex 32>
ALLOWED_ORIGINS=https://devisflow.votre-domaine.com

# Base de données (si PostgreSQL)
DATABASE_URL=postgresql://user:pass@host:5432/db

# Conversion PDF
SOFFICE_BIN=/usr/bin/soffice
```

Adapter `backend/config.py` pour lire ces variables :
```python
import os
CORS_ORIGINS = os.environ.get("ALLOWED_ORIGINS", "*").split(",")
```

### Headers de sécurité (Nginx)

Ajouter dans la config Nginx :
```nginx
add_header X-Frame-Options "SAMEORIGIN" always;
add_header X-Content-Type-Options "nosniff" always;
add_header Referrer-Policy "no-referrer-when-downgrade" always;
add_header Strict-Transport-Security "max-age=63072000" always;
```

### Pare-feu

```bash
# UFW (Ubuntu)
sudo ufw allow OpenSSH
sudo ufw allow 'Nginx Full'
sudo ufw enable
```

---

## Estimation de coûts mensuels (production)

| Setup | Coût | Capacité |
|---|---|---|
| VPS basique (OVH/Hetzner) + nom de domaine | ~8 €/mois | 1-10 utilisateurs concurrents, ~5 Go stockage |
| VPS moyen (Scaleway DEV1-M) + domaine | ~25 €/mois | 10-50 utilisateurs, ~25 Go stockage |
| Render Starter + domaine + sauvegardes | ~12 €/mois | Convient pour démarrer |
| AWS EC2 + RDS + S3 (full managed) | ~50-100 €/mois | Production sérieuse, scaling automatique |

Pour Marie Eugénie qui démarre, un **VPS Scaleway DEV1-S à 10 €/mois** + nom de domaine est largement suffisant.

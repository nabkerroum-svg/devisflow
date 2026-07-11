# DevisFlow — Plateforme de génération de devis Marie Eugénie

Application web complète pour produire des devis Word et PDF à partir de modèles maîtres uploadés. Conserve à l'identique la mise en page, les logos, les certifications, les polices et les marges. Remplace uniquement les zones explicitement marquées comme variables.

## Démarrage en 30 secondes

```bash
# Avec Docker (recommandé)
unzip devisflow.zip
cd devisflow
docker-compose up --build

# Ouvrir http://localhost:8000
```

Au premier lancement, le template Marie Eugénie est chargé automatiquement. Vous pouvez générer un devis immédiatement depuis l'onglet « Créer un devis ».

## Arborescence du projet

```
devisflow/
│
├── README.md                          ← ce fichier (point d'entrée principal)
├── Dockerfile                         ← image Docker (Python + LibreOffice)
├── docker-compose.yml                 ← orchestration (un seul service avec volume)
├── .gitignore
│
├── backend/                           ← API Python FastAPI
│   ├── requirements.txt               ← dépendances pip
│   ├── config.py                      ← paramètres (chemins, conversion PDF, DB)
│   ├── main.py                        ← entrypoint FastAPI
│   ├── models.py                      ← tables SQLModel (Template, Devis, Client, Equipement)
│   ├── template_service.py            ← cœur métier : analyse, annotation, génération, PDF
│   ├── routes_templates.py            ← endpoints CRUD modèles Word
│   ├── routes_devis.py                ← endpoints génération devis
│   ├── seed.py                        ← chargement initial du template Marie Eugénie
│   │
│   ├── seed_data/
│   │   └── copro_petite.docx          ← template Marie Eugénie pré-annoté (3,8 Mo)
│   │
│   └── storage/                       ← données persistantes (volume Docker)
│       ├── templates/                 ← .docx maîtres uploadés
│       ├── generated/                 ← devis générés (.docx + .pdf)
│       └── db/devisflow.db            ← base SQLite
│
├── frontend/                          ← Interface web
│   ├── index.html                     ← SPA 3 vues
│   ├── static/
│   │   ├── styles.css
│   │   └── app.js                     ← logique JS (appels API)
│   │
│   └── standalone/
│       └── devisflow.html             ← prototype HTML autonome riche (1,1 Mo)
│                                         à utiliser comme démo ou base d'inspiration
│
└── docs/                              ← Documentation
    ├── GUIDE_INSTALLATION.md          ← installation pas à pas (Docker + local)
    ├── GUIDE_DEPLOIEMENT.md           ← VPS, Docker, Render, Railway, etc.
    ├── GUIDE_ANNOTATION_TEMPLATES.md  ← comment préparer un .docx avec marqueurs
    ├── DOCUMENTATION_TECHNIQUE.md     ← architecture, flux, composants, schéma DB
    ├── ANALYSE_DEVIS_ORIGINAL.md      ← analyse typographique du devis Marie Eugénie
    ├── ARCHITECTURE_BACKEND_WORD.md   ← document de cadrage initial du projet
    │
    └── exemples/                      ← fichiers concrets de référence
        ├── template_marie_eugenie_annote.docx   ← exemple de template annoté
        ├── exemple_devis_genere.docx            ← exemple de devis produit
        ├── exemple_devis_genere.pdf
        ├── devis_original_page1.png             ← page 1 du devis Marie Eugénie original
        └── exemple_page1.png                    ← page 1 du devis généré par le système
```

## Lecture conseillée pour reprendre le projet

Si vous reprenez ce projet sans contexte préalable, lire dans l'ordre :

1. **Ce README** (5 min) — vue d'ensemble et démarrage rapide
2. **`docs/GUIDE_INSTALLATION.md`** (10 min) — installer et lancer le projet
3. **`docs/DOCUMENTATION_TECHNIQUE.md`** (20 min) — architecture et fonctionnement interne
4. **`docs/GUIDE_ANNOTATION_TEMPLATES.md`** (10 min) — pour préparer de nouveaux modèles Word
5. **`docs/GUIDE_DEPLOIEMENT.md`** (15 min) — quand vous serez prêt à déployer en production
6. **`docs/exemples/`** — examiner les fichiers concrets pour comprendre le résultat attendu

## Fonctionnalités

### Implémentées
- ✅ **Deux modèles maîtres** pré-chargés et pleinement fonctionnels : « Copropriété petite » (récurrent) et « Enlèvement d'encombrants » (ponctuel)
- ✅ **Calcul automatique des prix** : heures × agents × taux horaire × coefficient de technicité + frais → HT / TVA / TTC, avec agrégation mensuelle (récurrent) ou forfaitaire (ponctuel)
- ✅ **Override manuel** du prix final (le calcul auto reste pré-rempli)
- ✅ **Tableau d'options comparatives** dans le devis récurrent (1×, 2×, 3×/semaine… chiffrées automatiquement)
- ✅ **Boucle de prestations** dynamique dans le devis ponctuel (les anciennes prestations en dur sont retirées)
- ✅ Back-office « Modèles » : upload `.docx` **ou `.doc`** (conversion auto), prévisualisation, suppression, gestion du défaut
- ✅ **Table de substitution stockée en base** par modèle (importer un nouveau modèle sans toucher au code)
- ✅ Génération `.docx` + `.pdf` strictement identiques au modèle (logos, polices, marges, encadrés, pieds de page conservés)
- ✅ Aperçu prix temps réel (`POST /api/devis/calculer`)
- ✅ API REST documentée (Swagger sur `/docs`)
- ✅ Persistance SQLite + stockage disque · Containerisation Docker

### À développer (phase 2)
- 🔲 **Photos d'équipement** injectées à droite de la prestation (alignées, redimensionnées) dans le devis ponctuel — la bibliothèque `Equipement` et le dossier `storage/photos/` sont déjà en place
- 🔲 Masquage automatique des sections de fréquence non contractées (récurrent)
- 🔲 Authentification multi-utilisateurs · Gestion clients en back-office · Édition des devis générés · Migration PostgreSQL

## Stack technique

- **Backend** : Python 3.12, FastAPI, SQLModel
- **Génération Word** : docxtpl 0.20.2, python-docx
- **Conversion PDF** : LibreOffice headless
- **Frontend** : HTML/CSS/JS vanilla (zéro dépendance npm)
- **Base de données** : SQLite par défaut (migrable PostgreSQL)
- **Conteneurisation** : Docker + Docker Compose

## API REST

Documentation interactive : **`http://localhost:8000/docs`**

Endpoints principaux :
| Méthode | URL | Action |
|---|---|---|
| `GET` | `/api/health` | Statut de l'API |
| `GET` | `/api/templates` | Liste des modèles |
| `POST` | `/api/templates/upload` | Upload d'un `.docx` |
| `PUT` | `/api/templates/{code}` | Modifier un modèle |
| `DELETE` | `/api/templates/{code}` | Supprimer un modèle |
| `GET` | `/api/templates/{code}/preview` | Aperçu PDF du template |
| `POST` | `/api/devis/generer` | Générer un devis |
| `GET` | `/api/devis` | Liste des devis générés |
| `GET` | `/api/devis/{numero}/docx` | Télécharger un devis (Word) |
| `GET` | `/api/devis/{numero}/pdf` | Télécharger un devis (PDF) |

Exemple curl :
```bash
curl -X POST http://localhost:8000/api/devis/generer \
  -H "Content-Type: application/json" \
  -d '{
    "numero": "ME-6245",
    "template_code": "copro_petite",
    "variables": {
      "NUMERO_DEVIS": "ME-6245",
      "DATE_EMISSION": "19 juin 2026",
      "DEST_LIGNE1": "Madame Sophie MARCHAND",
      "DEST_LIGNE2": "Syndic Foncia Marseille",
      "DEST_LIGNE3": "12 rue Paradis",
      "DEST_LIGNE4": "13001 Marseille",
      "TYPE_PRESTATION": "Entretien des parties communes",
      "SITE_ADRESSE": "52 rue Louis Astruc",
      "SITE_CP_VILLE": "13005 Marseille"
    },
    "client_nom": "Syndic Foncia Marseille"
  }'
```

## Sauvegarde

Toutes les données persistantes sont dans **`backend/storage/`**. Sauvegarder ce dossier suffit. Voir `docs/GUIDE_INSTALLATION.md` § Sauvegarde et restauration.

## Licence et propriété

Projet développé pour la **SAS Marie-Eugénie** (1 rue Raspail, 13004 Marseille — SIRET 521 797 258). Code source propriété du client commanditaire.

## Support

Le projet est conçu pour être autonome et repris par tout développeur Python sans contexte préalable. Toute la documentation est dans `docs/`.

Si vous reprenez le projet, le point d'entrée du code métier est **`backend/template_service.py`** — c'est là que se joue toute la mécanique d'analyse, annotation et génération.

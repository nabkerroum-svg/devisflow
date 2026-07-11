# Documentation technique

Ce document explique le fonctionnement interne de DevisFlow pour un développeur qui reprend le projet.

## Sommaire

1. [Vue d'ensemble](#vue-densemble)
2. [Architecture](#architecture)
3. [Flux de données](#flux-de-données)
4. [Composants détaillés](#composants-détaillés)
5. [Schéma de la base de données](#schéma-de-la-base-de-données)
6. [Le cœur métier : template_service.py](#le-cœur-métier--template_servicepy)
7. [API REST](#api-rest)
8. [Frontend](#frontend)
9. [Points d'amélioration futurs](#points-damélioration-futurs)

---

## Vue d'ensemble

DevisFlow est une application web qui produit des devis Word et PDF **strictement identiques** à un modèle Word fourni par l'utilisateur, à l'exception de zones variables marquées avec la syntaxe Jinja2 `{{ NOM_VARIABLE }}`.

### Principe central

L'utilisateur fournit un fichier `.docx` qui devient le **modèle maître**. Ce fichier contient :
- La mise en page complète, logos, images, certifications, polices, marges
- Des marqueurs `{{ VARIABLE }}` aux endroits où les données doivent être injectées

À la génération d'un devis :
1. Le backend charge le modèle maître
2. Substitue les marqueurs par les valeurs reçues du formulaire
3. Sauve un `.docx` final qui conserve **tout** sauf les zones substituées
4. Convertit en `.pdf` via LibreOffice

Aucune mise en page n'est générée ou recalculée par le système — le modèle est utilisé tel quel.

### Pourquoi cette approche

Alternative écartée : reconstruire le devis en HTML/CSS puis le convertir. Problème : impossible de reproduire **strictement** une mise en page Word complexe (filigranes, zones de texte ancrées, en-têtes/pieds de page différents par page, etc.).

Solution adoptée : `docxtpl` (Jinja2 + python-docx) qui lit le XML interne du `.docx` et ne modifie que les nœuds texte contenant les marqueurs, en laissant **tout le reste** intact.

---

## Architecture

```
                    Utilisateur (navigateur)
                            │
                            │ HTTP
                            ▼
        ┌────────────────────────────────────────┐
        │       Frontend (HTML/CSS/JS)           │
        │  - index.html  (3 vues SPA)            │
        │  - styles.css                          │
        │  - app.js  (appels fetch vers /api)    │
        └────────────────────────────────────────┘
                            │ /api/*
                            ▼
        ┌────────────────────────────────────────┐
        │    Backend FastAPI (Python)            │
        │                                        │
        │   main.py     ← entrypoint            │
        │      │                                 │
        │      ├── routes_templates.py           │
        │      │     CRUD modèles Word           │
        │      │                                 │
        │      └── routes_devis.py               │
        │            génération devis            │
        │                                        │
        │   template_service.py  ← cœur métier  │
        │      - analyser_template()             │
        │      - annoter_auto()                  │
        │      - generer_devis()                 │
        │      - docx_to_pdf()                   │
        │                                        │
        │   models.py    (SQLModel/SQLAlchemy)   │
        │   config.py    (chemins, settings)     │
        └────────────────────────────────────────┘
                  │              │
                  ▼              ▼
        ┌────────────┐   ┌───────────────┐
        │  SQLite    │   │ Stockage      │
        │  storage/  │   │ disque        │
        │  db/       │   │  templates/   │
        │  *.db      │   │  generated/   │
        └────────────┘   └───────────────┘
                            │
                            ▼
                ┌───────────────────────┐
                │ LibreOffice headless  │
                │  conversion DOCX→PDF  │
                └───────────────────────┘
```

### Choix techniques

| Brique | Choix | Justification |
|---|---|---|
| Backend web | **FastAPI** | Asynchrone, autodocumenté (`/docs`), validation Pydantic |
| Templates Word | **docxtpl 0.20.2** | Jinja2 sur XML Word, mature, gère images dynamiques |
| ORM | **SQLModel** | Combine SQLAlchemy + Pydantic, syntaxe moderne |
| Base de données | **SQLite** par défaut | Zero-config ; migrable PostgreSQL |
| Conversion PDF | **LibreOffice headless** | Fidélité parfaite au Word source ; alternative : CloudConvert API |
| Frontend | **HTML/CSS/JS vanilla** | Pas de dépendance npm, projet portable, < 50 Ko |
| Conteneurisation | **Docker** | LibreOffice + Python + dépendances en une image |

---

## Flux de données

### Flux 1 : Upload d'un template Word

```
[1] Utilisateur sélectionne un .docx dans l'UI Modèles PDF
        │
[2] Frontend POST multipart vers /api/templates/upload
        │  (fichier, code, nom, famille, type_intervention, annoter)
        │
[3] routes_templates.py :
    - sauve le .docx dans storage/templates/{code}.docx
    - si annoter=True : appelle template_service.annoter_auto()
                        qui applique DEFAULT_SUBSTITUTIONS
    - appelle template_service.analyser_template() qui scanne le XML
      et retourne la liste des variables {{ X }} présentes
    - crée une entrée Template en DB (avec variables[] sérialisée en JSON)
        │
[4] Réponse : { ok, template_id, variables_detectees, fichier }
        │
[5] Frontend rafraîchit la liste des templates et le sélecteur
    du formulaire de création de devis
```

### Flux 2 : Génération d'un devis

```
[1] Utilisateur remplit le formulaire et clique "Générer"
        │
[2] Frontend construit le payload JSON :
    {
      numero: "ME-6245",
      template_code: "copro_petite",
      variables: { NUMERO_DEVIS, DATE_EMISSION, DEST_LIGNE1..4, ... }
    }
    POST /api/devis/generer
        │
[3] routes_devis.py.generer_devis() :
    - Charge le Template en DB par code
    - Vérifie qu'il est actif
    - Localise storage/templates/{fichier}
        │
[4] template_service.generer_devis() :
    - DocxTemplate(template_path)
    - Auto-complète les variables manquantes avec "" (évite Jinja undefined)
    - doc.render(data)
    - Sauve dans storage/generated/{numero}.docx
    - _nettoyer_doublons_zip() en post-traitement
        │
[5] template_service.docx_to_pdf() :
    - subprocess soffice --headless --convert-to pdf
    - Timeout 60 secondes
    - Retourne storage/generated/{numero}.pdf
        │
[6] Persiste un Devis en DB (numéro, template_code, fichiers, payload JSON)
        │
[7] Réponse :
    { ok, numero, docx_url, pdf_url }
        │
[8] Frontend affiche les liens de téléchargement
```

### Flux 3 : Téléchargement d'un devis généré

```
GET /api/devis/{numero}/docx   ou   /api/devis/{numero}/pdf
   │
   ├── Cherche Devis en DB par numero
   ├── Vérifie que le fichier disque existe
   └── FileResponse avec Content-Disposition: attachment
```

---

## Composants détaillés

### `backend/config.py`
Paramètres centralisés du projet. Tous les chemins de stockage, l'URL de la base de données, le binaire LibreOffice y sont définis. Lit certaines valeurs depuis l'environnement (`SOFFICE_BIN`, `DATABASE_URL` à ajouter).

### `backend/main.py`
Point d'entrée FastAPI. Configure CORS, monte les routers, démarre le seed initial via `lifespan`. Sert aussi le frontend statique sur `/` (SPA fallback).

### `backend/models.py`
Définit les 4 tables SQLModel :
- **Template** : modèles Word maîtres
- **Devis** : devis générés
- **Client** : carnet de clients (à brancher)
- **Equipement** : bibliothèque matériels (à brancher)

`init_db()` crée les tables au démarrage. `get_session()` est la dépendance FastAPI qui ouvre une session par requête.

### `backend/template_service.py`
**Cœur métier** — voir section dédiée plus bas.

### `backend/routes_templates.py`
Endpoints CRUD pour les modèles Word :
- `GET /api/templates` — liste
- `POST /api/templates/upload` — upload + analyse + annotation auto optionnelle
- `PUT /api/templates/{code}` — modifier métadonnées
- `DELETE /api/templates/{code}` — supprimer
- `GET /api/templates/{code}/download` — télécharger le .docx maître
- `GET /api/templates/{code}/preview` — aperçu PDF (avec marqueurs visibles)

### `backend/routes_devis.py`
Endpoints de génération et consultation :
- `POST /api/devis/generer` — produit un .docx + .pdf
- `GET /api/devis` — liste des devis produits
- `GET /api/devis/{numero}/docx` et `/pdf` — téléchargement

### `backend/seed.py`
Charge automatiquement le template Marie Eugénie au premier démarrage si la base est vide.

### `frontend/index.html` + `static/styles.css` + `static/app.js`
SPA légère en HTML/CSS/JS vanilla. Trois vues : Créer un devis, Mes devis, Modèles PDF. Appels fetch vers `/api/*`.

### `frontend/standalone/devisflow.html`
**Version autonome du prototype d'IHM** (~1,1 Mo). C'est le prototype initial développé avant le backend. Contient une UI très riche (gestion équipements avec photos, bibliothèque métier, multi-utilisateurs, autocomplete clients, modèles de devis administrables côté UI, etc.).

Cet HTML utilise localStorage pour la persistance et ne nécessite pas de serveur. Il peut servir de :
- **Démo standalone** (ouvrir le fichier dans un navigateur)
- **Référence des règles métier** (modèles de prestation, zones, conditions de paiement, etc.)
- **Base d'inspiration** pour étendre l'UI de la v1 backend

Pour basculer ses fonctionnalités dans le frontend v1, voir « Points d'amélioration futurs » plus bas.

---

## Schéma de la base de données

Diagramme entité-relation :

```
┌──────────────────────────┐         ┌──────────────────────────┐
│        Template          │         │         Devis            │
│──────────────────────────│         │──────────────────────────│
│ id (PK)                  │         │ id (PK)                  │
│ code (UNIQUE, indexé)    │◄────┐   │ numero (UNIQUE, indexé)  │
│ nom                      │     │   │ template_code (FK)       │
│ famille                  │     └───│ client_nom               │
│ fichier (path relatif)   │         │ site_adresse             │
│ type_intervention        │         │ date_emission            │
│ is_default (bool)        │         │ montant_ht / montant_ttc │
│ actif (bool)             │         │ fichier_docx / fichier_pdf│
│ variables (JSON string)  │         │ payload (JSON string)    │
│ created_at / updated_at  │         │ statut                   │
└──────────────────────────┘         │ created_at / updated_at  │
                                     └──────────────────────────┘

┌──────────────────────────┐         ┌──────────────────────────┐
│         Client           │         │       Equipement         │
│──────────────────────────│         │──────────────────────────│
│ id (PK)                  │         │ id (PK)                  │
│ nom (indexé)             │         │ code (UNIQUE, indexé)    │
│ civilite / contact       │         │ label                    │
│ email / telephone        │         │ categorie                │
│ adresse / code_postal    │         │ description              │
│ ville                    │         │ photo_path               │
│ site_nom / site_adresse  │         │ actif                    │
│ created_at               │         │ created_at               │
└──────────────────────────┘         └──────────────────────────┘
```

### Scripts SQL équivalents

SQLModel génère ces tables automatiquement à `init_db()`. Si vous voulez les créer manuellement (ex : PostgreSQL externe) :

```sql
CREATE TABLE template (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    code VARCHAR NOT NULL UNIQUE,
    nom VARCHAR NOT NULL,
    famille VARCHAR DEFAULT 'contrat',
    fichier VARCHAR NOT NULL,
    type_intervention VARCHAR,
    is_default BOOLEAN DEFAULT FALSE,
    actif BOOLEAN DEFAULT TRUE,
    created_at DATETIME,
    updated_at DATETIME,
    variables TEXT  -- JSON
);
CREATE INDEX ix_template_code ON template(code);

CREATE TABLE devis (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    numero VARCHAR NOT NULL UNIQUE,
    template_code VARCHAR NOT NULL,
    client_nom VARCHAR NOT NULL,
    site_adresse VARCHAR NOT NULL,
    date_emission VARCHAR NOT NULL,
    montant_ht REAL DEFAULT 0.0,
    montant_ttc REAL DEFAULT 0.0,
    fichier_docx VARCHAR,
    fichier_pdf VARCHAR,
    payload TEXT,  -- JSON
    statut VARCHAR DEFAULT 'brouillon',
    created_at DATETIME,
    updated_at DATETIME
);
CREATE INDEX ix_devis_numero ON devis(numero);

CREATE TABLE client (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nom VARCHAR NOT NULL,
    civilite VARCHAR,
    contact VARCHAR,
    email VARCHAR,
    telephone VARCHAR,
    adresse VARCHAR,
    code_postal VARCHAR,
    ville VARCHAR,
    site_nom VARCHAR,
    site_adresse VARCHAR,
    created_at DATETIME
);
CREATE INDEX ix_client_nom ON client(nom);

CREATE TABLE equipement (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    code VARCHAR NOT NULL UNIQUE,
    label VARCHAR NOT NULL,
    categorie VARCHAR DEFAULT 'materiel',
    description VARCHAR,
    photo_path VARCHAR,
    actif BOOLEAN DEFAULT TRUE,
    created_at DATETIME
);
CREATE INDEX ix_equipement_code ON equipement(code);
```

### Relations
- `Devis.template_code` réfère à `Template.code` (clé étrangère logique, pas définie en strict en SQLite pour souplesse)
- Pas de relations strictes définies pour le moment ; à ajouter quand Clients et Équipements seront branchés à la génération.

---

## Le cœur métier : `template_service.py`

Quatre fonctions principales.

### `analyser_template(docx_path: Path) -> List[str]`

Ouvre le `.docx` en tant que ZIP, lit chaque fichier XML interne (document.xml, header*.xml, footer*.xml), applique la regex `\{\{\s*([A-Z][A-Z0-9_]*)\s*\}\}` pour extraire tous les noms de variables Jinja2.

Retourne la liste triée des variables détectées.

Pourquoi lire le XML directement plutôt que via python-docx ? Parce que python-docx ne voit pas certaines zones (text frames, shapes, headers/footers complexes). La lecture XML couvre **100% du contenu**.

### `annoter_auto(docx_in, docx_out, substitutions=None) -> int`

Charge le `.docx` avec python-docx, parcourt tous les paragraphes (corps + cellules de tableaux) et applique les substitutions de la table `DEFAULT_SUBSTITUTIONS` :

```python
DEFAULT_SUBSTITUTIONS = [
    # (texte_original, marqueur_remplacement, exact_match)
    ("Marseille, le 21 juin 2022", "Marseille, le {{ DATE_EMISSION }}", False),
    ("client", "{{ DEST_LIGNE1 }}", True),
    ("adresse 1", "{{ DEST_LIGNE2 }}", False),
    # ...
]
```

- `exact_match=False` : substitution sur la première occurrence du `old_text` dans un paragraphe.
- `exact_match=True` : le paragraphe **entier** doit être égal à `old_text` (utile pour les zones jaunes Word qui ne contiennent que le texte à remplacer).

La fonction `_replace_in_paragraph()` gère la fragmentation Word : un texte comme `{{ DATE }}` peut être réparti sur plusieurs `<w:r>` (runs). Elle concatène tous les runs du paragraphe, fait le remplacement, puis remet le résultat dans le **premier run** en préservant son formatage (police, taille, couleur), et vide les runs suivants.

**Limites** : ne traite que les paragraphes accessibles via `doc.paragraphs` et `doc.tables`. Les **zones de texte (text frames)** et **shapes ancrés** ne sont pas couverts. Pour les couvrir, il faudrait traiter directement le XML brut.

### `generer_devis(template_path, data, output_path) -> Path`

1. Charge le template via `DocxTemplate`
2. Auto-complète les variables manquantes du payload avec `""` pour éviter les erreurs Jinja `undefined`
3. `doc.render(data)` — substitution effective
4. `doc.save(output_path)`
5. `_nettoyer_doublons_zip()` — corrige le bug `Duplicate name: docProps/core.xml` qui peut être produit par python-docx et empêche LibreOffice d'ouvrir le fichier

### `docx_to_pdf(docx_path, pdf_dir=None) -> Path`

Lance LibreOffice headless en sous-processus :
```bash
soffice --headless --convert-to pdf --outdir {pdf_dir} {docx_path}
```

Timeout : 60 secondes. En cas d'échec, lève une `RuntimeError` avec le message LibreOffice.

**Note importante** : LibreOffice produit un PDF très fidèle au Word source, mais peut différer légèrement de Word lui-même sur certains détails (espacement, polices manquantes). Pour une fidélité absolue, utiliser Microsoft Word côté serveur (compliqué) ou une API tierce comme CloudConvert.

---

## API REST

Documentation auto-générée Swagger : **`http://localhost:8000/docs`**

Endpoints (tous préfixés par `/api`) :

### Health
```
GET /api/health
→ {"status": "ok"}
```

### Templates
```
GET    /api/templates
POST   /api/templates/upload          (multipart : fichier, code, nom, famille, type_intervention, annoter)
PUT    /api/templates/{code}          (form : nom, famille, is_default, actif)
DELETE /api/templates/{code}
GET    /api/templates/{code}/download  → fichier .docx
GET    /api/templates/{code}/preview   → fichier .pdf
```

### Devis
```
POST /api/devis/generer    (JSON : numero, template_code, variables{}, client_nom, ...)
GET  /api/devis             → liste des devis générés
GET  /api/devis/{numero}/docx → fichier .docx
GET  /api/devis/{numero}/pdf  → fichier .pdf
```

### Format du payload de génération

```json
{
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
  "client_nom": "Syndic Foncia Marseille",
  "site_adresse": "52 rue Louis Astruc, 13005 Marseille",
  "date_emission": "19 juin 2026"
}
```

Les clés de `variables` doivent correspondre aux marqueurs `{{ X }}` présents dans le template. Pour découvrir ces marqueurs, faire `GET /api/templates` qui retourne le champ `variables[]`.

---

## Frontend

### Architecture
Trois fichiers (≈ 40 Ko au total) :
- `index.html` : structure HTML + 3 sections `<section class="view">` (création / mes devis / modèles)
- `static/styles.css` : palette Marie Eugénie (or #984806, ink #1a1a1a), grille flexible, table listing
- `static/app.js` : navigation, appels API, rendu des listings

### Navigation
Pas de framework. Bascule entre les vues via `display: none` / `is-active` sur `<section>`. Les onglets de navigation manipulent les classes.

### Appels API
Helper `api(path, opts)` qui gère le préfixe et les erreurs. Tous les appels passent par cette fonction.

### Évolutions à prévoir
Le frontend v1 est intentionnellement minimal pour assurer le fonctionnement de bout en bout. Pour enrichir l'UX, deux options :
1. **Migrer vers le prototype HTML autonome riche** (`frontend/standalone/devisflow.html`) en remplaçant ses appels localStorage par des appels REST vers le backend
2. **Refactor en React/Vue** pour une SPA moderne avec routing, state management, etc.

---

## Points d'amélioration futurs

Liste priorisée des fonctionnalités à ajouter (estimations en jours-développeur senior) :

### Priorité haute

1. **Authentification multi-utilisateurs** (2-3 j)
   - `fastapi-users` ou middleware JWT
   - Rôles : admin, commercial, lecture seule
   - Lier chaque devis au commercial créateur

2. **Boucles dans les templates Word** (2-3 j)
   - Permettre `{% for zone in zones %}...{% endfor %}` dans le `.docx`
   - Adapter le payload pour envoyer des tableaux
   - Documenter la syntaxe docxtpl

3. **Injection de photos d'équipement** (2 j)
   - Utiliser `docxtpl.InlineImage` dans le rendu
   - Brancher la table `Equipement` à la génération
   - Permettre upload de photos via le back-office

4. **Robustesse de l'annotation auto** (1-2 j)
   - Gérer les zones de texte / text frames Word
   - Détecter et corriger les marqueurs fragmentés en plusieurs runs

### Priorité moyenne

5. **Migration PostgreSQL** + variables d'environnement (1 j)
6. **Gestion clients en back-office** + autocomplete dans le formulaire devis (2 j)
7. **Édition des devis générés** (rouvrir un devis pour modifier les variables) (1-2 j)
8. **Historique des envois mail** + bouton "Envoyer par mail" avec PJ (2 j)
9. **Versionning des templates** (garder l'historique des modifications) (2 j)

### Priorité basse

10. **Migration frontend vers React/Vue** (5-10 j)
11. **Export comptable** (Sage, EBP) (2-3 j)
12. **API externe pour CRM Marie Eugénie** (Salesforce) (3-5 j)

---

## Pour reprendre le projet

Si vous êtes un développeur qui reprend le projet sans contexte :

1. Lire ce document
2. Lire `README.md` à la racine et `docs/GUIDE_INSTALLATION.md`
3. Lancer en local avec Docker
4. Tester la génération d'un devis (le template Marie Eugénie est pré-chargé)
5. Ouvrir `backend/template_service.py` — c'est là que se joue toute la mécanique
6. Pour ajouter une fonctionnalité, créer un nouveau router dans `backend/routes_*.py`, le brancher dans `main.py`
7. Pour ajouter une table DB, l'ajouter dans `models.py` ; SQLModel crée la table automatiquement au prochain démarrage

Les fichiers d'exemples (`docs/exemples/`) montrent un cas concret de bout en bout : template annoté + devis généré + PDF.

## Contacts

Le projet a été développé pour la SAS Marie-Eugénie (Marseille). Coordonnées dans `backend/seed_data/copro_petite.docx` (pied de page).

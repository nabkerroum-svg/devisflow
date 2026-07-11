# Architecture cible — Génération de devis Word à partir d'un template maître

> Document préparatoire pour la mise en production de DevisFlow avec génération
> de fichiers Word fidèles à votre modèle d'origine. À transmettre à votre
> prestataire technique quand vous serez prêt à passer du prototype HTML
> autonome à une vraie application client/serveur.

## Objectif métier

Permettre à Marie Eugénie de **fournir un fichier Word de référence** (par exemple
`1_COPRO_petite.docx`) qui devient le **modèle maître** de génération. Les futurs
devis générés sont des fichiers Word **bit-pour-bit identiques** au modèle, sauf
sur des zones variables explicitement marquées (prestations, prix, client, dates).

## Architecture recommandée

```
┌──────────────────────────┐         ┌─────────────────────────┐
│   Frontend React/HTML    │  HTTPS  │   Backend Python        │
│   (interface DevisFlow)  │ ──────► │   (Flask ou FastAPI)    │
│                          │         │                         │
│   - formulaire devis     │         │   - lit le template .docx│
│   - sélection client     │         │   - injecte les valeurs │
│   - sélection prestations│         │   - retourne le .docx   │
│   - choix du gabarit     │         │     généré              │
└──────────────────────────┘         └────────────┬────────────┘
                                                  │
                                                  ▼
                                     ┌─────────────────────────┐
                                     │   Stockage S3 / disque  │
                                     │                         │
                                     │   - templates_word/     │
                                     │     copro_petite.docx   │
                                     │     copro_grande.docx   │
                                     │     vitrerie.docx       │
                                     │     ...                 │
                                     │                         │
                                     │   - devis_generes/      │
                                     │     ME-6245.docx        │
                                     │     ME-6246.docx        │
                                     │     ...                 │
                                     └─────────────────────────┘
```

## Stack technique recommandée

### Backend
- **Python 3.11+** avec **FastAPI** (asynchrone, performant, autodocumenté)
- **docxtpl** (https://docxtpl.readthedocs.io/) : moteur Jinja2 pour templates Word
- **python-docx** : pour les manipulations bas-niveau qui dépassent docxtpl
- **Pillow** : pour redimensionner les photos d'équipement avant injection
- **PostgreSQL** : base de données (clients, devis, gabarits, équipements)
- **SQLAlchemy** + **Alembic** : ORM et migrations

### Frontend
- Réutiliser le prototype HTML existant comme base, ou refactor en React/TypeScript
- Communication via fetch / Axios vers l'API REST du backend
- Affichage de l'aperçu : conserver le rendu HTML actuel (qui est une approximation),
  mais le bouton « Télécharger Word » appelle le backend qui retourne le vrai fichier

### Hébergement
- **Scaleway** (souveraineté française, RGPD) — recommandé
- Alternative : **OVHcloud** (Roubaix), **Clever Cloud**
- Conteneurs Docker, déploiement via GitHub Actions

## Préparation du template Word côté Marie Eugénie

Le travail le plus important côté métier : **annoter** votre fichier Word de référence
avec les variables Jinja2 que docxtpl reconnaîtra.

### Variables simples

Dans votre fichier `.docx`, vous remplacez les zones surlignées jaune par des balises :

| Avant (devis actuel) | Après (template) |
|---|---|
| `Marseille, le ` `21 juin 2022` | `Marseille, le {{date_emission}}` |
| `client` `adresse 1` `adresse 2` `contact` | `{{client.civilite}} {{client.contact}}` `{{client.raison_sociale}}` `{{client.adresse}}` `{{client.code_postal}} {{client.ville}}` |
| `52 rue Louis Astruc,` `13005 MARSEILLE.` | `{{site.adresse}},` `{{site.code_postal}} {{site.ville}}.` |
| `Proposition` `ME` | `Proposition {{numero_devis}}` |
| `Entretien des parties communes` | `{{type_prestation}}` |

### Boucle sur les zones (contrats récurrents)

Dans le tableau de détail des prestations :

```
{%tr for zone in zones %}
| {{zone.titre}} |
{% for ope in zone.operations_actives %}
| - {{ope}} |
{% endfor %}
| Fréquence : {{zone.frequence}} |
{%tr endfor %}
```

### Tableau financier

```
| Prestation | Prix HT | TVA 20% | Prix TTC |
| {{libelle}} | {{totaux.ht}} | {{totaux.tva}} | {{totaux.ttc}} |
```

### Images dynamiques (photos d'équipement)

```python
# Côté backend, on injecte les images avec docxtpl
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm

context = {
    "materiel_photos": [
        {"label": "Autolaveuse", "photo": InlineImage(tpl, "autolaveuse.jpg", width=Mm(40))},
        {"label": "Monobrosse", "photo": InlineImage(tpl, "monobrosse.jpg", width=Mm(40))},
    ]
}
```

Dans le template Word :
```
{% for m in materiel_photos %}
{{m.photo}} {{m.label}}
{% endfor %}
```

## Endpoints API à implémenter

### `POST /api/devis/generer`
**Entrée** (JSON) :
```json
{
  "numero": "ME-6245",
  "date_emission": "2026-06-18",
  "modele_code": "copro_petite",
  "gabarit_code": "classique",
  "client": {
    "civilite": "Madame",
    "contact": "Sophie MARCHAND",
    "raison_sociale": "Syndic Foncia Marseille",
    "adresse": "12 rue Paradis",
    "code_postal": "13001",
    "ville": "Marseille",
    "email": "s.marchand@foncia.fr"
  },
  "site": {
    "nom": "Résidence Le Prado",
    "adresse": "52 rue Louis Astruc",
    "code_postal": "13005",
    "ville": "Marseille",
    "batiment": "Bât. A",
    "etage": "RDC + 4 étages",
    "code_acces": "A1234B"
  },
  "zones": [...],
  "totaux": { "ht": 180, "tva": 36, "ttc": 216 },
  "conditions": "...",
  "materiel": ["autolaveuse", "monobrosse", "aspirateur", "vehicule"]
}
```

**Sortie** : fichier `.docx` binaire avec `Content-Disposition: attachment`

**Logique** :
1. Charger le template Word correspondant à `gabarit_code` (ou famille du modèle)
2. Charger les photos d'équipement depuis `MATERIEL_PAR_MODELE[modele_code]`
3. Préparer le contexte Jinja2
4. `DocxTemplate.render(context)` puis `.save(buffer)`
5. Retourner le buffer

### `POST /api/devis/{numero}/pdf`
Conversion `.docx → .pdf` via **LibreOffice headless** côté serveur :
```bash
soffice --headless --convert-to pdf input.docx --outdir /tmp/
```

### `POST /api/templates/upload`
Permet à Marie Eugénie d'uploader un nouveau template Word annoté pour
créer un gabarit. Le backend :
1. Reçoit le fichier `.docx`
2. Valide qu'il contient les marqueurs Jinja2 attendus
3. Sauvegarde dans S3 sous une clé unique
4. Crée une entrée en base associée à un gabarit

## Estimation budgétaire pour la piste B

| Poste | Coût estimé | Détail |
|---|---|---|
| Développement backend Python (FastAPI + docxtpl + PostgreSQL) | 8-12 j-h | Endpoints, base, tests |
| Annotation des templates Word existants (Marie Eugénie) | 2-3 j-h | Reprise des 15 modèles, ajout marqueurs Jinja2 |
| Refactor frontend pour brancher l'API | 3-5 j-h | Remplacer la génération HTML par appel API |
| Setup infra (Docker, CI/CD, Scaleway) | 2-3 j-h | Déploiement initial |
| Migration des données du prototype (localStorage → PostgreSQL) | 1-2 j-h | Script d'import |
| Tests, recette, formation utilisateurs | 3-4 j-h | |
| **TOTAL** | **19-29 j-h** | ~7-10k€ TTC pour un freelance senior |

Hébergement mensuel : ~30-80 € (Scaleway petit instance + DB + stockage).

## Étapes pour démarrer la piste B

1. **Choisir un prestataire** : freelance Python senior recommandé (chercher sur Malt avec mots-clés FastAPI + docxtpl). Donner ce document comme cahier des charges initial.
2. **Préparer les templates Word annotés** : reprendre vos 15 modèles `.doc`, les convertir en `.docx`, et ajouter les marqueurs Jinja2 dans Word directement (pas besoin de connaissance technique — c'est de la simple substitution texte).
3. **Lancer un MVP** : commencer par un seul modèle (`copro_petite`), un seul endpoint (`/api/devis/generer`), et valider que le `.docx` généré est bit-pour-bit identique au template sauf sur les variables.
4. **Élargir progressivement** : ajouter les autres modèles, le PDF, le multi-utilisateurs, etc.

## Le rôle du prototype actuel dans cette migration

Le prototype DevisFlow tel qu'il existe aujourd'hui sert de **maquette fonctionnelle**
et de **référence métier**. Toutes les règles de gestion (associations modèle → matériel,
gabarits, blocs métiers conditionnels, validation client, génération de mailto, etc.)
sont déjà codées et testées. Le passage en production consiste essentiellement à :

- Déporter la **persistance** de localStorage vers PostgreSQL
- Déporter la **génération du document** de HTML/CSS vers docxtpl + template Word
- Conserver l'**interface utilisateur** comme dans le prototype

Tout le travail métier accumulé (modèles de prestations, zones, conditions, mentions,
blocs métiers Action Nuisibles + PEGASE, bibliothèque d'équipements, gabarits) est
réutilisable tel quel.

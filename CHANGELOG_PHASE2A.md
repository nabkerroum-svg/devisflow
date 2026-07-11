# CHANGELOG — Phase 2A : Interface riche reconnectée au backend

Objectif de la phase : faire du prototype riche `frontend/standalone/devisflow.html`
l'**interface principale** de DevisFlow, et **brancher réellement** au backend les
fonctions essentielles (clients, équipe, devis), avec persistance en base.

---

## 1. Constat (vérifié dans le code d'origine)

Le projet contenait **deux interfaces** :
- `frontend/index.html` — interface simple (3 onglets) branchée au backend ;
- `frontend/standalone/devisflow.html` — **prototype riche** (6 vues : Création,
  Mes devis, Clients, Bibliothèque, Équipe, Paramètres), ~60 fonctions JS, mais
  **persistance localStorage uniquement, zéro appel API**.

C'est ce prototype riche qui correspond à « l'outil complet ». La Phase 2A le
remet au centre et le connecte au serveur.

---

## 2. Backend — nouvelles tables (`models.py`)

- `PrestationType` : zones de prestations types (les 8 zones EXACTES du modèle
  Word « Copro Petite » + prestations ponctuelles), avec leurs opérations détaillées.
- `Membre` : membres de l'équipe (nom, e-mail, rôle, actif/inactif).
- `Parametre` : paramètres globaux (taux horaire, TVA, coefficients de technicité,
  codes modèles récurrent/ponctuel, infos société) en clé/valeur.
- `Client` : ajout de `archive` et `updated_at`.
- `Equipement` : ajout de `updated_at`.

## 3. Backend — nouvelles routes (`routes_biblio.py`, nouveau fichier)

- **Clients** : `GET /clients` (avec recherche `?q=` et filtre archives),
  `GET /clients/{id}`, `POST /clients`, `PUT /clients/{id}`,
  `POST /clients/{id}/archive`, `DELETE /clients/{id}`.
- **Équipements** : `GET/POST/PUT/DELETE /equipements`, `GET /equipements/{code}/photo`
  (upload de photo multipart, stockée dans `storage/photos/`).
- **Prestations types** : `GET/POST/PUT/DELETE /prestations` (filtre `?famille=`).
- **Membres** : `GET/POST/PUT/DELETE /membres`.
- **Paramètres** : `GET /parametres`, `PUT /parametres/{cle}`.

## 4. Backend — seed métier (`seed_metier.py`, nouveau fichier)

Pré-remplit au premier démarrage :
- les **8 zones récurrentes extraites du modèle Word Copro** (Hall d'entrée et ses
  13 opérations, ascenseur, escaliers, caves, garage, abords, conteneur, OM) —
  source de référence pour les prestations détaillées ;
- 4 prestations ponctuelles types (encombrants) ;
- 9 équipements (machines, matériel, véhicules, spécifique) ;
- 3 clients de démonstration ;
- 4 membres d'équipe ;
- 10 paramètres de calcul / modèles / société.

Données extraites du `.doc` original et stockées dans `seed_data/zones_copro.json`.

## 5. Backend — câblage (`main.py`)

- Enregistrement du routeur `routes_biblio`.
- Exécution de `seed_metier()` au démarrage.
- La racine `/` sert désormais l'**interface riche** (`index_rich.html`).

---

## 6. Frontend — interface riche reconnectée

- `frontend/index_rich.html` : le prototype riche devient l'interface principale,
  avec le **pont API** injecté.
- `frontend/static/api-bridge.js` (nouveau) : pont entre l'interface et l'API.
  - Au démarrage, vérifie le backend (badge « Connecté au backend » / « Mode
    maquette hors ligne ») puis **charge les clients, l'équipe et les devis réels**
    dans l'état de l'application et rafraîchit les vues.
  - **Redirige les écritures vers l'API** : création/modification de client et de
    membre passent par le serveur (avec repli local si hors ligne).
  - La **sélection d'un client dans le devis prérenseigne** automatiquement les
    coordonnées et l'adresse du site (via les données serveur).

---

## 7. Ce qui est FONCTIONNEL à la fin de la Phase 2A

- Interface complète (6 vues) reconnue comme l'outil DevisFlow.
- **Clients** : lister, rechercher, créer, modifier, archiver — persistés en base.
- **Équipe** : lister, créer, modifier, activer/désactiver — persistés en base.
- **Création de devis** : sélection d'un client réel avec préremplissage auto ;
  génération Word/PDF fidèle aux modèles maîtres (acquis en Phase 1).
- **Mes devis** : liste réelle depuis la base.
- **Prestations détaillées** : les zones exactes du modèle Word sont en base et
  exposées par l'API (`window.ME_ZONES`) pour la création de devis.

## 8. Reporté en Phase 2B

- Bibliothèque métier pleinement éditable depuis l'UI (prestations, matériel).
- Upload de photos d'équipement depuis l'UI + **injection alignée** dans le Word.
- Paramètres avancés branchés au moteur de calcul (lecture des coefficients/​taux
  depuis la base au moment du chiffrage).

---

## 9. Fichiers ajoutés / modifiés

| Fichier | Nature |
|---|---|
| `backend/models.py` | + `PrestationType`, `Membre`, `Parametre` ; `Client`/`Equipement` étendus |
| `backend/routes_biblio.py` | **nouveau** — CRUD clients/équipements/prestations/membres/paramètres |
| `backend/seed_metier.py` | **nouveau** — seed bibliothèque métier |
| `backend/seed_data/zones_copro.json` | **nouveau** — 8 zones extraites du modèle Word |
| `backend/main.py` | routeur biblio + seed métier + sert l'interface riche |
| `frontend/index_rich.html` | **nouveau** — interface riche + pont API (page principale) |
| `frontend/static/api-bridge.js` | **nouveau** — pont interface ↔ backend |

## 10. Tests réalisés

- CRUD clients (création/recherche/modification/archivage/suppression) — OK.
- Toutes les routes GET (clients, prestations, équipements, membres, paramètres) — OK.
- Seed : 3 clients, 9 équipements, 12 prestations (8 zones Copro + 4 ponctuelles),
  4 membres, 10 paramètres — OK.
- Interface : navigation des 6 vues sans erreur JS ; badge de connexion ;
  population des vues Clients et Équipe depuis l'API ; création d'un client via
  modale persistée via l'API — OK (validé en navigateur).

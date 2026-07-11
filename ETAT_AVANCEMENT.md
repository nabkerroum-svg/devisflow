# DevisFlow — État réel d'avancement (audit factuel)

> Document établi en lisant directement le code (tables, routes, pont d'interface)
> et en capturant les 5 vues. Aucune projection : uniquement ce qui existe.

---

## 1. Modules réellement FONCTIONNELS (interface ↔ API ↔ base)

| Module | État | Détail |
|---|---|---|
| **Génération Word/PDF** | ✅ Fonctionnel | Fidèle aux modèles maîtres (prouvé en comparaison) |
| **Aperçu PDF réel** | ✅ Fonctionnel | Affiche le vrai PDF du modèle Word (route `/devis/apercu`) |
| **Calcul automatique des prix** | ✅ Fonctionnel | `pricing.py` — HT/TVA/TTC, options, override |
| **Bibliothèque clients** | ✅ Fonctionnel | Lister, rechercher, créer, modifier, archiver — persisté en base. Sélection + préremplissage dans le devis OK |
| **Équipe** | ✅ Fonctionnel | Lister, créer, modifier, activer/désactiver — persisté en base (capture le confirme : 4 membres réels) |
| **Mes devis** | ✅ Fonctionnel | Liste réelle depuis la base |

## 2. Modules encore en MAQUETTE (s'affichent, mais lisent le localStorage du prototype, pas l'API)

| Module | État | Ce qui manque |
|---|---|---|
| **Bibliothèque métier** (vue) | 🟡 Maquette affichée | La vue montre 15 modèles du prototype (localStorage). L'API `/prestations` existe et contient les 8 zones réelles du Word, mais la vue n'y est pas encore branchée |
| **Paramètres** (vue) | 🟡 Maquette affichée | La vue montre société/numérotation/gabarits (localStorage). L'API `/parametres` existe (taux, TVA, coefficients) mais la vue ne l'utilise pas encore |
| **Équipements / Matériel / Véhicules** | 🟡 Données en base + API, vue non branchée | Table `Equipement` + routes CRUD + 9 équipements seedés existent. La vue d'administration dédiée n'affiche pas encore ces données via l'API |

## 3. Modules PAS ENCORE COMMENCÉS

| Module | État | Note |
|---|---|---|
| **Photos d'équipement — upload depuis l'UI** | 🔴 À faire | La route d'upload existe (`POST /equipements` accepte une photo), mais aucun écran d'upload branché |
| **Photos — injection alignée dans le devis** | 🔴 À faire | C'est le cœur de la phase 2B (tableau 2 colonnes, image à droite de la prestation) |
| **Édition de la bibliothèque métier depuis l'UI** | 🔴 À faire | Créer/modifier une prestation type via l'API depuis l'écran |
| **Paramètres avancés lus par le moteur de calcul** | 🔴 À faire | Le calcul utilise des constantes ; il devra lire taux/coefficients depuis la base |

---

## 4. Tables réellement en base (vérifié dans `models.py`)

| Table | Existe | Utilisée par |
|---|---|---|
| `Template` | ✅ | modèles Word maîtres |
| `Devis` | ✅ | devis générés |
| `TemplateSubstitution` | ✅ | annotation auto (import modèle) |
| `Client` | ✅ | **branchée** (vue Clients) |
| `Equipement` | ✅ | API OK, **vue pas encore branchée** |
| `PrestationType` | ✅ | API OK (8 zones Word + 4 ponctuelles), **vue pas encore branchée** |
| `Membre` | ✅ | **branchée** (vue Équipe) |
| `Parametre` | ✅ | API OK, **vue pas encore branchée** |

## 5. Routes API qui EXISTENT déjà

**Devis** : `POST /devis/calculer`, `POST /devis/apercu`, `POST /devis/generer`,
`GET /devis`, `GET /devis/{n}/docx`, `GET /devis/{n}/pdf`
**Modèles Word** : `GET /templates`, `POST /templates/upload`, `PUT/DELETE /templates/{code}`,
`GET /templates/{code}/download`, `GET /templates/{code}/preview`
**Clients** : `GET /clients` (+recherche), `GET /clients/{id}`, `POST /clients`,
`PUT /clients/{id}`, `POST /clients/{id}/archive`, `DELETE /clients/{id}`
**Équipements** : `GET /equipements`, `POST /equipements` (+photo), `PUT /equipements/{code}`,
`DELETE /equipements/{code}`, `GET /equipements/{code}/photo`
**Prestations** : `GET /prestations`, `POST /prestations`, `PUT /prestations/{code}`, `DELETE /prestations/{code}`
**Membres** : `GET /membres`, `POST /membres`, `PUT /membres/{id}`, `DELETE /membres/{id}`
**Paramètres** : `GET /parametres`, `PUT /parametres/{cle}`

## 6. Routes API qui RESTENT à développer

- `POST /devis/{numero}/photos` ou équivalent : associer des photos à des prestations d'un devis (pour l'injection alignée).
- Lecture des paramètres de calcul par `pricing.py` (pas une route, un branchement interne).
- Éventuellement : import Excel de clients (le bouton existe dans l'UI mais sans route).

> Les routes CRUD principales sont déjà là. Le travail restant est surtout du
> **branchement interface ↔ routes existantes**, pas de la création de routes.

---

## 7. CHECKLIST

### ✅ Terminé
- [x] Génération Word/PDF fidèle aux 2 modèles maîtres
- [x] Aperçu = vrai PDF du modèle Word (plus de rendu HTML reconstruit)
- [x] Calcul automatique des prix (heures, agents, taux, technicité, frais, TVA, options, override)
- [x] Interface riche complète (6 vues) comme interface principale
- [x] Bibliothèque clients : CRUD + recherche + archivage + préremplissage devis
- [x] Équipe : CRUD + actif/inactif
- [x] Mes devis : liste réelle
- [x] Tables + routes API pour clients, équipements, prestations, membres, paramètres
- [x] 8 zones de prestations exactes extraites du modèle Word Copro

### 🟡 En cours
- [ ] Bibliothèque métier : brancher la vue sur l'API `/prestations` (données en base prêtes)
- [ ] Équipements/Matériel/Véhicules : brancher la vue d'administration sur l'API (données prêtes)
- [ ] Paramètres : brancher la vue sur l'API `/parametres` (données prêtes)

### 🔴 À faire
- [ ] Upload de photos d'équipement depuis l'interface
- [ ] **Injection alignée des photos** dans le devis (texte à gauche, photo à droite)
- [ ] Édition de la bibliothèque métier depuis l'UI (créer/modifier une prestation type)
- [ ] Paramètres de calcul lus par le moteur de prix
- [ ] Import Excel de clients (optionnel)

---

## 8. Pourcentage d'avancement global

Estimation par pondération des grands blocs :

| Bloc | Poids | Avancement |
|---|---|---|
| Moteur de génération Word/PDF fidèle | 25 % | 100 % |
| Aperçu PDF réel | 10 % | 100 % |
| Calcul automatique des prix | 15 % | 100 % |
| Clients (bout en bout) | 10 % | 100 % |
| Équipe (bout en bout) | 5 % | 100 % |
| Bibliothèque métier (vue branchée) | 10 % | 40 % (API + données prêtes, vue non branchée) |
| Équipements + matériel + véhicules (vue branchée) | 10 % | 40 % (API + données prêtes, vue non branchée) |
| Paramètres (vue branchée + lus par le calcul) | 5 % | 30 % |
| Photos d'équipement (upload + injection alignée) | 10 % | 5 % (route upload seulement) |

**Avancement global pondéré : ≈ 72 %**

Répartition : tout le cœur (génération fidèle, aperçu réel, calcul, clients, équipe)
est terminé. Le quart restant concerne le branchement de 3 vues d'administration
(dont les données et les API existent déjà) et la fonctionnalité photos (phase 2B).

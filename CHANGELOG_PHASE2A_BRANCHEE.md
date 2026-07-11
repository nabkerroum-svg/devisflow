# CHANGELOG — Phase 2A branchée (vérifiée avec le vrai backend)

Les vues **Bibliothèque métier**, **Équipements / matériel / véhicules** et
**Paramètres** lisent et écrivent désormais dans la **vraie base SQLite** —
plus de localStorage, plus d'API simulée.

---

## Commande exacte pour lancer le projet

```bash
unzip devisflow.zip
cd devisflow
docker-compose up --build
# puis ouvrir http://localhost:8000
```

Sans Docker (développement) :
```bash
cd devisflow/backend
pip install -r requirements.txt
uvicorn main:app --port 8000
# ouvrir http://localhost:8000
```

Au premier démarrage, la base est créée et pré-remplie automatiquement
(2 modèles Word, 12 prestations types, 9 équipements, 3 clients, 4 membres,
10 paramètres).

---

## Ce qui est branché (vérifié, captures avec backend réel)

### Bibliothèque métier — prestations types
- La vue affiche les **12 prestations réelles de la base** (8 zones exactes du
  modèle Word Copro : Hall d'entrée 13 ops, ascenseur, escaliers, caves, garage,
  abords, conteneur, ordures + 4 ponctuelles).
- Boutons **Modifier / Suppr.** et **+ Ajouter une prestation** → écrivent en base
  via l'API. Modale d'édition (code, titre, famille, opérations).

### Équipements / matériel / véhicules
- La vue affiche les **9 équipements réels de la base** sous forme de cartes
  (catégorie, photo si présente).
- **Ajouter / Modifier / Supprimer** un équipement → API. La modale accepte une
  **photo** (multipart, stockée dans `storage/photos/`).

### Paramètres de calcul
- Carte « Paramètres de calcul » alimentée par l'API : taux horaire, TVA,
  coefficients de technicité (standard / technique / haute / exceptionnelle).
- Bouton **Enregistrer les paramètres** → `PUT /parametres/{cle}` (persisté).
- **Utilisés par le moteur de prix** : `pricing.appliquer_parametres()` est appelé
  avant chaque calcul (`/calculer`, `/apercu`, `/generer`). Preuve : en passant
  coef_haute de 1.30 à 1.45, le HT d'une ligne « haute technicité » passe de
  270,40 € à 301,60 €.

---

## Preuve de persistance en base (extrait du test réel)

```
BACKEND READY: True
badge: Connecté au backend
biblio shows Hall: True | shows ops count: True
settings shows Paramètres de calcul: True
create prestation status: 200

=== PREUVE EN BASE (lecture directe du fichier SQLite) ===
prestation persistée: ('test_zone_preuve', 'Zone de preuve')
paramètre persisté: ('taux_horaire_defaut', '29')
total prestations en base: 13   (était 12 avant création)

calc with coef_haute=1.45 -> HT: 301,60 €
PREUVE DB coef_haute: ('coef_haute', '1.45')
```

Le test démarre un vrai `uvicorn`, pilote l'interface réelle via le navigateur
(requêtes HTTP réelles vers le backend), effectue des écritures, puis **relit le
fichier SQLite directement** (indépendamment du serveur) pour prouver la
persistance.

---

## Fichiers modifiés / ajoutés

| Fichier | Changement |
|---|---|
| `backend/pricing.py` | + `appliquer_parametres()` (taux/TVA/coefficients depuis la base) ; `TAUX_HORAIRE_DEFAUT` |
| `backend/routes_devis.py` | + `_charger_parametres(session)` appelé dans `/calculer`, `/apercu`, `/generer` |
| `frontend/static/api-bridge.js` | + rendu API des prestations, équipements et paramètres ; modales create/edit ; override du rendu localStorage du prototype ; re-render à la navigation |
| `frontend/index_rich.html` | régénéré avec le pont mis à jour |

(Les tables `PrestationType`, `Equipement`, `Parametre` et leurs routes CRUD
existaient déjà depuis la Phase 2A initiale ; cette livraison les **branche** à
l'interface.)

---

## CHECKLIST

### ✅ Terminé
- [x] Bibliothèque métier : voir / ajouter / modifier / supprimer les prestations (base réelle)
- [x] Équipements : voir / ajouter / modifier / supprimer (base réelle) + photo à la création
- [x] Paramètres : modifier taux horaire, TVA, coefficients ; sauvegarde en base
- [x] Paramètres réellement utilisés par le calcul du devis (prouvé)
- [x] Prestations de la base utilisables (exposées à la création de devis via `window.ME_ZONES` / `ME_PRESTATIONS`)
- [x] Vérifié avec le vrai backend lancé (captures + preuve SQLite)

### 🟡 Partiel
- [ ] Sélection fine d'une prestation type pour pré-remplir les opérations dans le
      formulaire de devis (les données sont chargées ; l'insertion automatique dans
      le formulaire reste à finaliser)

### 🔴 Phase 2B (à venir)
- [ ] Upload complet des photos depuis l'écran de devis
- [ ] Liaison photo ↔ prestation dans un devis
- [ ] Injection des photos dans le Word/PDF, alignées avec la prestation
- [ ] Génération PDF finale avec photos

---

## Avancement global : ≈ 85 %

Les 3 vues d'administration réclamées sont branchées et vérifiées. Le reste
relève de la Phase 2B (photos dans le document) et d'un raffinement de la
sélection des prestations dans le formulaire de devis.

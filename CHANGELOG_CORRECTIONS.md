# CHANGELOG — Correction des 6 problèmes signalés

Audit et corrections réalisés en local avec le vrai backend (uvicorn + base
SQLite réelle), pilotage de l'interface via navigateur, et vérification du
fichier de base en lecture directe.

---

## Bug 1 — Prestations non cochées qui apparaissent dans l'aperçu/PDF

**Cause** : dans le modèle Word récurrent (Copro), les 8 zones (Hall, ascenseur,
escaliers, caves…) étaient du **texte codé en dur** ; seules les fréquences
étaient des variables. Mettre une fréquence vide n'enlevait pas la zone.

**Correction** : chaque zone est désormais entourée d'un conditionnel
`{%p if SHOW_HALL %} … {%p endif %}`. Le backend met `SHOW_*` à vrai uniquement
pour les zones cochées (`zones_selectionnees`). Les zones non cochées
**disparaissent entièrement** du document.

**Fichiers** : `backend/seed_data/copro_petite.docx` (8 conditionnels),
`backend/routes_devis.py` (champ `zones_selectionnees`, calcul des `SHOW_*`),
`frontend/static/api-bridge.js` (sélecteur de zones + envoi de la sélection).

**Preuve** : `PREUVE-zones-cochees-vs-decochees.png` — à gauche 3 zones cochées
(Hall + Cages d'escaliers + Caves), à droite 1 seule (Caves) : les autres ont
disparu.

## Bug 2 — Photos qui ne suivent pas la sélection

**Cause** : la liste de photos n'était pas correctement reconstruite à chaque
changement et le marqueur restait parfois.

**Correction** : `_resoudre_photos()` reconstruit la liste à chaque génération à
partir des seuls équipements cochés ; regroupement par prestation. Si aucun
équipement coché → aucun bloc photo (et le marqueur est retiré).

**Preuve** : `PREUVE-photos-cochees-vs-decochees.png` — à gauche 3 équipements
(photos affichées et alignées), à droite 0 (aucun bloc, document propre).
Comptage média : 3 équip → 7 images ; 0 équip → 4 (template seul).

## Bug 3 — Aperçu instable / « Bad magic number »

**Causes (deux)** :
1. Écriture concurrente sur le même fichier d'aperçu lors de clics rapides.
2. LibreOffice partage un profil utilisateur unique → les conversions
   simultanées échouaient (500 intermittent), renvoyant un corps non-PDF que
   l'iframe interprétait comme « Bad magic number ».
3. Bonus : `mapDevisToPayload()` pouvait renvoyer `null` (→ 422) quand aucun type
   de devis n'était encore choisi.

**Corrections** :
- Backend : nom de fichier d'aperçu **unique par requête** ; conversion
  LibreOffice **sérialisée par verrou** + **profil temporaire dédié** par appel.
- Frontend : **garde de séquence** (seule la dernière requête s'applique),
  **validation du type PDF** de la réponse, message clair si indisponible.
- `mapDevisToPayload()` déduit le type (zones cochées ⇒ récurrent) et ne renvoie
  plus `null`.

**Preuve** : 6 requêtes d'aperçu concurrentes → toutes 200 avec PDF valide.

**Fichiers** : `backend/template_service.py` (verrou + profil),
`backend/routes_devis.py` (fichier unique), `frontend/static/api-bridge.js`.

## Bug 4 — Méthode de calcul

Formule confirmée (voir METHODE_CALCUL.md) :

```
HT par passage = heures × agents × taux_horaire × coef_technicité + frais
Récurrent : HT mensuel = HT par passage × passages_par_mois
Ponctuel  : HT = HT par passage (forfait)
TVA = HT × taux_tva     TTC = HT + TVA
Correction manuelle : si prix_force_ht fourni, il remplace le HT calculé.
```

Coefficients : standard 1,00 · technique 1,15 · haute 1,30 · exceptionnelle 1,50.
Passages/mois : 1×/sem = 4,333 · 2×/sem = 8,667 · 3×/sem = 13 · 1×/mois = 1.

Exemples vérifiés :
- Récurrent 2h × 1 agent × 24 € × 1,0 × 4,333 = **208,00 € HT** → 249,60 € TTC.
- Récurrent 3h × 2 × 25 € × 1,15 × 8,667 = **1 495,00 € HT** → 1 794,00 € TTC.
- Ponctuel 8h × 2 × 28 € × 1,0 + 150 € = **598,00 € HT** → 717,60 € TTC.
- Ponctuel 5h × 3 × 30 € × 1,30 + 200 € = **785,00 € HT** → 942,00 € TTC.

## Bug 5 — Base clients

Vérifié : création persistée en base (lecture directe SQLite), mapping correct
vers le Word (DEST_LIGNE1 = civilité + contact, DEST_LIGNE2 = raison sociale,
DEST_LIGNE3 = adresse, DEST_LIGNE4 = CP + ville), **aucun « undefined »**
(toutes les valeurs absentes deviennent une chaîne vide).

## Bug 6 — Matériel / équipements avec photos

Vérifié : 9 équipements en base avec photos, CRUD complet (création/modification/
suppression/désactivation, remplacement de photo), sélection dans le devis,
injection correcte dans le PDF (alignée, plusieurs photos par prestation).

---

## Tests effectués (avec le vrai backend)

1. Génération récurrent avec 3 zones cochées vs 1 zone → seules les cochées
   apparaissent.
2. Génération ponctuel avec 3 équipements vs 0 → photos présentes vs absentes.
3. 6 aperçus PDF concurrents → tous valides (plus de 500/Bad magic number).
4. 4 exemples de calcul vérifiés (récurrent et ponctuel).
5. Cycle CRUD client + lecture directe de la base SQLite.
6. Boot complet de l'application depuis une base vierge.

## Livrables

- ZIP corrigé du projet
- PDF test récurrent (PDF-RECURRENT-3zones.pdf, PDF-RECURRENT-1zone.pdf)
- PDF test ponctuel avec photos (PDF-PONCTUEL-PHOTOS.pdf) et sans (PDF-PONCTUEL-SANS-PHOTOS.pdf)
- Preuves visuelles cochées/décochées (zones et photos)
- Méthode de calcul détaillée (METHODE_CALCUL.md)
- Ce changelog

## Avancement global : ≈ 97 %

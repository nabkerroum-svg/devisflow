# CHANGELOG — Options par zone + sauts de page

## 1. Sauts de page des grandes sections — FAIT
« 1 - Présentation de la société » et « 3 - Prestations complémentaires »
commencent désormais TOUJOURS en haut de leur propre page.
Méthode : propriété `pageBreakBefore` ajoutée au paragraphe de titre de chaque
section. Aucun autre élément n'est déplacé — c'est une simple propriété de
paragraphe, la mise en page du modèle reste intacte.
Preuve : `PREUVE-sauts-de-page.png` (les deux sections démarrent en haut de page).

## 2 & 5. Tableau financier dynamique — FAIT
Avant : le tableau affichait systématiquement 3 options de fréquence (1×/2×/3×
par semaine), même si elles n'existaient pas. CORRIGÉ : le tableau est construit
UNIQUEMENT à partir des zones réellement cochées et de leurs options activées.
- 1 ligne par zone cochée (titre + fréquence renseignée + prix)
- 1 ligne par option activée (désignation + fréquence + prix)
- rien d'autre n'apparaît.
La fréquence affichée correspond exactement à celle saisie pour chaque zone.

## 3 & 4. Options par zone — FAIT
Dans le panneau dépliable de chaque zone (icône d'expansion existante), une case
discrète « + Activer une option complémentaire ». Quand on l'active, 3 champs
apparaissent : désignation, fréquence (menu), prix HT. Replié par défaut →
l'écran reste simple. Disponible pour toutes les zones (hall, ascenseur, caves,
garage, local poubelles, abords, etc.).

## 6. Vérifié (génération réelle)
- Zone simple sans option → 1 ligne, pas d'option fantôme.
- Zone avec option → 2 lignes (zone + option), prix de l'option respecté.
- 3 zones, fréquences différentes, 1 option → 4 lignes exactes.
- Aperçu = PDF = DOCX (DOCX « Microsoft Word 2007+ »).
- Sauts de page en place.

## Modèle Word
Seuls ajouts au modèle : 2 propriétés `pageBreakBefore` (sur les titres des 2
grandes sections). Aucun élément déplacé, aucune marge/espacement/image modifié.
La boucle du tableau financier (déjà présente dans le modèle) est désormais
alimentée dynamiquement.

## Fichiers modifiés
- `backend/seed_data/copro_petite.docx` — `pageBreakBefore` sur les 2 titres
- `backend/pricing.py` — `construire_tableau_zones()` (tableau dynamique) + helper
- `backend/routes_devis.py` — champ `zones_detail` + utilisation du tableau dynamique
- `frontend/standalone/devisflow.html` — UI option par zone (case + 3 champs) + handlers
- `frontend/static/api-bridge.js` — envoi de `zones_detail`, suppression des
  options de fréquence codées en dur
- `frontend/index_rich.html` — régénéré

## Vérification chez vous
`http://localhost:8000/api/version` → tout `true`.
Test : Contrat récurrent → modèle → cocher une zone → la déplier → activer une
option → renseigner désignation/fréquence/prix → aperçu → PDF → DOCX.

## Lancement
```bash
docker compose down -v
docker compose up --build
```

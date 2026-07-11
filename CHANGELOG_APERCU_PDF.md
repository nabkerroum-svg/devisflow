# CHANGELOG — Correctif : aperçu = vrai PDF du modèle Word

## Problème corrigé

L'**aperçu** affiché dans le panneau de droite de l'interface était une
**reconstruction HTML** du devis (classes `.pv-*`). Ce HTML appliquait sa propre
mise en page — logo repositionné, bloc « Depuis 1975 » stylisé, titres
différents — et affichait « undefined undefined » sur des champs mal nommés. Il
ne ressemblait donc pas au modèle Word.

Important : le **document généré** (.docx/.pdf) était, lui, déjà strictement
fidèle au modèle. Le défaut venait uniquement de l'aperçu, qui montrait une
maquette au lieu du vrai document.

## Solution

L'aperçu HTML reconstruit est **supprimé** et remplacé par l'**aperçu du vrai
PDF**, généré par le backend à partir du modèle Word maître (seules les
variables sont remplacées). C'est désormais la seule source de vérité affichée.

### Backend — `routes_devis.py`
- Nouvelle route **`POST /api/devis/apercu`** : remplit le modèle Word avec les
  variables, le convertit en PDF (LibreOffice) et le renvoie directement, **sans
  le persister** comme devis définitif. Dossier de travail : `storage/generated/_apercus/`.
- Refactorisation : la construction des variables (OPTIONS, PRESTATIONS,
  FORFAIT_*) est extraite dans `_construire_data()`, partagée par l'aperçu et la
  génération définitive — garantissant un aperçu identique au document final.

### Frontend — `static/api-bridge.js`
- `updatePreview()` est redéfini : il n'injecte plus de HTML reconstruit. Il
  appelle `/api/devis/apercu` (anti-rebond 700 ms) et affiche le **PDF réel dans
  un `<iframe>`**.
- `mapDevisToPayload()` convertit l'état du devis (forme prototype) vers le
  format attendu par l'API, **sans jamais produire « undefined »** (toutes les
  valeurs absentes deviennent une chaîne vide ; les lignes du destinataire sont
  construites proprement à partir de civilité + contact + adresse + ville).
- Bouton **« Rafraîchir l'aperçu »** ajouté ; le toggle « Rendu finalisé »
  (spécifique à l'ancien rendu HTML) est masqué.
- Repli clair si le backend est absent : message invitant à lancer l'application,
  plutôt qu'un faux rendu.

## Vérification (preuve)

Comparaison côte à côte **modèle Word original vs aperçu PDF** : cover page
strictement identique (logo, bandeau, images sépia, badges, pied de page,
polices, marges). Seuls les champs variables diffèrent (date, destinataire,
numéro de proposition). Plus aucun « undefined ». Capture : `apercu-vs-modele.png`.

L'aperçu dans l'interface charge bien ce PDF via un `<iframe>` (blob), vérifié en
navigateur.

## Fichiers modifiés

| Fichier | Changement |
|---|---|
| `backend/routes_devis.py` | + route `/devis/apercu` (PDF réel) ; helper `_construire_data()` |
| `frontend/static/api-bridge.js` | aperçu PDF réel + mapper anti-undefined + bouton rafraîchir |
| `frontend/index_rich.html` | régénéré avec le pont mis à jour |

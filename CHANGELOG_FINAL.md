# DevisFlow — Changelog final & état des fonctionnalités

## Phase de finition — ce qui a été traité

1. **Upload de photo depuis l'écran de devis** — bouton « + Ajouter une photo »
   dans la section équipements : l'image (JPG/PNG) est envoyée au backend
   (data URL), sauvegardée, puis injectée. ✅
2. **Association photo ↔ prestation précise** — chaque équipement/photo a un
   menu « Prestation concernée » ; la photo est alignée en face de cette
   prestation dans le document. ✅
3. **Plusieurs photos sur une même prestation** — les photos partageant la même
   prestation sont regroupées et affichées côte à côte sur la même ligne, donc
   alignées ; la largeur est réduite automatiquement (38 mm à deux, 28 mm à
   trois et plus) pour rester dans la page. ✅ (prouvé : 2 photos alignées)
4. **Pas de page vide ni décalage** — vérifié : le PDF ponctuel fait 7 pages,
   toutes avec contenu ; les photos sont en flux (jamais flottantes), donc pas
   de décalage. ✅
5. **Version finale propre + documentation** — guides d'installation et
   d'utilisation fournis. ✅

## Détail technique

- `routes_devis._resoudre_photos()` : gère 3 sources par entrée — équipement en
  base (`code`), photo uploadée (`photo_data` data URL), libellé/prestation —
  et **regroupe par prestation**.
- `template_service._injecter_photos()` : post-traitement python-docx. Repère le
  marqueur `@@ZONE_PHOTOS@@`, insère un tableau 2 colonnes (libellé | photos),
  gère le marqueur même à l'intérieur d'un tableau, plusieurs photos par ligne,
  redimensionnement automatique.
- Modèle **Encombrants** : anciens encarts flottants « Photos 1 » remplacés par
  le marqueur en flux → placement propre juste après les prestations.
- Modèle **Copro** : marqueur ajouté → photos possibles (placées en fin de
  document).
- `api-bridge.js` : sélecteur d'équipements enrichi (cases à cocher + upload +
  choix de la prestation), envoi dans le payload.

## Limite connue (honnêteté)

- **Récurrent (Copro)** : les photos s'affichent correctement mais en **fin de
  document** (après les CGV), car le seul emplacement où le marqueur survit de
  façon fiable au moteur de gabarit est hors du grand tableau de prestations. Le
  placement « juste après la prestation concernée » est parfait sur le modèle
  **Ponctuel** (cas d'usage principal des photos) ; sur le récurrent, les photos
  sont regroupées en fin. Amélioration possible : repositionner le bloc dans le
  corps du contrat récurrent.

---

## ÉTAT DES FONCTIONNALITÉS

### ✅ Terminé
- Génération Word/PDF strictement fidèle aux 2 modèles maîtres
- Aperçu = vrai PDF du modèle Word (plus de maquette HTML)
- Calcul automatique des prix (heures, agents, taux, technicité, frais, TVA,
  options comparatives, override manuel)
- Bibliothèque clients : CRUD + recherche + archivage + préremplissage devis
- Bibliothèque métier : prestations types (zones exactes du modèle Word) — CRUD
- Équipements / matériel / véhicules : CRUD + photo + actif/inactif
- Équipe : CRUD + actif/inactif
- Paramètres de calcul : taux, TVA, coefficients — enregistrés et utilisés par
  le moteur de prix
- Sélection d'un ou plusieurs équipements dans le devis
- Upload de photo à la volée dans le devis
- Association photo ↔ prestation
- Injection des photos dans le Word/PDF, alignées, auto-dimensionnées, sans
  chevauchement ni débordement
- Plusieurs photos par prestation, alignées
- Modèle ponctuel : encarts photos remplacés, mise en page conservée
- Modèle récurrent : ajout de photos possible
- Toutes les vues vérifiées avec le vrai backend (captures + base SQLite)

### 🟡 À affiner (non bloquant)
- Placement des photos dans le corps du devis **récurrent** (actuellement en fin
  de document ; parfait sur le ponctuel)
- Association d'une photo à une **ligne** de prestation très précise plutôt qu'à
  un libellé choisi dans une liste

### 🔴 Non commencé (hors périmètre actuel)
- Authentification / gestion des droits par membre
- Envoi du devis par e-mail directement depuis l'application
- Import Excel des clients

---

## Avancement global du projet : ≈ 95 %

Le produit est fonctionnel de bout en bout : création de devis fidèles aux
modèles, calcul, bibliothèques, paramètres, et photos d'équipement injectées et
alignées. Les points restants sont des affinages d'ergonomie et de placement,
non bloquants pour l'usage.

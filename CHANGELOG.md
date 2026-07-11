# CHANGELOG — Reprise du moteur DevisFlow (Phase 1)

Cette phase rend les **deux modèles maîtres** Marie Eugénie pleinement
fonctionnels et ajoute le **calcul automatique des prix**. Les photos
d'équipement sont planifiées en phase 2 (l'infrastructure est en place).

---

## 1. Modèles maîtres ré-annotés (le cœur du sujet)

Vos deux `.doc` officiels ont été convertis en `.docx` puis annotés **au niveau
du XML**, sans jamais reconstruire la mise en page : logos, polices Arial,
marges 25 mm, images, badges (FEP, Palme Verte, Petit Futé), encadrés et pieds
de page sont conservés **à l'identique**. Seules les zones variables sont
devenues des marqueurs Jinja2, et les surlignages jaunes/verts de repérage ont
été retirés du rendu final.

### `seed_data/copro_petite.docx` (récurrent — 20 variables)
- Page de garde : `DATE_EMISSION`, `NUMERO_DEVIS`, `DEST_LIGNE1..4`,
  `TYPE_PRESTATION`, `SITE_ADRESSE`, `SITE_CP_VILLE`
  - **Correction du bug « contact »** : la 4ᵉ ligne destinataire
    (`DEST_LIGNE4`), située dans un hyperlien, n'était jamais remplacée. Elle
    fonctionne désormais.
- Détail des prestations : les 8 « Fréquence : XXX » deviennent `FREQ_HALL`,
  `FREQ_ASCENSEUR`, `FREQ_ESCALIERS`, `FREQ_CAVES`, `FREQ_GARAGE`,
  `FREQ_ABORDS`, `FREQ_CONTENEUR`, `FREQ_OM`
- Proposition financière : le tableau de prix devient une **boucle multi-options**
  (`{%tr for opt in OPTIONS %}`) → 1 à N options comparatives chiffrées
- Durée / prise d'effet : `DATE_PRISE_EFFET`, `DUREE_CONTRAT`

### `seed_data/ponctuel_generique.docx` (ponctuel — 12 variables)
- Page de garde / titre : `DATE_EMISSION`, `NUMERO_DEVIS`, `DEST_LIGNE1..3`,
  `TYPE_PRESTATION`, `NOM_OPPORTUNITE`
- Détail des prestations : les 6 lignes en dur sont remplacées par une **boucle**
  (`{%p for p in PRESTATIONS %}`) → la liste s'adapte aux prestations choisies
- Tableau financier : `FORFAIT_LIBELLE`, `FORFAIT_HT`, `FORFAIT_TVA`,
  `FORFAIT_TTC`
- Les deux cadres « Photos 1 » sont **laissés en place** (réservés à la phase 2)

> Note technique : dans `docxtpl` 0.20.2, la répétition d'une ligne de tableau
> exige que `{%tr for %}` et `{%tr endfor %}` soient placés dans des **lignes
> d'encadrement séparées** autour de la ligne à répéter (et non en ligne avec
> les données). L'ancien moteur ne pouvait donc pas produire un tableau
> multi-options : cette partie n'avait jamais fonctionné.

---

## 2. Nouveau moteur de calcul des prix — `backend/pricing.py` (créé)

Calcul par ligne :
```
HT_par_passage = durée_h × nb_agents × taux_horaire × coef_technicité + frais
récurrent → HT_mensuel = HT_par_passage × passages_par_mois (fréquence)
ponctuel  → HT_forfait = somme des lignes
TVA = HT × taux_tva   ;   TTC = HT + TVA
```
- **Barème de technicité** modifiable (`standard` 1.00, `technique` 1.15,
  `haute` 1.30, `exceptionnelle` 1.50) **+ override** numérique libre par ligne
- **Fréquences → passages/mois** (1×/sem ≈ 4,33, etc.)
- **Override du total HT** : si fourni, prime sur le calcul ; la TVA/TTC est
  recalculée et un indicateur `override` signale l'écart avec le calcul auto
- Formatage français des montants : `1 234,56 €`
- `construire_options_recurrentes()` : génère les options comparatives du tableau
  récurrent à partir d'une ligne de base et d'une liste de fréquences

---

## 3. Backend — fichiers modifiés

### `models.py`
- `Devis` : ajout de `montant_tva`, `montant_ttc`, `prix_override`
- **Nouvelle table `TemplateSubstitution`** : table de substitution par modèle,
  stockée en base → importer un nouveau modèle sans modifier le code

### `template_service.py`
- **`convertir_en_docx()`** (nouveau) : conversion `.doc → .docx` via LibreOffice
  (permet d'importer directement vos `.doc`)
- **`_replace_in_paragraph()`** fiabilisé : gère les runs fragmentés par Word et
  les espaces insécables (origine du bug « contact »)
- **`_coerce_substitutions()`** : accepte des tuples **ou** des objets DB
- Détection de variables étendue aux **boucles** (`OPTIONS`, `PRESTATIONS`…)
- `generer_devis()` : les variables de type liste reçoivent `[]` par défaut
  (et non `""`) pour ne pas casser les boucles

### `routes_devis.py`
- **Correction du bug** ligne 48 (`... if False else ...`) — requête nettoyée
- Payload enrichi : `lignes`, `frequences_options`, `prix_force_ht`, `taux_tva`
- **`POST /api/devis/calculer`** (nouveau) : aperçu prix temps réel sans générer
- `POST /api/devis/generer` : intègre le calcul, construit `OPTIONS` /
  `PRESTATIONS` / `FORFAIT_*`, persiste HT/TVA/TTC + override

### `routes_templates.py`
- `POST /api/templates/upload` accepte désormais **`.doc`** (conversion auto)

### `seed.py`
- Charge **les deux** modèles maîtres au premier démarrage (au lieu d'un seul)

---

## 4. Frontend — fichiers modifiés

### `index.html`
- Nouvelle carte « Prestations & tarification automatique » : éditeur de lignes
  (heures, agents, taux, technicité, frais), options de fréquence comparatives,
  aperçu HT/TVA/TTC et champ d'override du total
- Upload de modèle : accepte `.doc` et `.docx`

### `static/app.js`
- Éditeur de prestations dynamique (ajout/suppression de lignes)
- Aperçu prix temps réel (`/devis/calculer`)
- Payload de génération mis à jour (lignes, fréquences, override)

### `static/styles.css`
- Styles de l'éditeur de prestations et du bloc de tarification

---

## 5. Nouveaux fichiers / dossiers

| Chemin | Rôle |
|---|---|
| `backend/pricing.py` | Moteur de calcul des prix |
| `backend/seed_data/copro_petite.docx` | Modèle récurrent ré-annoté |
| `backend/seed_data/ponctuel_generique.docx` | Modèle ponctuel ré-annoté |
| `backend/storage/photos/` | Réservé aux photos d'équipement (phase 2) |
| `CHANGELOG.md` | Ce document |

---

## 6. Tests réalisés (rendus PDF vérifiés)

- Devis récurrent : page de garde + tableau **3 options** (208 / 416 / 624 €/mois)
  calculées automatiquement, date de prise d'effet et durée injectées
- Devis ponctuel : liste de prestations dynamique + forfait calculé
  (HT/TVA/TTC), signature et CGV conservées
- Endpoints HTTP `/calculer` et `/generer` validés
- Upload d'un `.doc` original : conversion `.docx` automatique confirmée
- Override du total HT : recalcul TVA/TTC vérifié

---

## 7. Phase 2 (photos d'équipement) — approche prévue

Insertion via `docxtpl.InlineImage` (images **en flux**, pas flottantes), dans
une **ligne de tableau invisible à 2 colonnes** (texte à gauche, photo à droite)
pour garantir l'alignement vertical avec la prestation et empêcher les photos de
« remonter » ou de chevaucher le texte. Redimensionnement automatique par Pillow
à largeur fixe (≈ 55 mm), compatible Word **et** PDF.

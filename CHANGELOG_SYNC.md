# DevisFlow 11 — Synchronisation interface ↔ devis

Correctifs ciblés sur la liaison entre l'interface, l'aperçu, le DOCX et le PDF.
La source de vérité reste l'état courant de l'interface (`state.devis`), lu à
chaque rafraîchissement de l'aperçu et envoyé tel quel au backend.

## Bugs corrigés

### 1. Prix jamais transmis (point 5)
`mapDevisToPayload` lisait `dd.totaux.force_ht`, une clé inexistante : le
Montant HT saisi n'arrivait jamais au backend (`prix_force_ht` = `null`), donc le
devis affichait toujours le prix calculé automatiquement.
- Le Montant HT global est désormais transmis comme `prix_force_ht`.
- Ajout d'un champ **« Prix HT mensuel »** par zone (`zones_detail[].prix_ht`),
  mécanisme adapté au tableau financier multi-zones de la copro. Vide ⇒ calcul auto.

Fichiers : `frontend/index_rich.html`, `frontend/static/api-bridge.js`.

### 2. Aperçu instable — récursion infinie (points 7 & résultat attendu)
`updatePreview()` appelait `recalcTotaux()`, qui rappelle `updatePreview()` :
boucle infinie rattrapée par un `try/catch`, d'où des mises à jour aléatoires.
- La boucle est supprimée ; `recalcTotaux` (déclenché par le champ HT) appelle
  `updatePreview`, et plus l'inverse. Anti-rebond ramené à 600 ms.

### 3. Retours à la ligne écrasés (point 6)
- **Éditeur visuel** (`/templates/{code}/contenu`) : la sauvegarde écrivait le
  texte d'un bloc dans un seul run, ce qui faisait disparaître les `\n`. Ils sont
  désormais convertis en vrais sauts de ligne Word (`<w:br/>`).
- **Génération du devis** : toute valeur texte multi-lignes est enveloppée dans un
  `docxtpl.Listing`, garantissant la conversion `\n` → `<w:br/>` et la
  préservation des espaces, dans l'aperçu comme dans le DOCX/PDF.

Fichiers : `backend/routes_templates.py`, `backend/template_service.py`.

### 4. Diagnostic enrichi (point 9)
Le bouton « Diagnostic devis » affiche maintenant explicitement le payload réel :
`template_code`, `client_nom`, `adresse`, `site_adresse`, `numero_devis`,
`prix_force_ht`, `prix_ht`, `prix_ttc` — en plus des zones cochées, FREQ_*/SHOW_*
et `zones_detail`. Permet de voir d'un coup d'œil si un écart vient du frontend
ou du backend.

## Vérifié (tests automatisés, backend réel + LibreOffice)

- Test 1 — Client : nom, contact, adresse et n° de devis présents dans le DOCX.
- Test 2 — Zones : seules les zones cochées apparaissent (SHOW_* + opérations) ;
  l'ajout/retrait d'une zone se répercute immédiatement.
- Test 3 — Fréquences : chaque FREQ_* rendue sous la bonne zone ; tableau OPTIONS
  cohérent.
- Test 4 — Retours à la ligne : `\n` → `<w:br/>` dans le devis ET dans l'éditeur.
- Test 5 — Prix : override par zone (180 € / 45 €) et global ponctuel (350 €)
  repris dans le tableau, le DOCX et le PDF.
- Test 6 — Pipeline complet DOCX → PDF (10 pages, sans page blanche).

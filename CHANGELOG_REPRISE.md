# CHANGELOG — Reprise sur la base Devis_flow_2 (version claire)

Base de travail : votre fichier `Devis_flow_2.zip` (présentation que vous trouvez
claire). J'ai gardé cette présentation et appliqué UNIQUEMENT des correctifs ciblés.

## 1. Présentation conservée
La section « Prestations » native (zones avec fréquence à droite, détail des
opérations) est conservée telle quelle. J'ai **retiré le panneau de zones en
double** que cette base affichait encore en haut (« Zones du contrat ») — il
créait de la confusion. Il ne reste que la section native, claire.

## 2. Zones cochées qui n'apparaissaient pas → CORRIGÉ
Cause : les codes de zones du prototype (`hall_entree`, `cabine_ascenseur`…) ne
correspondaient pas aux codes du modèle Word (`hall`, `ascenseur`…). La sélection
n'était donc pas reconnue. Ajout d'une table de correspondance + lecture de la
sélection native. Chaque zone cochée envoie sa fréquence propre.

## 3. Grands blancs / pages vides → CORRIGÉ
Ajout du compactage des paragraphes vides après génération : les zones décochées
disparaissent sans laisser de grands espaces.

## 4. Aperçu → CORRIGÉ (affiche le vrai contenu)
L'aperçu est rendu avec PDF.js (dessin sur canvas), embarqué localement
(`frontend/static/vendor/`). Il affiche exactement le PDF final : zones cochées
avec détail + fréquence, zones décochées absentes. Vérifié : aperçu = PDF = DOCX.

## 5. Photos par zone (simple) → EN PLACE
Le système natif de photo par zone est conservé et branché : bouton « Ajouter une
photo » à côté de chaque zone, miniature, suppression, **redimensionnement
automatique** (max 1200 px, proportions conservées, JPEG qualité 85). Les photos
sont injectées à proximité de la zone via le marqueur prévu du modèle.

## 6. Modèle Word : INCHANGÉ
Aucun élément non surligné n'est modifié. Le modèle reste la référence : logo,
images, titres, marges, espacements, pagination, blocs fixes, pieds de page —
tout est figé. Seules les zones dynamiques (variables jaunes) sont remplies.

## Vérification infaillible
Après démarrage, ouvrez : `http://localhost:8000/api/version`
Tout doit être `true` (et le template `8`). Si un bug visuel persiste alors que
`tout_ok` = true → cache navigateur : Ctrl+Shift+R.

## Fichiers modifiés (par rapport à la base Devis_flow_2)
- `frontend/static/api-bridge.js` — mapping des codes de zones + fréquence par
  zone ; aperçu PDF.js ; photos de zone envoyées au backend ; panneau doublon
  désactivé
- `frontend/static/vendor/pdf.min.js` + `pdf.worker.min.js` — ajout (PDF.js local)
- `frontend/index_rich.html` — régénéré
- `backend/template_service.py` — compactage des paragraphes vides
- `backend/main.py` — route de diagnostic `/api/version`
- `docker-compose.yml` — montage du code en volume (toujours à jour)

(Le modèle `copro_petite.docx` de cette base n'a PAS été modifié.)

## Test (depuis l'interface, http://localhost:8000)
1. « Contrat récurrent » → modèle « Copropriété — petite ».
2. Section Prestations : cocher Hall d'entrée, choisir sa fréquence.
3. (option) Ajouter une photo à la zone.
4. L'aperçu à droite montre Hall + détail + fréquence (+ photo).
5. « Générer le PDF » / « Télécharger .docx » → identiques à l'aperçu.
6. Refaire avec Hall + Ascenseur + Cages, 3 fréquences différentes.

## Lancement
```bash
docker compose down -v
docker compose up --build
```

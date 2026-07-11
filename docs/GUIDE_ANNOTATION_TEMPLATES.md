# Guide d'annotation des templates Word

Ce document explique comment préparer un fichier Word (`.docx`) pour qu'il devienne un template DevisFlow utilisable.

## Principe

Un template DevisFlow est un fichier `.docx` normal dans lequel les zones variables ont été remplacées par des marqueurs au format Jinja2 :

```
{{ NOM_VARIABLE }}
```

À la génération, le système substitue ces marqueurs par les valeurs fournies, sans toucher au reste du document (logos, polices, couleurs, mise en page).

## Règles d'écriture des marqueurs

| Règle | Pourquoi | Exemple |
|---|---|---|
| **En MAJUSCULES** avec underscores | Éviter les conflits avec le texte français normal | `{{ NUMERO_DEVIS }}` |
| **Pas d'accents** | Compatibilité Jinja2 | `{{ DATE_EMISSION }}` |
| **Espaces autour** des barres `{{ }}` | Lisibilité, marges de sécurité | `{{ CLIENT }}` ✅ pas `{{CLIENT}}` |
| **Syntaxe valide** : commence par lettre majuscule, suit avec lettres/chiffres/underscores | Regex de détection du système | `{{ DEST_LIGNE_1 }}` ✅ pas `{{ 1_DEST }}` |

## Méthode recommandée pour préparer un template

### Méthode 1 : Annotation manuelle dans Word (la plus fiable)

1. **Ouvrir votre devis** dans Microsoft Word
2. **Identifier chaque zone variable** (généralement en surlignage jaune dans le modèle Marie Eugénie original)
3. **Sélectionner la zone variable entière** (le mot/groupe complet)
4. **Supprimer-la complètement** (touche `Suppr`)
5. **Retaper d'une seule traite** le marqueur, par exemple : `{{ NUMERO_DEVIS }}`

> ⚠️ **Important** : ne PAS copier-coller depuis un autre endroit, ne PAS utiliser l'auto-complétion Word. Taper le marqueur en une seule fois pour éviter que Word ne le fragmente en plusieurs runs (ce qui casse la détection Jinja2).

6. **Sauvegarder en `.docx`** (pas `.doc`)
7. Importer dans DevisFlow via l'onglet « Modèles PDF »

### Méthode 2 : Annotation automatique (rapide mais limitée)

Si vous fournissez un devis **non encore annoté** au format proche du devis Marie Eugénie original, le système peut tenter une annotation automatique en cochant la case **« Annoter automatiquement »** dans le formulaire d'upload.

La table de substitution appliquée se trouve dans `backend/template_service.py` :

```python
DEFAULT_SUBSTITUTIONS = [
    ("Marseille, le 21 juin 2022",       "Marseille, le {{ DATE_EMISSION }}", False),
    ("client",                            "{{ DEST_LIGNE1 }}",                  True),
    ("adresse 1",                         "{{ DEST_LIGNE2 }}",                  False),
    ("adresse 2",                         "{{ DEST_LIGNE3 }}",                  False),
    ("contact",                           "{{ DEST_LIGNE4 }}",                  True),
    ("Proposition ME",                    "Proposition {{ NUMERO_DEVIS }}",     False),
    ("Entretien des parties communes",    "{{ TYPE_PRESTATION }}",              False),
    ("52 rue Louis Astruc,",              "{{ SITE_ADRESSE }},",                False),
    ("13005 MARSEILLE.",                  "{{ SITE_CP_VILLE }}.",               False),
]
```

Pour adapter cette table à vos propres devis, modifier ce fichier et relancer le serveur.

## Variables conventionnelles

Pour une cohérence à travers vos templates, utilisez ces noms standard :

| Marqueur | Usage | Exemple de valeur |
|---|---|---|
| `{{ NUMERO_DEVIS }}` | Numéro unique du devis | `ME-6245` |
| `{{ DATE_EMISSION }}` | Date d'émission lisible | `19 juin 2026` |
| `{{ TYPE_PRESTATION }}` | Type de la prestation | `Entretien des parties communes` |
| `{{ DEST_LIGNE1 }}` | Destinataire — ligne 1 (Civilité + Nom) | `Madame Sophie MARCHAND` |
| `{{ DEST_LIGNE2 }}` | Destinataire — ligne 2 (Raison sociale) | `Syndic Foncia Marseille` |
| `{{ DEST_LIGNE3 }}` | Destinataire — ligne 3 (Adresse) | `12 rue Paradis` |
| `{{ DEST_LIGNE4 }}` | Destinataire — ligne 4 (CP + Ville) | `13001 Marseille` |
| `{{ SITE_ADRESSE }}` | Adresse du site d'intervention | `52 rue Louis Astruc` |
| `{{ SITE_CP_VILLE }}` | CP + ville du site | `13005 Marseille` |
| `{{ MONTANT_HT }}` | Montant HT | `180.00 €` |
| `{{ MONTANT_TVA }}` | Montant TVA | `36.00 €` |
| `{{ MONTANT_TTC }}` | Montant TTC | `216.00 €` |
| `{{ CONDITIONS_REGLEMENT }}` | Modalités de paiement | `Mensuel, à 30 jours fin de mois` |

## Syntaxes avancées (pour évolutions futures)

`docxtpl` (la bibliothèque utilisée) supporte des syntaxes plus puissantes. Elles ne sont pas encore branchées dans la v1 mais le seront dans les versions futures.

### Boucles `{% for %}`

Permettent de répéter un bloc Word pour chaque élément d'une liste, ex : pour les zones d'un contrat.

```
{%tr for zone in zones %}
| {{ zone.titre }} |
| Fréquence : {{ zone.frequence }} |
{%tr endfor %}
```

> Le préfixe `{%tr` (au lieu de `{%`) indique que le contrôle est sur une ligne de tableau. Idem `{%p` pour un paragraphe.

### Conditions `{% if %}`

```
{% if site.code_acces %}
Code d'accès : {{ site.code_acces }}
{% endif %}
```

### Images dynamiques

```
{{ photo_site }}
```

Où `photo_site` est un `InlineImage` côté backend :
```python
from docxtpl import InlineImage
from docx.shared import Mm
data = {
    "photo_site": InlineImage(doc, "site.jpg", width=Mm(80))
}
```

## Vérifier qu'un template est correctement annoté

Après upload, l'onglet « Modèles PDF » affiche le nombre de variables détectées. Vous pouvez aussi utiliser le bouton **« Aperçu PDF »** qui montre le template **avec les marqueurs visibles** — pratique pour vérifier qu'ils ne sont pas fragmentés.

Pour vérifier en ligne de commande :

```bash
cd backend
python -c "
from pathlib import Path
from template_service import analyser_template
vars = analyser_template(Path('storage/templates/mon_template.docx'))
print('Variables détectées :', vars)
"
```

## Problèmes courants

### Le marqueur `{{ X }}` apparaît tel quel dans le devis généré

Cause probable : Word a fragmenté le marqueur en plusieurs runs internes (ex : `{{`, ` NUMERO_DEVIS `, ` }}`).

**Solution** : ouvrir le `.docx` dans Word, supprimer le marqueur entier et le retaper d'un seul jet sans correction automatique.

Pour diagnostiquer :
```bash
cd backend
python -c "
from pathlib import Path
from template_service import analyser_template
print(analyser_template(Path('storage/templates/mon_template.docx')))
"
```
Si le marqueur attendu n'apparaît pas dans la liste, il est fragmenté.

### Une zone Word est dans une « zone de texte » (text frame)

Les text frames Word (zones de texte ancrées flottantes, souvent utilisées pour positionner du contenu hors flux normal) ne sont **pas accessibles** via `doc.paragraphs`. Les marqueurs Jinja2 qui s'y trouvent ne seront pas détectés ni substitués.

**Solution** : retaper la zone variable dans un paragraphe normal (pas une zone de texte). Si la mise en page exige une zone de texte, en discuter avec un développeur pour traiter le XML brut.

### Le PDF généré a une mise en page différente de votre Word d'origine

Cause probable : LibreOffice manque certaines polices.

**Solution** : installer les polices Microsoft sur le serveur :
```bash
sudo apt install fonts-liberation fonts-dejavu ttf-mscorefonts-installer
fc-cache -f -v
```

Ou : enregistrer le `.docx` dans Word avec les polices incorporées (Fichier → Options → Enregistrement → cocher « Incorporer les polices dans le fichier »).

## Aller plus loin

Documentation officielle docxtpl : https://docxtpl.readthedocs.io/en/latest/

Pour des cas complexes (tableaux dynamiques, fusions de cellules, en-têtes par section), c'est la référence à consulter.

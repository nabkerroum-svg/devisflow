# Analyse typographique du devis Marie Eugénie

> Valeurs extraites par analyse XML directe du fichier `1_COPRO_petite.doc`
> (converti en `.docx` puis décompressé). Ces valeurs servent de référence
> pour caler le gabarit « Classique Marie Eugénie » sur le rendu réel.

## Métadonnées du document

- **Format source** : Microsoft Word 97-2003 (`.doc`)
- **Conversion** : LibreOffice headless → `.docx` (Office Open XML)
- **Taille originale** : 3,7 Mo (avec 16 images embarquées)
- **Total paragraphes** : 431
- **Date d'analyse** : juin 2026

## Marges de page

Extrait de `<w:pgMar>` dans `word/document.xml` :

| Marge | Valeur en twips | Valeur en mm |
|---|---|---|
| Haut (`w:top`) | 1418 | **25 mm** |
| Bas (`w:bottom`) | 1418 | **25 mm** |
| Gauche (`w:left`) | 1418 | **25 mm** |
| Droite (`w:right`) | 1418 | **25 mm** |
| En-tête (`w:header`) | 709 | 12,5 mm |
| Pied de page (`w:footer`) | 397 | 7 mm |

→ Le preset « normal » du gabarit utilise désormais **25 mm** au lieu de 14 mm.

## Polices utilisées

Statistiques sur 743 occurrences de polices dans le corps du document :

| Police | Occurrences | Pourcentage |
|---|---|---|
| **Arial** | 739 | **99,5 %** |
| Times New Roman | 2 | 0,3 % |
| Segoe Print | 2 | 0,3 % |

**Conclusion** : Le devis Marie Eugénie est intégralement en **Arial**. Le gabarit
« Classique » utilise donc maintenant Arial par défaut (au lieu de Cormorant Garamond
qui était une hypothèse esthétique initiale).

## Tailles de police

Distribution des tailles dans le document :

| Taille | Usage typique | Occurrences |
|---|---|---|
| 16 pt | **Titre principal page de garde** (« Proposition ME ») | 1 |
| 14 pt | **Titres de section** + adresses sites importantes | 5 |
| 12 pt | Coordonnées (client, société, adresses) | 16 |
| **10 pt** | **Texte courant** (dates, opérations, mentions) | **58 (197 runs)** |
| 8-9 pt | Notes de bas de page, mentions légales | 6 |

Configuration retenue pour le gabarit :
- `titre_principal: 16` (« Devis ME-XXXX » page de garde)
- `titres_section: 14` (« 1 - Détail des prestations »)
- `corps: 10` (paragraphes courants)

## Couleurs

Seules **deux couleurs de texte** sont utilisées dans le devis :

| Couleur hex | Usage | Occurrences |
|---|---|---|
| **#984806** | **Accent doré-brun Marie Eugénie** (éléments importants) | 9 |
| #0000FF | Hyperliens (email, site web) | 1 |

→ La couleur d'accent du gabarit a été corrigée de `#b08b5c` (estimation) à
**`#984806`** (valeur exacte extraite). Cette nuance brun-orangé est plus
chaude et plus typée que le doré clair initialement supposé.

## Coordonnées société complètes

Extrait du pied de page du devis original :

```
SAS Marie-Eugénie au Capital de 5000 €  -  1 rue Raspail – 13004 Marseille
Tél 04 91 47 14 38  /  Fax 04 91 47 51 30  -  N° Siret : 521 797 258
code APE n°8121 Z
contact@marie-eugenie.fr        www.Marie-Eugenie.fr
```

Champs additionnels par rapport à ce qu'on avait :
- **Capital social** : 5 000 €
- **Fax** : 04 91 47 51 30
- **N° SIRET** : 521 797 258
- **Code APE** : 8121 Z
- **Site web** : www.Marie-Eugenie.fr

Ces données sont maintenant intégrées au gabarit « Classique » via le champ
`pied_page`.

## Structure de la page de garde (originale)

D'après la conversion PDF du Word d'origine, la page 1 est structurée en
**deux colonnes** :

**Colonne gauche** (~50% largeur) :
- Logo officiel M·E en haut
- Titre « Proposition ME » centré
- Type de prestation (« Entretien des parties communes »)
- Adresse du site

**Colonne droite** (~50% largeur) :
- Baseline « Nettoyage & Services Associés » + date
- Bloc destinataire (client + adresses + contact) en surlignage jaune
- 3 bannières sépia empilées :
  1. « Marie-Eugénie, depuis plus de 35 ans au service de nos clients... »
  2. « Notre vrai richesse, la qualité de notre personnel... »
  3. « Nettoyage des parties communes, gestion des ordures ménagères... »

**Pied de page** :
- 3 certifications côte à côte : FEP (gauche), Petit Futé 2026 (centre), Palme Verte (droite)
- Bloc « Société Marie Eugénie - Représentée par Laurent PREVERT - lprevert@marie-eugenie.fr - 1 rue Raspail 13004 Marseille - T : 04 91 47 14 38 »
- Pied de page complet (capital, siret, APE, web)

## Écarts entre le prototype actuel et le devis original

Le prototype HTML actuel a une page de garde **en une colonne centrée**, alors
que le vrai devis est en **deux colonnes** avec les bannières sépia à droite.

C'est une limite assumée de la piste A (refactor HTML) : reproduire fidèlement
ce layout deux colonnes en HTML/CSS est faisable mais demanderait plusieurs
heures de travail supplémentaire, et le rendu en impression PDF reste sensible
aux différences d'engine HTML vs Word.

**Pour un rendu bit-pour-bit identique**, la piste B (backend Python + docxtpl)
est l'approche recommandée. Voir `ARCHITECTURE_BACKEND_WORD.md`.

## Récapitulatif des modifications appliquées au gabarit « Classique »

| Paramètre | Avant (hypothèse) | Après (valeur extraite) |
|---|---|---|
| Police titres | Cormorant Garamond (serif) | **Arial** |
| Police corps | Inter (sans-serif) | **Arial** |
| Taille titre principal | 32 px | **16 pt** (~21 px) |
| Taille titres section | 16 px | **14 pt** (~19 px) |
| Taille corps texte | 11,5 px | **10 pt** (~13 px) |
| Couleur accent | #b08b5c (doré clair) | **#984806** (brun-orangé) |
| Couleur titres | #1a1a1a | **#000000** |
| Couleur texte | #2a2a2a | **#000000** |
| Marges page (preset normal) | 14 mm | **25 mm** |
| SIRET | (vide) | **521 797 258** |
| Pied de page | Court (adresse + tél + email) | **Complet** (capital, fax, SIRET, APE, site web) |

# DevisFlow — Méthode de calcul des prix

## Formule

Pour chaque ligne de prestation :

```
coût main d'œuvre = heures × nombre d'agents × taux_horaire × coefficient_technicité
HT par passage   = coût main d'œuvre + frais complémentaires
```

Puis selon le type de devis :

```
RÉCURRENT (contrat)  : HT mensuel = HT par passage × passages_par_mois
PONCTUEL (forfait)   : HT         = HT par passage  (un seul passage)
```

Et la TVA / le TTC :

```
TVA = HT × taux_tva          (20 % par défaut)
TTC = HT + TVA
```

**Correction manuelle** : si un prix HT est forcé (`prix_force_ht`), il remplace
le HT calculé ; la TVA et le TTC sont recalculés à partir de ce montant.

## Coefficients de technicité (paramétrables)

| Niveau | Coefficient |
|---|---|
| Standard | 1,00 |
| Technique | 1,15 |
| Haute | 1,30 |
| Exceptionnelle | 1,50 |

Ces valeurs sont modifiables dans **Paramètres** et appliquées immédiatement au
calcul.

## Passages par mois (selon la fréquence)

| Fréquence | Passages/mois |
|---|---|
| Quotidien | 30 |
| 5×/semaine | 21,67 |
| 4×/semaine | 17,33 |
| 3×/semaine | 13,00 |
| 2×/semaine | 8,667 |
| 1×/semaine | 4,333 |
| Bimensuel | 2,00 |
| 1×/mois | 1,00 |
| Trimestriel | 0,333 |

(base : 52 semaines / 12 mois ≈ 4,333 semaines par mois)

## Exemples vérifiés

### Exemple 1 — Récurrent, 1×/semaine
2 h × 1 agent × 24 € × 1,00 = 48,00 € par passage
48,00 € × 4,333 passages/mois = **208,00 € HT/mois**
TVA 20 % = 41,60 € → **249,60 € TTC/mois**

### Exemple 2 — Récurrent, technique, 2×/semaine
3 h × 2 agents × 25 € × 1,15 = 172,50 € par passage
172,50 € × 8,667 = **1 495,00 € HT/mois**
TVA 20 % = 299,00 € → **1 794,00 € TTC/mois**

### Exemple 3 — Ponctuel, standard, avec frais
8 h × 2 agents × 28 € × 1,00 = 448,00 €
+ 150 € de frais = **598,00 € HT**
TVA 20 % = 119,60 € → **717,60 € TTC**

### Exemple 4 — Ponctuel, haute technicité, avec frais
5 h × 3 agents × 30 € × 1,30 = 585,00 €
+ 200 € de frais = **785,00 € HT**
TVA 20 % = 157,00 € → **942,00 € TTC**

### Exemple 5 — Correction manuelle
Prix forcé : 1 000,00 € HT → TVA 200,00 € → **1 200,00 € TTC**

## Options comparatives (récurrent)

Pour un devis récurrent, l'outil peut présenter plusieurs options de fréquence
(1×, 2×, 3×/semaine) chiffrées au mois dans un tableau comparatif, à partir de la
même prestation de base.

# DevisFlow — Guide d'utilisation

## Vue d'ensemble

DevisFlow génère des devis Word/PDF **strictement fidèles aux modèles Marie
Eugénie**, en ne remplaçant que les variables. Deux types de devis :
- **Récurrent** (contrat) — modèle « Copropriété petite »
- **Ponctuel** (forfait) — modèle « Encombrants Caves »

L'interface comporte 6 vues : Création, Mes devis, Clients, Bibliothèque,
Équipe, Paramètres.

## 1. Créer un devis

1. Cliquer sur **Nouveau devis**.
2. Choisir le **type** : Contrat récurrent ou Intervention ponctuelle.
3. **Client** : taper les premières lettres dans « Raison sociale » pour
   sélectionner un client existant — ses coordonnées et l'adresse du site se
   remplissent automatiquement. Ou saisir un nouveau client.
4. **Prestations** : renseigner les prestations, fréquences (récurrent),
   heures, agents, technicité, frais.
5. **Aperçu** : le panneau de droite affiche le **vrai PDF** généré depuis le
   modèle Word (pas une maquette). Bouton « Rafraîchir l'aperçu » pour le mettre
   à jour.
6. **Générer** : « Télécharger .docx » ou « Générer le PDF ».

## 2. Ajouter des photos d'équipement

Dans l'écran de création, section **Photos des équipements** :
- **Cocher** un ou plusieurs équipements de la bibliothèque (Autolaveuse,
  Camion benne, etc.) : leur photo sera injectée dans le devis.
- **Ajouter une photo** : bouton « + Ajouter une photo » pour téléverser une
  image directement (JPG/PNG), sans passer par la bibliothèque.
- **Prestation concernée** : pour chaque équipement/photo, choisir la prestation
  dans le menu déroulant. La photo sera **alignée en face** de cette prestation
  dans le document. Plusieurs photos sur la même prestation s'affichent côte à
  côte, alignées.

Les photos sont **redimensionnées automatiquement**, ne chevauchent pas le texte
et ne débordent pas de la page. La mise en page du modèle Word est conservée.

## 3. Bibliothèque clients

Vue **Clients** : créer, modifier, archiver, rechercher. Les clients sont
réutilisables dans tous les devis (préremplissage automatique).

## 4. Bibliothèque métier

Vue **Bibliothèque** :
- **Prestations types** : les zones du modèle (Hall d'entrée et ses opérations,
  caves, etc.). Ajouter / modifier / supprimer.
- **Équipements / matériel / véhicules** : chaque équipement a un nom, une
  catégorie, une description, une photo et un statut actif/inactif. Ajouter /
  modifier / supprimer, avec upload de photo.

## 5. Équipe

Vue **Équipe** : gérer les membres (nom, e-mail, rôle, actif/inactif).

## 6. Paramètres

Vue **Paramètres** → carte « Paramètres de calcul » :
- taux horaire par défaut,
- taux de TVA,
- coefficients de technicité (standard, technique, haute, exceptionnelle).

Ces valeurs sont **utilisées par le moteur de calcul** des devis. Modifier puis
« Enregistrer les paramètres ».

## Bon à savoir

- L'aperçu et le document final utilisent la **même logique** : ce que vous voyez
  est ce qui sera généré.
- Les prix se calculent automatiquement ; un total HT peut être forcé
  manuellement si besoin.
- Toutes les données (clients, équipements, paramètres, devis) sont enregistrées
  en base et conservées.

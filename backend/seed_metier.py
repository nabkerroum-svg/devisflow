"""
Seed de la bibliothèque métier DevisFlow.

Pré-remplit au premier démarrage :
  - les prestations types récurrentes (zones EXACTES du modèle Word « Copro Petite »)
  - quelques prestations types ponctuelles (encombrants)
  - la bibliothèque d'équipements / matériels / véhicules
  - quelques clients de démonstration
  - les membres de l'équipe
  - les paramètres de calcul (taux horaire, TVA, coefficients, modèles)

Idempotent : ne réinsère pas si la table correspondante contient déjà des lignes.
"""
import json
from pathlib import Path

from sqlmodel import Session, select

from config import SEED_DIR, STORAGE_DIR
from models import (engine, Client, Equipement, PrestationType, Membre, Parametre)


EQUIPEMENTS = [
    ("autolaveuse", "Autolaveuse", "machine", "Lavage mécanisé des grandes surfaces"),
    ("monobrosse", "Monobrosse", "machine", "Décapage et lustrage des sols"),
    ("injecteur_extracteur", "Injecteur-extracteur", "machine", "Nettoyage moquettes et textiles"),
    ("haute_pression", "Nettoyeur haute pression", "machine", "Décrassage des surfaces extérieures"),
    ("aspirateur_eau", "Aspirateur eau et poussière", "materiel", "Aspiration liquides et poussières"),
    ("camion_benne", "Camion benne", "vehicule", "Évacuation des encombrants"),
    ("utilitaire", "Véhicule utilitaire", "vehicule", "Transport matériel et équipes"),
    ("echafaudage", "Échafaudage roulant", "materiel", "Travaux en hauteur"),
    ("nebulisateur", "Nébulisateur (DSVA)", "specifique", "Désinfection par voie aérienne"),
]

EQUIPEMENTS_CATALOGUE = [
    ("perche_h2o", "Perche H2O", "Vitrerie",
     "SystÃ¨me de nettoyage Ã  l'eau pure permettant d'intervenir en hauteur en toute sÃ©curitÃ©.",
     ["vitrerie", "vitres", "hauteur"]),
    ("fourgon_encombrants", "Camion / fourgon d'intervention", "Encombrants",
     "VÃ©hicule adaptÃ© au chargement, au transport et Ã  l'Ã©vacuation des encombrants vers un centre agrÃ©Ã©.",
     ["encombrants", "evacuation", "transport"]),
    ("injecteur_extracteur", "Injecteur-extracteur", "Moquette / tapis",
     "MatÃ©riel professionnel permettant le nettoyage en profondeur des tapis, moquettes et textiles.",
     ["tapis", "moquette", "textile"]),
    ("monobrosse", "Monobrosse", "Remise en Ã©tat",
     "MatÃ©riel utilisÃ© pour le dÃ©capage, le lavage mÃ©canisÃ© et la remise en Ã©tat des sols.",
     ["remise_etat", "sols", "decapage"]),
    ("autolaveuse", "Autolaveuse", "Remise en Ã©tat",
     "Lavage et sÃ©chage mÃ©canisÃ©s des surfaces de sol.",
     ["remise_etat", "parking", "sols"]),
    ("haute_pression", "Nettoyeur haute pression", "Remise en Ã©tat",
     "DÃ©crassage des surfaces extÃ©rieures, parkings, sols et zones techniques.",
     ["remise_etat", "parking", "exterieur"]),
    ("aspirateur_eau", "Aspirateur eau et poussiÃ¨re", "MatÃ©riel",
     "Aspiration professionnelle des liquides et poussiÃ¨res.",
     ["appartement", "remise_etat", "sols"]),
    ("camion_benne", "Camion benne", "Encombrants",
     "VÃ©hicule de chargement et d'Ã©vacuation des encombrants.",
     ["encombrants", "evacuation", "transport"]),
    ("utilitaire", "VÃ©hicule utilitaire", "Encombrants",
     "Transport du matÃ©riel, des Ã©quipes et des dÃ©chets vers les filiÃ¨res adaptÃ©es.",
     ["encombrants", "transport"]),
    ("echafaudage", "Ã‰chafaudage roulant", "Vitrerie",
     "AccÃ¨s sÃ©curisÃ© pour les travaux ponctuels en hauteur.",
     ["vitrerie", "hauteur"]),
    ("nebulisateur", "NÃ©bulisateur (DSVA)", "DÃ©sinfection",
     "DÃ©sinfection par voie aÃ©rienne selon les besoins de l'intervention.",
     ["desinfection", "sanitaire"]),
]

CLIENTS = [
    dict(nom="Syndic Foncia Marseille", contact="Sophie MARCHAND", civilite="Madame",
         email="s.marchand@foncia.fr", telephone="04 91 12 34 56",
         adresse="12 rue Paradis", code_postal="13001", ville="Marseille",
         site_nom="Résidence Le Prado", site_adresse="52 rue Louis Astruc, 13005 Marseille"),
    dict(nom="Syndic Citya Prado", contact="Laurence DUVAL", civilite="Madame",
         email="l.duval@citya.fr", telephone="04 91 22 33 44",
         adresse="180 avenue du Prado", code_postal="13008", ville="Marseille",
         site_nom="Le Castellet", site_adresse="24 rue Borde, 13008 Marseille"),
    dict(nom="Cabinet Lieutaud", contact="Jean PEYRONNET", civilite="Monsieur",
         email="j.peyronnet@lieutaud.fr", telephone="04 91 55 66 77",
         adresse="5 cours Pierre Puget", code_postal="13006", ville="Marseille",
         site_nom="Résidence Puget", site_adresse="5 cours Pierre Puget, 13006 Marseille"),
]

MEMBRES = [
    ("Laurent Prévert", "lprevert@marie-eugenie.fr", "Président", True),
    ("Sophie Martin", "s.martin@marie-eugenie.fr", "Commercial", True),
    ("Karim Benhaddou", "k.benhaddou@marie-eugenie.fr", "Commercial", True),
    ("Marie Lopez", "m.lopez@marie-eugenie.fr", "Assistante exploitation", False),
]

PARAMETRES = [
    ("taux_horaire_defaut", "24", "Taux horaire par défaut (€/h)", "calcul"),
    ("tva_defaut", "0.20", "Taux de TVA par défaut", "calcul"),
    ("coef_standard", "1.00", "Coefficient technicité — standard", "calcul"),
    ("coef_technique", "1.15", "Coefficient technicité — technique", "calcul"),
    ("coef_haute", "1.30", "Coefficient technicité — haute", "calcul"),
    ("coef_exceptionnelle", "1.50", "Coefficient technicité — exceptionnelle", "calcul"),
    ("modele_recurrent", "copro_petite", "Modèle utilisé pour les devis récurrents", "modeles"),
    ("modele_bureaux", "bureaux_petit", "Modèle utilisé pour les devis bureaux", "modeles"),
    ("modele_ponctuel", "ponctuel_generique", "Modèle utilisé pour les devis ponctuels", "modeles"),
    ("societe_nom", "Marie Eugénie", "Raison sociale", "societe"),
    ("societe_siret", "521 797 258", "SIRET", "societe"),
]

# Prestations ponctuelles types (encombrants & co.)
PRESTATIONS_PONCTUELLES = [
    {"code": "p_affichage", "titre": "Affichage et information des résidents",
     "operations": ["Affichage dans les caves et le hall de la date d'intervention",
                    "Information des résidents sur le débarras à venir"]},
    {"code": "p_debarras", "titre": "Débarras des encombrants",
     "operations": ["Débarras de tous les encombrants des parties communes des caves",
                    "Tri sur place selon la nature des déchets"]},
    {"code": "p_transport", "titre": "Chargement et transport",
     "operations": ["Chargement dans un fourgon / camion benne",
                    "Transport vers un centre de traitement agréé"]},
    {"code": "p_traitement", "titre": "Traitement et remise en état",
     "operations": ["Coût du traitement des encombrants",
                    "Balayage des parties communes des caves"]},
]


def seed_metier():
    with Session(engine) as s:
        # --- Prestations récurrentes (zones du modèle Word) ---
        if not s.exec(select(PrestationType)).first():
            zones_file = SEED_DIR / "zones_copro.json"
            if zones_file.exists():
                zones = json.loads(zones_file.read_text(encoding="utf-8"))
                for i, z in enumerate(zones):
                    _upsert_prestation_type(s, z, "contrat", i)
                print(f"[seed_metier] {len(zones)} prestations récurrentes (zones Copro)")
            zones_bureaux_file = SEED_DIR / "zones_bureaux.json"
            if zones_bureaux_file.exists():
                zones_bureaux = json.loads(zones_bureaux_file.read_text(encoding="utf-8"))
                for i, z in enumerate(zones_bureaux):
                    _upsert_prestation_type(s, z, "bureaux", i)
                print(f"[seed_metier] {len(zones_bureaux)} prestations bureaux")
            for i, p in enumerate(PRESTATIONS_PONCTUELLES):
                s.add(PrestationType(
                    code=p["code"], titre=p["titre"], famille="ponctuel", freq_var=None,
                    operations=json.dumps(p["operations"], ensure_ascii=False), ordre=i,
                ))
            print(f"[seed_metier] {len(PRESTATIONS_PONCTUELLES)} prestations ponctuelles")
        else:
            zones_bureaux_file = SEED_DIR / "zones_bureaux.json"
            if zones_bureaux_file.exists():
                zones_bureaux = json.loads(zones_bureaux_file.read_text(encoding="utf-8"))
                for i, z in enumerate(zones_bureaux):
                    _upsert_prestation_type(s, z, "bureaux", i)

        # --- Équipements ---
        if not s.exec(select(Equipement)).first():
            import shutil
            photos_src = SEED_DIR / "photos"
            photos_dst = STORAGE_DIR / "photos"
            photos_dst.mkdir(parents=True, exist_ok=True)
            for code, label, cat, desc in EQUIPEMENTS:
                photo_path = None
                src = photos_src / f"{code}.jpg"
                if src.exists():
                    shutil.copy(src, photos_dst / f"{code}.jpg")
                    photo_path = f"{code}.jpg"
                s.add(Equipement(code=code, label=label, categorie=cat,
                                 description=desc, photo_path=photo_path))
            print(f"[seed_metier] {len(EQUIPEMENTS)} équipements (avec photos)")

        # ComplÃ©ment catalogue matÃ©riel : ajout idempotent des exemples mÃ©tier
        # utilisÃ©s pour les suggestions dans les devis ponctuels.
        photos_src = SEED_DIR / "photos"
        photos_dst = STORAGE_DIR / "photos"
        photos_dst.mkdir(parents=True, exist_ok=True)
        added_catalogue = 0
        for code, label, cat, desc, tags in EQUIPEMENTS_CATALOGUE:
            item = s.exec(select(Equipement).where(Equipement.code == code)).first()
            src = photos_src / f"{code}.jpg"
            photo_path = None
            if src.exists():
                import shutil
                shutil.copy(src, photos_dst / f"{code}.jpg")
                photo_path = f"{code}.jpg"
            if not item:
                s.add(Equipement(
                    code=code, label=label, categorie=cat, description=desc,
                    tags=json.dumps(tags, ensure_ascii=False), photo_path=photo_path,
                ))
                added_catalogue += 1
            elif not getattr(item, "tags", None):
                item.tags = json.dumps(tags, ensure_ascii=False)
                if photo_path and not item.photo_path:
                    item.photo_path = photo_path
                s.add(item)
        if added_catalogue:
            print(f"[seed_metier] {added_catalogue} matÃ©riel(s) catalogue ajoutÃ©(s)")

        # --- Clients ---
        if not s.exec(select(Client)).first():
            for c in CLIENTS:
                s.add(Client(**c))
            print(f"[seed_metier] {len(CLIENTS)} clients")

        # --- Membres ---
        if not s.exec(select(Membre)).first():
            for nom, email, role, actif in MEMBRES:
                s.add(Membre(nom=nom, email=email, role=role, actif=actif))
            print(f"[seed_metier] {len(MEMBRES)} membres")

        # --- Paramètres ---
        if not s.exec(select(Parametre)).first():
            for cle, val, lib, grp in PARAMETRES:
                s.add(Parametre(cle=cle, valeur=val, libelle=lib, groupe=grp))
            print(f"[seed_metier] {len(PARAMETRES)} paramètres")

        s.commit()


def _upsert_prestation_type(session: Session, zone: dict, famille: str, ordre: int) -> None:
    item = session.exec(select(PrestationType).where(PrestationType.code == zone["code"])).first()
    if not item:
        item = PrestationType(code=zone["code"])
    item.titre = zone["titre"]
    item.famille = famille
    item.freq_var = zone.get("freq_var")
    item.operations = json.dumps(zone.get("operations", []), ensure_ascii=False)
    item.ordre = ordre
    item.actif = True
    session.add(item)


if __name__ == "__main__":
    from models import init_db
    init_db()
    seed_metier()

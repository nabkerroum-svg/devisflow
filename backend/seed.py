"""
Seed initial - pre-charge les deux modeles maitres Marie Eugenie au premier
demarrage :

  - copro_petite        (recurrent / contrat)
  - bureaux_petit       (recurrent / contrat)
  - ponctuel_generique  (ponctuel)

Si la base existe deja, on applique quand meme la migration minimale qui remplace
l'ancien template physique encombrants_caves par le template ponctuel generique.
"""
import json
import shutil

from sqlmodel import Session, select

from config import TEMPLATES_DIR, SEED_DIR
from models import Template, engine
from template_service import analyser_template


SEED_TEMPLATES = [
    {
        "fichier": "copro_petite.docx",
        "code": "copro_petite",
        "nom": "Copropriete - petite surface (Marie Eugenie)",
        "famille": "contrat",
        "type_intervention": "Entretien des parties communes",
        "is_default": True,
    },
    {
        "fichier": "bureaux_petit.docx",
        "code": "bureaux_petit",
        "nom": "Bureaux - petite surface (Marie Eugenie)",
        "famille": "contrat",
        "type_intervention": "Nettoyage des bureaux",
        "is_default": False,
    },
    {
        "fichier": "ponctuel_generique.docx",
        "code": "ponctuel_generique",
        "nom": "Devis ponctuel generique",
        "famille": "ponctuel",
        "type_intervention": "Prestation ponctuelle",
        "is_default": False,
    },
    {
        "fichier": "ponctuel_encombrants.docx",
        "code": "ponctuel_encombrants",
        "nom": "Ponctuel - encombrants divers",
        "famille": "ponctuel",
        "type_intervention": "Encombrants divers",
        "is_default": False,
    },
    {
        "fichier": "ponctuel_encombrants_caves.docx",
        "code": "ponctuel_encombrants_caves",
        "nom": "Ponctuel - encombrants caves",
        "famille": "ponctuel",
        "type_intervention": "Encombrants caves",
        "is_default": False,
    },
    {
        "fichier": "ponctuel_tag.docx",
        "code": "ponctuel_tag",
        "nom": "Ponctuel - tags",
        "famille": "ponctuel",
        "type_intervention": "Enlevement de tags",
        "is_default": False,
    },
    {
        "fichier": "ponctuel_tapis.docx",
        "code": "ponctuel_tapis",
        "nom": "Ponctuel - tapis",
        "famille": "ponctuel",
        "type_intervention": "Fourniture et mise en place tapis",
        "is_default": False,
    },
    {
        "fichier": "ponctuel_vitrerie.docx",
        "code": "ponctuel_vitrerie",
        "nom": "Ponctuel - vitrerie",
        "famille": "ponctuel",
        "type_intervention": "Prestation de nettoyage de la vitrerie",
        "is_default": False,
    },
    {
        "fichier": "ponctuel_relamping.docx",
        "code": "ponctuel_relamping",
        "nom": "Ponctuel - relamping",
        "famille": "ponctuel",
        "type_intervention": "Relamping - changement d'ampoules",
        "is_default": False,
    },
    {
        "fichier": "ponctuel_remise_etat.docx",
        "code": "ponctuel_remise_etat",
        "nom": "Ponctuel - remise en etat",
        "famille": "ponctuel",
        "type_intervention": "Remise en etat",
        "is_default": False,
    },
    {
        "fichier": "ponctuel_appartement.docx",
        "code": "ponctuel_appartement",
        "nom": "Ponctuel - appartement",
        "famille": "ponctuel",
        "type_intervention": "Nettoyage d'appartement",
        "is_default": False,
    },
]

BUREAUX_PETIT_VARIABLES = [
    "DEST_LIGNE1", "DEST_LIGNE2", "DEST_LIGNE3", "DEST_LIGNE4",
    "NUMERO_DEVIS", "DATE_EMISSION", "DATE_SIGNATURE", "TYPE_PRESTATION",
    "SITE_ADRESSE", "SITE_CP_VILLE", "DATE_PRISE_EFFET", "DUREE_CONTRAT",
    "CONDITIONS_REGLEMENT", "TOTAL_HT", "TOTAL_TVA", "TOTAL_TTC",
    "SURFACE_LOCAUX", "NB_BLOCS_SANITAIRES", "NB_COLLABORATEURS",
    "REVETEMENT_BUREAU", "REVETEMENT_SANITAIRE",
    "SHOW_ACCUEIL_BUREAUX", "SHOW_CIRCULATION", "SHOW_SANITAIRES",
    "SHOW_CUISINE", "SHOW_VITRERIE", "SHOW_CONSOMMABLES",
    "FREQ_ACCUEIL_BUREAUX", "FREQ_CIRCULATION", "FREQ_SANITAIRES",
    "FREQ_CUISINE", "FREQ_VITRERIE",
    "OPS_ACCUEIL_BUREAUX", "OPS_CIRCULATION", "OPS_SANITAIRES",
    "OPS_CUISINE", "OPS_VITRERIE", "OPS_CONSOMMABLES",
]


def seed_initial():
    with Session(engine) as session:
        existing = session.exec(select(Template)).first()
        if existing:
            for spec in SEED_TEMPLATES:
                _upsert_seed_template(session, spec)
            _migrer_template_ponctuel_generique(session)
            return

        for spec in SEED_TEMPLATES:
            _upsert_seed_template(session, spec)


def _upsert_seed_template(session: Session, spec: dict):
    seed_file = SEED_DIR / spec["fichier"]
    if not seed_file.exists():
        print(f"[seed] fichier seed introuvable : {seed_file}")
        return

    target = TEMPLATES_DIR / spec["fichier"]
    if not target.exists():
        shutil.copy(seed_file, target)

    variables = BUREAUX_PETIT_VARIABLES if spec["code"] == "bureaux_petit" else analyser_template(target)
    t = session.exec(select(Template).where(Template.code == spec["code"])).first()
    if not t:
        t = Template(code=spec["code"])
    t.nom = spec["nom"]
    t.famille = spec["famille"]
    t.fichier = spec["fichier"]
    t.type_intervention = spec["type_intervention"]
    t.is_default = spec["is_default"]
    t.actif = True
    t.variables = json.dumps(variables)
    session.add(t)
    session.commit()
    print(f"[seed] Template '{spec['code']}' charge ({len(variables)} variables)")


def _migrer_template_ponctuel_generique(session: Session):
    """Assure que les bases existantes pointent vers le template ponctuel generique."""
    spec = next(s for s in SEED_TEMPLATES if s["code"] == "ponctuel_generique")
    seed_file = SEED_DIR / spec["fichier"]
    target = TEMPLATES_DIR / spec["fichier"]
    if seed_file.exists() and not target.exists():
        shutil.copy(seed_file, target)

    current = session.exec(select(Template).where(Template.code == spec["code"])).first()
    legacy = session.exec(select(Template).where(Template.code == "encombrants_caves")).first()
    if current and legacy and current.id != legacy.id:
        session.delete(legacy)
    t = current or legacy
    if not t:
        _upsert_seed_template(session, spec)
        return

    variables = analyser_template(target) if target.exists() else []
    t.code = spec["code"]
    t.nom = spec["nom"]
    t.famille = spec["famille"]
    t.fichier = spec["fichier"]
    t.type_intervention = spec["type_intervention"]
    t.is_default = spec["is_default"]
    t.actif = True
    t.variables = json.dumps(variables)
    session.add(t)
    session.commit()
    print(f"[seed] Template ponctuel migre vers '{spec['code']}'")


if __name__ == "__main__":
    seed_initial()

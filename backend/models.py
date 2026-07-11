"""
Modèles de données — SQLModel (= SQLAlchemy + Pydantic).

Tables :
  - Template       : un modèle Word maître (uploadé par l'admin)
  - Devis          : un devis généré (référence + métadonnées)
  - Client         : carnet de clients
  - Equipement     : bibliothèque d'équipements/matériels avec photos
"""
from datetime import datetime
from typing import Optional, List
from sqlmodel import SQLModel, Field, create_engine, Session, select
from sqlalchemy import Column, JSON
from config import SQLITE_URL


class Template(SQLModel, table=True):
    """Un template Word maître (uploadé par l'admin via la page Modèles PDF)."""
    id: Optional[int] = Field(default=None, primary_key=True)
    code: str = Field(index=True, unique=True)        # ex : "copro_petite"
    nom: str                                           # ex : "Copropriété — petite surface"
    famille: str = "contrat"                           # contrat | ponctuel
    fichier: str                                       # nom du fichier .docx dans storage/templates/
    type_intervention: Optional[str] = None            # ex : "Entretien parties communes"
    description: Optional[str] = None                   # description libre du devis type
    is_default: bool = False
    actif: bool = True
    created_at: datetime = Field(default_factory=datetime.utcnow)
    updated_at: datetime = Field(default_factory=datetime.utcnow)
    # Variables détectées dans le template (liste de noms Jinja2)
    variables: Optional[str] = None                    # JSON-encoded list


class Devis(SQLModel, table=True):
    """Un devis généré."""
    id: Optional[int] = Field(default=None, primary_key=True)
    numero: str = Field(index=True, unique=True)       # ex : "ME-6245"
    template_code: str                                  # référence au Template utilisé
    client_nom: str
    site_adresse: str
    date_emission: str                                  # format ISO
    montant_ht: float = 0.0
    montant_tva: float = 0.0
    montant_ttc: float = 0.0
    prix_override: bool = False                          # True si le prix a été forcé manuellement
    fichier_docx: Optional[str] = None                 # nom du .docx dans storage/generated/
    fichier_pdf: Optional[str] = None
    payload: Optional[str] = None                       # JSON-encoded snapshot des données saisies
    statut: str = "brouillon"                           # brouillon | envoye | accepte
    created_at: datetime = Field(default_factory=datetime.utcnow)
    updated_at: datetime = Field(default_factory=datetime.utcnow)


class TemplateSubstitution(SQLModel, table=True):
    """Table de substitution propre à un template (annotation automatique).

    Permet d'importer un nouveau modèle Word/PDF et de définir ses zones
    variables EN BASE, sans modifier le code. Chaque ligne associe un texte
    présent dans le document d'origine au marqueur Jinja2 à injecter.
    """
    id: Optional[int] = Field(default=None, primary_key=True)
    template_code: str = Field(index=True)              # référence au Template.code
    texte_origine: str                                  # texte trouvé dans le .docx
    marqueur: str                                       # ex : "{{ DATE_EMISSION }}"
    exact_match: bool = False                           # le paragraphe doit être égal au texte
    ordre: int = 0                                      # ordre d'application


class Client(SQLModel, table=True):
    """Carnet de clients."""
    id: Optional[int] = Field(default=None, primary_key=True)
    nom: str = Field(index=True)
    civilite: Optional[str] = None
    contact: Optional[str] = None
    email: Optional[str] = None
    telephone: Optional[str] = None
    adresse: Optional[str] = None
    code_postal: Optional[str] = None
    ville: Optional[str] = None
    site_nom: Optional[str] = None
    site_adresse: Optional[str] = None
    archive: bool = False
    created_at: datetime = Field(default_factory=datetime.utcnow)
    updated_at: datetime = Field(default_factory=datetime.utcnow)


class Equipement(SQLModel, table=True):
    """Bibliothèque d'équipements / matériels."""
    id: Optional[int] = Field(default=None, primary_key=True)
    code: str = Field(index=True, unique=True)
    label: str
    categorie: str = "materiel"   # machine | materiel | vehicule | specifique
    description: Optional[str] = None
    photo_path: Optional[str] = None   # chemin relatif vers storage/photos/
    tags: str = "[]"                   # JSON list[str] : vitrerie, encombrants, tapis...
    actif: bool = True
    created_at: datetime = Field(default_factory=datetime.utcnow)
    updated_at: datetime = Field(default_factory=datetime.utcnow)


class PrestationType(SQLModel, table=True):
    """Zone / prestation type de la bibliothèque métier.

    Reprend les zones EXACTES du modèle Word « Copro Petite » (Hall d'entrée,
    cabine d'ascenseur, etc.) avec leur liste d'opérations détaillées. Utilisée
    directement dans la création d'un devis récurrent.
    """
    id: Optional[int] = Field(default=None, primary_key=True)
    code: str = Field(index=True, unique=True)          # ex : "hall"
    titre: str                                          # ex : "Hall d'Entrée"
    famille: str = "contrat"                            # contrat | ponctuel
    freq_var: Optional[str] = None                      # ex : "FREQ_HALL"
    operations: str = "[]"                              # JSON list[str] des opérations détaillées
    ordre: int = 0
    actif: bool = True


class Membre(SQLModel, table=True):
    """Membre de l'équipe Marie Eugénie (vue Équipe)."""
    id: Optional[int] = Field(default=None, primary_key=True)
    nom: str = Field(index=True)
    email: Optional[str] = None
    role: Optional[str] = None
    actif: bool = True
    created_at: datetime = Field(default_factory=datetime.utcnow)


class Parametre(SQLModel, table=True):
    """Paramètres globaux de calcul et de génération (vue Paramètres).

    Stockés en clé/valeur pour rester souples : taux horaire par défaut, TVA,
    coefficients de technicité, codes modèles récurrent/ponctuel, etc.
    """
    id: Optional[int] = Field(default=None, primary_key=True)
    cle: str = Field(index=True, unique=True)
    valeur: str
    libelle: Optional[str] = None
    groupe: str = "calcul"   # calcul | modeles | societe


# === Helper engine ===
engine = create_engine(SQLITE_URL, echo=False, connect_args={"check_same_thread": False})


def init_db():
    """Crée toutes les tables (idempotent) puis applique les migrations de colonnes."""
    SQLModel.metadata.create_all(engine)
    _migrer_colonnes_manquantes()


# Colonnes attendues par table, avec leur définition SQL et valeur par défaut.
# Permet d'ajouter automatiquement les colonnes manquantes sur une ancienne base
# SQLite, sans avoir à supprimer la base.
_MIGRATIONS = {
    "template": [
        ("description", "TEXT DEFAULT ''"),
        ("is_default", "BOOLEAN DEFAULT 0"),
        ("actif", "BOOLEAN DEFAULT 1"),
        ("type_intervention", "TEXT"),
        ("created_at", "DATETIME"),
        ("updated_at", "DATETIME"),
        ("variables", "TEXT DEFAULT '[]'"),
    ],
    "equipement": [
        ("tags", "TEXT DEFAULT '[]'"),
    ],
}


def _migrer_colonnes_manquantes():
    """Vérifie chaque table et ajoute via ALTER TABLE les colonnes manquantes.
    Idempotent : ne touche qu'aux colonnes réellement absentes."""
    from sqlalchemy import text
    with engine.begin() as conn:
        for table, cols in _MIGRATIONS.items():
            # table existe-t-elle ?
            exists = conn.execute(text(
                "SELECT name FROM sqlite_master WHERE type='table' AND name=:n"
            ), {"n": table}).first()
            if not exists:
                continue
            # colonnes présentes
            rows = conn.execute(text(f"PRAGMA table_info({table})")).fetchall()
            present = {r[1] for r in rows}  # r[1] = nom de colonne
            for col_name, col_def in cols:
                if col_name not in present:
                    try:
                        conn.execute(text(f"ALTER TABLE {table} ADD COLUMN {col_name} {col_def}"))
                        print(f"[migration] colonne ajoutée : {table}.{col_name}")
                    except Exception as e:
                        print(f"[migration] échec {table}.{col_name} : {e}")
            # backfill des dates si NULL (created_at/updated_at)
            try:
                if "created_at" in {c[0] for c in cols}:
                    conn.execute(text(
                        f"UPDATE {table} SET created_at = CURRENT_TIMESTAMP WHERE created_at IS NULL"))
                if "updated_at" in {c[0] for c in cols}:
                    conn.execute(text(
                        f"UPDATE {table} SET updated_at = CURRENT_TIMESTAMP WHERE updated_at IS NULL"))
                if "variables" in {c[0] for c in cols}:
                    conn.execute(text(
                        f"UPDATE {table} SET variables = '[]' WHERE variables IS NULL"))
            except Exception as e:
                print(f"[migration] backfill {table} : {e}")


def get_session():
    """Dépendance FastAPI : ouvre une session par requête."""
    with Session(engine) as session:
        yield session

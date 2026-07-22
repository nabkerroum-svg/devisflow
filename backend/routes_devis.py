"""
Routes API pour la génération de devis.
"""
import json
import base64
import io
import re
import uuid
from pathlib import Path
from typing import Dict, List, Optional, Any
from datetime import date, datetime

from fastapi import APIRouter, HTTPException, Depends
from fastapi.responses import FileResponse
from pydantic import BaseModel
from sqlmodel import Session, select

from config import TEMPLATES_DIR, GENERATED_DIR, SEED_DIR
from models import Template, Devis, get_session
from template_service import generer_devis as gen_devis_file, docx_to_pdf
import pricing as P


def _analyser_pdf(pdf_path):
    """Retourne (nombre_de_pages, page_blanche_detectee).
    Une page est considérée blanche si elle contient moins de 8 mots de texte."""
    import subprocess
    # nombre de pages via pdfinfo
    info = subprocess.run(["pdfinfo", str(pdf_path)], capture_output=True, text=True)
    pages = None
    for line in info.stdout.splitlines():
        if line.startswith("Pages:"):
            pages = int(line.split(":")[1].strip())
            break
    if not pages:
        return None, None
    blanche = False
    for p in range(1, pages + 1):
        t = subprocess.run(["pdftotext", "-f", str(p), "-l", str(p), str(pdf_path), "-"],
                           capture_output=True, text=True).stdout
        if len(t.split()) < 8:
            blanche = True
            break
    return pages, blanche

router = APIRouter(prefix="/devis", tags=["devis"])

TEMPLATE_PONCTUEL_GENERIQUE = "ponctuel_generique"
MODELES_COPRO = {"copro_petite", "copro_importante"}
MODELES_BUREAUX_CONTRAT = {"bureaux_petit", "bureaux_important"}
LOGO_FALLBACK_IMAGE_NAMES = {"perche_h2o_e03c5fe5.jpg"}
BUREAUX_ZONE_CODES = ["accueil_bureaux", "circulation", "sanitaires", "cuisine", "vitrerie", "consommables"]
BUREAUX_FREQ_VARS = {
    "accueil_bureaux": "FREQ_ACCUEIL_BUREAUX",
    "circulation": "FREQ_CIRCULATION",
    "sanitaires": "FREQ_SANITAIRES",
    "cuisine": "FREQ_CUISINE",
    "vitrerie": "FREQ_VITRERIE",
    "consommables": None,
}

TEMPLATE_PONCTUEL_PAR_MODELE: Dict[str, str] = {
    "encombrants_divers": "ponctuel_encombrants",
    "encombrants_caves": "ponctuel_encombrants_caves",
    "tags": "ponctuel_tag",
    "tag": "ponctuel_tag",
    "tapis": "ponctuel_tapis",
    "vitrerie": "ponctuel_vitrerie",
    "relamping": "ponctuel_relamping",
    "ampoules": "ponctuel_relamping",
    "remise_etat": "ponctuel_remise_etat",
    "appartement": "ponctuel_appartement",
}

PRESTATIONS_PONCTUELLES_AUTORISEES: Dict[str, List[str]] = {
    "bureaux_petit": [
        "Depoussierage des bureaux, plans de travail et surfaces accessibles",
        "Vidage des corbeilles et remplacement des sacs si necessaire",
        "Aspiration et lavage des sols des bureaux",
        "Nettoyage des points de contact : poignees, interrupteurs et portes",
        "Nettoyage des sanitaires et reapprovisionnement si prevu",
        "Nettoyage de la kitchenette ou zone cafe si presente",
    ],
    "bureaux_important": [
        "Depoussierage complet des bureaux, postes de travail et surfaces accessibles",
        "Vidage des corbeilles, tri et evacuation des dechets courants",
        "Aspiration et lavage des sols des bureaux, circulations internes et salles de reunion",
        "Nettoyage des points de contact : poignees, interrupteurs, portes et claviers accessibles",
        "Nettoyage approfondi des sanitaires et points d'eau",
        "Nettoyage de la kitchenette, plans de travail et zones de pause",
        "Essuyage des vitrages interieurs accessibles si prevu au devis",
    ],
    "vitrerie": [
        "Mise en oeuvre de protection sur le sol par bachage ou tout autre moyen adapte",
        "Nettoyage de la vitrerie interieure",
        "Nettoyage de la vitrerie exterieure",
        "Nettoyage et essuyage des encadrements",
        "Enlevement de la protection du sol et nettoyage du perimetre d'intervention",
    ],
    "encombrants_caves": [
        "Affichage dans les caves et dans le hall de la date de l'intervention",
        "Debarras des encombrants situes dans les parties communes des caves",
        "Chargement dans un fourgon",
        "Transport des encombrants vers un centre de traitement agree",
        "Cout du traitement de ces encombrants",
        "Balayage des parties communes des caves",
    ],
    "encombrants_divers": [
        "Debarras des encombrants designes par le client",
        "Chargement et transport vers un centre de traitement agree",
        "Traitement des encombrants",
        "Nettoyage sommaire de la zone d'intervention",
    ],
    "appartement": [
        "Depoussierage de toutes les surfaces accessibles",
        "Nettoyage des sols par aspiration, balayage et lavage",
        "Nettoyage complet de la salle de bain et des sanitaires",
        "Nettoyage complet de la cuisine hors interieur des meubles",
        "Nettoyage des vitres interieures accessibles",
        "Vidage des poubelles et aeration des pieces",
    ],
    "nettoyage_apres_travaux": [
        "Evacuation des dechets de chantier residuels",
        "Depoussierage de toutes les surfaces : plafonds, murs et sols",
        "Nettoyage des vitres, fenetres et menuiseries",
        "Nettoyage approfondi des sanitaires et points d'eau",
        "Nettoyage des sols avec technique adaptee au revetement",
        "Elimination des traces de platre, peinture et colle",
    ],
    "remise_etat": [
        "Depoussierage approfondi des surfaces accessibles",
        "Lessivage des surfaces lavables et points de contact",
        "Nettoyage approfondi des sols selon leur revetement",
        "Nettoyage des sanitaires, faiences et points d'eau",
        "Nettoyage de la cuisine ou des zones techniques",
        "Evacuation des dechets courants lies a l'intervention",
    ],
    "cristallisation": [
        "Protection des mobiliers et plinthes",
        "Decapage du sol marbre a la monobrosse avec produit alcalin",
        "Rincage a l'eau claire",
        "Egrainage a la monobrosse munie de pads diamantes grain croissant",
        "Application du produit cristallisant",
        "Travail a la monobrosse jusqu'a brillance",
        "Lustrage final",
        "Aspiration des residus",
        "Mise en place de protection si necessaire",
    ],
    "tags": [
        "Test prealable de compatibilite du produit avec le support",
        "Application du produit decapant adapte au support",
        "Temps de pose selon preconisations fabricant",
        "Brossage et rincage haute pression",
        "Application d'un produit anti-graffiti protecteur si demande",
        "Nettoyage du perimetre d'intervention",
    ],
    "tapis": [
        "Prise de mesures precises de la zone de pose",
        "Fourniture du tapis aux dimensions adaptees",
        "Mise en place du tapis",
        "Conseil d'entretien",
    ],
    "relamping": [
        "Recensement des points lumineux defectueux",
        "Fourniture et pose des ampoules de remplacement",
        "Evacuation et recyclage des ampoules usagees",
    ],
    "parking": [
        "Balayage mecanise ou manuel des zones de stationnement",
        "Ramassage des dechets et evacuation",
        "Nettoyage des circulations, rampes et acces",
        "Traitement localise des taches selon faisabilite",
        "Nettoyage haute pression si prevu au devis",
    ],
    "parking_garage": [
        "Balayage mecanise ou manuel des zones de stationnement",
        "Ramassage des dechets et evacuation",
        "Nettoyage des circulations, rampes et acces",
        "Traitement localise des taches selon faisabilite",
        "Nettoyage haute pression si prevu au devis",
    ],
    "haute_pression": [
        "Mise en place du perimetre de securite",
        "Balayage prealable des surfaces",
        "Nettoyage haute pression avec produit degraissant si necessaire",
        "Rincage haute pression",
        "Evacuation des eaux residuelles si presentes",
    ],
    "desinfection": [
        "Preparation et confinement de la zone",
        "Mise en place des protections",
        "Nebulisation avec produit virucide et bactericide certifie",
        "Respect du temps de contact requis",
        "Aeration de la zone et remise en etat",
    ],
}


def _est_image_logo_entreprise(path: str | Path | None) -> bool:
    """Evite d'utiliser le logo Marie-Eugenie comme photo metier ou materiel."""
    if not path:
        return False
    name = Path(path).name.lower()
    if name in LOGO_FALLBACK_IMAGE_NAMES:
        return True
    return "logo" in name and ("marie" in name or "eugenie" in name or "eugénie" in name)

ALIASES_MODELES_PONCTUELS = {
    "bureau": "bureaux_petit",
    "bureaux": "bureaux_petit",
    "bureaux_importants": "bureaux_important",
    "nettoyage_logement": "appartement",
    "logement": "appartement",
    "garage": "parking_garage",
}

TERMES_COPRO_INTERDITS_HORS_COPRO = (
    "hall", "ascenseur", "escalier", "cage", "palier",
    "parties communes", "porte d'entree", "porte d'entrée",
    "boites aux lettres", "boîtes aux lettres", "local poubelles",
)


def _modele_metier_payload(payload: "DevisPayload") -> str:
    raw = ""
    if payload.variables:
        raw = str(payload.variables.get("_modele_code", "") or "")
    raw = raw or str(payload.template_code or "")
    code = raw.strip().lower()
    return ALIASES_MODELES_PONCTUELS.get(code, code)


def _est_modele_ponctuel_metier(modele_code: str) -> bool:
    return (
        modele_code in PRESTATIONS_PONCTUELLES_AUTORISEES
        and modele_code not in MODELES_BUREAUX_CONTRAT
        and modele_code not in MODELES_COPRO
    )


def _resoudre_template_code(payload: "DevisPayload") -> str:
    """Le modele metier choisit le contenu ; le template physique reste separe."""
    modele_code = _modele_metier_payload(payload)
    if modele_code == "bureaux_petit":
        return "bureaux_petit"
    if modele_code == "bureaux_important":
        return "bureaux_important"
    if _est_modele_ponctuel_metier(modele_code):
        template_code = TEMPLATE_PONCTUEL_PAR_MODELE.get(modele_code)
        if not template_code:
            raise HTTPException(400, "Modèle ou prestations non configurés pour ce type de devis ponctuel.")
        return template_code
    return payload.template_code


def _charger_zones_bureaux() -> List[Dict[str, Any]]:
    zones_file = SEED_DIR / "zones_bureaux.json"
    if not zones_file.exists():
        return []
    try:
        return json.loads(zones_file.read_text(encoding="utf-8"))
    except Exception:
        return []


def _ligne_depuis_libelle(libelle: str, source: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    base = dict(source or {})
    base["libelle"] = libelle
    base.setdefault("duree_h", 1)
    base.setdefault("nb_agents", 1)
    base.setdefault("taux_horaire", 26)
    base.setdefault("niveau_technicite", "standard")
    base.setdefault("frais", 0)
    return base


def _nettoyer_lignes_ponctuelles(modele_code: str, lignes: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Filtre les prestations incoherentes pour eviter toute fuite entre modeles."""
    modele_code = ALIASES_MODELES_PONCTUELS.get(modele_code, modele_code)
    defaults = PRESTATIONS_PONCTUELLES_AUTORISEES.get(modele_code)
    if not defaults:
        return lignes

    cleaned: List[Dict[str, Any]] = []
    for ligne in lignes or []:
        libelle = str(ligne.get("libelle", "") or "").strip()
        if not libelle:
            continue
        low = libelle.lower()
        if modele_code not in MODELES_COPRO and any(term in low for term in TERMES_COPRO_INTERDITS_HORS_COPRO):
            continue
        if modele_code == "vitrerie" and "encombrant" in low:
            continue
        if modele_code.startswith("encombrants") and "vitrerie" in low:
            continue
        cleaned.append(_ligne_depuis_libelle(libelle, ligne))

    if cleaned:
        return cleaned
    return [_ligne_depuis_libelle(libelle) for libelle in defaults]


def _normaliser_image_upload(data_url: str, uploads_dir: Path) -> Optional[str]:
    """Convertit une image data URL en JPEG utilisable par python-docx/LibreOffice."""
    if not data_url or not isinstance(data_url, str) or "," not in data_url:
        return None
    head, b64 = data_url.split(",", 1)
    mime_match = re.search(r"data:([^;]+)", head, re.I)
    mime = (mime_match.group(1).lower() if mime_match else "")
    if not (mime.startswith("image/") or mime == "application/octet-stream" or not mime):
        return None
    try:
        raw = base64.b64decode(b64, validate=False)
    except Exception as exc:
        print(f"[warn] base64 photo invalide : {exc}")
        return None
    try:
        try:
            import pillow_heif  # type: ignore
            pillow_heif.register_heif_opener()
        except Exception:
            pass
        from PIL import Image, ImageOps
        img = Image.open(io.BytesIO(raw))
        img = ImageOps.exif_transpose(img)
        if img.mode not in ("RGB", "L"):
            img = img.convert("RGB")
        elif img.mode == "L":
            img = img.convert("RGB")
        max_side = 1800
        if max(img.size) > max_side:
            img.thumbnail((max_side, max_side))
        buf = io.BytesIO()
        img.save(buf, format="JPEG", quality=88, optimize=True)
        fname = f"upload_{uuid.uuid4().hex[:10]}.jpg"
        out = uploads_dir / fname
        out.write_bytes(buf.getvalue())
        return str(out)
    except Exception as exc:
        print(f"[warn] photo non convertible ({mime or 'type inconnu'}) : {exc}")
        return None


class DevisPayload(BaseModel):
    """
    Données envoyées par le frontend pour générer un devis.

    `variables` contient les valeurs simples à injecter (NUMERO_DEVIS, DEST_*…).
    `lignes` contient les prestations à chiffrer (durée, agents, taux, technicité…)
    `frequences_options` (récurrent) : liste de fréquences pour le tableau comparatif.
    `prix_force_ht` : si fourni, force le total HT (override manuel).
    """
    numero: str
    template_code: str
    variables: Dict[str, Any] = {}
    lignes: List[Dict[str, Any]] = []
    frequences_options: List[str] = []
    prix_force_ht: Optional[float] = None
    taux_tva: float = 0.20
    client_nom: str = ""
    site_adresse: str = ""
    date_emission: str = ""
    # Phase 2B : équipements sélectionnés à illustrer en photo dans le devis
    # [{code, libelle}] — la photo est résolue côté serveur depuis la base.
    equipements: List[Dict[str, Any]] = []
    # Proposition financiere enrichie. Compatibilite : si cette liste est vide,
    # l'ancien prix_force_ht est converti en une ligne simple.
    mode_financier: str = "simple"
    lignes_financieres: List[Dict[str, Any]] = []
    note_financiere: str = ""
    note_tva: str = ""
    # Finition : codes des zones de prestation cochées (récurrent). Les zones non
    # listées ici sont MASQUÉES dans le document (SHOW_* = False).
    zones_selectionnees: List[str] = []
    # Détail par zone pour le tableau financier dynamique :
    # [{code, titre, frequence, option_active, option_libelle, option_frequence, option_prix_ht}]
    # Seules les zones cochées y figurent ; le tableau est construit à partir de ça.
    zones_detail: List[Dict[str, Any]] = []


@router.post("/calculer")
def calculer_prix(payload: DevisPayload, session: Session = Depends(get_session)):
    """Calcule les prix SANS générer de document (aperçu temps réel côté frontend).

    Retourne les totaux HT/TVA/TTC et, pour un contrat récurrent avec
    `frequences_options`, les options comparatives chiffrées.
    """
    _charger_parametres(session)
    t_famille = payload.variables.get("_famille", "")  # optionnel
    recurrent = (t_famille in {"contrat", "bureaux"}) or bool(payload.frequences_options)

    res = P.chiffrer(
        payload.lignes, recurrent=recurrent,
        prix_force_ht=payload.prix_force_ht, taux_tva_global=payload.taux_tva,
    )
    out = {"totaux": res.as_dict()}
    if payload.frequences_options and payload.lignes:
        out["options"] = P.construire_options_recurrentes(
            payload.lignes[0], payload.frequences_options, taux_tva_global=payload.taux_tva
        )
    return out


def _charger_parametres(session) -> dict:
    """Charge les paramètres de calcul depuis la base et les applique à pricing."""
    from models import Parametre
    rows = session.exec(select(Parametre)).all()
    params = {p.cle: p.valeur for p in rows}
    P.appliquer_parametres(params)
    return params


def _resoudre_photos(payload: "DevisPayload", session) -> list:
    """Résout les photos du devis en liste [{libelle, photos:[chemins]}].

    Sources possibles pour chaque entrée de `payload.equipements` :
      - {code}            → photo de l'équipement en base
      - {photo_data}      → photo uploadée à la volée (data URL base64) sauvegardée
      - {libelle/prestation} → texte associé (alignement avec une prestation précise)

    Les entrées partageant la même `prestation` sont REGROUPÉES : plusieurs photos
    s'affichent côte à côte sur la même ligne, donc bien alignées (point 3).
    """
    from models import Equipement
    from config import STORAGE_DIR
    photos_dir = STORAGE_DIR / "photos"
    uploads_dir = photos_dir / "_uploads"
    uploads_dir.mkdir(parents=True, exist_ok=True)

    # Étape 1 : résoudre chaque entrée en (groupe, chemin_photo)
    items = []  # [(prestation_label, libelle_equipement, photo_path, image_meta)]
    for eq in (payload.equipements or []):
        if not isinstance(eq, dict):
            continue
        if eq.get("encart") or eq.get("materiel"):
            continue
        prestation = (eq.get("prestation") or "").strip()
        libelle = (eq.get("libelle") or "").strip()
        photo_path = None
        meta = {
            "image_size": str(eq.get("image_size") or eq.get("taille_image") or eq.get("size") or "grande").strip().lower(),
            "image_align": str(eq.get("image_align") or eq.get("alignement_image") or eq.get("align") or "droite").strip().lower(),
            "image_width_pct": eq.get("image_width_pct") or eq.get("largeur_pct") or eq.get("width_pct") or 40,
        }

        # (a) photo uploadée à la volée
        data_url = eq.get("photo_data")
        if data_url and isinstance(data_url, str):
            photo_path = _normaliser_image_upload(data_url, uploads_dir)

        # (b) photo d'un équipement en base
        explicit_catalog_photo = bool(
            eq.get("as_photo")
            or eq.get("photo_intervention")
            or eq.get("afficher_photo")
            or eq.get("use_as_photo")
        )
        if not photo_path and eq.get("code") and explicit_catalog_photo:
            row = session.exec(select(Equipement).where(Equipement.code == eq["code"])).first()
            if row and row.photo_path and (photos_dir / row.photo_path).exists():
                candidate = photos_dir / row.photo_path
                if not _est_image_logo_entreprise(candidate):
                    photo_path = str(candidate)
                    if not libelle:
                        libelle = row.label

        if photo_path:
            items.append((prestation or libelle or "Equipement", libelle, photo_path, meta))

    # Étape 2 : regrouper par prestation (point 3 : plusieurs photos alignées)
    groupes = {}
    metas = {}
    ordre = []
    for prestation, libelle, path, meta in items:
        if prestation not in groupes:
            groupes[prestation] = []
            metas[prestation] = meta
            ordre.append(prestation)
        groupes[prestation].append(path)

    return [{"libelle": p, "photos": groupes[p], **metas.get(p, {})} for p in ordre]


def _resoudre_materiels(payload: "DevisPayload", session) -> list:
    """RÃ©sout les matÃ©riels Ã  afficher dans l'encart @@MATERIEL_ENCART@@.

    Contrairement aux photos de zone, ces entrÃ©es ne sont jamais injectÃ©es Ã  un
    emplacement devinÃ© : le service DOCX les insÃ¨re uniquement si le marqueur
    explicite existe dans le template.
    """
    from models import Equipement
    from config import STORAGE_DIR

    photos_dir = STORAGE_DIR / "photos"
    resolved = []
    seen = set()
    for eq in (payload.equipements or []):
        if not isinstance(eq, dict) or not (eq.get("encart") or eq.get("materiel")):
            continue
        code = str(eq.get("code") or "").strip()
        if not code or code in seen:
            continue
        row = session.exec(select(Equipement).where(Equipement.code == code)).first()
        if not row or not row.actif:
            continue
        photo_path = None
        if row.photo_path and (photos_dir / row.photo_path).exists():
            candidate = photos_dir / row.photo_path
            if not _est_image_logo_entreprise(candidate):
                photo_path = str(candidate)
        tags = []
        try:
            parsed = json.loads(getattr(row, "tags", "[]") or "[]")
            if isinstance(parsed, list):
                tags = [str(t) for t in parsed if str(t).strip()]
        except Exception:
            tags = []
        resolved.append({
            "code": row.code,
            "label": row.label,
            "categorie": row.categorie,
            "description": row.description or "",
            "photo_path": photo_path,
            "tags": tags,
        })
        seen.add(code)
    return resolved


def _parse_float_fr(value: Any, default: float = 0.0) -> float:
    """Parse un nombre venant de l'interface (1234.5, 1 234,50 €, 20%)."""
    if value is None:
        return default
    if isinstance(value, (int, float)):
        try:
            return float(value)
        except (TypeError, ValueError):
            return default
    raw = str(value).strip()
    if not raw:
        return default
    raw = raw.replace("\u00a0", " ").replace("€", "").replace("%", "")
    raw = raw.replace(" ", "").replace(",", ".")
    try:
        return float(raw)
    except ValueError:
        return default


def _fmt_taux_tva(value: float) -> str:
    pct = value * 100 if value <= 1 else value
    if abs(pct - round(pct)) < 0.001:
        return f"{int(round(pct))} %"
    return f"{pct:.2f}".replace(".", ",").rstrip("0").rstrip(",") + " %"


def _fmt_optional_number(value: Any) -> str:
    n = _parse_float_fr(value, None)
    if n is None:
        return ""
    if abs(n - round(n)) < 0.001:
        return str(int(round(n)))
    return f"{n:.2f}".replace(".", ",")


MOIS_FR = {
    1: "janvier",
    2: "février",
    3: "mars",
    4: "avril",
    5: "mai",
    6: "juin",
    7: "juillet",
    8: "août",
    9: "septembre",
    10: "octobre",
    11: "novembre",
    12: "décembre",
}


def _extraire_date(value: Any) -> date:
    raw = str(value or "").strip()
    raw = re.sub(r"^(?:à\s+)?marseille\s*,?\s*le\s+", "", raw, flags=re.I).strip()
    mois_lookup = {v: k for k, v in MOIS_FR.items()}
    mois_match = re.fullmatch(r"(\d{1,2})\s+([a-zéû]+)\s+(\d{4})", raw.lower())
    if mois_match:
        day = int(mois_match.group(1))
        month = mois_lookup.get(mois_match.group(2))
        year = int(mois_match.group(3))
        if month:
            return date(year, month, day)
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d.%m.%Y"):
        try:
            return datetime.strptime(raw, fmt).date()
        except ValueError:
            pass
    return date.today()


def _format_date_emission_fr(value: Any) -> tuple[str, str]:
    dt = _extraire_date(value)
    courte = f"{dt.day} {MOIS_FR[dt.month]} {dt.year}"
    return courte, f"À Marseille, le {courte}"


def _normaliser_lignes_financieres(payload: "DevisPayload", res: P.ResultatDevis, data: Dict[str, Any]) -> Dict[str, Any]:
    """Construit le tableau financier final, avec compatibilite prix unique."""
    raw_lines = [l for l in (payload.lignes_financieres or []) if isinstance(l, dict)]

    if not raw_lines:
        designation = (
            str(data.get("FORFAIT_LIBELLE") or data.get("OPTION_LIBELLE") or data.get("TYPE_PRESTATION") or "").strip()
            or "Selon descriptif"
        )
        raw_lines = [{
            "designation": designation,
            "total_ht": payload.prix_force_ht if payload.prix_force_ht is not None else res.total_ht,
            "taux_tva": payload.taux_tva,
            "type_ligne": "prestation",
            "inclure_total": True,
        }]

    normalized = []
    total_ht = total_tva = total_ttc = 0.0
    uses_qty = uses_unit = uses_unit_price = uses_description = False

    for idx, line in enumerate(raw_lines):
        designation = str(line.get("designation") or line.get("libelle") or "").strip() or f"Ligne {idx + 1}"
        description = str(line.get("description") or "").strip()
        type_ligne = str(line.get("type_ligne") or line.get("type") or "prestation").strip().lower()
        if type_ligne == "option" or line.get("option") is True:
            is_option = True
        else:
            is_option = False
        include_total = line.get("inclure_total", line.get("include_total", True))
        include_total = False if str(include_total).lower() in {"false", "0", "non", "no"} else bool(include_total)

        qty_raw = line.get("quantite", line.get("qty", ""))
        unit = str(line.get("unite") or line.get("unit") or "").strip()
        pu_raw = line.get("prix_unitaire_ht", line.get("pu_ht", line.get("unit_price_ht", "")))
        qty = _parse_float_fr(qty_raw, 0.0)
        unit_price = _parse_float_fr(pu_raw, 0.0)
        total_line_ht = _parse_float_fr(line.get("total_ht", line.get("ht", "")), 0.0)

        if qty and unit_price:
            total_line_ht = qty * unit_price
        if type_ligne == "remise" and total_line_ht > 0:
            total_line_ht = -total_line_ht

        tva_rate = _parse_float_fr(line.get("taux_tva", line.get("tva_rate", payload.taux_tva)), payload.taux_tva)
        if tva_rate > 1:
            tva_rate = tva_rate / 100
        montant_tva = round(total_line_ht * tva_rate, 2)
        total_line_ttc = round(total_line_ht + montant_tva, 2)

        is_info = type_ligne in {"information", "info"}
        is_total_like = type_ligne in {"sous-total", "sous_total", "subtotal", "total", "total_general"}
        if include_total and not is_info and not is_total_like:
            total_ht += total_line_ht
            total_tva += montant_tva
            total_ttc += total_line_ttc

        uses_qty = uses_qty or bool(str(qty_raw or "").strip())
        uses_unit = uses_unit or bool(unit)
        uses_unit_price = uses_unit_price or bool(str(pu_raw or "").strip())
        uses_description = uses_description or bool(description)

        normalized.append({
            "designation": designation,
            "description": description,
            "quantite": _fmt_optional_number(qty_raw),
            "unite": unit,
            "prix_unitaire_ht": P.fmt_euros(unit_price) if str(pu_raw or "").strip() else "",
            "total_ht": round(total_line_ht, 2),
            "total_ht_fmt": P.fmt_euros(total_line_ht) if not is_info else "",
            "taux_tva": round(tva_rate, 4),
            "taux_tva_fmt": _fmt_taux_tva(tva_rate),
            "montant_tva": round(montant_tva, 2),
            "montant_tva_fmt": P.fmt_euros(montant_tva) if not is_info else "",
            "total_ttc": round(total_line_ttc, 2),
            "total_ttc_fmt": P.fmt_euros(total_line_ttc) if not is_info else "",
            "type_ligne": type_ligne,
            "is_option": is_option,
            "inclure_total": include_total,
            "excluded_from_total": not include_total,
        })

    total_ht = round(total_ht, 2)
    total_tva = round(total_tva, 2)
    total_ttc = round(total_ttc, 2)
    return {
        "mode": str(payload.mode_financier or "simple").strip().lower() or "simple",
        "lines": normalized,
        "total_ht": total_ht,
        "total_tva": total_tva,
        "total_ttc": total_ttc,
        "total_ht_fmt": P.fmt_euros(total_ht),
        "total_tva_fmt": P.fmt_euros(total_tva),
        "total_ttc_fmt": P.fmt_euros(total_ttc),
        "columns": {
            "quantite": uses_qty,
            "unite": uses_unit,
            "prix_unitaire_ht": uses_unit_price,
            "description": uses_description,
        },
        "note": str(payload.note_financiere or "").strip(),
        "note_tva": str(payload.note_tva or "").strip(),
    }


def _construire_data(payload: "DevisPayload", recurrent: bool, session=None):
    """Construit le dict de variables Jinja2 + les totaux à partir du payload.

    Logique partagée entre la génération définitive et l'aperçu.
    """
    data: Dict[str, Any] = dict(payload.variables)
    data.pop("_famille", None)
    data["_FAMILLE_TEMPLATE"] = "contrat" if recurrent else "ponctuel"
    modele_code = str(data.pop("_modele_code", "") or payload.template_code or "").lower()
    modele_code = ALIASES_MODELES_PONCTUELS.get(modele_code, modele_code)
    data["MODELE_CODE"] = modele_code
    date_courte, date_longue = _format_date_emission_fr(data.get("DATE_EMISSION") or payload.date_emission)
    data["DATE_EMISSION"] = date_courte
    data["DATE_EMISSION_LONGUE"] = date_longue
    lignes_calc = payload.lignes
    if not recurrent:
        lignes_calc = _nettoyer_lignes_ponctuelles(modele_code, payload.lignes)
    res = P.chiffrer(
        lignes_calc, recurrent=recurrent,
        prix_force_ht=payload.prix_force_ht, taux_tva_global=payload.taux_tva,
    )
    if not str(data.get("DATE_SIGNATURE", "")).strip():
        data["DATE_SIGNATURE"] = date_courte

    if recurrent:
        zones_detail_by_code = {
            str(z.get("code") or "").lower(): z
            for z in (payload.zones_detail or [])
            if isinstance(z, dict) and str(z.get("code") or "").strip()
        }
        if modele_code == "bureaux_petit" and not str(data.get("TYPE_PRESTATION", "")).strip():
            data["TYPE_PRESTATION"] = "Nettoyage des bureaux"
        elif not str(data.get("TYPE_PRESTATION", "")).strip():
            data["TYPE_PRESTATION"] = "Entretien des parties communes"

        if modele_code == "bureaux_petit":
            zones_bureaux = _charger_zones_bureaux()
            selection = {str(z).lower() for z in (payload.zones_selectionnees or [])}
            for zone in zones_bureaux:
                code = str(zone.get("code", "")).lower()
                key = code.upper()
                freq_var = zone.get("freq_var")
                show = code in selection if selection else bool(freq_var and str(payload.variables.get(freq_var, "")).strip())
                data[f"SHOW_{key}"] = bool(show)
                ops_key = f"OPS_{key}"
                ops = payload.variables.get(ops_key)
                detail = zones_detail_by_code.get(code)
                if detail and isinstance(detail.get("operations"), list):
                    ops = detail.get("operations")
                if not isinstance(ops, list):
                    ops = zone.get("operations", [])
                data[ops_key] = [str(op) for op in ops if str(op).strip()]
                flags = payload.variables.get(f"OPS_ENABLED_{key}")
                if detail and isinstance(detail.get("ops_enabled"), list):
                    flags = detail.get("ops_enabled")
                if isinstance(flags, list):
                    data[f"OPS_ENABLED_{key}"] = flags
                if freq_var and freq_var not in data:
                    data[freq_var] = ""
        else:
            # Zones sélectionnées → flags SHOW_* (les zones non cochées disparaissent du document)
            ZONES = ["HALL", "ASCENSEUR", "ESCALIERS", "CAVES", "GARAGE", "ABORDS", "CONTENEUR", "OM"]
            selection = payload.zones_selectionnees or []
            if selection:
                sel_upper = {str(z).upper() for z in selection}
                for z in ZONES:
                    data[f"SHOW_{z}"] = z in sel_upper
            else:
                # Aucune sélection transmise → on n'affiche que les zones ayant une fréquence renseignée
                for z in ZONES:
                    freq = payload.variables.get(f"FREQ_{z}", "")
                    data[f"SHOW_{z}"] = bool(str(freq).strip())
            for code, detail in zones_detail_by_code.items():
                key = code.upper()
                if isinstance(detail.get("operations"), list):
                    data[f"OPS_{key}"] = [str(op) for op in detail.get("operations") if str(op).strip()]
                if isinstance(detail.get("ops_enabled"), list):
                    data[f"OPS_ENABLED_{key}"] = detail.get("ops_enabled")
        # Tableau financier : DYNAMIQUE à partir des zones réellement cochées
        # (et leurs options activées). On ne montre QUE ce qui est sélectionné.
        if payload.zones_detail:
            base_l = payload.lignes[0] if payload.lignes else {"duree_h": 2, "nb_agents": 1, "taux_horaire": 24}
            data["OPTIONS"] = P.construire_tableau_zones(
                payload.zones_detail, base_l, taux_tva_global=payload.taux_tva
            )
        elif payload.frequences_options and payload.lignes:
            data["OPTIONS"] = P.construire_options_recurrentes(
                payload.lignes[0], payload.frequences_options, taux_tva_global=payload.taux_tva
            )
        elif "OPTIONS" not in data:
            data["OPTIONS"] = [{
                "libelle": payload.variables.get("OPTION_LIBELLE", "Prestation"),
                "ht": res.total_ht_fmt, "tva": res.total_tva_fmt, "ttc": res.total_ttc_fmt,
            }]
    else:
        if "PRESTATIONS" not in data:
            data["PRESTATIONS"] = [l.get("libelle", "") for l in lignes_calc if l.get("libelle")]
        data.setdefault("FORFAIT_LIBELLE", payload.variables.get("FORFAIT_LIBELLE", "Forfait prestation"))
        modeles_encombrants = {"encombrants_caves", "encombrants_divers"}
        is_encombrants = modele_code in modeles_encombrants
        if is_encombrants:
            data.setdefault(
                "MENTION_VALIDITE",
                "Compte tenu de l'évolution potentielle des dépôts d'encombrants, le devis est valable 15 jours",
            )
            data.setdefault(
                "MENTION_SPECIFIQUE",
                "La société MARIE EUGENIE ne pourra être aucunement tenue responsable de l'évacuation d'objets situés dans les parties communes ; le client ne peut entamer le moindre recours contre la société MARIE EUGENIE dans le cadre de l'enlèvement et de la destruction des objets présents dans les parties communes.",
            )
        else:
            data.setdefault("MENTION_VALIDITE", "")
            data.setdefault("MENTION_SPECIFIQUE", "")
        data["FORFAIT_HT"] = res.total_ht_fmt
        data["FORFAIT_TVA"] = res.total_tva_fmt
        data["FORFAIT_TTC"] = res.total_ttc_fmt
    financial = _normaliser_lignes_financieres(payload, res, data)
    res.total_ht = financial["total_ht"]
    res.total_tva = financial["total_tva"]
    res.total_ttc = financial["total_ttc"]
    res.total_ht_fmt = financial["total_ht_fmt"]
    res.total_tva_fmt = financial["total_tva_fmt"]
    res.total_ttc_fmt = financial["total_ttc_fmt"]
    data["FINANCIAL_MODE"] = financial["mode"]
    data["FINANCIAL_LINES"] = financial["lines"]
    data["FINANCIAL_COLUMNS"] = financial["columns"]
    data["FINANCIAL_NOTE"] = financial["note"]
    data["TVA_NOTE"] = financial["note_tva"]
    data["TOTAL_HT"] = financial["total_ht_fmt"]
    data["TOTAL_TVA"] = financial["total_tva_fmt"]
    data["TOTAL_TTC"] = financial["total_ttc_fmt"]
    data["FORFAIT_HT"] = financial["total_ht_fmt"]
    data["FORFAIT_TVA"] = financial["total_tva_fmt"]
    data["FORFAIT_TTC"] = financial["total_ttc_fmt"]
    data.setdefault("PRIX_HT", financial["total_ht_fmt"])
    data.setdefault("TVA", financial["total_tva_fmt"])
    data.setdefault("PRIX_TTC", financial["total_ttc_fmt"])
    # Phase 2B : photos des équipements sélectionnés (texte + image alignés)
    if session is not None:
        photos = _resoudre_photos(payload, session)
        if photos:
            data["PRESTATIONS_PHOTOS"] = photos
        materiels = _resoudre_materiels(payload, session)
        if materiels:
            data["MATERIEL_SELECTIONNE"] = materiels

    return data, res


def _date_signature_depuis_emission(value: Any) -> str:
    """Retourne une date jj/mm/aaaa pour le bloc Bon pour accord."""
    dt = _extraire_date(value)
    return dt.strftime("%d/%m/%Y")


@router.post("/apercu")
def apercu_devis(payload: DevisPayload, session: Session = Depends(get_session)):
    """
    Génère un APERÇU PDF RÉEL à partir du modèle Word maître, sans le persister
    comme devis définitif.

    C'est cet aperçu — le vrai document Word converti en PDF — qui doit être
    affiché dans l'interface, et non une reconstruction HTML. Le rendu est donc
    strictement celui du modèle fourni, seules les variables étant remplacées.
    """
    template_code = _resoudre_template_code(payload)
    t = session.exec(select(Template).where(Template.code == template_code)).first()
    if not t:
        if template_code == "bureaux_important":
            raise HTTPException(404, "Modèle source Bureaux important manquant")
        raise HTTPException(404, f"Template '{template_code}' introuvable")
    template_path = TEMPLATES_DIR / t.fichier
    if not template_path.exists():
        raise HTTPException(500, f"Fichier template manquant : {t.fichier}")

    _charger_parametres(session)
    recurrent = (t.famille in {"contrat", "bureaux"})
    data, _ = _construire_data(payload, recurrent, session)

    # Générer dans un dossier d'aperçus temporaire (écrasé à chaque fois)
    apercu_dir = GENERATED_DIR / "_apercus"
    apercu_dir.mkdir(parents=True, exist_ok=True)
    # Nom de fichier UNIQUE par requête : évite la corruption quand plusieurs
    # aperçus sont générés en parallèle (cause du "Bad magic number").
    import uuid as _uuid
    safe = "".join(c for c in (payload.numero or "apercu") if c.isalnum() or c in "-_") or "apercu"
    uniq = _uuid.uuid4().hex[:8]
    out_docx = apercu_dir / f"{safe}_{uniq}.docx"
    try:
        gen_devis_file(template_path, data, out_docx)
        out_pdf = docx_to_pdf(out_docx, apercu_dir)
    except Exception as e:
        raise HTTPException(500, f"Aperçu PDF impossible : {e}")
    # Nettoyer les anciens aperçus pour ne pas accumuler (garder le courant)
    try:
        import time as _t
        for f in apercu_dir.glob("*"):
            if f != out_pdf and f != out_docx and (_t.time() - f.stat().st_mtime) > 120:
                f.unlink()
    except Exception:
        pass

    return FileResponse(out_pdf, media_type="application/pdf",
                        filename=f"apercu_{safe}.pdf",
                        headers={"Cache-Control": "no-store"})


@router.post("/generer")
def generer_devis(payload: DevisPayload, session: Session = Depends(get_session)):
    """
    Génère un devis .docx + .pdf à partir du template choisi, avec calcul
    automatique des prix.
    """
    # Charger le template
    template_code = _resoudre_template_code(payload)
    t = session.exec(select(Template).where(Template.code == template_code)).first()
    if not t:
        if template_code == "bureaux_important":
            raise HTTPException(404, "Modèle source Bureaux important manquant")
        raise HTTPException(404, f"Template '{template_code}' introuvable")
    if not t.actif:
        raise HTTPException(400, f"Template '{payload.template_code}' désactivé")

    template_path = TEMPLATES_DIR / t.fichier
    if not template_path.exists():
        raise HTTPException(500, f"Fichier template manquant sur disque : {t.fichier}")

    _charger_parametres(session)
    recurrent = (t.famille in {"contrat", "bureaux"})

    # --- Calcul des prix + variables (logique partagée avec l'aperçu) ---
    data, res = _construire_data(payload, recurrent, session)

    # Générer le .docx
    output_docx = GENERATED_DIR / f"{payload.numero}.docx"
    try:
        gen_devis_file(template_path, data, output_docx)
    except Exception as e:
        raise HTTPException(500, f"Erreur de génération .docx : {e}")

    # Générer le PDF.
    # Correctif : on ne masque plus l'échec derrière un pdf_url null silencieux.
    # docx_to_pdf gère déjà warm-up + retry unique ; s'il échoue encore, on
    # remonte une erreur explicite contenant le message exact de LibreOffice.
    try:
        output_pdf = docx_to_pdf(output_docx, GENERATED_DIR)
    except Exception as e:
        raise HTTPException(500, f"Conversion PDF impossible : {e}")

    # Diagnostic : nombre de pages + détection de page blanche (page < 8 mots)
    pdf_pages = None
    page_blanche = None
    if output_pdf:
        try:
            pdf_pages, page_blanche = _analyser_pdf(output_pdf)
        except Exception as e:
            print(f"[warn] analyse PDF impossible : {e}")

    # Persister en DB
    existing = session.exec(select(Devis).where(Devis.numero == payload.numero)).first()
    snapshot = json.dumps(payload.model_dump(), default=str)
    if existing:
        d = existing
        d.template_code = template_code
        d.client_nom = payload.client_nom
        d.site_adresse = payload.site_adresse
        d.date_emission = payload.date_emission
        d.montant_ht = res.total_ht
        d.montant_tva = res.total_tva
        d.montant_ttc = res.total_ttc
        d.prix_override = res.override
        d.fichier_docx = output_docx.name
        d.fichier_pdf = output_pdf.name if output_pdf else None
        d.payload = snapshot
        d.updated_at = datetime.utcnow()
    else:
        d = Devis(
            numero=payload.numero,
            template_code=template_code,
            client_nom=payload.client_nom,
            site_adresse=payload.site_adresse,
            date_emission=payload.date_emission,
            montant_ht=res.total_ht,
            montant_tva=res.total_tva,
            montant_ttc=res.total_ttc,
            prix_override=res.override,
            fichier_docx=output_docx.name,
            fichier_pdf=output_pdf.name if output_pdf else None,
            payload=snapshot,
        )
        session.add(d)
    session.commit()

    return {
        "ok": True,
        "numero": payload.numero,
        "totaux": res.as_dict(),
        "docx_url": f"/api/devis/{payload.numero}/docx",
        "pdf_url": f"/api/devis/{payload.numero}/pdf" if output_pdf else None,
        "pdf_pages": pdf_pages,
        "page_blanche": page_blanche,
    }


@router.get("")
def list_devis(session: Session = Depends(get_session)):
    """Liste tous les devis générés."""
    devis = session.exec(select(Devis).order_by(Devis.created_at.desc())).all()
    return [
        {
            "id": d.id, "numero": d.numero, "template_code": d.template_code,
            "client_nom": d.client_nom, "site_adresse": d.site_adresse,
            "date_emission": d.date_emission,
            "montant_ht": d.montant_ht, "montant_tva": d.montant_tva,
            "montant_ttc": d.montant_ttc, "prix_override": d.prix_override,
            "statut": d.statut,
            "has_docx": bool(d.fichier_docx),
            "has_pdf": bool(d.fichier_pdf),
            "created_at": d.created_at.isoformat(),
        } for d in devis
    ]


@router.get("/{numero}/docx")
def download_devis_docx(numero: str, session: Session = Depends(get_session)):
    d = session.exec(select(Devis).where(Devis.numero == numero)).first()
    if not d or not d.fichier_docx:
        raise HTTPException(404, "Devis ou fichier introuvable")
    fp = GENERATED_DIR / d.fichier_docx
    if not fp.exists():
        raise HTTPException(404, "Fichier disque manquant")
    safe_numero = "".join(c for c in numero if c.isalnum() or c in ("-", "_")) or "devis"
    return FileResponse(
        fp,
        filename=f"{safe_numero}.docx",
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={
            "Content-Disposition": f'attachment; filename="{safe_numero}.docx"',
            "X-Content-Type-Options": "nosniff",
        },
    )


@router.get("/{numero}/pdf")
def download_devis_pdf(numero: str, session: Session = Depends(get_session)):
    d = session.exec(select(Devis).where(Devis.numero == numero)).first()
    if not d or not d.fichier_pdf:
        raise HTTPException(404, "Devis ou PDF introuvable")
    fp = GENERATED_DIR / d.fichier_pdf
    if not fp.exists():
        raise HTTPException(404, "Fichier disque manquant")
    return FileResponse(fp, filename=f"{numero}.pdf", media_type="application/pdf")

"""
Routes API pour la bibliothèque métier et la gestion DevisFlow :
  - /clients        : carnet de clients (CRUD + recherche + archivage)
  - /equipements    : matériels / équipements / véhicules (CRUD + photo)
  - /prestations    : zones de prestations types (issues du modèle Word)
  - /membres        : équipe (CRUD léger)
  - /parametres     : paramètres de calcul et de génération

Ces routes alimentent les vues riches de l'interface (Clients, Bibliothèque
métier, Équipe, Paramètres) et la création de devis.
"""
import json
import shutil
import uuid
import io
from pathlib import Path
from typing import Optional, List

from fastapi import APIRouter, HTTPException, Depends, UploadFile, File, Form
from fastapi.responses import FileResponse
from pydantic import BaseModel
from sqlmodel import Session, select

from config import STORAGE_DIR
from models import Client, Equipement, PrestationType, Membre, Parametre, get_session

router = APIRouter(tags=["bibliotheque"])

PHOTOS_DIR = STORAGE_DIR / "photos"
PHOTOS_DIR.mkdir(parents=True, exist_ok=True)


def _sauver_photo_equipement(photo: UploadFile, code: str) -> Optional[str]:
    """Convertit une photo d'équipement en JPEG pour garantir l'insertion Word/PDF."""
    if not photo or not photo.filename:
        return None
    raw = photo.file.read()
    try:
        try:
            import pillow_heif  # type: ignore
            pillow_heif.register_heif_opener()
        except Exception:
            pass
        from PIL import Image, ImageOps
        img = Image.open(io.BytesIO(raw))
        img = ImageOps.exif_transpose(img)
        if img.mode != "RGB":
            img = img.convert("RGB")
        max_side = 1800
        if max(img.size) > max_side:
            img.thumbnail((max_side, max_side))
        fname = f"{code}_{uuid.uuid4().hex[:8]}.jpg"
        img.save(PHOTOS_DIR / fname, format="JPEG", quality=88, optimize=True)
        return fname
    except Exception as exc:
        raise HTTPException(400, f"Photo non convertible en JPG/PNG exploitable : {exc}")


# ============================================================
# CLIENTS
# ============================================================
class ClientIn(BaseModel):
    nom: str
    civilite: Optional[str] = None
    contact: Optional[str] = None
    email: Optional[str] = None
    telephone: Optional[str] = None
    adresse: Optional[str] = None
    code_postal: Optional[str] = None
    ville: Optional[str] = None
    site_nom: Optional[str] = None
    site_adresse: Optional[str] = None


def _client_dict(c: Client):
    return {
        "id": c.id, "nom": c.nom, "civilite": c.civilite, "contact": c.contact,
        "email": c.email, "telephone": c.telephone, "adresse": c.adresse,
        "code_postal": c.code_postal, "ville": c.ville,
        "site_nom": c.site_nom, "site_adresse": c.site_adresse,
        "archive": c.archive,
    }


@router.get("/clients")
def list_clients(q: Optional[str] = None, inclure_archives: bool = False,
                 session: Session = Depends(get_session)):
    """Liste les clients. `q` filtre par nom/contact/ville ; archives masquées par défaut."""
    stmt = select(Client)
    if not inclure_archives:
        stmt = stmt.where(Client.archive == False)  # noqa: E712
    clients = session.exec(stmt.order_by(Client.nom)).all()
    if q:
        ql = q.lower()
        clients = [c for c in clients if ql in (c.nom or "").lower()
                   or ql in (c.contact or "").lower() or ql in (c.ville or "").lower()]
    return [_client_dict(c) for c in clients]


@router.get("/clients/{client_id}")
def get_client(client_id: int, session: Session = Depends(get_session)):
    c = session.get(Client, client_id)
    if not c:
        raise HTTPException(404, "Client introuvable")
    return _client_dict(c)


@router.post("/clients")
def create_client(data: ClientIn, session: Session = Depends(get_session)):
    c = Client(**data.model_dump())
    session.add(c)
    session.commit()
    session.refresh(c)
    return _client_dict(c)


@router.put("/clients/{client_id}")
def update_client(client_id: int, data: ClientIn, session: Session = Depends(get_session)):
    c = session.get(Client, client_id)
    if not c:
        raise HTTPException(404, "Client introuvable")
    for k, v in data.model_dump().items():
        setattr(c, k, v)
    from datetime import datetime
    c.updated_at = datetime.utcnow()
    session.add(c)
    session.commit()
    return _client_dict(c)


@router.post("/clients/{client_id}/archive")
def archive_client(client_id: int, archive: bool = True, session: Session = Depends(get_session)):
    c = session.get(Client, client_id)
    if not c:
        raise HTTPException(404, "Client introuvable")
    c.archive = archive
    session.add(c)
    session.commit()
    return {"ok": True, "archive": c.archive}


@router.delete("/clients/{client_id}")
def delete_client(client_id: int, session: Session = Depends(get_session)):
    c = session.get(Client, client_id)
    if not c:
        raise HTTPException(404, "Client introuvable")
    session.delete(c)
    session.commit()
    return {"ok": True}


# ============================================================
# EQUIPEMENTS / MATERIEL
# ============================================================
def _tags_to_list(value) -> List[str]:
    if value is None:
        return []
    if isinstance(value, list):
        raw = value
    else:
        text = str(value or "").strip()
        if not text:
            return []
        try:
            parsed = json.loads(text)
            raw = parsed if isinstance(parsed, list) else [text]
        except Exception:
            raw = [x.strip() for x in text.split(",")]
    return [str(x).strip() for x in raw if str(x).strip()]


def _tags_to_storage(value) -> str:
    return json.dumps(_tags_to_list(value), ensure_ascii=False)


def _equip_dict(e: Equipement):
    return {
        "id": e.id, "code": e.code, "label": e.label, "categorie": e.categorie,
        "description": e.description, "actif": e.actif,
        "tags": _tags_to_list(getattr(e, "tags", "[]")),
        "photo_url": f"/api/equipements/{e.code}/photo" if e.photo_path else None,
    }


@router.get("/equipements")
def list_equipements(categorie: Optional[str] = None, q: Optional[str] = None,
                     tag: Optional[str] = None, inclure_inactifs: bool = True,
                     session: Session = Depends(get_session)):
    stmt = select(Equipement)
    if categorie:
        stmt = stmt.where(Equipement.categorie == categorie)
    if not inclure_inactifs:
        stmt = stmt.where(Equipement.actif == True)  # noqa: E712
    items = session.exec(stmt.order_by(Equipement.label)).all()
    if q:
        ql = q.strip().lower()
        items = [
            e for e in items
            if ql in (e.code or "").lower()
            or ql in (e.label or "").lower()
            or ql in (e.description or "").lower()
            or any(ql in t.lower() for t in _tags_to_list(getattr(e, "tags", "[]")))
        ]
    if tag:
        tl = tag.strip().lower()
        items = [e for e in items if tl in {t.lower() for t in _tags_to_list(getattr(e, "tags", "[]"))}]
    return [_equip_dict(e) for e in items]


@router.post("/equipements")
async def create_equipement(
    code: str = Form(...), label: str = Form(...), categorie: str = Form("materiel"),
    description: Optional[str] = Form(None), tags: Optional[str] = Form(None),
    photo: Optional[UploadFile] = File(None),
    session: Session = Depends(get_session),
):
    if session.exec(select(Equipement).where(Equipement.code == code)).first():
        raise HTTPException(409, f"Équipement '{code}' existe déjà")
    photo_path = None
    if photo and photo.filename:
        photo_path = _sauver_photo_equipement(photo, code)
    e = Equipement(code=code, label=label, categorie=categorie,
                   description=description, tags=_tags_to_storage(tags),
                   photo_path=photo_path)
    session.add(e)
    session.commit()
    session.refresh(e)
    return _equip_dict(e)


@router.put("/equipements/{code}")
async def update_equipement(
    code: str, label: Optional[str] = Form(None), categorie: Optional[str] = Form(None),
    description: Optional[str] = Form(None), tags: Optional[str] = Form(None),
    actif: Optional[bool] = Form(None),
    photo: Optional[UploadFile] = File(None), session: Session = Depends(get_session),
):
    e = session.exec(select(Equipement).where(Equipement.code == code)).first()
    if not e:
        raise HTTPException(404, "Équipement introuvable")
    if label is not None: e.label = label
    if categorie is not None: e.categorie = categorie
    if description is not None: e.description = description
    if tags is not None: e.tags = _tags_to_storage(tags)
    if actif is not None: e.actif = actif
    if photo and photo.filename:
        e.photo_path = _sauver_photo_equipement(photo, code)
    from datetime import datetime
    e.updated_at = datetime.utcnow()
    session.add(e)
    session.commit()
    return _equip_dict(e)


@router.delete("/equipements/{code}")
def delete_equipement(code: str, session: Session = Depends(get_session)):
    e = session.exec(select(Equipement).where(Equipement.code == code)).first()
    if not e:
        raise HTTPException(404, "Équipement introuvable")
    session.delete(e)
    session.commit()
    return {"ok": True}


@router.get("/equipements/{code}/photo")
def equipement_photo(code: str, session: Session = Depends(get_session)):
    e = session.exec(select(Equipement).where(Equipement.code == code)).first()
    if not e or not e.photo_path:
        raise HTTPException(404, "Photo introuvable")
    fp = PHOTOS_DIR / e.photo_path
    if not fp.exists():
        raise HTTPException(404, "Fichier photo manquant")
    return FileResponse(fp)


# ============================================================
# PRESTATIONS TYPES (zones)
# ============================================================
def _presta_dict(p: PrestationType):
    return {
        "id": p.id, "code": p.code, "titre": p.titre, "famille": p.famille,
        "freq_var": p.freq_var, "operations": json.loads(p.operations or "[]"),
        "ordre": p.ordre, "actif": p.actif,
    }


@router.get("/prestations")
def list_prestations(famille: Optional[str] = None, session: Session = Depends(get_session)):
    stmt = select(PrestationType)
    if famille:
        stmt = stmt.where(PrestationType.famille == famille)
    items = session.exec(stmt.order_by(PrestationType.ordre)).all()
    return [_presta_dict(p) for p in items]


class PrestationIn(BaseModel):
    code: str
    titre: str
    famille: str = "contrat"
    freq_var: Optional[str] = None
    operations: List[str] = []
    ordre: int = 0


@router.post("/prestations")
def create_prestation(data: PrestationIn, session: Session = Depends(get_session)):
    if session.exec(select(PrestationType).where(PrestationType.code == data.code)).first():
        raise HTTPException(409, f"Prestation '{data.code}' existe déjà")
    p = PrestationType(code=data.code, titre=data.titre, famille=data.famille,
                       freq_var=data.freq_var, operations=json.dumps(data.operations),
                       ordre=data.ordre)
    session.add(p)
    session.commit()
    session.refresh(p)
    return _presta_dict(p)


@router.put("/prestations/{code}")
def update_prestation(code: str, data: PrestationIn, session: Session = Depends(get_session)):
    p = session.exec(select(PrestationType).where(PrestationType.code == code)).first()
    if not p:
        raise HTTPException(404, "Prestation introuvable")
    p.titre = data.titre
    p.famille = data.famille
    p.freq_var = data.freq_var
    p.operations = json.dumps(data.operations)
    p.ordre = data.ordre
    session.add(p)
    session.commit()
    return _presta_dict(p)


@router.delete("/prestations/{code}")
def delete_prestation(code: str, session: Session = Depends(get_session)):
    p = session.exec(select(PrestationType).where(PrestationType.code == code)).first()
    if not p:
        raise HTTPException(404, "Prestation introuvable")
    session.delete(p)
    session.commit()
    return {"ok": True}


# ============================================================
# MEMBRES (équipe)
# ============================================================
def _membre_dict(m: Membre):
    return {"id": m.id, "nom": m.nom, "email": m.email, "role": m.role, "actif": m.actif}


class MembreIn(BaseModel):
    nom: str
    email: Optional[str] = None
    role: Optional[str] = None
    actif: bool = True


@router.get("/membres")
def list_membres(session: Session = Depends(get_session)):
    return [_membre_dict(m) for m in session.exec(select(Membre).order_by(Membre.nom)).all()]


@router.post("/membres")
def create_membre(data: MembreIn, session: Session = Depends(get_session)):
    m = Membre(**data.model_dump())
    session.add(m)
    session.commit()
    session.refresh(m)
    return _membre_dict(m)


@router.put("/membres/{membre_id}")
def update_membre(membre_id: int, data: MembreIn, session: Session = Depends(get_session)):
    m = session.get(Membre, membre_id)
    if not m:
        raise HTTPException(404, "Membre introuvable")
    for k, v in data.model_dump().items():
        setattr(m, k, v)
    session.add(m)
    session.commit()
    return _membre_dict(m)


@router.delete("/membres/{membre_id}")
def delete_membre(membre_id: int, session: Session = Depends(get_session)):
    m = session.get(Membre, membre_id)
    if not m:
        raise HTTPException(404, "Membre introuvable")
    session.delete(m)
    session.commit()
    return {"ok": True}


# ============================================================
# PARAMETRES
# ============================================================
@router.get("/parametres")
def list_parametres(session: Session = Depends(get_session)):
    items = session.exec(select(Parametre).order_by(Parametre.groupe, Parametre.cle)).all()
    return [{"cle": p.cle, "valeur": p.valeur, "libelle": p.libelle, "groupe": p.groupe} for p in items]


class ParametreIn(BaseModel):
    valeur: str


@router.put("/parametres/{cle}")
def update_parametre(cle: str, data: ParametreIn, session: Session = Depends(get_session)):
    p = session.exec(select(Parametre).where(Parametre.cle == cle)).first()
    if not p:
        raise HTTPException(404, "Paramètre introuvable")
    p.valeur = data.valeur
    session.add(p)
    session.commit()
    return {"ok": True, "cle": cle, "valeur": p.valeur}

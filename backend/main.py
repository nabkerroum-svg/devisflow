"""
DevisFlow — backend FastAPI.

Lance le serveur avec :
    uvicorn main:app --reload --port 8000
"""
from contextlib import asynccontextmanager
from pathlib import Path

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse

from config import API_PREFIX, CORS_ORIGINS, FRONTEND_DIR
from models import init_db
from routes_templates import router as templates_router
from routes_devis import router as devis_router
from routes_biblio import router as biblio_router


@asynccontextmanager
async def lifespan(app: FastAPI):
    # Au démarrage : créer les tables + seed initial si DB vide
    init_db()
    try:
        from seed import seed_initial
        seed_initial()
    except Exception as e:
        print(f"[seed] {e}")
    try:
        from seed_metier import seed_metier
        seed_metier()
    except Exception as e:
        print(f"[seed_metier] {e}")
    yield


app = FastAPI(
    title="DevisFlow API",
    description="Génération de devis Word à partir de templates maîtres",
    version="1.0.0",
    lifespan=lifespan,
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=CORS_ORIGINS,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


# Toute erreur non gérée sous /api renvoie du JSON propre (jamais du HTML
# "Internal Server Error" que le frontend ne saurait pas parser).
from fastapi import Request
from fastapi.responses import JSONResponse


@app.exception_handler(Exception)
async def json_exception_handler(request: Request, exc: Exception):
    if request.url.path.startswith(API_PREFIX):
        return JSONResponse(status_code=500, content={
            "error": True,
            "message": "Erreur serveur",
            "details": str(exc),
        })
    raise exc

# Routers API
app.include_router(templates_router, prefix=API_PREFIX)
app.include_router(devis_router, prefix=API_PREFIX)
app.include_router(biblio_router, prefix=API_PREFIX)


@app.get(f"{API_PREFIX}/health")
def health():
    return {"status": "ok"}


@app.get(f"{API_PREFIX}/version")
def version():
    """Diagnostic : confirme que le code servi est bien la dernière version."""
    import zipfile as _zip
    info = {"version": "DevisFlow 11 — synchronisation interface/devis — 2026-06-23",
            "base": "Devis_flow_2 + correctifs", "checks": {}}
    try:
        bridge = (FRONTEND_DIR / "static" / "api-bridge.js").read_text(encoding="utf-8")
        info["checks"]["mapping_zones"] = "ZONE_CODE_MAP" in bridge
        info["checks"]["apercu_pdfjs"] = "_renderPdfCanvas" in bridge
        info["checks"]["photos_par_zone"] = "z.photo" in bridge
        info["checks"]["panneau_doublon_retire"] = "_installZonePicker_OFF" in bridge
        info["checks"]["badge_version_diagnostic"] = "installVersionBadge" in bridge
    except Exception as e:
        info["checks"]["bridge"] = f"ERREUR: {e}"
    try:
        idx = (FRONTEND_DIR / "index_rich.html").read_text(encoding="utf-8")
        info["checks"]["index_rich_a_jour"] = "ZONE_CODE_MAP" in idx and "installVersionBadge" in idx
    except Exception as e:
        info["checks"]["index_rich"] = f"ERREUR: {e}"
    try:
        from config import SEED_DIR
        x = _zip.ZipFile(SEED_DIR / "copro_petite.docx").read("word/document.xml").decode("utf-8")
        info["checks"]["template_conditionnels (8)"] = x.count("{%p if SHOW_")
        # sauts de page = sauts manuels (w:br type=page) + pageBreakBefore
        info["checks"]["sauts_de_page"] = x.count('w:type="page"') + x.count("<w:pageBreakBefore/>")
        info["checks"]["pageBreakBefore (doit etre 0)"] = x.count("<w:pageBreakBefore/>")
    except Exception as e:
        info["checks"]["template"] = f"ERREUR: {e}"
    info["checks"]["pdfjs_bundle"] = (FRONTEND_DIR / "static" / "vendor" / "pdf.min.js").exists()
    _intchecks = {"template_conditionnels (8)": 8}
    def _ok(k, v):
        if k in _intchecks: return v == _intchecks[k]
        if k in ("sauts_de_page", "pageBreakBefore (doit etre 0)"): return True  # informatif
        return v is True
    info["tout_ok"] = all(_ok(k, v) for k, v in info["checks"].items())
    return info


# === Servir le frontend ===
# Le frontend est en HTML/CSS/JS statique dans /frontend
if FRONTEND_DIR.exists():
    app.mount("/static", StaticFiles(directory=str(FRONTEND_DIR / "static")), name="static")

    @app.get("/")
    def root():
        rich = FRONTEND_DIR / "index_rich.html"
        if rich.exists():
            return FileResponse(rich, media_type="text/html; charset=utf-8")
        return FileResponse(FRONTEND_DIR / "index.html", media_type="text/html; charset=utf-8")

    @app.get("/{path:path}")
    def spa_fallback(path: str):
        # Si le chemin demandé existe en statique, le servir
        candidate = FRONTEND_DIR / path
        if candidate.exists() and candidate.is_file():
            if candidate.suffix.lower() == ".html":
                return FileResponse(candidate, media_type="text/html; charset=utf-8")
            return FileResponse(candidate)
        # Sinon, retomber sur index.html (SPA)
        return FileResponse(FRONTEND_DIR / "index.html", media_type="text/html; charset=utf-8")

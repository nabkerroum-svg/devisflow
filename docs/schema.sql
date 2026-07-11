-- ============================================================
-- DevisFlow — Schéma SQL
-- ============================================================
-- Ce fichier est fourni à titre de référence. Les tables sont créées
-- automatiquement par SQLModel au démarrage de l'application (voir
-- backend/models.py).
--
-- Compatible : SQLite, PostgreSQL (avec quelques ajustements de types
-- mineurs : INTEGER → SERIAL pour les PK, BOOLEAN natif, etc.)
-- ============================================================

-- ------------------------------------------------------------
-- Table : template
-- ------------------------------------------------------------
-- Stocke les modèles Word maîtres uploadés par l'admin.
-- Le fichier .docx physique est dans backend/storage/templates/
CREATE TABLE IF NOT EXISTS template (
    id                  INTEGER PRIMARY KEY AUTOINCREMENT,
    code                VARCHAR NOT NULL UNIQUE,
    nom                 VARCHAR NOT NULL,
    famille             VARCHAR DEFAULT 'contrat',       -- contrat | ponctuel
    fichier             VARCHAR NOT NULL,                -- nom du .docx dans storage/templates/
    type_intervention   VARCHAR,
    is_default          BOOLEAN DEFAULT 0,
    actif               BOOLEAN DEFAULT 1,
    variables           TEXT,                            -- JSON-encoded list des marqueurs détectés
    created_at          DATETIME DEFAULT CURRENT_TIMESTAMP,
    updated_at          DATETIME DEFAULT CURRENT_TIMESTAMP
);
CREATE INDEX IF NOT EXISTS ix_template_code ON template(code);

-- ------------------------------------------------------------
-- Table : devis
-- ------------------------------------------------------------
-- Trace les devis générés. Référence un template (par code, pas par FK
-- pour souplesse). Les fichiers .docx/.pdf sont dans storage/generated/
CREATE TABLE IF NOT EXISTS devis (
    id                  INTEGER PRIMARY KEY AUTOINCREMENT,
    numero              VARCHAR NOT NULL UNIQUE,
    template_code       VARCHAR NOT NULL,
    client_nom          VARCHAR NOT NULL DEFAULT '',
    site_adresse        VARCHAR NOT NULL DEFAULT '',
    date_emission       VARCHAR DEFAULT '',
    montant_ht          REAL DEFAULT 0.0,
    montant_ttc         REAL DEFAULT 0.0,
    fichier_docx        VARCHAR,
    fichier_pdf         VARCHAR,
    payload             TEXT,                            -- JSON snapshot des données saisies
    statut              VARCHAR DEFAULT 'brouillon',     -- brouillon | envoye | accepte
    created_at          DATETIME DEFAULT CURRENT_TIMESTAMP,
    updated_at          DATETIME DEFAULT CURRENT_TIMESTAMP
);
CREATE INDEX IF NOT EXISTS ix_devis_numero ON devis(numero);

-- ------------------------------------------------------------
-- Table : client
-- ------------------------------------------------------------
-- Carnet de clients (non encore branché au formulaire de devis dans la v1)
CREATE TABLE IF NOT EXISTS client (
    id                  INTEGER PRIMARY KEY AUTOINCREMENT,
    nom                 VARCHAR NOT NULL,
    civilite            VARCHAR,
    contact             VARCHAR,
    email               VARCHAR,
    telephone           VARCHAR,
    adresse             VARCHAR,
    code_postal         VARCHAR,
    ville               VARCHAR,
    site_nom            VARCHAR,
    site_adresse        VARCHAR,
    created_at          DATETIME DEFAULT CURRENT_TIMESTAMP
);
CREATE INDEX IF NOT EXISTS ix_client_nom ON client(nom);

-- ------------------------------------------------------------
-- Table : equipement
-- ------------------------------------------------------------
-- Bibliothèque d'équipements avec photo (non encore branchée à la
-- génération Word dans la v1 - prévu via docxtpl.InlineImage)
CREATE TABLE IF NOT EXISTS equipement (
    id                  INTEGER PRIMARY KEY AUTOINCREMENT,
    code                VARCHAR NOT NULL UNIQUE,
    label               VARCHAR NOT NULL,
    categorie           VARCHAR DEFAULT 'materiel',      -- machine | materiel | vehicule | specifique
    description         VARCHAR,
    photo_path          VARCHAR,                          -- chemin vers storage/photos/
    actif               BOOLEAN DEFAULT 1,
    created_at          DATETIME DEFAULT CURRENT_TIMESTAMP
);
CREATE INDEX IF NOT EXISTS ix_equipement_code ON equipement(code);

-- ============================================================
-- Seed : template Marie Eugénie (chargé automatiquement par seed.py)
-- ============================================================
-- INSERT INTO template (code, nom, famille, fichier, type_intervention, is_default, actif, variables)
-- VALUES (
--   'copro_petite',
--   'Copropriété — petite surface (Marie Eugénie)',
--   'contrat',
--   'copro_petite.docx',
--   'Entretien des parties communes',
--   1,
--   1,
--   '["DATE_EMISSION", "DEST_LIGNE1", "DEST_LIGNE2", "DEST_LIGNE3", "NUMERO_DEVIS", "SITE_ADRESSE", "SITE_CP_VILLE", "TYPE_PRESTATION"]'
-- );

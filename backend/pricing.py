"""
Moteur de calcul des prix DevisFlow.

Calcule automatiquement HT / TVA / TTC pour chaque ligne de prestation à partir
de paramètres métier (durée, nombre d'agents, taux horaire, technicité, frais),
puis agrège selon le type de devis :

  - récurrent  → prix MENSUEL (la fréquence est convertie en passages/mois)
  - ponctuel   → prix FORFAITAIRE (somme simple des lignes)

Toutes les valeurs renvoyées sont à la fois numériques (pour la persistance et
les recalculs) et formatées en chaîne « 1 234,56 € » (pour l'injection Word).

Le prix final calculé reste TOUJOURS surchargeable manuellement : si un appelant
fournit `prix_force_ht`, c'est cette valeur qui prime (un indicateur `override`
signale l'écart avec le calcul automatique).
"""
from __future__ import annotations

from dataclasses import dataclass, field, asdict
from typing import Dict, List, Optional


# ============================================================
# Barème de technicité (modifiable)
# ============================================================
# Coefficient appliqué au coût de base. Le barème est volontairement simple et
# centralisé ici pour pouvoir être ajusté sans toucher au reste du code.
# L'utilisateur peut aussi fournir un coefficient libre (override) par ligne.
BAREME_TECHNICITE: Dict[str, float] = {
    "standard": 1.00,        # entretien courant
    "technique": 1.15,       # prestation nécessitant matériel/compétence spécifique
    "haute": 1.30,           # haute technicité (cristallisation, travaux en hauteur…)
    "exceptionnelle": 1.50,  # intervention exceptionnelle / contraintes fortes
}

# Conversion d'une fréquence en nombre de passages par mois.
# 52 semaines / 12 mois = 4.333… ; on garde la valeur exacte pour la précision.
SEMAINES_PAR_MOIS = 52 / 12  # ≈ 4.3333

FREQUENCE_PAR_MOIS: Dict[str, float] = {
    "quotidien": 30.0,
    "5/semaine": 5 * SEMAINES_PAR_MOIS,
    "4/semaine": 4 * SEMAINES_PAR_MOIS,
    "3/semaine": 3 * SEMAINES_PAR_MOIS,
    "2/semaine": 2 * SEMAINES_PAR_MOIS,
    "1/semaine": 1 * SEMAINES_PAR_MOIS,
    "bimensuel": 2.0,
    "1/mois": 1.0,
    "trimestriel": 1 / 3,
    "ponctuel": 1.0,
}

TVA_DEFAUT = 0.20  # 20 %

# Taux horaire par défaut (peut être surchargé par les paramètres en base)
TAUX_HORAIRE_DEFAUT = 24.0


def appliquer_parametres(params: dict):
    """Surcharge les constantes de calcul à partir des paramètres stockés en base.

    `params` est un dict {cle: valeur(str)} issu de la table Parametre. Les clés
    reconnues : tva_defaut, taux_horaire_defaut, coef_standard, coef_technique,
    coef_haute, coef_exceptionnelle. Toute clé absente garde sa valeur par défaut.
    """
    global TVA_DEFAUT, TAUX_HORAIRE_DEFAUT
    def f(cle, defaut):
        try:
            return float(str(params.get(cle, defaut)).replace(",", "."))
        except (TypeError, ValueError):
            return defaut
    TVA_DEFAUT = f("tva_defaut", TVA_DEFAUT)
    TAUX_HORAIRE_DEFAUT = f("taux_horaire_defaut", TAUX_HORAIRE_DEFAUT)
    BAREME_TECHNICITE["standard"] = f("coef_standard", BAREME_TECHNICITE["standard"])
    BAREME_TECHNICITE["technique"] = f("coef_technique", BAREME_TECHNICITE["technique"])
    BAREME_TECHNICITE["haute"] = f("coef_haute", BAREME_TECHNICITE["haute"])
    BAREME_TECHNICITE["exceptionnelle"] = f("coef_exceptionnelle", BAREME_TECHNICITE["exceptionnelle"])


def coef_technicite(niveau: Optional[str], override: Optional[float] = None) -> float:
    """Retourne le coefficient de technicité.

    Priorité à l'override numérique s'il est fourni (cas « valeur libre »),
    sinon on lit le barème. Niveau inconnu → 1.0 (neutre).
    """
    if override is not None:
        return float(override)
    if not niveau:
        return 1.0
    return BAREME_TECHNICITE.get(niveau.strip().lower(), 1.0)


def passages_par_mois(frequence: Optional[str]) -> float:
    """Convertit une fréquence (clé connue) en passages/mois. Inconnu → 1.0."""
    if not frequence:
        return 1.0
    return FREQUENCE_PAR_MOIS.get(frequence.strip().lower(), 1.0)


def fmt_euros(montant: float) -> str:
    """Formate un montant en euros à la française : 1 234,56 €."""
    s = f"{montant:,.2f}"                 # 1,234.56
    s = s.replace(",", " ").replace(".", ",")  # 1 234,56
    return f"{s} €"


@dataclass
class LignePrestation:
    """Une ligne de prestation à chiffrer.

    Champs d'entrée (saisis par l'utilisateur) :
      libelle, duree_h, nb_agents, taux_horaire, niveau_technicite,
      coef_override, frais, taux_tva, frequence

    Champs calculés (remplis par calculer()) :
      cout_base_ht, ht_unitaire, ht, tva, ttc, passages_mois
    """
    libelle: str = ""
    duree_h: float = 0.0
    nb_agents: float = 1.0
    taux_horaire: float = 0.0
    niveau_technicite: Optional[str] = "standard"
    coef_override: Optional[float] = None
    frais: float = 0.0
    taux_tva: float = TVA_DEFAUT
    frequence: Optional[str] = "ponctuel"

    # calculés
    coef: float = field(default=0.0, init=False)
    passages_mois: float = field(default=0.0, init=False)
    ht_par_passage: float = field(default=0.0, init=False)
    ht: float = field(default=0.0, init=False)
    tva: float = field(default=0.0, init=False)
    ttc: float = field(default=0.0, init=False)

    def calculer(self, recurrent: bool) -> "LignePrestation":
        """Calcule HT/TVA/TTC de la ligne.

        - HT par passage = (durée × agents × taux × coef technicité) + frais
        - récurrent : HT mensuel = HT par passage × passages/mois
        - ponctuel  : HT = HT par passage (un seul « passage »)
        """
        self.coef = coef_technicite(self.niveau_technicite, self.coef_override)
        main_oeuvre = self.duree_h * self.nb_agents * self.taux_horaire * self.coef
        self.ht_par_passage = main_oeuvre + self.frais

        if recurrent:
            self.passages_mois = passages_par_mois(self.frequence)
            self.ht = self.ht_par_passage * self.passages_mois
        else:
            self.passages_mois = 1.0
            self.ht = self.ht_par_passage

        self.tva = self.ht * self.taux_tva
        self.ttc = self.ht + self.tva
        return self

    def as_dict(self) -> Dict:
        d = asdict(self)
        # versions formatées pour l'injection Word
        d["ht_fmt"] = fmt_euros(self.ht)
        d["tva_fmt"] = fmt_euros(self.tva)
        d["ttc_fmt"] = fmt_euros(self.ttc)
        return d


@dataclass
class ResultatDevis:
    """Résultat agrégé d'un devis chiffré."""
    recurrent: bool
    lignes: List[Dict]
    total_ht: float
    total_tva: float
    total_ttc: float
    total_ht_fmt: str
    total_tva_fmt: str
    total_ttc_fmt: str
    override: bool = False
    total_ht_calcule: Optional[float] = None  # valeur auto avant override (si override)

    def as_dict(self) -> Dict:
        return asdict(self)


def chiffrer(
    lignes: List[Dict],
    recurrent: bool,
    prix_force_ht: Optional[float] = None,
    taux_tva_global: float = TVA_DEFAUT,
) -> ResultatDevis:
    """Chiffre un ensemble de lignes et agrège les totaux.

    Args:
        lignes : liste de dicts compatibles avec LignePrestation
        recurrent : True pour un contrat récurrent (calcul mensuel),
                    False pour un devis ponctuel (forfait)
        prix_force_ht : si fourni, force le total HT (override manuel). La TVA et
                        le TTC sont recalculés sur cette base avec taux_tva_global.
        taux_tva_global : taux de TVA appliqué en cas d'override

    Returns:
        ResultatDevis avec lignes calculées et totaux (numériques + formatés).
    """
    calculees: List[LignePrestation] = []
    for l in lignes:
        ligne = LignePrestation(
            libelle=l.get("libelle", ""),
            duree_h=float(l.get("duree_h", 0) or 0),
            nb_agents=float(l.get("nb_agents", 1) or 1),
            taux_horaire=float(l.get("taux_horaire", 0) or 0),
            niveau_technicite=l.get("niveau_technicite", "standard"),
            coef_override=(float(l["coef_override"]) if l.get("coef_override") not in (None, "") else None),
            frais=float(l.get("frais", 0) or 0),
            taux_tva=float(l.get("taux_tva", taux_tva_global) or taux_tva_global),
            frequence=l.get("frequence", "ponctuel" if not recurrent else "1/semaine"),
        )
        ligne.calculer(recurrent)
        calculees.append(ligne)

    total_ht = sum(l.ht for l in calculees)
    total_tva = sum(l.tva for l in calculees)
    total_ttc = total_ht + total_tva

    override = False
    total_ht_calcule = None
    if prix_force_ht is not None and prix_force_ht != "":
        override = True
        total_ht_calcule = total_ht
        total_ht = float(prix_force_ht)
        total_tva = total_ht * taux_tva_global
        total_ttc = total_ht + total_tva

    return ResultatDevis(
        recurrent=recurrent,
        lignes=[l.as_dict() for l in calculees],
        total_ht=round(total_ht, 2),
        total_tva=round(total_tva, 2),
        total_ttc=round(total_ttc, 2),
        total_ht_fmt=fmt_euros(total_ht),
        total_tva_fmt=fmt_euros(total_tva),
        total_ttc_fmt=fmt_euros(total_ttc),
        override=override,
        total_ht_calcule=round(total_ht_calcule, 2) if total_ht_calcule is not None else None,
    )


def construire_tableau_zones(
    zones_detail: List[Dict],
    base_ligne: Dict,
    taux_tva_global: float = TVA_DEFAUT,
) -> List[Dict]:
    """Construit le tableau financier DYNAMIQUE à partir des zones réellement
    cochées et de leurs options éventuelles.

    Une ligne par zone (libellé = titre de zone, fréquence renseignée, prix).
    Une ligne supplémentaire par option ACTIVÉE (désignation, fréquence, prix).
    Les zones/options non sélectionnées n'apparaissent jamais.
    """
    lignes = []
    for z in (zones_detail or []):
        titre = (z.get("titre") or z.get("code") or "Zone").strip()
        freq = (z.get("frequence") or "").strip()
        # Prix de la zone : prix forcé si fourni, sinon calcul standard
        prix_force = z.get("prix_ht")
        l = dict(base_ligne); l["frequence"] = _freq_to_key(freq)
        res = chiffrer([l], recurrent=True, taux_tva_global=taux_tva_global,
                       prix_force_ht=(float(prix_force) if prix_force not in (None, "") else None))
        lib = titre + (f" — {freq}" if freq else "")
        lignes.append({"libelle": lib, "ht": res.total_ht_fmt,
                       "tva": res.total_tva_fmt, "ttc": res.total_ttc_fmt})
        # Option complémentaire de la zone, si activée
        if z.get("option_active"):
            o_lib = (z.get("option_libelle") or "Option complémentaire").strip()
            o_freq = (z.get("option_frequence") or "").strip()
            o_prix = z.get("option_prix_ht")
            ol = dict(base_ligne); ol["frequence"] = _freq_to_key(o_freq)
            ores = chiffrer([ol], recurrent=True, taux_tva_global=taux_tva_global,
                            prix_force_ht=(float(o_prix) if o_prix not in (None, "") else None))
            o_full = f"{titre} — Option : {o_lib}" + (f" ({o_freq})" if o_freq else "")
            lignes.append({"libelle": o_full, "ht": ores.total_ht_fmt,
                           "tva": ores.total_tva_fmt, "ttc": ores.total_ttc_fmt})
    return lignes


def _freq_to_key(freq_label: str) -> str:
    """Convertit un libellé de fréquence en clé connue de FREQUENCE_PAR_MOIS."""
    if not freq_label:
        return "1/mois"
    f = freq_label.strip().lower()
    table = {
        "quotidien": "quotidien", "tous les jours": "quotidien",
        "6 fois par semaine": "5/semaine", "5 fois par semaine": "5/semaine",
        "4 fois par semaine": "4/semaine", "3 fois par semaine": "3/semaine",
        "2 fois par semaine": "2/semaine", "1 fois par semaine": "1/semaine",
        "2 fois par mois": "bimensuel", "1 fois par mois": "1/mois",
        "trimestriel": "trimestriel", "ponctuel": "ponctuel",
    }
    return table.get(f, "1/semaine")


def construire_options_recurrentes(
    base_ligne: Dict,
    frequences: List[str],
    taux_tva_global: float = TVA_DEFAUT,
) -> List[Dict]:
    """Construit une liste d'OPTIONS comparatives pour le tableau récurrent.

    Pour chaque fréquence fournie, recalcule le prix mensuel et renvoie un dict
    prêt à injecter dans le template ({libelle, ht, tva, ttc}).

    Exemple : base_ligne décrivant l'entretien d'un hall, frequences =
    ["1/semaine","2/semaine","3/semaine"] → 3 lignes d'options chiffrées.
    """
    libelles_freq = {
        "1/semaine": "1 fois par semaine",
        "2/semaine": "2 fois par semaine",
        "3/semaine": "3 fois par semaine",
        "4/semaine": "4 fois par semaine",
        "5/semaine": "5 fois par semaine",
        "quotidien": "tous les jours",
        "bimensuel": "2 fois par mois",
        "1/mois": "1 fois par mois",
    }
    options = []
    for i, freq in enumerate(frequences, start=1):
        l = dict(base_ligne)
        l["frequence"] = freq
        res = chiffrer([l], recurrent=True, taux_tva_global=taux_tva_global)
        lib = libelles_freq.get(freq, freq)
        options.append({
            "libelle": f"Option {i} : {lib}",
            "ht": res.total_ht_fmt,
            "tva": res.total_tva_fmt,
            "ttc": res.total_ttc_fmt,
        })
    return options

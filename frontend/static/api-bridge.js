/* =====================================================================
   DevisFlow — Pont API (Phase 2A)
   Connecte l'interface riche au backend FastAPI.

   Stratégie : l'interface d'origine fonctionnait sur localStorage. Ce pont
   charge les données réelles depuis l'API au démarrage, peuple `state`, et
   redirige les opérations de création/modification/suppression vers l'API.
   Si le backend est injoignable (ex. fichier ouvert seul, sans serveur), on
   bascule en mode local pour que la maquette reste consultable.
   ===================================================================== */
(function () {
  const API = (location.origin && location.origin.startsWith("http"))
    ? location.origin + "/api" : "/api";

  const ME_API = {
    online: false,

    async _get(path) {
      const r = await fetch(API + path);
      if (!r.ok) throw new Error(r.status + " " + path);
      return r.json();
    },
    async _send(path, method, body, isForm) {
      const opts = { method };
      if (isForm) { opts.body = body; }
      else { opts.headers = { "Content-Type": "application/json" }; opts.body = JSON.stringify(body); }
      const r = await fetch(API + path, opts);
      if (!r.ok) throw new Error((await r.text().catch(() => "")) || r.status);
      return r.headers.get("content-type")?.includes("json") ? r.json() : r;
    },

    // --- Clients ---
    listClients: (q) => ME_API._get("/clients" + (q ? "?q=" + encodeURIComponent(q) : "")),
    createClient: (c) => ME_API._send("/clients", "POST", c),
    updateClient: (id, c) => ME_API._send("/clients/" + id, "PUT", c),
    archiveClient: (id, a) => ME_API._send("/clients/" + id + "/archive?archive=" + a, "POST", {}),
    deleteClient: (id) => ME_API._send("/clients/" + id, "DELETE", {}),

    // --- Membres (équipe) ---
    listMembres: () => ME_API._get("/membres"),
    createMembre: (m) => ME_API._send("/membres", "POST", m),
    updateMembre: (id, m) => ME_API._send("/membres/" + id, "PUT", m),
    deleteMembre: (id) => ME_API._send("/membres/" + id, "DELETE", {}),

    // --- Prestations / équipements / paramètres ---
    listPrestations: (fam) => ME_API._get("/prestations" + (fam ? "?famille=" + fam : "")),
    createPrestation: (p) => ME_API._send("/prestations", "POST", p),
    updatePrestation: (code, p) => ME_API._send("/prestations/" + code, "PUT", p),
    deletePrestation: (code) => ME_API._send("/prestations/" + code, "DELETE", {}),
    listEquipements: (params = "") => ME_API._get("/equipements" + (params ? "?" + params : "")),
    deleteEquipement: (code) => ME_API._send("/equipements/" + code, "DELETE", {}),
    listParametres: () => ME_API._get("/parametres"),
    updateParametre: (cle, v) => ME_API._send("/parametres/" + encodeURIComponent(cle), "PUT", { valeur: v }),

    // --- Devis ---
    listDevis: () => ME_API._get("/devis"),
    genererDevis: (payload) => ME_API._send("/devis/generer", "POST", payload),
  };
  window.ME_API = ME_API;

  // Mapping API client -> forme attendue par l'interface (state.clients[*])
  function mapClient(c) {
    return {
      id: c.id, nom: c.nom, contact: c.contact || "", email: c.email || "",
      tel: c.telephone || "", adresse: c.adresse || "", code_postal: c.code_postal || "",
      ville: c.ville || "", site_nom: c.site_nom || "", site_adresse: c.site_adresse || "",
      civilite: c.civilite || "", archive: c.archive, devis: 0,
    };
  }
  function mapMembre(m) {
    return { id: m.id, nom: m.nom, email: m.email || "", role: m.role || "",
             avatar: (m.nom || "?").split(" ").map(s => s[0]).join("").slice(0, 2).toUpperCase(),
             actif: m.actif };
  }
  window.ME_MAP = { client: mapClient, membre: mapMembre };

  /* Bootstrap : charge les données réelles puis rafraîchit les listings.
     Appelé après le boot d'origine (qui a déjà défini state + fonctions render). */
  async function bootFromBackend() {
    if (typeof state === "undefined") return;
    try {
      // ping
      await ME_API._get("/health");
      ME_API.online = true;
    } catch (e) {
      console.warn("[DevisFlow] Backend hors ligne — mode maquette local.", e.message);
      const badge = document.getElementById("me-conn");
      if (badge) { badge.textContent = "Mode maquette (hors ligne)"; badge.className = "me-conn off"; }
      return;
    }

    const badge = document.getElementById("me-conn");
    if (badge) { badge.textContent = "Connecté au backend"; badge.className = "me-conn on"; }

    // Clients
    try {
      const clients = await ME_API.listClients();
      state.clients = clients.map(mapClient);
      if (typeof renderClientsListing === "function") renderClientsListing();
      if (typeof setupClientAutocomplete === "function") setupClientAutocomplete();
    } catch (e) { console.warn("clients:", e.message); }

    // Membres / équipe
    try {
      const membres = await ME_API.listMembres();
      state.users = membres.map(mapMembre);
      if (typeof renderUsersListing === "function") renderUsersListing();
    } catch (e) { console.warn("membres:", e.message); }

    // Devis émis
    try {
      const devis = await ME_API.listDevis();
      state.devisList = devis.map(d => ({
        id: d.numero, numero: d.numero, type: d.template_code,
        recurrent: (d.template_code === "copro_petite"),
        client: d.client_nom, site: d.site_adresse, date: d.date_emission,
        ht: d.montant_ht, ttc: d.montant_ttc, statut: d.statut || "brouillon",
        has_docx: d.has_docx, has_pdf: d.has_pdf,
      }));
      if (typeof renderDevisListing === "function") renderDevisListing();
    } catch (e) { console.warn("devis:", e.message); }

    // Prestations réelles (zones exactes du modèle Word) pour la création de devis
    try { window.ME_ZONES = await ME_API.listPrestations("contrat"); }
    catch (e) { console.warn("prestations:", e.message); }

    try { installCatalogueMaterielPage(); } catch (e) { console.warn("catalogue materiel:", e.message); }
    // Vues d'administration branchées sur la base
    try { await renderAdminViews(); } catch (e) { console.warn("admin views:", e.message); }
    // Sélecteur d'équipements (photos) dans la création de devis
    try { await installEquipPicker(); } catch (e) { console.warn("equip picker:", e.message); }
    // Sélecteur de zones (récurrent) — pilote l'affichage des prestations
    try { await installZonePicker(); } catch (e) { console.warn("zone picker:", e.message); }

    // Re-render les vues admin à chaque navigation vers elles
    if (typeof window.switchView === "function" && !window.__switchWrapped) {
      const _sv = window.switchView;
      window.switchView = function (v) {
        _sv(v);
        if (["biblio", "materiel", "settings"].includes(v) && ME_API.online) renderAdminViews();
      };
      window.__switchWrapped = true;
    }
  }
  window.bootFromBackend = bootFromBackend;

  /* ------------------------------------------------------------------
     Branchement des écritures (création/modification) sur l'API.
     On redéfinit certaines fonctions globales du prototype APRÈS son
     chargement : si le backend est en ligne, on persiste via l'API ;
     sinon on laisse le comportement local d'origine.
     ------------------------------------------------------------------ */
  function wireWrites() {
    if (typeof window.openModal !== "function") return;

    // --- Création de client ---
    window.openClientModal = function (clientId) {
      const existing = clientId && typeof state !== "undefined"
        ? state.clients.find(c => String(c.id) === String(clientId)) : null;
      const v = existing || {};
      const body = `
        <div class="field"><label class="field__label">Raison sociale / Nom<span class="req">*</span></label>
          <input class="field__input" id="cli-nom" value="${(v.nom||"").replace(/"/g,'&quot;')}" placeholder="ex : Syndic Foncia"></div>
        <div class="field"><label class="field__label">Contact principal</label>
          <input class="field__input" id="cli-contact" value="${(v.contact||"").replace(/"/g,'&quot;')}"></div>
        <div class="field__row field__row--2">
          <div class="field"><label class="field__label">Email</label>
            <input class="field__input" id="cli-email" value="${(v.email||"")}"></div>
          <div class="field"><label class="field__label">Téléphone</label>
            <input class="field__input" id="cli-tel" value="${(v.tel||"")}"></div>
        </div>
        <div class="field__row field__row--2">
          <div class="field"><label class="field__label">Adresse</label>
            <input class="field__input" id="cli-adr" value="${(v.adresse||"").replace(/"/g,'&quot;')}"></div>
          <div class="field"><label class="field__label">Ville</label>
            <input class="field__input" id="cli-ville" value="${(v.ville||"").replace(/"/g,'&quot;')}"></div>
        </div>
        <div class="field"><label class="field__label">Adresse du site d'intervention</label>
          <input class="field__input" id="cli-site" value="${(v.site_adresse||"").replace(/"/g,'&quot;')}"></div>
      `;
      openModal(existing ? "Modifier le client" : "Nouveau client",
                "Carnet d'adresses — enregistré sur le serveur", body, async () => {
        const nom = document.getElementById("cli-nom").value.trim();
        if (!nom) return showToast("Nom requis", "danger");
        const payload = {
          nom,
          contact: document.getElementById("cli-contact").value.trim(),
          email: document.getElementById("cli-email").value.trim(),
          telephone: document.getElementById("cli-tel").value.trim(),
          adresse: document.getElementById("cli-adr").value.trim(),
          ville: document.getElementById("cli-ville").value.trim(),
          site_adresse: document.getElementById("cli-site").value.trim(),
        };
        try {
          if (ME_API.online) {
            if (existing) await ME_API.updateClient(existing.id, payload);
            else await ME_API.createClient(payload);
            const fresh = await ME_API.listClients();
            state.clients = fresh.map(mapClient);
          } else {
            if (existing) Object.assign(existing, payload, { tel: payload.telephone });
            else state.clients.unshift({ id: "c" + Date.now(), ...payload, tel: payload.telephone, devis: 0 });
          }
          renderClientsListing();
          if (typeof setupClientAutocomplete === "function") setupClientAutocomplete();
          closeModal();
          showToast(existing ? "Client modifié" : `Client « ${nom} » ajouté`, "success");
        } catch (e) { showToast("Erreur : " + e.message, "danger"); }
      });
    };

    // --- Création / édition de membre d'équipe ---
    window.openUserModal = function (userId) {
      const existing = userId && typeof state !== "undefined"
        ? state.users.find(u => String(u.id) === String(userId)) : null;
      const v = existing || {};
      const roles = ["Président", "Directeur", "Commercial", "Assistante exploitation", "Technicien"];
      const body = `
        <div class="field"><label class="field__label">Nom<span class="req">*</span></label>
          <input class="field__input" id="usr-nom" value="${(v.nom||"").replace(/"/g,'&quot;')}"></div>
        <div class="field"><label class="field__label">Email</label>
          <input class="field__input" id="usr-email" value="${(v.email||"")}"></div>
        <div class="field"><label class="field__label">Rôle</label>
          <select class="field__input" id="usr-role">${roles.map(r=>`<option ${r===v.role?"selected":""}>${r}</option>`).join("")}</select></div>
        <div class="field"><label class="field__label"><input type="checkbox" id="usr-actif" ${v.actif!==false?"checked":""}> Membre actif</label></div>
      `;
      openModal(existing ? "Modifier le membre" : "Nouveau membre",
                "Équipe Marie Eugénie", body, async () => {
        const nom = document.getElementById("usr-nom").value.trim();
        if (!nom) return showToast("Nom requis", "danger");
        const payload = {
          nom, email: document.getElementById("usr-email").value.trim(),
          role: document.getElementById("usr-role").value,
          actif: document.getElementById("usr-actif").checked,
        };
        try {
          if (ME_API.online) {
            if (existing) await ME_API.updateMembre(existing.id, payload);
            else await ME_API.createMembre(payload);
            const fresh = await ME_API.listMembres();
            state.users = fresh.map(mapMembre);
          } else {
            if (existing) Object.assign(existing, payload);
            else state.users.push({ id: "u" + Date.now(), ...payload,
              avatar: nom.split(" ").map(s=>s[0]).join("").slice(0,2).toUpperCase() });
          }
          renderUsersListing();
          if (typeof initInfoBlock === "function") initInfoBlock();
          closeModal();
          showToast(existing ? "Membre modifié" : "Membre ajouté", "success");
        } catch (e) { showToast("Erreur : " + e.message, "danger"); }
      });
    };
  }
  window.wireWrites = wireWrites;

  /* ------------------------------------------------------------------
     APERÇU PDF RÉEL
     L'interface affichait une reconstruction HTML du devis (≠ modèle Word).
     On la remplace par l'aperçu du VRAI document : le modèle Word rempli puis
     converti en PDF par le backend. C'est la seule source de vérité.
     ------------------------------------------------------------------ */


function getCurrentTemplateIssue() {
  const code = window.state && state.devis && state.devis.modele_code;
  const m = code && state.biblio && state.biblio.modeles ? state.biblio.modeles[code] : null;
  if (m && m.template_missing) return m.missing_reason || "Mod?le source manquant";
  return "";
}

  // Convertit l'état du devis (forme prototype) vers le format attendu par l'API.
  function mapDevisToPayload() {
    const d = (typeof state !== "undefined") ? state.devis : {};
    const dd = d || {};
    // Le modele metier prime sur le type d'onglet : un modele bureau reste
    // ponctuel meme si un ancien brouillon indique encore "contrat".
    const zonesSel = window.ME_DEVIS_ZONES || [];
    const modeleCode = dd.modele_code || dd.modele || (dd.type === "contrat" ? "copro_petite" : "ponctuel_generique");
    const modeleMetier = (window.state && state.biblio && state.biblio.modeles)
      ? state.biblio.modeles[modeleCode]
      : null;
    let recurrent;
    if (modeleMetier && ["contrat", "bureaux"].includes(modeleMetier.famille)) recurrent = true;
    else if (modeleMetier && modeleMetier.famille === "ponctuel") recurrent = false;
    else if (dd.type === "contrat") recurrent = true;
    else if (dd.type === "ponctuel") recurrent = false;
    else recurrent = zonesSel.length > 0;

    const PONCTUEL_TEMPLATE_BY_MODELE = {
      encombrants_caves: "ponctuel_encombrants_caves",
      encombrants_divers: "ponctuel_encombrants",
      tags: "ponctuel_tag",
      tapis: "ponctuel_tapis",
      vitrerie: "ponctuel_vitrerie",
      relamping: "ponctuel_relamping",
      remise_etat: "ponctuel_remise_etat",
      appartement: "ponctuel_appartement",
    };
    const tplCode = ["bureaux_petit", "bureaux_important"].includes(modeleCode) ? modeleCode : (recurrent ? "copro_petite" : (PONCTUEL_TEMPLATE_BY_MODELE[modeleCode] || modeleCode));
    const getPonctuelOps = () => {
      if (recurrent) return [];
      const modelOps = (modeleMetier && Array.isArray(modeleMetier.operations)) ? modeleMetier.operations : [];
      const sourceOps = Array.isArray(dd.operations) ? dd.operations : [];
      const flags = Array.isArray(dd.ops_enabled) ? dd.ops_enabled : [];
      const sameModel = !dd.operations_model_code || dd.operations_model_code === modeleCode;
      const ops = sameModel && sourceOps.length ? sourceOps : modelOps;
      return ops.filter((op, i) => (!sameModel || flags[i] !== false) && v(op).trim());
    };

    // Destinataire : on remplit DEST_LIGNE1..4 sans jamais produire "undefined".
    const cli = dd.client || {};
    const site = dd.site || {};
    const v = (x) => (x === undefined || x === null) ? "" : String(x);
    const parseMoney = (x) => {
      if (x === undefined || x === null || x === "") return 0;
      if (typeof x === "number") return isFinite(x) ? x : 0;
      const n = parseFloat(String(x).replace(/\s/g, "").replace("€", "").replace("%", "").replace(",", "."));
      return isFinite(n) ? n : 0;
    };
    const calcLine = (line) => {
      const qty = parseMoney(line.quantite);
      const pu = parseMoney(line.prix_unitaire_ht);
      let ht = parseMoney(line.total_ht ?? line.ht);
      if (qty && pu) ht = qty * pu;
      if (line.type_ligne === "remise" && ht > 0) ht = -ht;
      const taux = parseMoney(line.taux_tva || 20);
      const tva = Math.round(ht * (taux / 100) * 100) / 100;
      return { ht: Math.round(ht * 100) / 100, tva, ttc: Math.round((ht + tva) * 100) / 100 };
    };
    const pfSource = dd.proposition_financiere || {};
    const rawFinancialLines = Array.isArray(pfSource.lignes) && pfSource.lignes.length
      ? pfSource.lignes
      : [{ designation: (dd.info && dd.info.type_prestation) || "Selon descriptif", total_ht: dd.totaux?.ht || 0, taux_tva: 20, type_ligne: "prestation", inclure_total: true }];
    let financialTotalHt = 0, financialTotalTva = 0, financialTotalTtc = 0;
    const financialLines = rawFinancialLines.map((line, idx) => {
      const c = calcLine(line || {});
      const include = line && line.inclure_total !== false && !["information", "info", "sous_total", "sous-total", "total_general"].includes(line.type_ligne);
      if (include) {
        financialTotalHt += c.ht;
        financialTotalTva += c.tva;
        financialTotalTtc += c.ttc;
      }
      return {
        designation: v(line.designation || line.libelle || `Ligne ${idx + 1}`),
        description: v(line.description || ""),
        quantite: v(line.quantite || ""),
        unite: v(line.unite || ""),
        prix_unitaire_ht: v(line.prix_unitaire_ht || ""),
        total_ht: c.ht,
        taux_tva: parseMoney(line.taux_tva || 20),
        montant_tva: c.tva,
        total_ttc: c.ttc,
        type_ligne: v(line.type_ligne || line.type || "prestation"),
        inclure_total: line.inclure_total !== false,
      };
    });
    const nomComplet = [v(cli.civilite), v(cli.contact)].filter(Boolean).join(" ").trim();

    const variables = {
      NUMERO_DEVIS: v(dd.numero),
      DATE_EMISSION: v(dd.date_emission),
      TYPE_PRESTATION: v((dd.info && dd.info.type_prestation) || (modeleMetier && modeleMetier.objet_principal) || ""),
      DEST_LIGNE1: nomComplet || v(cli.raison_sociale),
      DEST_LIGNE2: nomComplet ? v(cli.raison_sociale) : v(cli.adresse),
      DEST_LIGNE3: v(cli.adresse) || v([cli.code_postal, cli.ville].filter(Boolean).join(" ")),
      DEST_LIGNE4: v([cli.code_postal, cli.ville].filter(Boolean).join(" ")) || v(cli.email),
      SITE_ADRESSE: v(site.adresse),
      SITE_CP_VILLE: v([site.code_postal, site.ville].filter(Boolean).join(" ")),
      DATE_PRISE_EFFET: v(dd.date_effet),
      DUREE_CONTRAT: v(dd.duree_contrat),
      NOM_OPPORTUNITE: v(dd.zone_concernee) || v(cli.raison_sociale),
      FORFAIT_LIBELLE: v((modeleMetier && modeleMetier.libelle) || (dd.info && dd.info.type_prestation) || "Forfait prestation"),
      MENTION_VALIDITE: (!recurrent && modeleMetier && ["encombrants_caves", "encombrants_divers"].includes(v(modeleCode)))
        ? v(modeleMetier.mention_rouge || "")
        : "",
      MENTION_SPECIFIQUE: (!recurrent && ["encombrants_caves", "encombrants_divers"].includes(v(modeleCode)))
        ? "La société MARIE EUGENIE ne pourra être aucunement tenue responsable de l'évacuation d'objets situés dans les parties communes ; le client ne peut entamer le moindre recours contre la société MARIE EUGENIE dans le cadre de l'enlèvement et de la destruction des objets présents dans les parties communes."
        : "",
      _modele_code: v(modeleCode),
      FINANCIAL_NOTE: v(pfSource.note || ""),
      TVA_NOTE: v(pfSource.note_tva || ""),
    };

    if (modeleCode === "bureaux_petit") {
      const typologie = dd.typologie_bureaux || {};
      ["SURFACE_LOCAUX", "NB_BLOCS_SANITAIRES", "NB_COLLABORATEURS", "REVETEMENT_BUREAU", "REVETEMENT_SANITAIRE"]
        .forEach(k => { variables[k] = v(typologie[k] || variables[k] || ""); });
    }

    // Fréquences par zone (récurrent) + collecte des zones cochées.
    // SOURCE DE VÉRITÉ : la section Prestations native du prototype (dd.zones),
    // que l'utilisateur coche réellement. On mappe les codes prototype
    // (hall_entree...) vers les codes du modèle Word (hall...) et on injecte la
    // fréquence PROPRE de chaque zone dans sa variable FREQ_* dédiée.
    const zonesSelectionnees = [];
    var zonesDetail = [];
    if (recurrent) {
      const ZONE_CODE_MAP = {
        hall_entree: "hall", cabine_ascenseur: "ascenseur", cages_escaliers: "escaliers",
        caves_descentes_garage: "caves", garage_parking: "garage",
        abords_exterieurs: "abords", local_poubelles: "conteneur", ordures_menageres: "om",
        hall: "hall", ascenseur: "ascenseur", escaliers: "escaliers", caves: "caves",
        garage: "garage", abords: "abords", conteneur: "conteneur", om: "om",
        accueil_bureaux: "accueil_bureaux", circulation: "circulation", sanitaires: "sanitaires",
        cuisine: "cuisine", vitrerie: "vitrerie", consommables: "consommables",
      };
      const freqVar = { hall: "FREQ_HALL", ascenseur: "FREQ_ASCENSEUR", escaliers: "FREQ_ESCALIERS",
        caves: "FREQ_CAVES", garage: "FREQ_GARAGE", abords: "FREQ_ABORDS",
        conteneur: "FREQ_CONTENEUR", om: "FREQ_OM",
        accueil_bureaux: "FREQ_ACCUEIL_BUREAUX", circulation: "FREQ_CIRCULATION",
        sanitaires: "FREQ_SANITAIRES", cuisine: "FREQ_CUISINE", vitrerie: "FREQ_VITRERIE" };

      const protoZones = Array.isArray(dd.zones) ? dd.zones : [];
      protoZones.forEach(z => {
        if (!z || !z.selected) return;
        const back = ZONE_CODE_MAP[z.code];
        if (!back) return;
        const hasFreq = !!freqVar[back] && !z.no_frequency;
        const freq = hasFreq ? (v(z.frequence) || "1 fois par semaine") : "";
        if (hasFreq) variables[freqVar[back]] = freq;
        zonesSelectionnees.push(back);
        variables[`OPS_${back.toUpperCase()}`] = (z.operations || []).filter((op, idx) => z.ops_enabled[idx] !== false && v(op).trim());
        // Détail pour le tableau financier dynamique (+ option par zone si activée)
        zonesDetail.push({
          code: back, titre: z.titre || back, frequence: freq,
          prix_ht: (z.prix_ht != null && z.prix_ht !== "") ? z.prix_ht : null,
          option_active: !!z.option_active,
          option_libelle: v(z.option_libelle || ""),
          option_frequence: v(z.option_frequence || ""),
          option_prix_ht: (z.option_prix_ht != null && z.option_prix_ht !== "") ? z.option_prix_ht : null,
        });
      });
      // repli éventuel sur un sélecteur dédié
      if (!zonesSelectionnees.length && window.ME_DEVIS_ZONES && window.ME_DEVIS_ZONES.length) {
        window.ME_DEVIS_ZONES.forEach(z => {
          const back = ZONE_CODE_MAP[z.code] || z.code;
          if (freqVar[back]) {
            const freq = v(z.frequence) || "1 fois par semaine";
            variables[freqVar[back]] = freq; zonesSelectionnees.push(back);
            zonesDetail.push({ code: back, titre: z.titre || back, frequence: freq });
          }
        });
      }
      Object.values(freqVar).forEach(fv => { if (fv && !(fv in variables)) variables[fv] = ""; });
    }

    // Lignes de prestation pour le calcul
    let lignes = [];
    if (recurrent) {
      lignes = [{ duree_h: 2, nb_agents: 1, taux_horaire: 24, niveau_technicite: "standard", frais: 0 }];
    } else {
      const ops = getPonctuelOps();
      lignes = ops.length ? ops.map(o => ({ libelle: v(o), duree_h: 1, nb_agents: 1,
        taux_horaire: 26, niveau_technicite: "standard", frais: 0 }))
        : [{ libelle: "Prestation", duree_h: 1, nb_agents: 1, taux_horaire: 26 }];
    }

    return {
      numero: v(dd.numero) || "APERCU",
      template_code: tplCode,
      variables,
      lignes,
      frequences_options: [],
      // Prix : si l'utilisateur a saisi un Montant HT global (section Tarif),
      // on le transmet comme override (prix_force_ht). Pour le récurrent, chaque
      // zone peut aussi porter son propre prix (zones_detail[].prix_ht).
      prix_force_ht: financialTotalHt > 0 ? Math.round(financialTotalHt * 100) / 100 : (function () {
        const t = dd.totaux || {};
        const ht = (t.force_ht != null && t.force_ht !== "") ? t.force_ht
                 : (t.ht != null && t.ht !== "" ? t.ht : null);
        const n = parseFloat(ht);
        return (isFinite(n) && n > 0) ? n : null;
      })(),
      mode_financier: v(pfSource.mode || (financialLines.length > 1 ? "detaillee" : "simple")),
      lignes_financieres: financialLines,
      note_financiere: v(pfSource.note || ""),
      note_tva: v(pfSource.note_tva || ""),
      equipements: (function () {
        // Photos par zone (système natif simple) : chaque zone cochée avec une
        // photo est envoyée comme item photo, associé au titre de la zone pour
        // être injecté à proximité de cette zone dans le document.
        const items = (window.ME_DEVIS_EQUIPEMENTS || []).slice();
        if (recurrent && Array.isArray(dd.zones)) {
          dd.zones.forEach(z => {
            if (z && z.selected && Array.isArray(z.photos)) {
              z.photos.forEach(ph => {
                if (ph && (ph.photo_data || ph.dataUrl)) {
                  items.push({
                    libelle: z.titre || "Zone",
                    prestation: z.titre || "",
                    photo_data: ph.photo_data || ph.dataUrl,
                    image_size: ph.image_size || z.image_size || "grande",
                    image_align: ph.image_align || z.image_align || "droite",
                    image_width_pct: ph.image_width_pct || z.image_width_pct || 40,
                  });
                }
              });
            }
            if (z && z.selected && z.photo) {
              items.push({
                libelle: z.titre || "Zone",
                prestation: z.titre || "",
                photo_data: z.photo,
                image_size: z.image_size || "grande",
                image_align: z.image_align || "droite",
                image_width_pct: z.image_width_pct || 40,
              });
            }
          });
        }
        if (Array.isArray(dd.photos)) {
          const selectedZone = recurrent && Array.isArray(dd.zones) ? dd.zones.find(z => z && z.selected) : null;
          const fallbackPrestation = dd.zone_concernee || (selectedZone && selectedZone.titre) || getPonctuelOps()[0] || "Prestation principale";
          dd.photos.forEach(p => {
            if (p && (p.photo_data || p.dataUrl)) {
              items.push({
                libelle: p.name || fallbackPrestation,
                prestation: p.prestation || fallbackPrestation,
                photo_data: p.photo_data || p.dataUrl,
                image_size: p.image_size || "grande",
                image_align: p.image_align || "droite",
                image_width_pct: p.image_width_pct || 40,
              });
            }
          });
        }
        return items;
      })(),
      zones_selectionnees: zonesSelectionnees,
      zones_detail: zonesDetail,
    };
  }

  window.mapDevisToPayload = mapDevisToPayload;
  let _apercuTimer = null, _apercuUrl = null, _apercuSeq = 0;

  // PDF.js (embarqué localement) : rend l'aperçu sur canvas, sans plugin.
  let _pdfjsLoading = null;
  function _ensurePdfJs() {
    if (window.pdfjsLib) return Promise.resolve(window.pdfjsLib);
    if (_pdfjsLoading) return _pdfjsLoading;
    _pdfjsLoading = new Promise((resolve, reject) => {
      const s = document.createElement("script");
      s.src = "/static/vendor/pdf.min.js";
      s.onload = () => {
        try { window.pdfjsLib.GlobalWorkerOptions.workerSrc = "/static/vendor/pdf.worker.min.js"; } catch (e) {}
        resolve(window.pdfjsLib);
      };
      s.onerror = () => reject(new Error("pdf.js indisponible"));
      document.head.appendChild(s);
    });
    return _pdfjsLoading;
  }
  async function _renderPdfCanvas(buffer, container, seq) {
    try {
      const lib = await _ensurePdfJs();
      if (seq !== _apercuSeq) return true;
      const pdf = await lib.getDocument({ data: buffer }).promise;
      if (seq !== _apercuSeq) return true;
      const wrap = document.createElement("div");
      wrap.className = "preview-pdf-scroll";
      container.innerHTML = ""; container.appendChild(wrap);
      const availableW = Math.max(320, (container.clientWidth || 760) - 44);
      const targetW = Math.min(availableW, 820);
      for (let n = 1; n <= pdf.numPages; n++) {
        if (seq !== _apercuSeq) return true;
        const page = await pdf.getPage(n);
        const base = page.getViewport({ scale: 1 });
        const vp = page.getViewport({ scale: targetW / base.width });
        const c = document.createElement("canvas");
        c.width = vp.width; c.height = vp.height;
        c.className = "preview-pdf-page";
        wrap.appendChild(c);
        await page.render({ canvasContext: c.getContext("2d"), viewport: vp }).promise;
      }
      return true;
    } catch (e) { return false; }
  }
  window._renderPdfCanvas = _renderPdfCanvas;

  async function rafraichirApercuPDF() {
    const content = document.getElementById("preview-content");
    if (!content) return;
    if (!ME_API.online) {
      content.innerHTML = '<div style="padding:40px;text-align:center;color:#8c877c;font:14px Arial">'
        + "L'aperçu du document réel nécessite le backend en marche.<br>"
        + "Lancez l'application (docker-compose up) pour voir le PDF généré depuis le modèle Word.</div>";
      return;
    }
    const templateIssue = (typeof getCurrentTemplateIssue === "function") ? getCurrentTemplateIssue() : "";
    if (templateIssue) {
      content.innerHTML = '<div style="padding:30px;color:#b42318;font:13px Arial">Aper?u indisponible pour ce mod?le.<br><span style="color:#8c877c;font-size:11px">'
        + templateIssue + "</span></div>";
      return;
    }
    const payload = mapDevisToPayload();
    if (!payload || !payload.template_code) return;

    const seq = ++_apercuSeq;   // garde anti-concurrence : seule la dernière requête s'applique
    content.innerHTML = '<div id="apercu-loading" style="padding:40px;text-align:center;color:#8c877c;font:14px Arial">Génération de l\'aperçu depuis le modèle Word…</div>';
    try {
      const r = await fetch(API + "/devis/apercu", {
        method: "POST", headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });
      if (seq !== _apercuSeq) return;            // une requête plus récente a pris le relais
      if (!r.ok) {
        let msg = r.status;
        try { const j = await r.json(); msg = j.detail || msg; } catch (e) {}
        throw new Error(msg);
      }
      const raw = await r.arrayBuffer();
      if (seq !== _apercuSeq) return;
      const ct = r.headers.get("content-type") || "";
      const head = new Uint8Array(raw.slice(0, 5));
      const isPdf = head[0] === 0x25 && head[1] === 0x50 && head[2] === 0x44 && head[3] === 0x46; // %PDF
      if (!isPdf && !ct.includes("pdf")) throw new Error("réponse inattendue du serveur");
      if (_apercuUrl) URL.revokeObjectURL(_apercuUrl);
      const blob = new Blob([raw], { type: "application/pdf" });
      _apercuUrl = URL.createObjectURL(blob);
      // Rendu universel via PDF.js (canvas) — fonctionne sans plugin PDF.
      const ok = await _renderPdfCanvas(raw.slice(0), content, seq);
      if (!ok && seq === _apercuSeq) {
        content.innerHTML = '<object data="' + _apercuUrl + '#view=FitH" type="application/pdf" '
          + 'style="width:100%;height:100%;min-height:80vh;border:0;background:#525659">'
          + '<div style="padding:24px;font:13px Arial"><a href="' + _apercuUrl + '" target="_blank">Ouvrir l\'aperçu PDF</a></div></object>';
      }
    } catch (e) {
      if (seq !== _apercuSeq) return;
      content.innerHTML = '<div style="padding:30px;color:#b42318;font:13px Arial">Aperçu momentanément indisponible. Cliquez sur « Rafraîchir l\'aperçu ».<br><span style="color:#8c877c;font-size:11px">'
        + (e.message || e) + "</span></div>";
    }
  }
  window.rafraichirApercuPDF = rafraichirApercuPDF;

  // Remplace l'aperçu HTML reconstruit par l'aperçu PDF réel (anti-rebond).
  function installRealPreview() {
    window.updatePreview = function () {
      // NB : ne PAS rappeler recalcTotaux() ici — recalcTotaux() appelle déjà
      // updatePreview() à la fin : cela créait une récursion infinie qui faisait
      // que l'aperçu ne se mettait plus à jour de façon fiable.
      if (document.querySelector('.preview-tab[data-tab="json"].is-active')) return;
      clearTimeout(_apercuTimer);
      _apercuTimer = setTimeout(rafraichirApercuPDF, 600);
    };
    // bouton "Rafraîchir" + retrait du toggle "rendu finalisé" devenu inutile
    const bar = document.querySelector(".preview-actions");
    if (bar && !document.getElementById("btn-refresh-apercu")) {
      const b = document.createElement("button");
      b.id = "btn-refresh-apercu"; b.textContent = "Rafraîchir l'aperçu";
      b.style.cssText = "border:1px solid #161513;background:#161513;color:#fff;border-radius:7px;padding:7px 12px;font:600 12px Arial;cursor:pointer";
      b.onclick = rafraichirApercuPDF;
      bar.insertBefore(b, bar.firstChild);
    }
    const toggle = document.querySelector(".preview-toggle");
    if (toggle) toggle.style.display = "none";  // le rendu vient du Word, plus de "rendu finalisé"
  }
  window.installRealPreview = installRealPreview;

  /* ------------------------------------------------------------------
     PHASE 2B — Sélecteur d'équipements dans la création de devis.
     L'utilisateur coche un ou plusieurs équipements ; leurs photos seront
     injectées dans le Word/PDF, alignées avec leur libellé.
     ------------------------------------------------------------------ */
  window.ME_DEVIS_EQUIPEMENTS = [];
  const ME_PHOTO_ACCEPT = "image/jpeg,image/png,image/webp,image/heic,image/heif,.jpg,.jpeg,.png,.webp,.heic,.heif";
  function meIsAcceptedPhoto(file) {
    const name = (file && file.name || "").toLowerCase();
    const type = (file && file.type || "").toLowerCase();
    return /^image\/(jpeg|png|webp|heic|heif)$/.test(type) || /\.(jpe?g|png|webp|heic|heif)$/.test(name);
  }

  function meReadPhotoForBackend(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onerror = () => reject(new Error("Lecture du fichier impossible"));
      reader.onload = () => {
        const rawDataUrl = reader.result;
        const img = new Image();
        img.onload = () => resolve({ photo_data: rawDataUrl, thumb: rawDataUrl, preview_ok: true });
        img.onerror = () => resolve({ photo_data: rawDataUrl, thumb: "", preview_ok: false });
        img.src = rawDataUrl;
      };
      reader.readAsDataURL(file);
    });
  }

  async function installEquipPicker() {
    if (!ME_API.online) return;
    const view = document.getElementById("view-creation");
    if (!view || document.getElementById("me-equip-picker")) return;
    let equips = [];
    try { equips = (await ME_API.listEquipements()).filter(e => e.actif); } catch (e) { return; }
    window.ME_EQUIP_ALL = equips;

    const sect = document.createElement("section");
    sect.id = "me-equip-picker";
    sect.style.cssText = "background:#fff;border:1px solid #e4e0d7;border-radius:12px;padding:18px;margin:18px 0";
    sect.innerHTML = `
      <div style="font:700 14px Arial;margin-bottom:4px">Matériel à présenter dans le devis</div>
      <div style="font:12px Arial;color:#8c877c;margin-bottom:12px">Cochez le matériel du catalogue à ajouter dans l'encart @@MATERIEL_ENCART@@. Les photos ajoutées manuellement restent associées aux prestations.</div>
      <div id="me-equip-list" style="display:grid;grid-template-columns:1fr;gap:8px"></div>
      <div style="margin-top:12px;display:flex;gap:10px;flex-wrap:wrap;align-items:center">
        <button id="me-add-photo" class="me-btn-primary" type="button">+ Ajouter des photos</button>
        <input id="me-photo-file" type="file" accept="${ME_PHOTO_ACCEPT}" multiple style="display:none">
        <span style="font:12px Arial;color:#8c877c">JPG, PNG, WEBP, HEIC/HEIF acceptés. S&eacute;lection multiple possible.</span>
      </div>`;
    const anchor = view.querySelector(".workspace__body, .workspace, form") || view;
    anchor.appendChild(sect);

    function prestationOptions() {
      // Propose les prestations du devis en cours (zones cochées ou opérations)
      const d = (typeof state !== "undefined") ? state.devis : null;
      let opts = [];
      if (d) {
        if (Array.isArray(d.zones)) opts = d.zones.filter(z => z.selected).map(z => z.titre || z.code);
        {
          const code = d.modele_code || d.modele || "";
          const model = state && state.biblio && state.biblio.modeles ? state.biblio.modeles[code] : null;
          const sourceOps = Array.isArray(d.operations) ? d.operations : [];
          const sameModel = !d.operations_model_code || d.operations_model_code === code;
          const ops = sameModel && sourceOps.length ? sourceOps : ((model && model.operations) || []);
          const flags = Array.isArray(d.ops_enabled) ? d.ops_enabled : [];
          opts = opts.concat(ops.filter((op, i) => (!sameModel || flags[i] !== false) && op));
        }
      }
      if (!opts.length) opts = ["Prestation principale"];
      return opts;
    }

    function rowHtml(item, idx) {
      const presta = prestationOptions();
      const thumb = item.thumb || item.photo_data
        ? (item.thumb ? `<img src="${item.thumb}" alt="" style="width:54px;height:40px;object-fit:cover;border-radius:6px;border:1px solid #d8d3c8">`
          : `<div style="width:54px;height:40px;display:grid;place-items:center;border-radius:6px;border:1px solid #d8d3c8;background:#f6f1e8;color:#7a6a58;font:700 10px Arial;text-align:center">Image<br>ajoutée</div>`)
        : "";
      return `<div class="me-equip-row" data-idx="${idx}" style="display:flex;gap:10px;align-items:center;border:1px solid #e4e0d7;border-radius:8px;padding:8px 10px">
        <input type="checkbox" class="me-eq-on" ${item.on?"checked":""}>
        ${thumb}
        <div style="flex:1">
          <div style="font:600 13px Arial">${esc(item.libelle)}</div>
          <div style="font:11px Arial;color:#8c877c">${item.code?esc(catLabelOf(item.code)):"photo ajoutée"}${item.photo_data?" · photo perso":""}</div>
        </div>
        <label style="font:11px Arial;color:#8c877c">Prestation :
          <select class="me-eq-presta" style="font:12px Arial;padding:5px;border:1px solid #d8d3c8;border-radius:6px;max-width:220px">
            ${presta.map(p=>`<option ${item.prestation===p?"selected":""}>${esc(p)}</option>`).join("")}
          </select></label>
        ${item.photo_data ? `<div style="display:flex;gap:4px">
          <button type="button" class="me-eq-up me-btn-mini" title="Monter">↑</button>
          <button type="button" class="me-eq-down me-btn-mini" title="Descendre">↓</button>
          <button type="button" class="me-eq-del me-btn-mini danger" title="Supprimer">Suppr.</button>
        </div>` : ""}
      </div>`;
    }

    function scoreSuggestion(e) {
      const d = (typeof state !== "undefined") ? state.devis : {};
      const code = String((d && (d.modele_code || d.modele || d.type)) || "").toLowerCase();
      const tags = Array.isArray(e.tags) ? e.tags.map(t => String(t).toLowerCase()) : [];
      const hay = [e.code, e.label, e.categorie, ...(e.tags || [])].join(" ").toLowerCase();
      let score = 0;
      if (code.includes("vitrerie") && /vitrerie|vitres|hauteur/.test(hay)) score += 10;
      if (code.includes("encombr") && /encombrants|evacuation|transport|fourgon|camion/.test(hay)) score += 10;
      if (code.includes("tapis") && /tapis|moquette|textile/.test(hay)) score += 10;
      if (code.includes("remise") && /remise_etat|remise en/.test(hay)) score += 10;
      if (tags.some(t => code.includes(t) || t.includes(code))) score += 5;
      return score;
    }

    // état interne : une ligne par équipement + lignes ajoutées (photos perso)
    const model = equips
      .map(e => ({ code: e.code, libelle: e.label, categorie: e.categorie, tags: e.tags || [], thumb: e.photo_url ? API.replace("/api","") + e.photo_url : null, suggestion: scoreSuggestion(e), on: false, prestation: "", photo_data: null }))
      .sort((a, b) => (b.suggestion - a.suggestion) || String(a.libelle).localeCompare(String(b.libelle), "fr"));
    window.__equipModel = model;

    function renderRows() {
      const host = sect.querySelector("#me-equip-list");
      host.innerHTML = model.map((it, i) => rowHtml(it, i)).join("");
      host.querySelectorAll(".me-equip-row").forEach(r => {
        const i = +r.dataset.idx;
        r.querySelector(".me-eq-on").addEventListener("change", e => { model[i].on = e.target.checked; syncSelection(); });
        r.querySelector(".me-eq-presta").addEventListener("change", e => { model[i].prestation = e.target.value; syncSelection(); });
        const up = r.querySelector(".me-eq-up");
        const down = r.querySelector(".me-eq-down");
        const del = r.querySelector(".me-eq-del");
        if (up) up.addEventListener("click", () => { if (i > 0) { const x = model.splice(i, 1)[0]; model.splice(i - 1, 0, x); renderRows(); syncSelection(); } });
        if (down) down.addEventListener("click", () => { if (i < model.length - 1) { const x = model.splice(i, 1)[0]; model.splice(i + 1, 0, x); renderRows(); syncSelection(); } });
        if (del) del.addEventListener("click", () => { model.splice(i, 1); renderRows(); syncSelection(); });
      });
    }
    function syncSelection() {
      window.ME_DEVIS_EQUIPEMENTS = model.filter(m => m.on || m.photo_data).map(m => ({
        code: m.code || undefined,
        libelle: m.libelle,
        prestation: m.prestation,
        photo_data: m.photo_data || undefined,
        encart: m.code && !m.photo_data ? true : undefined,
        materiel: m.code && !m.photo_data ? true : undefined,
      }));
      if (typeof rafraichirApercuPDF === "function") rafraichirApercuPDF();
    }
    window.__syncEquip = syncSelection;
    renderRows();

    // Upload d'une photo à la volée
    sect.querySelector("#me-add-photo").addEventListener("click", () => sect.querySelector("#me-photo-file").click());
    sect.querySelector("#me-photo-file").addEventListener("change", e => {
      const rejected = [];
      const files = Array.from(e.target.files || []).filter(f => {
        const ok = meIsAcceptedPhoto(f);
        if (!ok) rejected.push(f.name);
        return ok;
      });
      if (rejected.length && typeof showToast === "function") showToast("Format non support&eacute; : " + rejected.join(", "), "danger");
      if (!files.length) return;
      let pending = files.length;
      files.forEach(f => {
        meReadPhotoForBackend(f).then(photo => {
          model.push({ code: null, libelle: f.name.replace(/\.[^.]+$/, ""), on: true, prestation: "", photo_data: photo.photo_data, thumb: photo.thumb });
          pending -= 1;
          if (!pending) { renderRows(); syncSelection(); e.target.value = ""; }
          if (!photo.preview_ok && typeof showToast === "function") showToast("Image ajoutée, conversion au moment de la génération : " + f.name, "success");
        }).catch(() => {
          pending -= 1;
          if (typeof showToast === "function") showToast("Lecture impossible : " + f.name, "danger");
          if (!pending) { renderRows(); syncSelection(); e.target.value = ""; }
        });
      });
    });
  }
  window.installEquipPicker = installEquipPicker;

  /* ------------------------------------------------------------------
     Sélecteur de ZONES (récurrent) : l'utilisateur coche les prestations
     du contrat à inclure. Les zones non cochées disparaissent de l'aperçu
     et du PDF (SHOW_* côté backend).
     ------------------------------------------------------------------ */
  window.ME_DEVIS_ZONES = [];

  async function installZonePicker() {
    // Désactivé : on garde uniquement la section "Prestations" native du prototype
    // (cases à cocher + fréquence par zone à droite), plus claire et sans doublon.
    return;
  }
  async function _installZonePicker_OFF() {
    if (!ME_API.online) return;
    const view = document.getElementById("view-creation");
    if (!view || document.getElementById("me-zone-picker")) return;
    let zones = [];
    try { zones = await ME_API.listPrestations("contrat"); } catch (e) { return; }
    // codes backend connus pour le modèle Copro
    const KNOWN = ["hall", "ascenseur", "escaliers", "caves", "garage", "abords", "conteneur", "om"];
    zones = zones.filter(z => KNOWN.includes(z.code));

    const sect = document.createElement("section");
    sect.id = "me-zone-picker";
    sect.style.cssText = "background:#fff;border:1px solid #e4e0d7;border-radius:12px;padding:18px;margin:18px 0";
    sect.innerHTML = `
      <div style="font:700 14px Arial;margin-bottom:4px">Zones du contrat (récurrent)</div>
      <div style="font:12px Arial;color:#8c877c;margin-bottom:12px">Cochez les zones à inclure. Les zones décochées n'apparaîtront ni dans l'aperçu ni dans le PDF.</div>
      <div style="display:grid;grid-template-columns:1fr;gap:8px">
        ${zones.map(z => `
          <div class="me-zone-row" data-code="${esc(z.code)}" style="display:flex;gap:10px;align-items:center;border:1px solid #e4e0d7;border-radius:8px;padding:8px 10px">
            <input type="checkbox" class="me-zone-on">
            <div style="flex:1;font:600 13px Arial">${esc(z.titre)} <span style="font-weight:400;color:#8c877c">(${(z.operations||[]).length} opérations)</span></div>
            <label style="font:11px Arial;color:#8c877c">Fréquence :
              <input class="me-zone-freq" value="1 fois / semaine" style="font:12px Arial;padding:5px;border:1px solid #d8d3c8;border-radius:6px;width:140px"></label>
          </div>`).join("")}
      </div>`;
    const anchor = view.querySelector(".workspace__body, .workspace, form") || view;
    anchor.insertBefore(sect, anchor.firstChild);

    function sync() {
      window.ME_DEVIS_ZONES = Array.from(sect.querySelectorAll(".me-zone-row"))
        .filter(r => r.querySelector(".me-zone-on").checked)
        .map(r => ({ code: r.dataset.code, frequence: r.querySelector(".me-zone-freq").value }));
      if (typeof rafraichirApercuPDF === "function") rafraichirApercuPDF();
    }
    sect.querySelectorAll(".me-zone-on").forEach(c => c.addEventListener("change", sync));
    sect.querySelectorAll(".me-zone-freq").forEach(i => i.addEventListener("input", () => {
      clearTimeout(window.__zoneFreqT); window.__zoneFreqT = setTimeout(sync, 600);
    }));
    window.__syncZones = sync;
  }
  window.installZonePicker = installZonePicker;
  function catLabelOf(code) {
    const e = (window.ME_EQUIP_ALL || []).find(x => x.code === code);
    return e ? catLabel(e.categorie) : "";
  }

  /* ==================================================================
     PHASE 2A — Vues d'administration branchées sur la VRAIE base
     (Bibliothèque métier, Équipements, Paramètres de calcul)
     Ces fonctions vident les conteneurs du prototype et les re-rendent
     depuis l'API ; les créations/modifications passent par l'API.
     ================================================================== */

  const CATS = [
    ["Vitrerie","Vitrerie"],
    ["Encombrants","Encombrants"],
    ["Moquette / tapis","Moquette / tapis"],
    ["Remise en état","Remise en état"],
    ["Désinfection","Désinfection"],
    ["Matériel","Matériel"],
    ["machine","Machine"],
    ["materiel","Matériel technique"],
    ["vehicule","Véhicule"],
    ["specifique","Équipement spécifique"],
  ];

  // ---- Prestations types (Bibliothèque métier) ----
  async function renderPrestationsAPI() {
    // Conteneur principal visible ("MODÈLES DE PRESTATION")
    const cont = document.getElementById("biblio-modeles") || document.getElementById("biblio-zones");
    if (!cont || !ME_API.online) return;
    let items = [];
    try { items = await ME_API.listPrestations(); } catch (e) { return; }
    window.ME_PRESTATIONS = items;
    const hintEl = document.querySelector("#view-biblio .section__hint")
      || document.getElementById("biblio-modeles-hint");
    if (hintEl) hintEl.textContent = items.length + " prestation(s) type(s) en base — administrables sans code";
    // En-tête de section : compteur
    const counter = document.querySelector("#view-biblio .section__count, #view-biblio .card__count");
    if (counter) counter.textContent = items.length + " prestations";
    cont.innerHTML = `
      <div class="listing__row listing__row--header" style="grid-template-columns:2fr 1fr 90px 120px">
        <div>Prestation</div><div>Famille</div><div>Opérations</div><div style="text-align:right">Actions</div>
      </div>` + items.map(p => `
      <div class="listing__row" style="grid-template-columns:2fr 1fr 90px 120px">
        <div><div class="listing__cell--strong">${esc(p.titre)}</div>
          <div class="listing__cell--meta">${esc(p.code)}${p.actif?"":" · inactif"}</div></div>
        <div>${p.famille==="contrat"?"Contrat":"Ponctuel"}</div>
        <div>${(p.operations||[]).length} ops</div>
        <div style="text-align:right">
          <button class="me-btn-mini" onclick="editPrestationAPI('${esc(p.code)}')">Modifier</button>
          <button class="me-btn-mini danger" onclick="deletePrestationAPI('${esc(p.code)}')">Suppr.</button>
        </div>
      </div>`).join("") +
      `<div style="padding:12px"><button class="me-btn-primary" onclick="openPrestationAPI()">+ Ajouter une prestation</button></div>`;
  }

  window.openPrestationAPI = function (code) {
    const ex = code ? (window.ME_PRESTATIONS||[]).find(p => p.code === code) : null;
    const v = ex || {};
    const body = `
      <div class="field"><label class="field__label">Code<span class="req">*</span></label>
        <input class="field__input" id="pr-code" value="${esc(v.code||"")}" ${ex?"disabled":""} placeholder="ex : hall"></div>
      <div class="field"><label class="field__label">Titre<span class="req">*</span></label>
        <input class="field__input" id="pr-titre" value="${esc(v.titre||"")}" placeholder="ex : Hall d'entrée"></div>
      <div class="field"><label class="field__label">Famille</label>
        <select class="field__input" id="pr-fam">
          <option value="contrat" ${v.famille==="contrat"?"selected":""}>Contrat (récurrent)</option>
          <option value="ponctuel" ${v.famille==="ponctuel"?"selected":""}>Ponctuel</option></select></div>
      <div class="field"><label class="field__label">Opérations (une par ligne)</label>
        <textarea class="field__input" id="pr-ops" rows="6">${esc((v.operations||[]).join("\n"))}</textarea></div>`;
    openModal(ex ? "Modifier la prestation" : "Nouvelle prestation", "Bibliothèque métier — base de données", body, async () => {
      const code2 = document.getElementById("pr-code").value.trim();
      const titre = document.getElementById("pr-titre").value.trim();
      if (!code2 || !titre) return showToast("Code et titre requis", "danger");
      const payload = { code: code2, titre, famille: document.getElementById("pr-fam").value,
        operations: document.getElementById("pr-ops").value.split("\n").map(s=>s.trim()).filter(Boolean), ordre: 0 };
      try {
        if (ex) await ME_API.updatePrestation(code2, payload); else await ME_API.createPrestation(payload);
        await renderPrestationsAPI(); closeModal();
        showToast(ex ? "Prestation modifiée" : "Prestation ajoutée", "success");
      } catch (e) { showToast("Erreur : " + e.message, "danger"); }
    });
  };
  window.editPrestationAPI = (c) => openPrestationAPI(c);
  window.deletePrestationAPI = async (c) => {
    if (!confirm("Supprimer la prestation « " + c + " » ?")) return;
    try { await ME_API.deletePrestation(c); await renderPrestationsAPI(); showToast("Prestation supprimée", "success"); }
    catch (e) { showToast("Erreur : " + e.message, "danger"); }
  };

  // ---- Équipements / matériel / véhicules ----
  async function renderEquipementsAPI() {
    const cont = document.getElementById("biblio-materiel-grid");
    if (!cont || !ME_API.online) return;
    let items = [];
    try { items = await ME_API.listEquipements(); } catch (e) { return; }
    window.ME_EQUIPEMENTS = items;
    const hint = document.getElementById("biblio-materiel-hint");
    if (hint) hint.textContent = items.length + " équipement(s) en base — administrables";
    cont.innerHTML = items.map(e => `
      <div class="materiel-card" style="${e.actif?"":"opacity:.5"}">
        <div class="materiel-card__thumb">${e.photo_url
          ? `<img src="${API.replace("/api","")}${e.photo_url}" style="width:100%;height:100%;object-fit:cover;border-radius:8px">`
          : `<div style="display:grid;place-items:center;height:100%;color:#b9b3a6">📷</div>`}</div>
        <div class="materiel-card__body">
          <div class="materiel-card__name">${esc(e.label)}</div>
          <div class="materiel-card__cat">${esc(catLabel(e.categorie))}</div>
        </div>
        <div style="display:flex;gap:6px;margin-top:6px">
          <button class="me-btn-mini" onclick="editEquipementAPI('${esc(e.code)}')">Modifier</button>
          <button class="me-btn-mini danger" onclick="deleteEquipementAPI('${esc(e.code)}')">Suppr.</button>
        </div>
      </div>`).join("") +
      `<div class="materiel-card" style="border:1px dashed #cfc8b8;display:grid;place-items:center;cursor:pointer;min-height:120px"
            onclick="openEquipementAPI()"><div style="text-align:center;color:#8c877c">ï¼‹<br>Ajouter</div></div>`;
  }

  window.openEquipementAPI = function (code) {
    const ex = code ? (window.ME_EQUIPEMENTS||[]).find(e => e.code === code) : null;
    const v = ex || {};
    const body = `
      <div class="field"><label class="field__label">Code<span class="req">*</span></label>
        <input class="field__input" id="eq-code" value="${esc(v.code||"")}" ${ex?"disabled":""} placeholder="ex : autolaveuse"></div>
      <div class="field"><label class="field__label">Libellé<span class="req">*</span></label>
        <input class="field__input" id="eq-label" value="${esc(v.label||"")}"></div>
      <div class="field"><label class="field__label">Catégorie</label>
        <select class="field__input" id="eq-cat">${CATS.map(([c,l])=>`<option value="${c}" ${v.categorie===c?"selected":""}>${l}</option>`).join("")}</select></div>
      <div class="field"><label class="field__label">Description</label>
        <input class="field__input" id="eq-desc" value="${esc(v.description||"")}"></div>
      <div class="field"><label class="field__label">Photo (optionnel)</label>
        <input class="field__input" id="eq-photo" type="file" accept="image/jpeg,image/png,image/webp,image/heic,image/heif,.jpg,.jpeg,.png,.webp,.heic,.heif"></div>
      ${ex?`<div class="field"><label class="field__label"><input type="checkbox" id="eq-actif" ${v.actif?"checked":""}> Actif</label></div>`:""}`;
    openModal(ex ? "Modifier l'équipement" : "Nouvel équipement", "Bibliothèque — matériel / véhicules", body, async () => {
      const code2 = document.getElementById("eq-code").value.trim();
      const label = document.getElementById("eq-label").value.trim();
      if (!code2 || !label) return showToast("Code et libellé requis", "danger");
      const fd = new FormData();
      if (!ex) fd.append("code", code2);
      fd.append("label", label);
      fd.append("categorie", document.getElementById("eq-cat").value);
      fd.append("description", document.getElementById("eq-desc").value);
      if (ex) fd.append("actif", document.getElementById("eq-actif").checked);
      const pf = document.getElementById("eq-photo").files[0];
      if (pf) fd.append("photo", pf);
      try {
        const path = ex ? "/equipements/" + code2 : "/equipements";
        await ME_API._send(path, ex ? "PUT" : "POST", fd, true);
        await renderEquipementsAPI(); closeModal();
        showToast(ex ? "Équipement modifié" : "Équipement ajouté", "success");
      } catch (e) { showToast("Erreur : " + e.message, "danger"); }
    });
  };
  window.editEquipementAPI = (c) => openEquipementAPI(c);
  window.deleteEquipementAPI = async (c) => {
    if (!confirm("Supprimer l'équipement « " + c + " » ?")) return;
    try { await ME_API.deleteEquipement(c); await renderEquipementsAPI(); showToast("Équipement supprimé", "success"); }
    catch (e) { showToast("Erreur : " + e.message, "danger"); }
  };

  function installCatalogueMaterielPage() {
    if (document.getElementById("view-materiel")) return;
    const nav = document.querySelector('.nav__item[data-view="biblio"]');
    if (nav && nav.parentElement && !document.querySelector('.nav__item[data-view="materiel"]')) {
      const li = document.createElement("li");
      li.className = "nav__item";
      li.setAttribute("data-view", "materiel");
      li.innerHTML = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><rect x="3" y="3" width="18" height="18" rx="2"/><circle cx="9" cy="9" r="2"/><path d="M21 15l-5-5L5 21"/></svg> Catalogue matériel';
      nav.parentElement.insertBefore(li, nav.nextSibling);
    }
    const main = document.querySelector("main") || document.querySelector("#main") || document.body;
    const view = document.createElement("div");
    view.className = "view";
    view.id = "view-materiel";
    view.innerHTML = `
      <section class="workspace">
        <header class="workspace__topbar">
          <div class="workspace__heading">
            <h1 class="workspace__title">Catalogue matériel</h1>
            <div class="workspace__sub">Matériels, photos, descriptions et tags utilisables dans les devis ponctuels</div>
          </div>
          <div class="workspace__actions">
            <button class="btn btn--primary btn--sm" onclick="openEquipementAPI()">+ Ajouter un matériel</button>
          </div>
        </header>
        <div class="workspace__body">
          <section class="section">
            <div class="section__header">
              <div class="section__title">Matériels disponibles</div>
              <div class="section__hint" id="catalogue-materiel-hint">— matériels</div>
            </div>
            <div class="section__body">
              <div class="materiel-filters" style="display:flex;gap:8px;flex-wrap:wrap;align-items:center;margin-bottom:14px">
                <input class="field__input" id="materiel-search" placeholder="Rechercher un matériel, une catégorie ou un tag" style="max-width:320px">
                <select class="field__input" id="materiel-category" style="max-width:220px">
                  <option value="">Toutes les catégories</option>
                  ${CATS.map(([c,l])=>`<option value="${esc(c)}">${esc(l)}</option>`).join("")}
                </select>
                <label style="font:12px Arial;color:#6f6a60"><input type="checkbox" id="materiel-show-inactif-page"> Inclure les inactifs</label>
              </div>
              <div class="materiel-grid" id="catalogue-materiel-grid"></div>
            </div>
          </section>
        </div>
      </section>`;
    main.appendChild(view);
    view.querySelector("#materiel-search").addEventListener("input", () => renderEquipementsAPI());
    view.querySelector("#materiel-category").addEventListener("change", () => renderEquipementsAPI());
    view.querySelector("#materiel-show-inactif-page").addEventListener("change", () => renderEquipementsAPI());
  }

  function tagsOf(e) {
    return Array.isArray(e.tags) ? e.tags : [];
  }

  function renderEquipementCards(items) {
    if (!items.length) {
      return `<div style="padding:18px;color:#8c877c;font:13px Arial">Aucun matériel trouvé.</div>`;
    }
    return items.map(e => `
      <div class="materiel-card" style="${e.actif?"":"opacity:.5"}">
        <div class="materiel-card__thumb">${e.photo_url
          ? `<img src="${API.replace("/api","")}${e.photo_url}" style="width:100%;height:100%;object-fit:cover;border-radius:8px">`
          : `<div style="display:grid;place-items:center;height:100%;color:#b9b3a6">Photo</div>`}</div>
        <div class="materiel-card__body">
          <div class="materiel-card__name">${esc(e.label)}</div>
          <div class="materiel-card__cat">${esc(catLabel(e.categorie))}</div>
          <div style="font:11px Arial;color:#8c877c;margin-top:4px">${esc(e.description || "")}</div>
          ${tagsOf(e).length ? `<div style="font:10px Arial;color:#7a6a58;margin-top:6px">${tagsOf(e).map(t=>`#${esc(t)}`).join(" ")}</div>` : ""}
        </div>
        <div style="display:flex;gap:6px;margin-top:8px;flex-wrap:wrap">
          <button class="me-btn-mini" onclick="editEquipementAPI('${esc(e.code)}')">Modifier</button>
          <button class="me-btn-mini danger" onclick="deleteEquipementAPI('${esc(e.code)}')">Suppr.</button>
        </div>
      </div>`).join("") +
      `<div class="materiel-card" style="border:1px dashed #cfc8b8;display:grid;place-items:center;cursor:pointer;min-height:120px"
            onclick="openEquipementAPI()"><div style="text-align:center;color:#8c877c">+<br>Ajouter</div></div>`;
  }

  renderEquipementsAPI = async function () {
    if (!ME_API.online) return;
    let items = [];
    try { items = await ME_API.listEquipements(); } catch (e) { return; }
    window.ME_EQUIPEMENTS = items;
    window.ME_EQUIP_ALL = items;
    const q = (document.getElementById("materiel-search")?.value || "").trim().toLowerCase();
    const cat = document.getElementById("materiel-category")?.value || "";
    const showInactivePage = !!document.getElementById("materiel-show-inactif-page")?.checked;
    let visible = items.slice();
    if (cat) visible = visible.filter(e => e.categorie === cat);
    if (!showInactivePage) visible = visible.filter(e => e.actif !== false);
    if (q) visible = visible.filter(e =>
      (e.code || "").toLowerCase().includes(q)
      || (e.label || "").toLowerCase().includes(q)
      || (e.categorie || "").toLowerCase().includes(q)
      || (e.description || "").toLowerCase().includes(q)
      || tagsOf(e).some(t => String(t).toLowerCase().includes(q))
    );
    const html = renderEquipementCards(visible);
    const biblio = document.getElementById("biblio-materiel-grid");
    if (biblio) biblio.innerHTML = renderEquipementCards(items.filter(e => e.actif !== false));
    const catalogue = document.getElementById("catalogue-materiel-grid");
    if (catalogue) catalogue.innerHTML = html;
    const hintB = document.getElementById("biblio-materiel-hint");
    if (hintB) hintB.textContent = items.length + " matériel(s) en base";
    const hintC = document.getElementById("catalogue-materiel-hint");
    if (hintC) hintC.textContent = visible.length + " matériel(s) affiché(s)";
  };

  window.openEquipementAPI = function (code) {
    const ex = code ? (window.ME_EQUIPEMENTS||[]).find(e => e.code === code) : null;
    const v = ex || {};
    const body = `
      <div class="field"><label class="field__label">Code matériel<span class="req">*</span></label>
        <input class="field__input" id="eq-code" value="${esc(v.code||"")}" ${ex?"disabled":""} placeholder="ex : perche_h2o"></div>
      <div class="field"><label class="field__label">Nom du matériel<span class="req">*</span></label>
        <input class="field__input" id="eq-label" value="${esc(v.label||"")}"></div>
      <div class="field"><label class="field__label">Catégorie</label>
        <select class="field__input" id="eq-cat">${CATS.map(([c,l])=>`<option value="${esc(c)}" ${v.categorie===c?"selected":""}>${esc(l)}</option>`).join("")}</select></div>
      <div class="field"><label class="field__label">Description courte</label>
        <textarea class="field__input" id="eq-desc" rows="3">${esc(v.description||"")}</textarea></div>
      <div class="field"><label class="field__label">Tags / prestations associées</label>
        <input class="field__input" id="eq-tags" value="${esc(tagsOf(v).join(", "))}" placeholder="ex : vitrerie, vitres, hauteur"></div>
      <div class="field"><label class="field__label">Photo</label>
        <input class="field__input" id="eq-photo" type="file" accept="image/*,.jpg,.jpeg,.png,.webp,.heic,.heif"></div>
      ${ex?`<div class="field"><label class="field__label"><input type="checkbox" id="eq-actif" ${v.actif?"checked":""}> Actif</label></div>`:""}`;
    openModal(ex ? "Modifier le matériel" : "Nouveau matériel", "Catalogue matériel", body, async () => {
      const code2 = document.getElementById("eq-code").value.trim();
      const label = document.getElementById("eq-label").value.trim();
      if (!code2 || !label) return showToast("Code et nom du matériel requis", "danger");
      const fd = new FormData();
      if (!ex) fd.append("code", code2);
      fd.append("label", label);
      fd.append("categorie", document.getElementById("eq-cat").value);
      fd.append("description", document.getElementById("eq-desc").value);
      fd.append("tags", document.getElementById("eq-tags").value);
      if (ex) fd.append("actif", document.getElementById("eq-actif").checked);
      const pf = document.getElementById("eq-photo").files[0];
      if (pf) fd.append("photo", pf);
      try {
        const path = ex ? "/equipements/" + code2 : "/equipements";
        await ME_API._send(path, ex ? "PUT" : "POST", fd, true);
        await renderEquipementsAPI(); closeModal();
        showToast(ex ? "Matériel modifié" : "Matériel ajouté", "success");
      } catch (e) { showToast("Erreur : " + e.message, "danger"); }
    });
  };
  window.editEquipementAPI = (c) => openEquipementAPI(c);

  // ---- Paramètres de calcul ----
  async function renderParametresAPI() {
    if (!ME_API.online) return;
    const view = document.getElementById("view-settings");
    if (!view) return;
    let params = [];
    try { params = await ME_API.listParametres(); } catch (e) { return; }
    window.ME_PARAMS = {}; params.forEach(p => window.ME_PARAMS[p.cle] = p.valeur);
    const calc = params.filter(p => p.groupe === "calcul");
    let card = document.getElementById("me-params-card");
    if (!card) {
      card = document.createElement("section");
      card.id = "me-params-card"; card.className = "settings-card";
      card.style.cssText = "background:#fff;border:1px solid #e4e0d7;border-radius:12px;padding:20px;margin:18px;display:block;width:auto";
      // Insérer en tout début de la zone de contenu de la vue Paramètres
      const anchor = view.querySelector(".workspace, .workspace__body, .view__body, section");
      if (anchor && anchor.parentNode) anchor.parentNode.insertBefore(card, anchor);
      else view.appendChild(card);
    }
    card.innerHTML = `
      <div style="font:700 14px Arial;margin-bottom:4px">Paramètres de calcul</div>
      <div style="font:12px Arial;color:#8c877c;margin-bottom:14px">Utilisés par le moteur de prix — enregistrés en base.</div>
      <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:12px">
        ${calc.map(p => `
          <div class="field"><label class="field__label">${esc(p.libelle||p.cle)}</label>
            <input class="field__input me-param" data-cle="${esc(p.cle)}" value="${esc(p.valeur)}"></div>`).join("")}
      </div>
      <div style="margin-top:14px"><button id="me-save-params" class="me-btn-primary">Enregistrer les paramètres</button>
        <span id="me-params-msg" style="margin-left:10px;font:12px Arial;color:#1f7a44"></span></div>`;
    document.getElementById("me-save-params").onclick = async () => {
      const inputs = card.querySelectorAll(".me-param");
      try {
        for (const inp of inputs) await ME_API.updateParametre(inp.dataset.cle, inp.value.trim());
        document.getElementById("me-params-msg").textContent = "Enregistré ✓";
        showToast("Paramètres enregistrés en base", "success");
      } catch (e) { showToast("Erreur : " + e.message, "danger"); }
    };
  }

  function esc(s) { return String(s == null ? "" : s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;"); }
  function catLabel(c) { const f = CATS.find(x => x[0] === c); return f ? f[1] : c; }

  async function renderAdminViews() {
    await renderPrestationsAPI();
    await renderEquipementsAPI();
    await renderParametresAPI();
  }
  window.renderAdminViews = renderAdminViews;

  // Empêche le rendu localStorage du prototype d'écraser nos vues API.
  function overrideProtoRenders() {
    if (typeof window.renderBiblioListings === "function") {
      const _orig = window.renderBiblioListings;
      window.renderBiblioListings = function () {
        if (ME_API.online) { renderPrestationsAPI(); renderEquipementsAPI(); }
        else { try { _orig(); } catch (e) {} }
      };
    }
  }
  window.overrideProtoRenders = overrideProtoRenders;

  // styles utilitaires pour les boutons injectés
  (function () {
    const st = document.createElement("style");
    st.textContent = `.me-btn-mini{border:1px solid #d8d3c8;background:#fff;border-radius:6px;padding:5px 9px;font:600 11px Arial;cursor:pointer;margin-left:4px}
      .me-btn-mini.danger{color:#b42318;border-color:#f0c9c4}
      .me-btn-primary{border:1px solid #161513;background:#161513;color:#fff;border-radius:8px;padding:9px 16px;font:600 13px Arial;cursor:pointer}`;
    document.head.appendChild(st);
  })();

  // ---- Version visible + bouton Diagnostic ----
  const ME_VERSION = "DevisFlow 11 — synchronisation interface/devis — 2026-06-23";

  function installVersionBadge() {
    if (document.getElementById("me-version")) return;
    const v = document.createElement("div");
    v.id = "me-version";
    v.style.cssText = "position:fixed;bottom:12px;left:14px;z-index:9999;font:600 11px/1 Arial;"
      + "padding:7px 12px;border-radius:999px;background:#1f2430;color:#fff;opacity:.92;cursor:default";
    v.textContent = ME_VERSION;
    document.body.appendChild(v);

    const btn = document.createElement("button");
    btn.id = "me-diag-btn";
    btn.textContent = "Diagnostic devis";
    btn.style.cssText = "position:fixed;bottom:12px;left:330px;z-index:9999;font:600 11px/1 Arial;"
      + "padding:7px 12px;border-radius:8px;background:#b3741f;color:#fff;border:0;cursor:pointer";
    btn.onclick = runDiagnostic;
    document.body.appendChild(btn);

    const tbtn = document.createElement("button");
    tbtn.id = "me-tpl-btn";
    tbtn.textContent = "Modèle Word";
    tbtn.style.cssText = "position:fixed;bottom:12px;left:448px;z-index:9999;font:600 11px/1 Arial;"
      + "padding:7px 12px;border-radius:8px;background:#2d6a4f;color:#fff;border:0;cursor:pointer";
    tbtn.onclick = openTemplateManager;
    document.body.appendChild(tbtn);
  }

  // ---- Importer / remplacer le modèle Word maître ----
  function openTemplateManager() {
    const code = (window.state && state.devis && state.devis.modele_code) ? state.devis.modele_code : "copro_petite";
    const html = ""
      + "<h3 style='margin:0 0 10px'>Modèle Word maître</h3>"
      + "<p style='font:13px Arial;color:#444;margin:0 0 14px'>Téléchargez le modèle, corrigez la mise en page dans Word "
      + "(en gardant les variables {{...}}), puis réimportez-le. Les devis suivants utiliseront votre mise en page.</p>"
      + "<div style='display:flex;gap:10px;align-items:center;flex-wrap:wrap'>"
      + "<a id='me-tpl-dl' href='" + API + "/templates/" + code + "/download' "
      + "style='padding:8px 16px;border-radius:8px;background:#1f2430;color:#fff;text-decoration:none;font:600 13px Arial'>"
      + "⬇ Télécharger le modèle actuel</a>"
      + "<label style='padding:8px 16px;border-radius:8px;background:#2d6a4f;color:#fff;cursor:pointer;font:600 13px Arial'>"
      + "⬆ Importer mon modèle corrigé"
      + "<input id='me-tpl-file' type='file' accept='.docx,.doc' style='display:none'></label>"
      + "<span style='font:13px Arial;color:#2d6a4f' id='me-tpl-code'>modèle : " + code + "</span>"
      + "</div>"
      + "<div id='me-tpl-images' style='margin-top:18px'></div>"
      + "<div id='me-tpl-status' style='font:13px Arial;margin-top:14px'></div>";
    showDiagModal(html);
    const input = document.getElementById("me-tpl-file");
    if (input) input.onchange = () => uploadTemplate(code, input.files[0]);
    loadTemplateImages(code, "me-tpl-images");
  }

  async function uploadTemplate(code, file) {
    const status = document.getElementById("me-tpl-status");
    if (!file) return;
    if (status) status.innerHTML = "Import en cours…";
    try {
      const fd = new FormData();
      fd.append("fichier", file);
      const r = await fetch(API + "/templates/" + code + "/replace", { method: "POST", body: fd });
      const j = await r.json();
      if (r.ok && j.ok) {
        if (status) status.innerHTML = "<b style='color:#2d6a4f'>✓ Modèle remplacé.</b> "
          + (j.variables_detectees ? j.variables_detectees.length : 0) + " variables détectées. "
          + "Les prochains devis utiliseront votre mise en page. "
          + "<br>Cliquez sur « Rafraîchir l'aperçu » pour voir le résultat.";
        if (typeof updatePreview === "function") setTimeout(updatePreview, 600);
      } else {
        if (status) status.innerHTML = "<b style='color:#b00'>Échec :</b> " + (j.detail || JSON.stringify(j));
      }
    } catch (e) {
      if (status) status.innerHTML = "<b style='color:#b00'>Erreur :</b> " + e;
    }
  }

  async function loadTemplateImages(code, targetId) {
    const target = document.getElementById(targetId);
    if (!target) return;
    target.innerHTML = "<div style='font:13px Arial;color:#777'>Chargement des images du modèle…</div>";
    try {
      const r = await fetch(API + "/templates/" + code + "/images");
      const j = await r.json();
      if (!r.ok || !j.ok) {
        target.innerHTML = "<div style='font:13px Arial;color:#b42318'>Images indisponibles : " + esc(j.detail || j.message || "erreur") + "</div>";
        return;
      }
      const images = j.images || [];
      if (!images.length) {
        target.innerHTML = "<div style='font:13px Arial;color:#777'>Aucune image embarquée dans ce modèle.</div>";
        return;
      }
      target.innerHTML = "<h4 style='margin:0 0 10px;font:700 14px Arial'>Images du modèle</h4>"
        + "<div style='display:grid;grid-template-columns:repeat(auto-fill,minmax(170px,1fr));gap:12px'>"
        + images.map(img => {
            const meta = (img.width && img.height) ? img.width + " × " + img.height + " px" : Math.round((img.bytes || 0) / 1024) + " Ko";
            return "<div style='border:1px solid #e3ded3;border-radius:10px;padding:10px;background:#fff'>"
              + "<div style='height:95px;display:flex;align-items:center;justify-content:center;background:#f7f4ee;border-radius:8px;overflow:hidden'>"
              + "<img src='" + API + "/templates/" + code + "/images/" + encodeURIComponent(img.name) + "?v=" + Date.now() + "' style='max-width:100%;max-height:95px;object-fit:contain'>"
              + "</div>"
              + "<div style='font:600 12px Arial;margin-top:8px;word-break:break-all'>" + esc(img.name) + "</div>"
              + "<div style='font:11px Arial;color:#888;margin-top:2px'>" + esc(meta) + "</div>"
              + "<label style='display:inline-block;margin-top:8px;padding:6px 10px;border-radius:7px;background:#2d6a4f;color:#fff;font:600 11px Arial;cursor:pointer'>Remplacer"
              + "<input type='file' accept='.jpg,.jpeg,.png,.webp,image/*' style='display:none' onchange='window.dtReplaceImage(\"" + code + "\",\"" + img.name + "\",this)'></label>"
              + "</div>";
          }).join("")
        + "</div>";
    } catch (e) {
      target.innerHTML = "<div style='font:13px Arial;color:#b42318'>Images indisponibles : " + esc(e) + "</div>";
    }
  }

  window.dtReplaceImage = async (code, imageName, input) => {
    const file = input && input.files && input.files[0];
    if (!file) return;
    const fd = new FormData();
    fd.append("fichier", file);
    const r = await fetch(API + "/templates/" + code + "/images/" + encodeURIComponent(imageName), { method: "POST", body: fd });
    const j = await r.json();
    if (r.ok && j.ok) {
      alert("Image remplacée dans le modèle.");
      loadTemplateImages(code, "me-tpl-images");
      loadTemplateImages(code, "dt-images-panel");
    } else {
      alert("Échec : " + (j.detail || j.message || JSON.stringify(j)));
    }
  };

  window.dtImages = (code) => {
    showDiagModal("<h3 style='margin:0 0 10px'>Images du modèle " + esc(code) + "</h3><div id='dt-images-panel'></div>");
    loadTemplateImages(code, "dt-images-panel");
  };

  async function runDiagnostic() {
    let html = "<h3 style='margin:0 0 10px'>Diagnostic devis</h3>";
    try {
      // 1. version backend (template, sauts de page)
      let ver = {};
      try { ver = await (await fetch(API + "/version")).json(); } catch (e) {}
      // 2. payload courant (zones, fréquences, SHOW_*)
      const payload = window.mapDevisToPayload ? window.mapDevisToPayload() : {};
      const showVars = {};
      Object.keys(payload.variables || {}).forEach(k => { if (k.startsWith("SHOW_") || k.startsWith("FREQ_")) showVars[k] = payload.variables[k]; });
      // 3. générer pour obtenir nom DOCX + nb pages + page blanche
      let gen = {};
      try {
        const r = await fetch(API + "/devis/generer", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(payload) });
        gen = await r.json();
      } catch (e) { gen = { erreur: String(e) }; }

      const row = (k, v) => "<tr><td style='padding:3px 10px 3px 0;color:#666;vertical-align:top'>" + k + "</td><td style='padding:3px 0'><b>" + v + "</b></td></tr>";
      html += "<table style='font:13px Arial;border-collapse:collapse'>";
      html += row("Version interface", ME_VERSION);
      html += row("Modèle Word utilisé", payload.template_code || "?");
      html += row("Template à jour (backend)", ver.tout_ok === true ? "OUI ✓" : JSON.stringify(ver.checks || {}));
      html += row("Sauts de page (modèle)", (ver.checks && ver.checks["sauts_de_page"] != null) ? ver.checks["sauts_de_page"] : "voir /api/version");
      html += row("Zones cochées", (payload.zones_selectionnees || []).join(", ") || "(aucune)");
      html += row("Fréquences / SHOW envoyés", "<pre style='margin:0;white-space:pre-wrap'>" + JSON.stringify(showVars, null, 1) + "</pre>");
      html += row("zones_detail", "<pre style='margin:0;white-space:pre-wrap'>" + JSON.stringify(payload.zones_detail || [], null, 1) + "</pre>");
      html += row("Fichier DOCX généré", gen.docx_url ? gen.docx_url.split("/").pop() : (gen.erreur || "?"));
      html += row("Nombre de pages PDF", gen.pdf_pages != null ? gen.pdf_pages : "?");
      html += row("Page blanche détectée", gen.page_blanche === true ? "OUI ✗" : (gen.page_blanche === false ? "NON ✓" : "?"));
      html += "</table>";
    } catch (e) {
      html += "<p style='color:#b00'>Erreur diagnostic : " + e + "</p>";
    }
    showDiagModal(html);
  }

  function showDiagModal(html) {
    let ov = document.getElementById("me-diag-modal");
    if (ov) ov.remove();
    ov = document.createElement("div");
    ov.id = "me-diag-modal";
    ov.style.cssText = "position:fixed;inset:0;z-index:10000;background:rgba(0,0,0,.4);display:flex;align-items:center;justify-content:center";
    ov.onclick = e => { if (e.target === ov) ov.remove(); };
    const box = document.createElement("div");
    box.style.cssText = "background:#fff;max-width:640px;max-height:80vh;overflow:auto;padding:22px 26px;border-radius:12px;box-shadow:0 10px 40px rgba(0,0,0,.3)";
    box.innerHTML = html + "<div style='text-align:right;margin-top:14px'><button onclick=\"document.getElementById('me-diag-modal').remove()\" style='padding:7px 16px;border:0;border-radius:8px;background:#1f2430;color:#fff;cursor:pointer'>Fermer</button></div>";
    ov.appendChild(box);
    document.body.appendChild(ov);
  }
  window.runDiagnostic = runDiagnostic;
  window.openTemplateManager = openTemplateManager;
  window.renderDevisTypes = renderDevisTypes;

  // ---- Page "Devis types" (gestion des modèles) ----
  function installDevisTypesPage() {
    if (document.getElementById("view-devistypes")) return;
    // 1. Entrée de navigation (après "Bibliothèque")
    const nav = document.querySelector('.nav__item[data-view="biblio"]');
    if (nav && nav.parentElement) {
      const li = document.createElement("li");
      li.className = "nav__item";
      li.setAttribute("data-view", "devistypes");
      li.style.cssText = "cursor:pointer";
      li.innerHTML = '<span style="display:inline-block;width:18px">▦</span> Devis types';
      li.onclick = () => { switchView("devistypes"); renderDevisTypes(); };
      nav.parentElement.insertBefore(li, nav.nextSibling);
    }
    // 2. Conteneur de vue
    const main = document.querySelector("#main") || document.body;
    const view = document.createElement("div");
    view.className = "view";
    view.id = "view-devistypes";
    view.style.cssText = "padding:32px 40px;overflow:auto";
    view.innerHTML = "<div id='me-dt-content'>Chargement…</div>";
    main.appendChild(view);
  }

  async function renderDevisTypes() {
    const box = document.getElementById("me-dt-content");
    if (!box) return;
    box.innerHTML = "Chargement des devis types…";
    let list = [];
    try {
      const resp = await fetch(API + "/templates");
      const data = await resp.json();   // toujours du JSON désormais
      if (!resp.ok || (data && data.error)) {
        box.innerHTML = "<div style='font:14px Arial;color:#b00;padding:20px;border:1px solid #f0c0c0;border-radius:10px;background:#fdf2f2'>"
          + "<b>" + esc((data && data.message) || "Erreur de chargement") + "</b>"
          + (data && data.details ? "<br><span style='color:#a55;font-size:12px'>" + esc(data.details) + "</span>" : "")
          + "</div>";
        return;
      }
      list = data;
    } catch (e) {
      box.innerHTML = "<div style='font:14px Arial;color:#b00;padding:20px'>Erreur de chargement : " + esc(String(e)) + "</div>";
      return;
    }

    let h = "<div style='display:flex;align-items:center;justify-content:space-between;margin-bottom:20px'>"
      + "<h1 style='font:700 24px Arial;margin:0'>Devis types</h1>"
      + "<span style='font:13px Arial;color:#777'>" + list.length + " modèle(s) — base de travail réutilisable</span></div>"
      + "<p style='font:13px Arial;color:#666;margin:0 0 18px'>Ce sont vos <b>modèles</b> (pas les devis clients). "
      + "Corrigez-les une fois, puis créez des devis à partir d'eux.</p>";

    h += "<div style='display:grid;gap:14px'>";
    list.forEach(t => {
      const dnatecreate = t.created_at ? t.created_at.split("T")[0] : "?";
      const datemaj = t.updated_at ? t.updated_at.split("T")[0] : "?";
      const fam = t.famille === "contrat" ? "Récurrent" : "Ponctuel";
      h += "<div style='border:1px solid #e3ded3;border-radius:12px;padding:16px 18px;background:#fff'>"
        + "<div style='display:flex;justify-content:space-between;align-items:flex-start;gap:12px'>"
        + "<div><div style='font:700 16px Arial'>" + esc(t.nom) + "</div>"
        + "<div style='font:12px Arial;color:#888;margin-top:3px'>"
        + "<span style='background:#f0ece2;padding:2px 8px;border-radius:6px'>" + fam + "</span> "
        + " &nbsp;code : <code>" + esc(t.code) + "</code> &nbsp; fichier : <code>" + esc(t.fichier) + "</code></div>"
        + (t.description ? "<div style='font:13px Arial;color:#555;margin-top:6px'>" + esc(t.description) + "</div>" : "")
        + "<div style='font:11px Arial;color:#aaa;margin-top:6px'>créé le " + dnatecreate + " · modifié le " + datemaj
        + " · " + (t.variables ? t.variables.length : 0) + " variables</div>"
        + "</div></div>"
        + "<div style='display:flex;flex-wrap:wrap;gap:8px;margin-top:14px'>"
        + btn("Créer un devis", "#b3741f", "window.dtCreateDevis('" + t.code + "')")
        + btn("Aperçu du modèle", "#1f2430", "window.dtPreview('" + t.code + "')")
        + btn("Modifier", "#2d6a4f", "window.dtEdit('" + t.code + "')")
        + btn("Images", "#7a5c2e", "window.dtImages('" + t.code + "')")
        + btn("Dupliquer", "#5a4ea3", "window.dtDuplicate('" + t.code + "')")
        + btn("Télécharger Word", "#444", "window.dtDownload('" + t.code + "')")
        + btn("Importer un Word", "#0a6", "window.dtReplace('" + t.code + "')")
        + "</div></div>";
    });
    h += "</div>";
    box.innerHTML = h;
  }
  function esc(s) { return String(s == null ? "" : s).replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c])); }
  function btn(label, color, onclick) {
    return "<button onclick=\"" + onclick + "\" style='font:600 12px Arial;padding:8px 14px;border:0;border-radius:8px;background:" + color + ";color:#fff;cursor:pointer'>" + label + "</button>";
  }

  window.dtDownload = (code) => { window.open(API + "/templates/" + code + "/download", "_blank"); };
  window.dtPreview = (code) => { window.open(API + "/templates/" + code + "/preview-sample", "_blank"); };

  window.dtReplace = (code) => {
    const inp = document.createElement("input");
    inp.type = "file"; inp.accept = ".docx,.doc";
    inp.onchange = async () => {
      const fd = new FormData(); fd.append("fichier", inp.files[0]);
      const r = await fetch(API + "/templates/" + code + "/replace", { method: "POST", body: fd });
      const j = await r.json();
      alert(r.ok && j.ok ? "Modèle Word remplacé pour « " + code + " ». Les nouveaux devis utiliseront ce fichier." : "Échec : " + (j.detail || JSON.stringify(j)));
      renderDevisTypes();
    };
    inp.click();
  };

  window.dtDuplicate = async (code) => {
    const nouveau_nom = prompt("Nom du nouveau modèle :", "Copie de " + code);
    if (!nouveau_nom) return;
    const nouveau_code = prompt("Code unique du nouveau modèle (sans espaces) :", code + "_copie");
    if (!nouveau_code) return;
    const fd = new FormData(); fd.append("nouveau_code", nouveau_code); fd.append("nouveau_nom", nouveau_nom);
    const r = await fetch(API + "/templates/" + code + "/duplicate", { method: "POST", body: fd });
    const j = await r.json();
    alert(r.ok && j.ok ? j.message : "Échec : " + (j.detail || JSON.stringify(j)));
    renderDevisTypes();
  };

  function dtToolButton(label, cmd, value, title) {
    const val = value == null ? "" : String(value).replace(/'/g, "&#39;");
    return "<button type='button' class='dt-ribbon-btn' data-cmd='" + cmd + "' data-value='" + val + "' title='" + esc(title || label) + "'>" + label + "</button>";
  }
  function dtRibbonGroup(label, content) {
    return "<div class='dt-ribbon-group'>" + content + "<span class='dt-ribbon-label'>" + label + "</span></div>";
  }
  function dtEditorToolbar() {
    return "<div class='dt-ribbon' id='me-dt-toolbar' aria-label='Banni&egrave;re de mise en forme'>"
      + dtRibbonGroup("Fichier", "<button type='button' class='dt-ribbon-primary' onclick='window.dtSaveContent()'>Enregistrer</button>" + "<button type='button' class='dt-ribbon-btn dt-ribbon-wide' onclick='window.dtRefreshPreview()'>Actualiser l&rsquo;aper&ccedil;u</button>")
      + dtRibbonGroup("Police", "<select class='dt-ribbon-select dt-ribbon-font' data-format='fontName'><option>Aptos</option><option>Arial</option><option>Calibri</option><option>Times New Roman</option><option>Verdana</option><option>Tahoma</option></select>" + "<select class='dt-ribbon-select' data-format='block'><option value='P'>Normal</option><option value='H1'>Titre 1</option><option value='H2'>Titre 2</option><option value='H3'>Titre 3</option></select>" + "<select class='dt-ribbon-select dt-ribbon-size' data-format='size'><option value='2'>10</option><option value='3' selected>11</option><option value='4'>12</option><option value='5'>14</option><option value='6'>16</option><option value='7'>18</option></select>" + "<select class='dt-ribbon-select dt-ribbon-color' data-format='foreColor' title='Couleur du texte'><option value='#000000'>Noir</option><option value='#666666'>Gris</option><option value='#c00000'>Rouge</option><option value='#1f4e79'>Bleu</option><option value='#2e7d32'>Vert</option><option value='#c55a11'>Orange</option></select><input class='dt-ribbon-colorbox' data-format='foreColorCustom' type='color' value='#000000' title='Couleur personnalis&eacute;e'>" + "<select class='dt-ribbon-select dt-ribbon-highlight' data-format='hiliteColor' title='Surlignage'><option value=''>Sans surlignage</option><option value='#fff200'>Jaune</option><option value='#d9ead3'>Vert clair</option><option value='#cfe2f3'>Bleu clair</option><option value='#f4cccc'>Rouge clair</option></select>" + dtToolButton("<strong>G</strong>", "bold", null, "Gras") + dtToolButton("<em>I</em>", "italic", null, "Italique") + dtToolButton("<u>S</u>", "underline", null, "Soulign&eacute;"))
      + dtRibbonGroup("Paragraphe", dtToolButton("Gauche", "justifyLeft", null, "Aligner &agrave; gauche") + dtToolButton("Centre", "justifyCenter", null, "Centrer") + dtToolButton("Droite", "justifyRight", null, "Aligner &agrave; droite") + dtToolButton("&bull; Liste", "insertUnorderedList", null, "Liste &agrave; puces") + dtToolButton("1. Liste", "insertOrderedList", null, "Liste num&eacute;rot&eacute;e") + dtToolButton("- Retrait", "outdent", null, "Diminuer le retrait") + dtToolButton("+ Retrait", "indent", null, "Augmenter le retrait"))
      + dtRibbonGroup("&Eacute;dition", dtToolButton("Annuler", "undo", null, "Annuler") + dtToolButton("R&eacute;tablir", "redo", null, "R&eacute;tablir"))
      + "</div>";
  }
  function dtApplyTypingStyle() {
    const editor = document.getElementById("me-word-editor");
    const st = window._dtActiveStyle || {};
    if (!editor || !document.activeElement || !editor.contains(document.activeElement)) return;
    if (st.fontName) document.execCommand("fontName", false, st.fontName);
    if (st.fontSize) document.execCommand("fontSize", false, st.fontSize);
    if (st.foreColor) document.execCommand("foreColor", false, st.foreColor);
    try {
      if (st.hiliteColor) document.execCommand("hiliteColor", false, st.hiliteColor);
    } catch (e) {}
  }
  function dtSetStatus(message, type) {
    const el = document.getElementById("me-dt-status");
    if (!el) return;
    el.textContent = message || "";
    el.className = "dt-status" + (type ? " dt-status--" + type : "");
  }
  async function dtRenderPreview(code) {
    const cont = document.getElementById("me-dt-preview");
    if (!cont) return;
    cont.innerHTML = "<div class='dt-preview-loading'>G&eacute;n&eacute;ration de l&rsquo;aper&ccedil;u PDF...</div>";
    try {
      const resp = await fetch(API + "/templates/" + code + "/preview-sample?t=" + Date.now());
      if (!resp.ok) throw new Error(await resp.text());
      const buf = await resp.arrayBuffer();
      cont.innerHTML = "";
      await _renderPdfCanvas(buf, cont, _apercuSeq);
      dtSetStatus("Aper\u00e7u actualis\u00e9.", "ok");
    } catch (e) {
      cont.innerHTML = "<div class='dt-preview-error'>Aper&ccedil;u indisponible : " + esc(String(e)) + "</div>";
      dtSetStatus("Impossible d\u2019actualiser l\u2019aper\u00e7u.", "error");
    }
  }
  window.dtRefreshPreview = async () => {
    if (!window._dtEditCode) return;
    await dtRenderPreview(window._dtEditCode);
  };
  function dtWireToolbar() {
    const toolbar = document.getElementById("me-dt-toolbar");
    const editor = document.getElementById("me-word-editor");
    if (!toolbar || !editor) return;
    toolbar.addEventListener("mousedown", e => { if (e.target.closest("button")) e.preventDefault(); });
    window._dtActiveStyle = window._dtActiveStyle || { fontName: "Aptos", fontSize: "3", foreColor: "#000000", hiliteColor: "" };
    try { document.execCommand("styleWithCSS", false, true); document.execCommand("defaultParagraphSeparator", false, "p"); } catch (e) {}
    toolbar.querySelectorAll("[data-cmd]").forEach(b => {
      b.onclick = () => { editor.focus(); document.execCommand(b.dataset.cmd, false, b.dataset.value || null); dtMarkSelectionChanged(); };
    });
    toolbar.querySelectorAll("select[data-format]").forEach(sel => {
      sel.onchange = () => {
        editor.focus();
        const fmt = sel.dataset.format;
        if (fmt === "block") document.execCommand("formatBlock", false, sel.value);
        if (fmt === "size") { window._dtActiveStyle.fontSize = sel.value; document.execCommand("fontSize", false, sel.value); }
        if (fmt === "fontName") { window._dtActiveStyle.fontName = sel.value; document.execCommand("fontName", false, sel.value); }
        if (fmt === "foreColor") { window._dtActiveStyle.foreColor = sel.value; document.execCommand("foreColor", false, sel.value); const c = toolbar.querySelector("[data-format='foreColorCustom']"); if (c) c.value = sel.value; }
        if (fmt === "hiliteColor") { window._dtActiveStyle.hiliteColor = sel.value; document.execCommand("hiliteColor", false, sel.value || "transparent"); }
        dtMarkSelectionChanged();
      };
    });
    toolbar.querySelectorAll("input[data-format='foreColorCustom']").forEach(inp => {
      inp.oninput = () => { editor.focus(); window._dtActiveStyle.foreColor = inp.value; document.execCommand("foreColor", false, inp.value); dtMarkSelectionChanged(); };
    });
  }
  function dtMarkSelectionChanged() {
    const editor = document.getElementById("me-word-editor");
    if (!editor) return;
    editor.querySelectorAll("[data-ref][contenteditable='true']").forEach(el => { window._dtEditMods[el.dataset.ref] = el.innerHTML; });
    dtSetStatus("Modifications non enregistr\u00e9es.", "dirty");
  }

  window.dtEdit = async (code) => {
    switchView("devistypes");
    const box = document.getElementById("me-dt-content");
    box.innerHTML = "<div class='dt-loading'>Ouverture de l&rsquo;&eacute;diteur...</div>";
    let data;
    try {
      const r = await fetch(API + "/templates/" + code + "/contenu");
      data = await r.json();
      if (!r.ok || data.error) throw new Error(data.message || "Erreur");
    } catch (e) { box.innerHTML = "<div class='dt-error'>Erreur : " + esc(String(e)) + "</div>"; return; }
    window._dtEditCode = code;
    window._dtEditMods = {};
    const items = (data.paragraphs && data.paragraphs.length) ? data.paragraphs : data.blocs;
    let html = "<section class='dt-workbench'><header class='dt-editor-head'><div><span class='dt-kicker'>Devis type</span><h1>" + esc(data.nom) + "</h1><p>Modifiez le contenu du mod&egrave;le &agrave; gauche et comparez le rendu PDF &agrave; droite.</p></div><div class='dt-editor-actions'><button class='dt-action dt-action--primary' onclick='window.dtSaveContent()'>Enregistrer</button><button class='dt-action' onclick='window.dtRefreshPreview()'>Pr&eacute;visualiser</button><button class='dt-action' onclick='window.dtPreview(\"" + code + "\")'>Ouvrir le PDF</button><button class='dt-action' onclick='window.dtCreateDevis(\"" + code + "\")'>G&eacute;n&eacute;rer un test</button></div></header>" + dtEditorToolbar() + "<div id='me-dt-status' class='dt-status'>Pr&ecirc;t.</div><div class='dt-editor-layout'><section class='dt-panel dt-panel--editor'><div class='dt-panel-head'><strong>Contenu modifiable</strong><span>Variables conserv&eacute;es</span></div><main id='me-word-editor' class='dt-word-page' contenteditable='false'>";
    items.forEach(item => {
      const tag = item.type === "titre" ? "h2" : "p";
      if (item.editable) html += "<" + tag + " contenteditable='true' data-ref='" + esc(item.ref) + "' class='dt-edit-line dt-edit-" + esc(item.type || "text") + "'>" + esc(item.texte).replace(/\n/g, "<br>") + "</" + tag + ">";
      else html += "<p contenteditable='false' data-ref='" + esc(item.ref) + "' class='dt-protected' title='" + esc(item.protege_raison || "balise technique") + "'>" + esc(item.texte).replace(/\n/g, "<br>") + "</p>";
    });
    html += "</main></section><aside class='dt-panel dt-panel--preview'><div class='dt-panel-head'><strong>Aper&ccedil;u du devis</strong><button type='button' onclick='window.dtRefreshPreview()'>Actualiser</button></div><div id='me-dt-preview' class='dt-pdf-preview'></div></aside></div></section>";
    box.innerHTML = html;
    dtWireToolbar();
    const editor = document.getElementById("me-word-editor");
    editor.addEventListener("input", e => { const el = e.target.closest("[data-ref][contenteditable='true']"); if (el) window._dtEditMods[el.dataset.ref] = el.innerHTML; dtSetStatus("Modifications non enregistr\u00e9es.", "dirty"); });
    editor.addEventListener("keydown", e => { if (e.key === "Enter") setTimeout(dtApplyTypingStyle, 0); });
    editor.addEventListener("keyup", e => { if (["Enter","ArrowLeft","ArrowRight","ArrowUp","ArrowDown"].includes(e.key)) setTimeout(dtApplyTypingStyle, 0); });
    editor.addEventListener("mouseup", () => setTimeout(dtApplyTypingStyle, 0));
    editor.addEventListener("focusin", () => setTimeout(dtApplyTypingStyle, 0));
    editor.addEventListener("paste", e => { e.preventDefault(); const text = (e.clipboardData || window.clipboardData).getData("text/plain"); document.execCommand("insertText", false, text); });
    await dtRenderPreview(code);
  };

  window.dtSaveContent = async () => {
    const code = window._dtEditCode;
    const mods = window._dtEditMods || {};
    if (!Object.keys(mods).length) { dtSetStatus("Aucune modification \u00e0 enregistrer.", "ok"); return; }
    dtSetStatus("Enregistrement du mod\u00e8le...", "dirty");
    try {
      const r = await fetch(API + "/templates/" + code + "/contenu", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ html_modifications: mods }) });
      const j = await r.json();
      if (r.ok && j.ok) { window._dtEditMods = {}; const ignored = j.refusees && j.refusees.length ? " " + j.refusees.length + " bloc(s) prot\u00e9g\u00e9(s) ignor\u00e9(s)." : ""; dtSetStatus("Mod\u00e8le enregistr\u00e9." + ignored, "ok"); await dtRenderPreview(code); }
      else dtSetStatus("\u00c9chec de l\u2019enregistrement : " + (j.message || JSON.stringify(j)), "error");
    } catch (e) { dtSetStatus("Erreur d\u2019enregistrement : " + e, "error"); }
  };

  window.dtCreateDevis = (code) => {
    // Ouvre le formulaire de création avec ce modèle comme base
    switchView("creation");
    try {
      if (typeof selectType === "function") {
        // déduire la famille
        fetch(API + "/templates").then(r => r.json()).then(list => {
          const t = list.find(x => x.code === code);
          const fam = (t && t.famille === "ponctuel") ? "ponctuel" : "contrat";
          selectType(fam === "contrat" ? "contrat" : "ponctuel");
          if (typeof selectModele === "function") {
            // trouver le code prototype correspondant si mappé, sinon utiliser tel quel
            try { selectModele(code); } catch (e) {}
          }
          if (typeof updatePreview === "function") setTimeout(updatePreview, 500);
        });
      }
    } catch (e) {}
  };

  // Lancer après le DOMContentLoaded d'origine
  function _boot() { wireWrites(); installRealPreview(); overrideProtoRenders(); bootFromBackend(); installVersionBadge(); installDevisTypesPage(); setTimeout(rafraichirApercuPDF, 1200); }
  if (document.readyState === "complete" || document.readyState === "interactive") {
    setTimeout(_boot, 300);
  } else {
    document.addEventListener("DOMContentLoaded", () => setTimeout(_boot, 300));
  }
})();

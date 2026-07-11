/**
 * DevisFlow — JS frontend
 *
 * Logique :
 *  - Navigation entre 3 vues : Création / Mes devis / Modèles
 *  - Création : récupère les variables du template sélectionné, construit
 *    le payload, appelle POST /api/devis/generer, propose les téléchargements
 *  - Mes devis : GET /api/devis avec liens DOCX + PDF
 *  - Modèles : GET /api/templates + upload (multipart) + actions CRUD
 */

const API = "/api";

// ============================================================
// Helpers
// ============================================================
async function api(path, opts = {}) {
  const url = path.startsWith("http") ? path : API + path;
  const res = await fetch(url, opts);
  if (!res.ok) {
    const txt = await res.text().catch(() => "");
    throw new Error(`${res.status} ${res.statusText} — ${txt}`);
  }
  if (res.headers.get("content-type")?.includes("application/json")) {
    return res.json();
  }
  return res;
}

function setStatus(elId, msg, kind = "") {
  const el = document.getElementById(elId);
  if (!el) return;
  el.textContent = msg;
  el.className = kind ? `hint gen-status--${kind}` : "hint";
}

// ============================================================
// Navigation
// ============================================================
document.querySelectorAll(".nav-tab").forEach(tab => {
  tab.addEventListener("click", () => {
    const target = tab.dataset.view;
    document.querySelectorAll(".nav-tab").forEach(t => t.classList.toggle("is-active", t === tab));
    document.querySelectorAll(".view").forEach(v => v.classList.toggle("is-active", v.id === "view-" + target));
    if (target === "devis-list") renderDevisList();
    if (target === "templates")   renderTemplatesList();
  });
});

// ============================================================
// VUE : Création — éditeur de prestations + tarification
// ============================================================
const TECHNICITE = ["standard", "technique", "haute", "exceptionnelle"];

function ligneRowHTML() {
  const opts = TECHNICITE.map(t => `<option value="${t}">${t}</option>`).join("");
  return `
    <tr class="ligne-row">
      <td><input class="l-libelle" placeholder="Ex : Débarras des caves"></td>
      <td><input class="l-duree" type="number" min="0" step="0.25" value="1" style="width:64px"></td>
      <td><input class="l-agents" type="number" min="1" step="1" value="1" style="width:56px"></td>
      <td><input class="l-taux" type="number" min="0" step="0.5" value="24" style="width:72px"></td>
      <td><select class="l-tech">${opts}</select></td>
      <td><input class="l-frais" type="number" min="0" step="1" value="0" style="width:72px"></td>
      <td><button class="btn btn--sm btn--danger l-del" type="button">×</button></td>
    </tr>`;
}

function addLigne() {
  const body = document.getElementById("lignes-body");
  body.insertAdjacentHTML("beforeend", ligneRowHTML());
  body.querySelectorAll(".l-del").forEach(b => b.onclick = (e) => { e.target.closest("tr").remove(); });
}

function collectLignes() {
  return Array.from(document.querySelectorAll("#lignes-body .ligne-row")).map(tr => ({
    libelle: tr.querySelector(".l-libelle").value.trim(),
    duree_h: parseFloat(tr.querySelector(".l-duree").value) || 0,
    nb_agents: parseFloat(tr.querySelector(".l-agents").value) || 1,
    taux_horaire: parseFloat(tr.querySelector(".l-taux").value) || 0,
    niveau_technicite: tr.querySelector(".l-tech").value,
    frais: parseFloat(tr.querySelector(".l-frais").value) || 0,
  }));
}

function selectedFrequences() {
  return Array.from(document.querySelectorAll(".freq-opt:checked")).map(c => c.value);
}

function currentTemplateFamille(list, code) {
  const t = (list || []).find(x => x.code === code);
  return t ? t.famille : "contrat";
}

async function recalcPrix() {
  const code = document.getElementById("f-template-code").value;
  if (!code) return;
  let list = [];
  try { list = await api("/templates"); } catch (e) {}
  const famille = currentTemplateFamille(list, code);
  const isRec = famille === "contrat";
  document.getElementById("freq-options-group").style.display = isRec ? "" : "none";

  const force = document.getElementById("f-prix-force").value.trim();
  const payload = {
    numero: document.getElementById("f-numero").value.trim() || "TMP",
    template_code: code,
    variables: { _famille: famille },
    lignes: collectLignes(),
    frequences_options: isRec ? selectedFrequences() : [],
    prix_force_ht: force === "" ? null : parseFloat(force),
  };
  try {
    const res = await api("/devis/calculer", {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    const tot = res.totaux;
    document.getElementById("t-ht").textContent = tot.total_ht_fmt;
    document.getElementById("t-tva").textContent = tot.total_tva_fmt;
    document.getElementById("t-ttc").textContent = tot.total_ttc_fmt;
  } catch (e) {
    document.getElementById("t-ht").textContent = "erreur";
  }
}

document.getElementById("btn-add-ligne").addEventListener("click", addLigne);
document.getElementById("btn-recalc").addEventListener("click", recalcPrix);

// ============================================================
// VUE : Création — génération
// ============================================================
async function loadTemplatesForSelect() {
  const sel = document.getElementById("f-template-code");
  try {
    const list = await api("/templates");
    if (list.length === 0) {
      sel.innerHTML = `<option value="">(aucun modèle — importez-en un dans l'onglet "Modèles PDF")</option>`;
      return;
    }
    sel.innerHTML = list.filter(t => t.actif).map(t =>
      `<option value="${t.code}" ${t.is_default ? "selected" : ""}>${t.nom}${t.is_default ? " (défaut)" : ""}</option>`
    ).join("");
    updateTemplateHint();
  } catch (e) {
    sel.innerHTML = `<option value="">Erreur de chargement</option>`;
    console.error(e);
  }
}

async function updateTemplateHint() {
  const code = document.getElementById("f-template-code").value;
  if (!code) return;
  try {
    const list = await api("/templates");
    const t = list.find(x => x.code === code);
    if (!t) return;
    const vars = t.variables || [];
    document.getElementById("f-template-hint").textContent =
      vars.length > 0
        ? `${vars.length} variable(s) détectée(s) : ${vars.join(", ")}`
        : "Aucune variable détectée — le template n'est pas annoté.";
  } catch (e) { /* silent */ }
}

document.getElementById("f-template-code").addEventListener("change", () => { updateTemplateHint(); recalcPrix(); });

document.getElementById("btn-generer").addEventListener("click", async () => {
  const code = document.getElementById("f-template-code").value;
  if (!code) return setStatus("gen-status", "Sélectionnez un modèle", "err");

  const numero = document.getElementById("f-numero").value.trim();
  if (!numero) return setStatus("gen-status", "Numéro de devis obligatoire", "err");

  let list = [];
  try { list = await api("/templates"); } catch (e) {}
  const famille = currentTemplateFamille(list, code);
  const isRec = famille === "contrat";

  // Variables simples (le backend complète OPTIONS / PRESTATIONS / FORFAIT_* via le calcul)
  const variables = {
    NUMERO_DEVIS:    numero,
    DATE_EMISSION:   document.getElementById("f-date-emission").value.trim(),
    TYPE_PRESTATION: document.getElementById("f-type-prestation").value.trim(),
    DEST_LIGNE1:     document.getElementById("f-dest1").value.trim(),
    DEST_LIGNE2:     document.getElementById("f-dest2").value.trim(),
    DEST_LIGNE3:     document.getElementById("f-dest3").value.trim(),
    DEST_LIGNE4:     document.getElementById("f-dest4").value.trim(),
    SITE_ADRESSE:    document.getElementById("f-site-adresse").value.trim(),
    SITE_CP_VILLE:   document.getElementById("f-site-cp-ville").value.trim(),
  };

  const force = document.getElementById("f-prix-force").value.trim();
  const payload = {
    numero,
    template_code: code,
    variables,
    lignes: collectLignes(),
    frequences_options: isRec ? selectedFrequences() : [],
    prix_force_ht: force === "" ? null : parseFloat(force),
    client_nom:    document.getElementById("f-dest2").value.trim(),
    site_adresse:  document.getElementById("f-site-adresse").value.trim(),
    date_emission: document.getElementById("f-date-emission").value.trim(),
  };

  setStatus("gen-status", "Génération en cours…");
  try {
    const result = await api("/devis/generer", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    const tot = result.totaux || {};
    let html = `Devis ${numero} généré (HT ${tot.total_ht_fmt || "—"}, TTC ${tot.total_ttc_fmt || "—"}). `;
    html += `<a href="${result.docx_url}" download>Télécharger .docx</a>`;
    if (result.pdf_url) html += ` &nbsp;·&nbsp; <a href="${result.pdf_url}" download>Télécharger .pdf</a>`;
    document.getElementById("gen-status").innerHTML = html;
    document.getElementById("gen-status").className = "hint gen-status--ok";
  } catch (e) {
    setStatus("gen-status", "Erreur : " + e.message, "err");
  }
});

// ============================================================
// VUE : Mes devis
// ============================================================
async function renderDevisList() {
  const tbody = document.querySelector("#devis-table tbody");
  try {
    const list = await api("/devis");
    if (list.length === 0) {
      tbody.innerHTML = `<tr><td colspan="6" class="hint" style="text-align:center;padding:20px;">Aucun devis généré pour le moment.</td></tr>`;
      return;
    }
    tbody.innerHTML = list.map(d => `
      <tr>
        <td class="code-cell">${d.numero}</td>
        <td>${d.date_emission || "—"}</td>
        <td>${d.client_nom || "—"}</td>
        <td>${d.site_adresse || "—"}</td>
        <td class="code-cell">${d.template_code}</td>
        <td class="actions">
          ${d.has_docx ? `<a class="btn btn--sm" href="${API}/devis/${d.numero}/docx" download>.docx</a>` : ""}
          ${d.has_pdf  ? `<a class="btn btn--sm" href="${API}/devis/${d.numero}/pdf"  download>.pdf</a>`  : ""}
        </td>
      </tr>
    `).join("");
  } catch (e) {
    tbody.innerHTML = `<tr><td colspan="6" class="gen-status--err">Erreur : ${e.message}</td></tr>`;
  }
}

// ============================================================
// VUE : Modèles PDF (back-office)
// ============================================================
async function renderTemplatesList() {
  const tbody = document.querySelector("#templates-table tbody");
  try {
    const list = await api("/templates");
    if (list.length === 0) {
      tbody.innerHTML = `<tr><td colspan="7" class="hint" style="text-align:center;padding:20px;">Aucun modèle enregistré. Importez votre premier modèle ci-dessus.</td></tr>`;
      return;
    }
    tbody.innerHTML = list.map(t => `
      <tr>
        <td class="code-cell">${t.code}</td>
        <td>
          <strong>${t.nom}</strong>
          ${t.is_default ? '<span class="badge badge--default" style="margin-left:6px;">Défaut</span>' : ""}
        </td>
        <td>${t.famille}</td>
        <td>${t.type_intervention || "—"}</td>
        <td>${(t.variables || []).length}</td>
        <td>${t.actif ? '<span class="badge badge--success">Actif</span>' : '<span class="badge badge--neutral">Inactif</span>'}</td>
        <td class="actions">
          <a class="btn btn--sm" href="${API}/templates/${t.code}/download" download>Télécharger</a>
          <a class="btn btn--sm" href="${API}/templates/${t.code}/preview" target="_blank">Aperçu PDF</a>
          ${!t.is_default ? `<button class="btn btn--sm" onclick="setDefault('${t.code}')">Définir défaut</button>` : ""}
          <button class="btn btn--sm btn--danger" onclick="deleteTemplate('${t.code}')">Supprimer</button>
        </td>
      </tr>
    `).join("");
  } catch (e) {
    tbody.innerHTML = `<tr><td colspan="7" class="gen-status--err">Erreur : ${e.message}</td></tr>`;
  }
}

async function setDefault(code) {
  const fd = new FormData();
  fd.append("is_default", "true");
  try {
    await fetch(`${API}/templates/${code}`, { method: "PUT", body: fd });
    renderTemplatesList();
    loadTemplatesForSelect();
  } catch (e) { alert("Erreur : " + e.message); }
}

async function deleteTemplate(code) {
  if (!confirm(`Supprimer définitivement le modèle "${code}" ?`)) return;
  try {
    await fetch(`${API}/templates/${code}`, { method: "DELETE" });
    renderTemplatesList();
    loadTemplatesForSelect();
  } catch (e) { alert("Erreur : " + e.message); }
}

document.getElementById("form-upload-template").addEventListener("submit", async (e) => {
  e.preventDefault();
  const file = document.getElementById("upload-file").files[0];
  if (!file) return setStatus("upload-status", "Sélectionnez un fichier .docx", "err");

  const fd = new FormData();
  fd.append("fichier", file);
  fd.append("code",    document.getElementById("upload-code").value.trim());
  fd.append("nom",     document.getElementById("upload-nom").value.trim());
  fd.append("famille", document.getElementById("upload-famille").value);
  fd.append("type_intervention", document.getElementById("upload-type").value.trim());
  fd.append("annoter", document.getElementById("upload-annoter").checked ? "true" : "false");

  setStatus("upload-status", "Upload en cours…");
  try {
    const res = await fetch(`${API}/templates/upload`, { method: "POST", body: fd });
    if (!res.ok) {
      const t = await res.text();
      throw new Error(t);
    }
    const result = await res.json();
    setStatus("upload-status",
      `Modèle "${result.code}" importé. Variables détectées : ${result.variables_detectees.join(", ") || "aucune"}`,
      "ok");
    renderTemplatesList();
    loadTemplatesForSelect();
    document.getElementById("form-upload-template").reset();
  } catch (e) {
    setStatus("upload-status", "Erreur : " + e.message, "err");
  }
});

// ============================================================
// Init
// ============================================================
addLigne();           // une première ligne de prestation prête à remplir
loadTemplatesForSelect();

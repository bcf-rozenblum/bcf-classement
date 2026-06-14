/*
 * BCF Classement — logique du volet (version DÉMO)
 * ------------------------------------------------------------------
 * La couche "données" (objet DataSource ci-dessous) est volontairement
 * isolée. En version démo, elle renvoie des dossiers et un message
 * fictifs. À l'étape suivante, on remplacera UNIQUEMENT cet objet par
 * des appels à Microsoft Graph — le reste de l'interface ne change pas.
 * ------------------------------------------------------------------
 */

/* ====================== COUCHE DONNÉES (DÉMO) ====================== */
const DataSource = {
  // Liste complète des dossiers (fictive en démo).
  listFolders() {
    return [
      "Clients / Belvaux SA",
      "Clients / Martin & Co",
      "Comptabilité / Factures 2026",
      "Comptabilité / Notes de frais",
      "Fournisseurs",
      "Projets / Migration Outlook",
      "RH / Contrats",
      "Archive 2025",
      "Banque / BNP",
      "TVA / Déclarations"
    ];
  },

  // Dossiers favoris épinglés (fictif en démo).
  listFavorites() {
    return ["Comptabilité / Factures 2026", "Clients / Belvaux SA", "Archive 2025"];
  },

  // Message actuellement sélectionné. En démo : valeur fixe.
  // En version réelle : lecture via Office.context.mailbox.item.
  currentMessage() {
    return {
      subject: "Facture Q2 — Société Belvaux SA",
      from: "comptabilite@belvaux.be"
    };
  },

  // Déplacement d'un message vers un dossier.
  // En démo : simulation. En réel : POST /messages/{id}/move via Graph.
  // Renvoie une promesse résolue avec le dossier d'origine (pour l'undo).
  fileMessage(folderName) {
    return Promise.resolve({ ok: true, previousFolder: "Boîte de réception" });
  },

  // Annulation : redéplace vers le dossier d'origine.
  undoMessage(previousFolder) {
    return Promise.resolve({ ok: true });
  }
};

/* ============== MOTEUR DE SUGGESTION (DÉMO simplifié) ============== */
// En démo : scores fixes plausibles. En réel : calcul à partir de
// l'historique stocké dans la boîte (expéditeur, domaine, mots-clés…).
function computeSuggestions(message, folders) {
  return [
    { name: "Comptabilité / Factures 2026", score: 94 },
    { name: "Clients / Belvaux SA", score: 81 },
    { name: "TVA / Déclarations", score: 63 }
  ];
}

/* ====================== ÉTAT & UTILITAIRES ======================== */
let lastAction = null; // { folder, previousFolder } pour l'undo

function folderIconSvg() {
  return '<svg class="folder-icon" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 7a2 2 0 0 1 2-2h4l2 2h8a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"/></svg>';
}
function chevronSvg() {
  return '<svg class="chevron" width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M9 18l6-6-6-6"/></svg>';
}
function shortName(full) {
  const parts = full.split(" / ");
  return parts[parts.length - 1];
}

/* ====================== RENDU DE L'INTERFACE ====================== */
function makeFolderButton(name, score) {
  const btn = document.createElement("button");
  btn.className = "folder-btn";
  const left = document.createElement("span");
  left.className = "folder-left";
  left.innerHTML = folderIconSvg() + '<span class="folder-name"></span>';
  left.querySelector(".folder-name").textContent = name;
  btn.appendChild(left);
  if (typeof score === "number") {
    const s = document.createElement("span");
    s.className = "score";
    s.textContent = score + "%";
    btn.appendChild(s);
  } else {
    btn.insertAdjacentHTML("beforeend", chevronSvg());
  }
  btn.addEventListener("click", () => fileInto(name));
  return btn;
}

function renderMessage() {
  const m = DataSource.currentMessage();
  document.getElementById("msgSubject").textContent = m.subject;
  document.getElementById("msgFrom").textContent = "de : " + m.from;
}

function renderSuggestions() {
  const box = document.getElementById("suggestions");
  box.innerHTML = "";
  const sugg = computeSuggestions(DataSource.currentMessage(), DataSource.listFolders());
  sugg.forEach(s => box.appendChild(makeFolderButton(s.name, s.score)));
}

function renderFavorites() {
  const box = document.getElementById("favs");
  box.innerHTML = "";
  DataSource.listFavorites().forEach(f => {
    const chip = document.createElement("button");
    chip.className = "chip";
    chip.textContent = shortName(f);
    chip.title = f;
    chip.addEventListener("click", () => fileInto(f));
    box.appendChild(chip);
  });
}

function setupSearch() {
  const input = document.getElementById("search");
  const results = document.getElementById("results");
  const folders = DataSource.listFolders();

  input.addEventListener("input", () => {
    const q = input.value.trim().toLowerCase();
    results.innerHTML = "";
    if (!q) return;
    const matches = folders.filter(f => f.toLowerCase().includes(q)).slice(0, 5);
    if (matches.length === 0) {
      results.innerHTML = '<p class="empty-hint">Aucun dossier ne correspond. (En version réelle, vous pourrez en créer un.)</p>';
      return;
    }
    matches.forEach(f => results.appendChild(makeFolderButton(f, null)));
  });

  input.addEventListener("keydown", (e) => {
    if (e.key === "Enter") {
      const first = results.querySelector(".folder-btn");
      if (first) first.click();
    }
  });
}

/* ====================== ACTIONS ====================== */
function fileInto(folderName) {
  DataSource.fileMessage(folderName).then(res => {
    if (res && res.ok) {
      lastAction = { folder: folderName, previousFolder: res.previousFolder };
      showStatus("Classé dans « " + folderName + " »");
    }
  });
}

function showStatus(text) {
  document.getElementById("statusText").textContent = text;
  document.getElementById("status").classList.add("show");
}
function hideStatus() {
  document.getElementById("status").classList.remove("show");
}

function setupUndo() {
  document.getElementById("undoBtn").addEventListener("click", () => {
    if (!lastAction) return;
    DataSource.undoMessage(lastAction.previousFolder).then(res => {
      if (res && res.ok) {
        showStatus("Classement annulé — message revenu dans « " + lastAction.previousFolder + " »");
        lastAction = null;
        setTimeout(hideStatus, 2200);
      }
    });
  });
}

/* ====================== DÉMARRAGE ====================== */
function init() {
  renderMessage();
  renderSuggestions();
  renderFavorites();
  setupSearch();
  setupUndo();
}

// Office.onReady garantit que l'hôte Outlook est prêt.
// En démo, on initialise même si Office n'est pas présent (test navigateur).
if (typeof Office !== "undefined") {
  Office.onReady(() => init());
} else {
  document.addEventListener("DOMContentLoaded", init);
}

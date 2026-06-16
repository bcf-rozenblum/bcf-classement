/*
 * BCF Classement — Volet de classement (VERSION RÉELLE)
 * ====================================================================
 * Se connecte à la boîte via MSAL/NAA, lit le message sélectionné,
 * charge l'arborescence des dossiers, suggère un classement, déplace
 * réellement le message, permet de créer des dossiers, et d'annuler.
 *
 * L'apprentissage (favoris + historique de classement) est stocké dans
 * la boîte (roamingSettings d'Office), donc synchronisé entre postes.
 * ====================================================================
 */

// ====== Configuration Entra ID ======
const CLIENT_ID = "ea5e7d18-727f-486d-a939-56a2332c3420";
const TENANT_ID = "fbbe2873-c64f-491c-aa80-7453e77a18c7";
const GRAPH_SCOPES = ["Mail.ReadWrite", "User.Read"];
const GRAPH = "https://graph.microsoft.com/v1.0";

// ====== État global ======
let pca = null;
let cachedToken = null;
let allFolders = [];        // [{id, name, path}]
let currentMessage = null;  // {restId, subject, fromAddress, fromName, conversationId, parentFolderId}
let lastAction = null;      // conservé pour compat, non utilisé pour l'affichage
let sessionMoves = [];      // pile de session : [{id, messageRestId, fromFolderId, subject, toFolderName, undone}]
let learnModel = { bySender: {}, favorites: [], history: [] };

// ====== Démarrage ======
Office.onReady(() => {
  bindUi();
  // Recharger automatiquement quand l'utilisateur change de message (volet épinglé)
  try {
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, onItemChanged);
  } catch (_) {}
  // Lancement automatique : on tente la connexion et le chargement dès l'ouverture
  start();
});

// Appelé quand le message sélectionné change (si le volet reste ouvert/épinglé)
async function onItemChanged() {
  try {
    currentMessage = await readSelectedMessage();
    if (currentMessage) {
      try {
        const m = await graph("GET", "/me/messages/" + currentMessage.restId + "?$select=parentFolderId");
        currentMessage.parentFolderId = m.parentFolderId || null;
      } catch (_) {}
    }
    renderMessage();
    renderSuggestions();
    const results = document.getElementById("results");
    const search = document.getElementById("search");
    if (results) results.innerHTML = "";
    if (search) search.value = "";
  } catch (e) {
    setStatus("Impossible de recharger le message : " + (e.message || e), "err");
  }
}

function bindUi() {
  const search = document.getElementById("search");
  if (search) {
    search.addEventListener("input", onSearchInput);
    search.addEventListener("keydown", (e) => {
      if (e.key === "Enter") {
        const first = document.querySelector("#results .folder-btn");
        if (first) first.click();
      }
    });
  }
  const retry = document.getElementById("retryBtn");
  if (retry) retry.addEventListener("click", start);
  const createBtn = document.getElementById("createBtn");
  if (createBtn) createBtn.addEventListener("click", onCreateFolder);
}

function setStatus(text, kind) {
  const el = document.getElementById("status");
  if (!el) return;
  el.textContent = text;
  el.className = "status-line " + (kind === "ok" ? "status-ok" : kind === "err" ? "status-err" : "status-info");
  el.style.display = text ? "block" : "none";
}

function showBanner() { /* obsolète : remplacé par la liste des classements récents */ }
function hideBanner() { /* obsolète */ }

// ====== MSAL / NAA ======
async function ensureMsal() {
  if (pca) return pca;
  const msalConfig = {
    auth: {
      clientId: CLIENT_ID,
      authority: "https://login.microsoftonline.com/" + TENANT_ID,
      supportsNestedAppAuth: true
    },
    cache: { cacheLocation: "localStorage" }
  };
  if (msal.createNestablePublicClientApplication) {
    pca = await msal.createNestablePublicClientApplication(msalConfig);
  } else {
    pca = new msal.PublicClientApplication(msalConfig);
    if (pca.initialize) await pca.initialize();
  }
  return pca;
}

async function getToken() {
  if (cachedToken) return cachedToken;
  await ensureMsal();
  const request = { scopes: GRAPH_SCOPES };
  try {
    const r = await pca.acquireTokenSilent(request);
    cachedToken = r.accessToken;
  } catch (e) {
    const r = await pca.acquireTokenPopup(request);
    cachedToken = r.accessToken;
  }
  return cachedToken;
}

// ====== Appels Graph ======
async function graph(method, path, body) {
  const token = await getToken();
  const opts = { method, headers: { Authorization: "Bearer " + token } };
  if (body) { opts.headers["Content-Type"] = "application/json"; opts.body = JSON.stringify(body); }
  const resp = await fetch(GRAPH + path, opts);
  if (!resp.ok) {
    const txt = await resp.text();
    throw new Error("Graph " + resp.status + " : " + txt.slice(0, 300));
  }
  if (resp.status === 204) return null;
  return resp.json();
}

// ====== Lecture du message sélectionné ======
function readSelectedMessage() {
  return new Promise((resolve) => {
    const item = Office.context.mailbox.item;
    if (!item || !item.itemId) { resolve(null); return; }
    const restId = Office.context.mailbox.convertToRestId(
      item.itemId, Office.MailboxEnums.RestVersion.v2_0
    );
    const subject = item.subject || "(sans objet)";
    let fromAddress = "", fromName = "";
    if (item.from) { fromAddress = item.from.emailAddress || ""; fromName = item.from.displayName || ""; }
    const conversationId = item.conversationId || null;
    resolve({ restId, subject, fromAddress, fromName, conversationId, parentFolderId: null });
  });
}

// ====== Chargement de l'arborescence des dossiers ======
// On parcourt récursivement avec $expand=childFolders pour limiter les appels.
async function loadAllFolders() {
  const result = [];
  async function walk(parentPath, fetchUrl) {
    const data = await graph("GET", fetchUrl);
    const items = (data.value || []);
    for (const f of items) {
      if (f.isHidden) continue;
      const path = parentPath ? (parentPath + " / " + f.displayName) : f.displayName;
      result.push({ id: f.id, name: f.displayName, path });
      if (f.childFolderCount && f.childFolderCount > 0) {
        await walk(path, "/me/mailFolders/" + f.id + "/childFolders?$top=200&$select=id,displayName,childFolderCount,isHidden");
      }
    }
  }
  await walk("", "/me/mailFolders?$top=200&$select=id,displayName,childFolderCount,isHidden");
  result.sort((a, b) => a.path.localeCompare(b.path, "fr"));
  return result;
}

// ====== Apprentissage (roamingSettings) ======
function loadLearnModel() {
  try {
    const raw = Office.context.roamingSettings.get("bcfLearnModel");
    if (raw && typeof raw === "object") learnModel = Object.assign(learnModel, raw);
  } catch (_) {}
}
function saveLearnModel() {
  try {
    Office.context.roamingSettings.set("bcfLearnModel", learnModel);
    Office.context.roamingSettings.saveAsync(() => {});
  } catch (_) {}
}

// ====== Moteur de suggestion (règles pondérées par historique) ======
function computeSuggestions(message) {
  const scores = {}; // folderName -> score
  const add = (name, pts) => { if (name) scores[name] = (scores[name] || 0) + pts; };

  // 1) Historique par expéditeur exact
  const sender = (message.fromAddress || "").toLowerCase();
  if (sender && learnModel.bySender[sender]) {
    for (const [folder, count] of Object.entries(learnModel.bySender[sender])) add(folder, count * 10);
  }
  // 2) Historique par domaine
  const domain = sender.split("@")[1] || "";
  if (domain) {
    for (const [s, folders] of Object.entries(learnModel.bySender)) {
      if (s.endsWith("@" + domain)) {
        for (const [folder, count] of Object.entries(folders)) add(folder, count * 3);
      }
    }
  }
  // 3) Correspondance de mots de l'objet avec un nom de dossier
  const subjectWords = (message.subject || "").toLowerCase().split(/[^a-zà-ÿ0-9]+/).filter(w => w.length >= 4);
  for (const f of allFolders) {
    const fname = f.name.toLowerCase();
    for (const w of subjectWords) if (fname.includes(w)) add(f.path, 2);
  }

  // Normaliser en pourcentages indicatifs
  const entries = Object.entries(scores).sort((a, b) => b[1] - a[1]).slice(0, 4);
  if (entries.length === 0) return [];
  const max = entries[0][1];
  return entries.map(([name, sc]) => ({
    name,
    score: Math.max(35, Math.min(98, Math.round((sc / max) * 95)))
  }));
}

// ====== Rendu ======
function folderIcon() {
  return '<svg class="folder-icon" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 7a2 2 0 0 1 2-2h4l2 2h8a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"/></svg>';
}
function chevron() {
  return '<svg class="chevron" width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M9 18l6-6-6-6"/></svg>';
}
function shortName(path) { const p = path.split(" / "); return p[p.length - 1]; }

function makeFolderButton(path, score) {
  const btn = document.createElement("button");
  btn.className = "folder-btn";
  const left = document.createElement("span");
  left.className = "folder-left";
  left.innerHTML = folderIcon() + '<span class="folder-name"></span>';
  left.querySelector(".folder-name").textContent = path;
  btn.appendChild(left);
  if (typeof score === "number") {
    const s = document.createElement("span"); s.className = "score"; s.textContent = score + "%"; btn.appendChild(s);
  } else {
    btn.insertAdjacentHTML("beforeend", chevron());
  }
  btn.addEventListener("click", () => fileInto(path));
  return btn;
}

function renderMessage() {
  const subj = document.getElementById("msgSubject");
  const from = document.getElementById("msgFrom");
  if (currentMessage) {
    subj.textContent = currentMessage.subject;
    from.textContent = "de : " + (currentMessage.fromName || currentMessage.fromAddress || "—");
  } else {
    subj.textContent = "Aucun message sélectionné";
    from.textContent = "Sélectionnez un message pour le classer.";
  }
}

function renderSuggestions() {
  const box = document.getElementById("suggestions");
  box.innerHTML = "";
  if (!currentMessage) return;
  const sugg = computeSuggestions(currentMessage);
  if (sugg.length === 0) {
    box.innerHTML = '<p class="empty-hint">Pas encore de suggestion (l\'apprentissage se construit à mesure que vous classez). Utilisez la recherche ci-dessous.</p>';
    return;
  }
  sugg.forEach(s => box.appendChild(makeFolderButton(s.name, s.score)));
}

function renderFavorites() {
  const box = document.getElementById("favs");
  if (!box) return;
  box.innerHTML = "";
  const favs = (learnModel.favorites || []).slice(0, 8);
  if (favs.length === 0) { box.innerHTML = '<p class="empty-hint">Vos dossiers les plus utilisés apparaîtront ici.</p>'; return; }
  favs.forEach(path => {
    const chip = document.createElement("button");
    chip.className = "chip"; chip.textContent = shortName(path); chip.title = path;
    chip.addEventListener("click", () => fileInto(path));
    box.appendChild(chip);
  });
}

function onSearchInput() {
  const q = document.getElementById("search").value.trim().toLowerCase();
  const results = document.getElementById("results");
  results.innerHTML = "";
  if (!q) return;
  const matches = allFolders.filter(f => f.path.toLowerCase().includes(q)).slice(0, 8);
  if (matches.length === 0) {
    results.innerHTML = '<p class="empty-hint">Aucun dossier ne correspond. Vous pouvez en créer un ci-dessous.</p>';
    return;
  }
  matches.forEach(f => results.appendChild(makeFolderButton(f.path, null)));
}

// ====== Action : classer (déplacement réel) ======
async function fileInto(folderPath) {
  // Toujours relire le message réellement sélectionné juste avant d'agir,
  // pour éviter d'utiliser un identifiant périmé (volet épinglé, changement de message).
  const fresh = await readSelectedMessage();
  if (fresh) {
    if (!currentMessage || fresh.restId !== currentMessage.restId) {
      currentMessage = fresh;
      try {
        const m = await graph("GET", "/me/messages/" + currentMessage.restId + "?$select=parentFolderId");
        currentMessage.parentFolderId = m.parentFolderId || null;
      } catch (_) {}
      renderMessage();
    }
  }
  if (!currentMessage) { setStatus("Aucun message sélectionné.", "err"); return; }
  const target = allFolders.find(f => f.path === folderPath);
  if (!target) { setStatus("Dossier introuvable : " + folderPath, "err"); return; }
  setStatus("Classement en cours…", "info");
  try {
    // Mémoriser l'origine pour l'undo
    const fromFolderId = currentMessage.parentFolderId;
    await graph("POST", "/me/messages/" + currentMessage.restId + "/move", { destinationId: target.id });

    // Apprentissage : incrémente le compteur expéditeur -> dossier
    const sender = (currentMessage.fromAddress || "").toLowerCase();
    if (sender) {
      learnModel.bySender[sender] = learnModel.bySender[sender] || {};
      learnModel.bySender[sender][target.path] = (learnModel.bySender[sender][target.path] || 0) + 1;
    }
    // Favoris = dossiers les plus utilisés (recalcul simple)
    bumpFavorite(target.path);
    // Historique
    learnModel.history.unshift({ when: Date.now(), subject: currentMessage.subject, to: target.path });
    learnModel.history = learnModel.history.slice(0, 200);
    saveLearnModel();

    lastAction = { messageRestId: currentMessage.restId, fromFolderId, toFolderName: target.path };
    // Empiler dans la pile de session (annulation individuelle possible)
    sessionMoves.unshift({
      id: "m" + Date.now() + Math.random().toString(36).slice(2, 6),
      messageRestId: currentMessage.restId,
      fromFolderId: fromFolderId,
      subject: currentMessage.subject,
      toFolderName: target.path,
      undone: false
    });
    sessionMoves = sessionMoves.slice(0, 15);
    renderSessionMoves();
    setStatus("Classé dans « " + target.path + " ».", "ok");
    setTimeout(() => setStatus("", "info"), 1800);
  } catch (e) {
    setStatus("Échec du classement : " + (e.message || e), "err");
  }
}

function bumpFavorite(path) {
  const counts = {};
  for (const h of learnModel.history) counts[h.to] = (counts[h.to] || 0) + 1;
  counts[path] = (counts[path] || 0) + 1;
  learnModel.favorites = Object.entries(counts).sort((a, b) => b[1] - a[1]).map(e => e[0]).slice(0, 8);
}

// ====== Annulation d'une entrée précise de la pile ======
async function undoMove(moveId) {
  const mv = sessionMoves.find(m => m.id === moveId);
  if (!mv || mv.undone) return;
  setStatus("Annulation…", "info");
  try {
    const dest = mv.fromFolderId || "inbox";
    await graph("POST", "/me/messages/" + mv.messageRestId + "/move", { destinationId: dest });
    mv.undone = true;
    renderSessionMoves();
    setStatus("Classement annulé.", "ok");
    setTimeout(() => setStatus("", "info"), 1600);
  } catch (e) {
    setStatus("Impossible d'annuler ce classement (le message a peut-être été déplacé entre-temps) : " + (e.message || e), "err");
  }
}

// ====== Rendu de la liste des classements récents (session) ======
function renderSessionMoves() {
  const wrap = document.getElementById("sessionWrap");
  const list = document.getElementById("sessionList");
  if (!wrap || !list) return;
  const active = sessionMoves;
  if (active.length === 0) { wrap.style.display = "none"; return; }
  wrap.style.display = "block";
  list.innerHTML = "";
  active.forEach(mv => {
    const row = document.createElement("div");
    row.className = "move-row" + (mv.undone ? " move-undone" : "");
    const info = document.createElement("div");
    info.className = "move-info";
    const subj = document.createElement("div");
    subj.className = "move-subject";
    subj.textContent = mv.subject || "(sans objet)";
    const dest = document.createElement("div");
    dest.className = "move-dest";
    dest.textContent = (mv.undone ? "annulé · " : "→ ") + mv.toFolderName;
    info.appendChild(subj); info.appendChild(dest);
    row.appendChild(info);
    if (!mv.undone) {
      const btn = document.createElement("button");
      btn.className = "move-undo";
      btn.textContent = "Annuler";
      btn.addEventListener("click", () => undoMove(mv.id));
      row.appendChild(btn);
    }
    list.appendChild(row);
  });
}

// ====== Création de dossier ======
async function onCreateFolder() {
  const input = document.getElementById("newFolderName");
  const name = (input.value || "").trim();
  if (!name) { setStatus("Indiquez un nom de dossier.", "info"); return; }
  setStatus("Création du dossier…", "info");
  try {
    // Crée à la racine de la boîte (niveau supérieur)
    const created = await graph("POST", "/me/mailFolders", { displayName: name });
    allFolders.push({ id: created.id, name: created.displayName, path: created.displayName });
    allFolders.sort((a, b) => a.path.localeCompare(b.path, "fr"));
    input.value = "";
    setStatus("Dossier « " + name + " » créé. Recherchez-le ci-dessus pour y classer le message.", "ok");
  } catch (e) {
    setStatus("Échec de la création : " + (e.message || e), "err");
  }
}

// ====== Démarrage principal ======
async function start() {
  const retry = document.getElementById("retryBtn");
  if (retry) retry.style.display = "none";
  setStatus("Connexion…", "info");
  try {
    await getToken();
    loadLearnModel();

    setStatus("Lecture du message…", "info");
    currentMessage = await readSelectedMessage();

    setStatus("Chargement de vos dossiers…", "info");
    allFolders = await loadAllFolders();

    // Déterminer le dossier parent du message (pour un undo précis)
    if (currentMessage) {
      try {
        const m = await graph("GET", "/me/messages/" + currentMessage.restId + "?$select=parentFolderId");
        currentMessage.parentFolderId = m.parentFolderId || null;
      } catch (_) {}
    }

    renderMessage();
    renderSuggestions();
    renderFavorites();
    renderSessionMoves();
    document.getElementById("mainUi").style.display = "block";
    setStatus(allFolders.length + " dossiers chargés.", "ok");
    setTimeout(() => setStatus("", "info"), 1500);
  } catch (e) {
    setStatus("Problème : " + (e.message || e), "err");
    if (retry) retry.style.display = "inline-flex";
  }
}

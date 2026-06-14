/*
 * BCF Classement — Test de connexion (version minimale)
 * --------------------------------------------------------------
 * Objectif : se connecter au compte Microsoft de l'utilisateur via
 * MSAL.js (authentification d'application imbriquée / NAA), puis lire
 * et afficher ses vrais dossiers de messagerie via Microsoft Graph.
 *
 * Aucune action de classement ici : c'est un test d'accès.
 * --------------------------------------------------------------
 */

// ====== Configuration (vos identifiants Entra ID) ======
const CLIENT_ID = "ea5e7d18-727f-486d-a939-56a2332c3420";
const TENANT_ID = "fbbe2873-c64f-491c-aa80-7453e77a18c7";
const GRAPH_SCOPES = ["Mail.ReadWrite", "User.Read"];

let pca = null; // instance MSAL

// ====== Démarrage ======
Office.onReady(() => {
  document.getElementById("connectBtn").addEventListener("click", connectAndList);
});

function setStatus(text, kind) {
  const el = document.getElementById("status");
  el.textContent = text;
  el.className = "status-line " + (kind === "ok" ? "status-ok" : kind === "err" ? "status-err" : "status-info");
}

// ====== Initialisation de MSAL avec NAA ======
async function ensureMsal() {
  if (pca) return pca;
  const msalConfig = {
    auth: {
      clientId: CLIENT_ID,
      authority: "https://login.microsoftonline.com/" + TENANT_ID,
      // supportsNestedAppAuth active le mode NAA dans les hôtes Office compatibles
      supportsNestedAppAuth: true
    },
    cache: { cacheLocation: "localStorage" }
  };
  // createNestablePublicClientApplication = chemin recommandé pour Office add-ins
  if (msal.createNestablePublicClientApplication) {
    pca = await msal.createNestablePublicClientApplication(msalConfig);
  } else {
    pca = new msal.PublicClientApplication(msalConfig);
    if (pca.initialize) await pca.initialize();
  }
  return pca;
}

// ====== Obtenir un jeton d'accès (silencieux, sinon popup) ======
async function getToken() {
  await ensureMsal();
  const request = { scopes: GRAPH_SCOPES };
  try {
    const result = await pca.acquireTokenSilent(request);
    return result.accessToken;
  } catch (e) {
    // Pas de session silencieuse : on demande une connexion interactive
    const result = await pca.acquireTokenPopup(request);
    return result.accessToken;
  }
}

// ====== Appel Microsoft Graph ======
async function graphGet(path, token) {
  const resp = await fetch("https://graph.microsoft.com/v1.0" + path, {
    headers: { Authorization: "Bearer " + token }
  });
  if (!resp.ok) {
    const body = await resp.text();
    throw new Error("Graph " + resp.status + " : " + body.slice(0, 300));
  }
  return resp.json();
}

// ====== Action principale ======
async function connectAndList() {
  const btn = document.getElementById("connectBtn");
  btn.disabled = true;
  setStatus("Connexion en cours…", "info");

  try {
    const token = await getToken();
    setStatus("Connecté. Lecture de vos dossiers…", "info");

    // Identité (pour confirmer qui est connecté)
    try {
      const me = await graphGet("/me", token);
      const who = me.displayName ? (me.displayName + " — " + (me.mail || me.userPrincipalName || "")) : "";
      document.getElementById("user").textContent = who ? ("Connecté en tant que : " + who) : "";
    } catch (_) { /* non bloquant */ }

    // Dossiers de messagerie (niveau racine), triés, avec nombre d'éléments
    const data = await graphGet("/me/mailFolders?$top=100&$select=displayName,totalItemCount,childFolderCount", token);
    const folders = (data.value || []).slice().sort((a, b) => a.displayName.localeCompare(b.displayName, "fr"));

    const wrap = document.getElementById("foldersWrap");
    const list = document.getElementById("folders");
    list.innerHTML = "";
    folders.forEach(f => {
      const row = document.createElement("div");
      row.className = "folder-item";
      row.innerHTML =
        '<svg class="folder-icon" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 7a2 2 0 0 1 2-2h4l2 2h8a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"/></svg>' +
        '<span class="folder-name"></span>' +
        '<span class="folder-count"></span>';
      row.querySelector(".folder-name").textContent = f.displayName;
      const bits = [];
      if (typeof f.totalItemCount === "number") bits.push(f.totalItemCount + " élt");
      if (f.childFolderCount) bits.push(f.childFolderCount + " sous-doss.");
      row.querySelector(".folder-count").textContent = bits.join(" · ");
      list.appendChild(row);
    });
    wrap.style.display = "block";
    setStatus("Connexion réussie : " + folders.length + " dossier(s) racine trouvé(s).", "ok");
  } catch (e) {
    setStatus("Échec : " + (e && e.message ? e.message : String(e)), "err");
  } finally {
    btn.disabled = false;
  }
}

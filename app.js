// Myra Dashboard - App-Logik
// Ablauf: Login via MSAL (Entra ID) -> Access Token für Graph -> SharePoint-Site
// auflösen -> Listen abrufen -> in die drei Dashboard-Bereiche rendern.

const msalInstance = new msal.PublicClientApplication(MSAL_CONFIG);
let account = null;

const loadState = document.getElementById("load-state");
const userInfo = document.getElementById("user-info");
const loginBtn = document.getElementById("login-btn");

function setStatus(text) {
  loadState.textContent = text;
}

async function init() {
  await msalInstance.initialize();

  const response = await msalInstance.handleRedirectPromise().catch(err => {
    setStatus("Login-Fehler: " + err.message);
    return null;
  });

  if (response && response.account) {
    account = response.account;
  } else {
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length > 0) account = accounts[0];
  }

  if (account) {
    userInfo.textContent = "Angemeldet als " + account.username;
    loginBtn.textContent = "Abmelden";
    loginBtn.onclick = () => msalInstance.logoutRedirect();
    await loadDashboard();
  } else {
    userInfo.textContent = "Nicht angemeldet";
    loginBtn.textContent = "Anmelden";
    loginBtn.onclick = () => msalInstance.loginRedirect({ scopes: GRAPH_SCOPES });
    setStatus("Bitte anmelden, um die Daten zu laden.");
  }
}

async function getAccessToken() {
  const request = { scopes: GRAPH_SCOPES, account };
  try {
    const result = await msalInstance.acquireTokenSilent(request);
    return result.accessToken;
  } catch (e) {
    const result = await msalInstance.acquireTokenRedirect(request);
    return null; // redirect wird ausgeführt, Funktion endet hier faktisch
  }
}

async function graphGet(path) {
  const token = await getAccessToken();
  const res = await fetch("https://graph.microsoft.com/v1.0" + path, {
    headers: { Authorization: "Bearer " + token }
  });
  if (!res.ok) {
    throw new Error("Graph-Anfrage fehlgeschlagen (" + res.status + "): " + path);
  }
  return res.json();
}

async function getSiteId() {
  // Root-Site: kein Unterpfad -> einfacher Hostname-Aufruf.
  // Unterseite (falls sitePath gesetzt ist): Hostname + ":/" + Pfad
  const path = SHAREPOINT_CONFIG.sitePath
    ? "/sites/" + SHAREPOINT_CONFIG.siteHostname + ":/" + SHAREPOINT_CONFIG.sitePath
    : "/sites/" + SHAREPOINT_CONFIG.siteHostname;
  const data = await graphGet(path);
  return data.id;
}

async function getListItems(siteId, listName) {
  const data = await graphGet(
    "/sites/" + siteId + "/lists/" + encodeURIComponent(listName) + "/items?expand=fields"
  );
  return data.value.map(item => item.fields);
}

// --- Rendering ---

function renderOrgChart(items) {
  const container = document.getElementById("orgchart");
  container.innerHTML = "";

  if (!items.length) {
    container.innerHTML = '<span class="placeholder-note">Noch keine Org-Chart-Daten hinterlegt</span>';
    return;
  }

  const byManager = {};
  items.forEach(p => {
    const mgr = p.ReportsTo || "__root__";
    if (!byManager[mgr]) byManager[mgr] = [];
    byManager[mgr].push(p);
  });

  function renderLevel(managerName) {
    const people = byManager[managerName];
    if (!people) return null;
    const level = document.createElement("div");
    level.className = "org-level";
    people.forEach(p => {
      const wrap = document.createElement("div");
      const node = document.createElement("div");
      node.className = "org-node";
      node.innerHTML = '<div class="name">' + p.Name + '</div><div class="role">' + (p.Rolle || "") + "</div>";
      wrap.appendChild(node);
      const children = renderLevel(p.Name);
      if (children) {
        const childWrap = document.createElement("div");
        childWrap.className = "org-children";
        childWrap.appendChild(children);
        wrap.appendChild(childWrap);
      }
      level.appendChild(wrap);
    });
    return level;
  }

  const root = renderLevel("__root__");
  if (root) container.appendChild(root);
  else container.innerHTML = '<span class="placeholder-note">Keine Person ohne ReportsTo gefunden (Root fehlt)</span>';
}

function renderSummaries(items) {
  const hr = items.find(i => (i.Kategorie || "").toLowerCase().startsWith("hr"));
  const myra = items.find(i => (i.Kategorie || "").toLowerCase().startsWith("myra"));

  if (hr) {
    document.getElementById("hr-summary").textContent = hr.Text || "–";
    document.getElementById("hr-updated").textContent = hr.LetztesUpdate
      ? "Stand: " + new Date(hr.LetztesUpdate).toLocaleDateString("de-DE")
      : "";
  }
  if (myra) {
    document.getElementById("myra-summary").textContent = myra.Text || "–";
    document.getElementById("myra-updated").textContent = myra.LetztesUpdate
      ? "Stand: " + new Date(myra.LetztesUpdate).toLocaleDateString("de-DE")
      : "";
  }
}

function renderKPIs(items) {
  const tbody = document.querySelector("#kpi-table tbody");
  if (!items.length) {
    tbody.innerHTML = '<tr><td colspan="5" class="placeholder-note">Noch keine KPI-Werte — folgt, sobald Systemzugriff besteht</td></tr>';
    return;
  }
  tbody.innerHTML = items.map(k => {
    const statusClass = (k.Status || "Offen").toLowerCase() === "erreicht" ? "status-erreicht" : "status-offen";
    return "<tr>" +
      "<td>" + (k.Name || "") + "</td>" +
      "<td>" + (k.Definition || "") + "</td>" +
      "<td>" + (k.Zielwert ?? "–") + "</td>" +
      "<td>" + (k.AktuellerWert ?? "–") + "</td>" +
      '<td><span class="status-pill ' + statusClass + '">' + (k.Status || "Offen") + "</span></td>" +
      "</tr>";
  }).join("");
}

// --- Hauptablauf ---

async function loadDashboard() {
  setStatus("Lade Daten aus SharePoint …");
  try {
    const siteId = await getSiteId();

    const [orgChartItems, summaryItems, kpiItems] = await Promise.all([
      getListItems(siteId, SHAREPOINT_CONFIG.lists.orgChart).catch(() => []),
      getListItems(siteId, SHAREPOINT_CONFIG.lists.summaries).catch(() => []),
      getListItems(siteId, SHAREPOINT_CONFIG.lists.kpis).catch(() => [])
    ]);

    renderOrgChart(orgChartItems);
    renderSummaries(summaryItems);
    renderKPIs(kpiItems);

    setStatus("");
  } catch (err) {
    setStatus("Fehler beim Laden: " + err.message);
  }
}

init();

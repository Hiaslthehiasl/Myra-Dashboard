// Konfiguration für Myra Dashboard
// Client ID + Tenant ID sind öffentliche Identifikatoren, keine Geheimnisse.
// Ein Client Secret gehört NIE in Frontend-Code - wird hier auch nicht gebraucht,
// da der Login per MSAL.js im Browser läuft (Authorization Code Flow mit PKCE,
// kein Secret nötig für eine Single-Page-App).

const MSAL_CONFIG = {
  auth: {
    clientId: "0dcfb7a9-4f11-4abb-991c-be434d049720",
    authority: "https://login.microsoftonline.com/5756f6c7-6b68-4cea-90c0-7f12786882ad",
    redirectUri: window.location.origin + window.location.pathname
  },
  cache: {
    cacheLocation: "sessionStorage"
  }
};

// Berechtigungen, die die App beim Login anfragt
const GRAPH_SCOPES = ["User.Read", "Sites.Read.All"];

// SharePoint-Zielstruktur — Werte hier eintragen, sobald die Listen angelegt sind
const SHAREPOINT_CONFIG = {
  // Hostname + Site-Pfad deiner SharePoint-Seite, z.B. "hofmuth.sharepoint.com" / "sites/MyraDashboard"
  siteHostname: "PLACEHOLDER_HOSTNAME",
  sitePath: "PLACEHOLDER_SITE_PATH",

  // Listennamen exakt wie in SharePoint angelegt
  lists: {
    orgChart: "OrgChart",
    summaries: "Summaries",
    kpis: "KPIs"
  }
};

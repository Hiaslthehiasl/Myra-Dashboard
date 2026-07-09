// Konfiguration für Myra Dashboard
// Diese Werte NICHT mit echten Secrets füllen (Client Secret gehört niemals in Frontend-Code).
// Nur Client ID + Tenant ID sind hier korrekt - das sind öffentliche Identifikatoren, keine Geheimnisse.

const MSAL_CONFIG = {
  auth: {
    clientId: "PLACEHOLDER_APP_CLIENT_ID",       // Application (client) ID aus der Entra-App-Registrierung
    authority: "https://login.microsoftonline.com/PLACEHOLDER_TENANT_ID",
    redirectUri: window.location.origin
  }
};

// SharePoint-Zielliste (wird gesetzt, sobald Listenstruktur steht)
const SHAREPOINT_CONFIG = {
  siteId: "PLACEHOLDER_SITE_ID",
  lists: {
    kpis: "PLACEHOLDER_LIST_ID_KPIS",
    insights: "PLACEHOLDER_LIST_ID_INSIGHTS",
    questions: "PLACEHOLDER_LIST_ID_QUESTIONS"
  }
};

// Konfiguration fuer die SharePoint-Anbindung. Diese Werte sind keine Geheimnisse
// (Mandanten-ID und Client-ID einer Single-Page-App duerfen im Repo stehen).
window.SP_CONFIG = {
  tenantId: "07cf9a53-cac1-46a9-9f4a-e9075df36c17",          // Entra > Uebersicht > Verzeichnis-ID (Mandanten-ID)
  clientId: "5460a931-94d3-4316-a01e-5827ae1139b4",         // Entra > App-Registrierung > Anwendungs-ID (Client-ID)
  redirectUri: "",                         // leer = aktuelle Seite (muss in Entra als SPA-Redirect-URI stehen)
  hostname: "claimmgmbh.sharepoint.com",
  sitePath: "/sites/claim.mGmbH",
  folderPath: "General/03_Projekte/EXPO_REAL_2026_Nachbereitung",   // Ordner in "Freigegebene Dokumente"
  fileName: "nachbereitung-daten.json",
  pollSeconds: 20                          // Abgleich mit Aenderungen der anderen
};

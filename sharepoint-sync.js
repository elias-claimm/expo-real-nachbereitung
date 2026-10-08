/*
 * sharepoint-sync.js
 * Synchronisiert die Eintraege der EXPO-REAL-Nachbereitung mit einer JSON-Datei in SharePoint.
 * Anmeldung: Microsoft (MSAL, delegierte Berechtigungen). Speicherung: Microsoft Graph.
 * Konfliktbehandlung: pro Eintrag gewinnt der Stand mit dem juengeren "updatedAt".
 */
(function () {
  "use strict";

  const cfg = window.SP_CONFIG || {};
  const GRAPH = "https://graph.microsoft.com/v1.0";
  const SCOPES = ["Files.ReadWrite.All", "Sites.Read.All"];

  let app = null, account = null, siteId = null;
  let api = null;               // { getRecords, applyRecords, onStatus }
  let etag = null;              // eTag der Datei beim letzten Abgleich
  let saving = false, dirty = false, saveTimer = null, pollTimer = null;
  const SP = { active: false };
  window.SPSync = SP;

  const say = (msg, error) => { if (api && api.onStatus) api.onStatus(msg, !!error); };
  const clock = () => new Date().toLocaleTimeString("de-DE", { hour: "2-digit", minute: "2-digit" });
  const encPath = p => p.split("/").filter(Boolean).map(encodeURIComponent).join("/");
  const filePath = () => encPath(cfg.folderPath) + "/" + encodeURIComponent(cfg.fileName);

  function configured() {
    return cfg.tenantId && cfg.clientId && !/DEINE/i.test(cfg.tenantId + cfg.clientId)
      && cfg.hostname && cfg.sitePath && cfg.folderPath && cfg.fileName;
  }

  async function getToken() {
    try {
      const r = await app.acquireTokenSilent({ scopes: SCOPES, account });
      return r.accessToken;
    } catch (e) {
      await app.acquireTokenRedirect({ scopes: SCOPES, account });
      throw new Error("Weiterleitung zur Anmeldung");
    }
  }

  async function graph(path, opts) {
    const token = await getToken();
    const o = opts || {};
    return fetch(GRAPH + path, { ...o, headers: { Authorization: "Bearer " + token, ...(o.headers || {}) } });
  }

  async function resolveSite() {
    if (siteId) return siteId;
    const res = await graph("/sites/" + cfg.hostname + ":" + cfg.sitePath);
    if (!res.ok) throw new Error("SharePoint-Site nicht gefunden (" + res.status + ")");
    siteId = (await res.json()).id;
    return siteId;
  }

  async function getMeta() {
    const id = await resolveSite();
    const res = await graph("/sites/" + id + "/drive/root:/" + filePath());
    if (res.status === 404) return null;
    if (!res.ok) throw new Error("Datei nicht lesbar (" + res.status + ")");
    return res.json();
  }

  async function download() {
    const id = await resolveSite();
    const res = await graph("/sites/" + id + "/drive/root:/" + filePath() + ":/content");
    if (!res.ok) throw new Error("Download fehlgeschlagen (" + res.status + ")");
    const data = JSON.parse(await res.text());
    return Array.isArray(data.records) ? data.records : [];
  }

  async function upload(records, currentEtag) {
    const id = await resolveSite();
    const body = JSON.stringify({ version: 1, schemaVersion: 3, savedAt: new Date().toISOString(), records }, null, 1);
    const headers = { "Content-Type": "application/json" };
    let url = "/sites/" + id + "/drive/root:/" + filePath() + ":/content";
    if (currentEtag) headers["If-Match"] = currentEtag;
    else url += "?@microsoft.graph.conflictBehavior=fail";   // erste Anlage darf nichts ueberschreiben
    return graph(url, { method: "PUT", headers, body });
  }

  // Pro Eintrag gewinnt der juengere Stand (updatedAt); Eintraege, die nur eine Seite kennt, bleiben erhalten.
  function merge(local, remote) {
    const map = new Map();
    for (const r of remote) map.set(r.id, r);
    for (const l of local) {
      const r = map.get(l.id);
      if (!r || String(l.updatedAt || "") > String(r.updatedAt || "")) map.set(l.id, l);
    }
    return [...map.values()];
  }

  async function sync() {
    if (saving) { dirty = true; return; }
    saving = true;
    say("Synchronisiere mit SharePoint …");
    try {
      for (let attempt = 0; attempt < 5; attempt++) {
        const meta = await getMeta();
        const local = api.getRecords();
        if (!meta) {                                   // Datei existiert noch nicht: aus aktuellem Stand anlegen
          if (!local.length) throw new Error("In SharePoint liegt noch keine Datendatei und lokal sind keine Eintraege vorhanden");
          const res = await upload(local, null);
          if (res.status === 409) continue;            // jemand war schneller -> neu laden
          if (!res.ok) throw new Error("Anlegen fehlgeschlagen (" + res.status + ")");
          etag = (await res.json()).eTag;
          break;
        }
        const remote = await download();
        const merged = merge(local, remote);
        const sameAsRemote = JSON.stringify(merged) === JSON.stringify(remote);
        if (!sameAsRemote) {
          const res = await upload(merged, meta.eTag);
          if (res.status === 412) continue;            // zwischenzeitlich geaendert -> erneut abgleichen
          if (!res.ok) throw new Error("Speichern fehlgeschlagen (" + res.status + ")");
          etag = (await res.json()).eTag;
        } else {
          etag = meta.eTag;
        }
        if (JSON.stringify(merged) !== JSON.stringify(local)) api.applyRecords(merged);
        break;
      }
      dirty = false;
      say("Mit SharePoint synchronisiert um " + clock());
    } catch (e) {
      if (String(e.message).indexOf("Weiterleitung") === -1) say("SharePoint nicht erreichbar – lokal gespeichert (" + e.message + ")", true);
    } finally {
      saving = false;
      if (dirty) queueSave();
    }
  }

  function queueSave() {
    if (!SP.active) return;
    dirty = true;
    say("Speichere …");
    clearTimeout(saveTimer);
    saveTimer = setTimeout(sync, 1500);
  }
  SP.queueSave = queueSave;

  async function poll() {
    if (saving || dirty) return;
    try {
      const meta = await getMeta();
      if (!meta || meta.eTag !== etag) await sync();
    } catch (_) { /* naechster Versuch beim naechsten Intervall */ }
  }

  SP.start = async function (callbacks) {
    api = callbacks;
    if (!configured()) { say("SharePoint nicht konfiguriert – nur lokal gespeichert", true); return; }
    if (location.protocol === "file:") { say("Datei bitte ueber https oder http://localhost oeffnen – nur lokal gespeichert", true); return; }
    if (!window.msal) { say("Anmeldebibliothek nicht geladen – nur lokal gespeichert", true); return; }
    try {
      app = new msal.PublicClientApplication({
        auth: {
          clientId: cfg.clientId,
          authority: "https://login.microsoftonline.com/" + cfg.tenantId,
          redirectUri: cfg.redirectUri || (location.origin + location.pathname)
        },
        cache: { cacheLocation: "localStorage" }
      });
      await app.initialize();
      const res = await app.handleRedirectPromise();
      account = (res && res.account) || app.getAllAccounts()[0];
      if (!account) { await app.loginRedirect({ scopes: SCOPES }); return; }
      app.setActiveAccount(account);
      SP.active = true;
      await sync();
      pollTimer = setInterval(poll, (cfg.pollSeconds || 20) * 1000);
      window.addEventListener("beforeunload", function (e) {
        if (dirty || saving) { e.preventDefault(); e.returnValue = ""; }
      });
    } catch (e) {
      say("Anmeldung fehlgeschlagen – lokal gespeichert (" + e.message + ")", true);
    }
  };

  SP._merge = merge;   // fuer Tests
})();

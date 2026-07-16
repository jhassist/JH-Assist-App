'use strict';

(function exposeOneDriveSync(global) {
  const CONFIG_KEY = 'jh-assist-cloud-config-v1';
  const DEVICE_KEY = 'jh-assist-device-id-v1';
  const LAST_SYNC_KEY = 'jh-assist-last-sync-v1';
  const CLOUD_SCHEMA = 'jh-assist-cloud-v1';
  const DATA_FILE = 'data.json';
  const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';
  const SCOPES = ['Files.ReadWrite.AppFolder'];

  let hooks = null;
  let client = null;
  let clientSignature = '';
  let initializing = null;
  let activeSyncPromise = null;
  let syncQueued = false;
  let syncTimer = null;
  let currentStatus = { state: 'local', message: 'Nur lokal gespeichert.', account: '', lastSyncAt: '', appFolderUrl: '' };

  function safeParse(text, fallback = null) {
    try { return text ? JSON.parse(text) : fallback; }
    catch { return fallback; }
  }

  function defaultRedirectUri() {
    if (!global.location || !/^https?:$/.test(global.location.protocol)) return '';
    return new URL('./', global.location.href).href;
  }

  function getConfig() {
    const stored = safeParse(global.localStorage?.getItem(CONFIG_KEY), {}) || {};
    const embedded = global.JH_ASSIST_CLOUD_CONFIG || {};
    let redirectUri = String(stored.redirectUri || embedded.redirectUri || defaultRedirectUri()).trim();
    // Vorabversionen verwendeten eine separate redirect.html. Ab 1.1.0 ist
    // die veröffentlichte App-Startadresse selbst die registrierte SPA-URI.
    if (/\/redirect\.html(?:[?#].*)?$/i.test(redirectUri)) redirectUri = defaultRedirectUri();
    const autoSync = stored.autoSync === undefined ? embedded.autoSync !== false : stored.autoSync !== false;
    return {
      clientId: String(stored.clientId || embedded.clientId || '').trim(),
      tenantId: String(stored.tenantId || embedded.tenantId || 'organizations').trim() || 'organizations',
      redirectUri,
      autoSync
    };
  }

  function saveConfig(config) {
    const normalized = {
      clientId: String(config?.clientId || '').trim(),
      tenantId: String(config?.tenantId || 'organizations').trim() || 'organizations',
      redirectUri: String(config?.redirectUri || defaultRedirectUri()).trim(),
      autoSync: config?.autoSync !== false
    };
    global.localStorage.setItem(CONFIG_KEY, JSON.stringify(normalized));
    return normalized;
  }

  function getDeviceId() {
    let id = global.localStorage?.getItem(DEVICE_KEY);
    if (!id) {
      id = global.crypto?.randomUUID?.() || `${Date.now()}-${Math.random().toString(16).slice(2)}`;
      global.localStorage?.setItem(DEVICE_KEY, id);
    }
    return id;
  }

  function accountLabel() {
    const account = getAccount();
    return account?.username || account?.name || '';
  }

  function updateStatus(next) {
    currentStatus = {
      ...currentStatus,
      ...next,
      account: next?.account ?? accountLabel(),
      lastSyncAt: next?.lastSyncAt ?? global.localStorage?.getItem(LAST_SYNC_KEY) ?? ''
    };
    try { hooks?.onStatus?.({ ...currentStatus }); }
    catch (error) { console.warn('Synchronisationsstatus konnte nicht angezeigt werden.', error); }
    return currentStatus;
  }

  function getStatus() { return { ...currentStatus }; }
  function configured() { return Boolean(getConfig().clientId); }
  function signedIn() { return Boolean(getAccount()); }
  function supportedEnvironment() { return Boolean(global.msal && global.location && /^https?:$/.test(global.location.protocol)); }

  function getAccount() {
    if (!client) return null;
    return client.getActiveAccount?.() || client.getAllAccounts?.()[0] || null;
  }

  function resetClient() {
    client = null;
    clientSignature = '';
    initializing = null;
  }

  async function ensureClient() {
    const config = getConfig();
    if (!config.clientId) throw new Error('Für die Microsoft-Verbindung fehlt die Application (Client) ID.');
    if (!supportedEnvironment()) throw new Error('Die OneDrive-Synchronisierung benötigt die veröffentlichte HTTPS-App.');

    const signature = `${config.clientId}|${config.tenantId}|${config.redirectUri}`;
    if (client && clientSignature === signature) return client;
    if (initializing && clientSignature === signature) return initializing;

    clientSignature = signature;
    initializing = (async () => {
      client = new global.msal.PublicClientApplication({
        auth: {
          clientId: config.clientId,
          authority: `https://login.microsoftonline.com/${encodeURIComponent(config.tenantId)}`,
          redirectUri: config.redirectUri,
          postLogoutRedirectUri: config.redirectUri,
          navigateToLoginRequestUrl: false
        },
        cache: { cacheLocation: 'localStorage' }
      });
      await client.initialize();
      const redirectResult = await client.handleRedirectPromise();
      const account = redirectResult?.account || client.getAllAccounts()[0] || null;
      if (account) client.setActiveAccount(account);
      updateStatus(account
        ? { state: 'ready', message: 'Mit Microsoft verbunden.' }
        : { state: 'signed_out', message: 'Microsoft-Verbindung eingerichtet; Anmeldung ausstehend.' });
      return client;
    })();

    try { return await initializing; }
    finally { initializing = null; }
  }

  async function reconfigure(config) {
    saveConfig(config);
    resetClient();
    updateStatus({ state: config?.clientId ? 'signed_out' : 'local', message: config?.clientId ? 'Konfiguration gespeichert.' : 'Nur lokal gespeichert.', account: '', appFolderUrl: '' });
    if (config?.clientId) await ensureClient();
    return getConfig();
  }

  async function signIn() {
    const instance = await ensureClient();
    updateStatus({ state: 'syncing', message: 'Microsoft-Anmeldung wird geöffnet …' });
    await instance.loginRedirect({ scopes: SCOPES, redirectUri: getConfig().redirectUri, prompt: 'select_account' });
  }

  async function signOut() {
    const instance = await ensureClient();
    const account = getAccount();
    if (!account) {
      updateStatus({ state: 'signed_out', message: 'Es besteht keine Microsoft-Anmeldung.', account: '' });
      return;
    }
    await instance.logoutRedirect({ account, postLogoutRedirectUri: getConfig().redirectUri });
  }

  async function accessToken(interactive = false) {
    const instance = await ensureClient();
    const account = getAccount();
    if (!account) {
      if (interactive) {
        await signIn();
        return null;
      }
      throw new Error('Microsoft-Anmeldung erforderlich.');
    }
    try {
      const result = await instance.acquireTokenSilent({ account, scopes: SCOPES });
      return result.accessToken;
    } catch (error) {
      const interactionNeeded = error instanceof global.msal.InteractionRequiredAuthError ||
        /interaction_required|login_required|consent_required/i.test(String(error?.errorCode || error?.message || ''));
      if (interactive && interactionNeeded) {
        await instance.acquireTokenRedirect({ account, scopes: SCOPES, redirectUri: getConfig().redirectUri });
        return null;
      }
      throw error;
    }
  }

  async function graphFetch(token, path, options = {}) {
    const headers = new global.Headers(options.headers || {});
    headers.set('Authorization', `Bearer ${token}`);
    return global.fetch(`${GRAPH_BASE}${path}`, { ...options, headers });
  }

  async function getAppFolder(token) {
    const response = await graphFetch(token, '/me/drive/special/approot');
    if (!response.ok) throw new Error(`Der OneDrive-App-Ordner konnte nicht geöffnet werden (${response.status}).`);
    return response.json();
  }

  async function readRemote(token) {
    const folder = await getAppFolder(token);
    const rootId = encodeURIComponent(folder.id);
    const fileName = encodeURIComponent(DATA_FILE);
    const metadataPath = `/me/drive/items/${rootId}:/${fileName}?$select=id,eTag,@microsoft.graph.downloadUrl`;
    const metadataResponse = await graphFetch(token, metadataPath, { cache: 'no-store' });
    if (metadataResponse.status === 404) {
      return { rootId: folder.id, envelope: null, etag: '', appFolderUrl: folder.webUrl || '' };
    }
    if (!metadataResponse.ok) {
      throw new Error(`OneDrive-Metadaten konnten nicht gelesen werden (${metadataResponse.status}).`);
    }
    const metadata = await metadataResponse.json();
    const downloadUrl = metadata['@microsoft.graph.downloadUrl'];
    if (!downloadUrl) throw new Error('OneDrive hat keine Download-Adresse für die Datendatei geliefert.');

    // Browser-Anwendungen dürfen die Graph-/content-Route mit Authorization
    // wegen des 302-CORS-Verhaltens nicht direkt verwenden. Die von Graph
    // gelieferte, kurzlebige Download-Adresse ist bereits authentifiziert.
    const contentResponse = await global.fetch(downloadUrl, { cache: 'no-store' });
    if (!contentResponse.ok) throw new Error(`OneDrive-Daten konnten nicht geladen werden (${contentResponse.status}).`);
    const envelope = await contentResponse.json();
    if (envelope?.schema !== CLOUD_SCHEMA || !envelope?.state) {
      throw new Error('Die OneDrive-Datei enthält kein gültiges JH-Assist-Datenformat.');
    }
    return {
      rootId: folder.id,
      envelope,
      etag: metadata.eTag || metadata['@odata.etag'] || '',
      appFolderUrl: folder.webUrl || ''
    };
  }

  async function writeRemote(token, rootId, envelope, etag = '') {
    const headers = { 'Content-Type': 'application/json; charset=utf-8' };
    if (etag) headers['If-Match'] = etag;
    else headers['If-None-Match'] = '*';
    const fileName = encodeURIComponent(DATA_FILE);
    const response = await graphFetch(token, `/me/drive/items/${encodeURIComponent(rootId)}:/${fileName}:/content`, {
      method: 'PUT',
      headers,
      body: JSON.stringify(envelope, null, 2)
    });
    if (response.status === 412) {
      const error = new Error('OneDrive-Daten wurden zwischenzeitlich geändert.');
      error.code = 'etag_conflict';
      throw error;
    }
    if (!response.ok) {
      let details = '';
      try { details = (await response.json())?.error?.message || ''; } catch {}
      throw new Error(`OneDrive-Daten konnten nicht gespeichert werden (${response.status})${details ? `: ${details}` : '.'}`);
    }
    return response.json();
  }

  function makeEnvelope(state) {
    return {
      schema: CLOUD_SCHEMA,
      version: hooks?.version || '1.1.0',
      updatedAt: new Date().toISOString(),
      updatedByDevice: getDeviceId(),
      state
    };
  }

  async function performSync(token, attempt = 0) {
    const remote = await readRemote(token);
    const merged = hooks.mergeStates(hooks.getState(), remote.envelope?.state || null);
    hooks.setState(merged);
    const envelope = makeEnvelope(merged);
    try {
      await writeRemote(token, remote.rootId, envelope, remote.etag || '');
      return {
        merged,
        appFolderUrl: remote.appFolderUrl || '',
        cloudUpdatedAt: envelope.updatedAt
      };
    } catch (error) {
      if (error.code === 'etag_conflict' && attempt < 2) return performSync(token, attempt + 1);
      throw error;
    }
  }

  async function runSingleSync(options = {}) {
    if (global.navigator?.onLine === false) {
      updateStatus({ state: 'offline', message: 'Offline – Änderungen bleiben lokal und werden später synchronisiert.' });
      return { status: 'offline' };
    }
    if (!configured()) {
      updateStatus({ state: 'local', message: 'OneDrive-Synchronisierung ist noch nicht eingerichtet.' });
      return { status: 'unconfigured' };
    }

    updateStatus({ state: 'syncing', message: 'Synchronisierung läuft …' });
    try {
      const token = await accessToken(Boolean(options.interactive));
      if (!token) return { status: 'auth_pending' };
      const result = await performSync(token);
      const timestamp = new Date().toISOString();
      global.localStorage.setItem(LAST_SYNC_KEY, timestamp);
      updateStatus({
        state: 'synced',
        message: `Synchronisiert: ${new Date(timestamp).toLocaleString('de-DE')}`,
        lastSyncAt: timestamp,
        appFolderUrl: result.appFolderUrl
      });
      return { status: 'synced', lastSyncAt: timestamp, cloudUpdatedAt: result.cloudUpdatedAt, state: result.merged };
    } catch (error) {
      console.error('OneDrive-Synchronisierung fehlgeschlagen.', error);
      const message = String(error?.message || error);
      updateStatus({ state: /Anmeldung erforderlich/i.test(message) ? 'signed_out' : 'error', message });
      if (options.throwOnError) throw error;
      return { status: 'error', error };
    }
  }

  async function drainSyncQueue(options = {}) {
    let result;
    let nextOptions = { ...options };
    do {
      syncQueued = false;
      result = await runSingleSync(nextOptions);
      nextOptions = { ...nextOptions, interactive: false };
    } while (syncQueued && result?.status === 'synced');
    return result;
  }

  function syncNow(options = {}) {
    if (activeSyncPromise) {
      // Der Aufrufer wartet auf denselben Abgleich. Gleichzeitig wird ein
      // zweiter Durchlauf vorgemerkt, damit Änderungen während des laufenden
      // Abgleichs sicher vor einem Excel-Export in OneDrive landen.
      syncQueued = true;
      return activeSyncPromise;
    }

    activeSyncPromise = drainSyncQueue(options).finally(() => {
      activeSyncPromise = null;
      if (syncQueued) {
        syncQueued = false;
        scheduleSync(50);
      }
    });
    return activeSyncPromise;
  }

  function scheduleSync(delay = 1200) {
    const config = getConfig();
    if (!config.autoSync || !configured() || !getAccount()) return;
    if (activeSyncPromise) {
      syncQueued = true;
      return;
    }
    clearTimeout(syncTimer);
    syncTimer = setTimeout(() => syncNow(), delay);
  }

  async function initialize(initHooks) {
    hooks = initHooks;
    if (!configured()) {
      updateStatus({ state: 'local', message: 'Nur lokal gespeichert. OneDrive ist noch nicht eingerichtet.' });
      return getStatus();
    }
    try {
      await ensureClient();
      if (getAccount()) await syncNow();
      return getStatus();
    } catch (error) {
      console.error('Microsoft-Verbindung konnte nicht initialisiert werden.', error);
      updateStatus({ state: 'error', message: String(error?.message || error) });
      return getStatus();
    }
  }

  global.addEventListener?.('online', () => scheduleSync(250));
  global.addEventListener?.('focus', () => scheduleSync(500));
  global.document?.addEventListener?.('visibilitychange', () => {
    if (global.document.visibilityState === 'visible') scheduleSync(500);
  });

  global.JHOneDrive = {
    initialize,
    reconfigure,
    getConfig,
    getStatus,
    defaultRedirectUri,
    configured,
    signedIn,
    signIn,
    signOut,
    syncNow,
    scheduleSync,
    supportedEnvironment
  };
})(globalThis);

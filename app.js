// ─── STATE ────────────────────────────────────────────────────────────────────
const state = {
  msalInstance:     null,
  account:          null,
  siteId:           null,
  listId:           null,
  items:            [],
  selectedItem:     null,
  workItemField:    null,
  statusNotesField: null,
  inTeams:          false,   // true when running inside Teams client
  teamsToken:       null,    // raw Teams SSO token (exchanged for Graph token)
};

// ─── DOM REFS ─────────────────────────────────────────────────────────────────
const $ = id => document.getElementById(id);

const el = {
  signinScreen:  $('signin-screen'),
  appScreen:     $('app-screen'),
  signinBtn:     $('signin-btn'),
  signinError:   $('signin-error'),
  signoutBtn:    $('signout-btn'),
  userName:      $('user-name'),
  userAvatar:    $('user-avatar'),
  projectSelect: $('project-select'),
  loadError:     $('load-error'),
  detailPanel:   $('detail-panel'),
  statusHistory: $('status-history'),
  refreshBtn:    $('refresh-btn'),
  updateText:    $('update-text'),
  submitBtn:     $('submit-btn'),
  submitStatus:  $('submit-status'),
  submitError:   $('submit-error'),
  loading:       $('loading'),
  loadingText:   $('loading-text'),
};

// ─── LOADING HELPERS ──────────────────────────────────────────────────────────
function showLoading(msg = 'Loading…') {
  el.loadingText.textContent = msg;
  el.loading.classList.remove('hidden');
}
function hideLoading() {
  el.loading.classList.add('hidden');
}
function showError(elRef, msg) {
  elRef.textContent = msg;
  elRef.classList.remove('hidden');
}
function clearError(elRef) {
  elRef.textContent = '';
  elRef.classList.add('hidden');
}

// ─── TEAMS SDK INIT ───────────────────────────────────────────────────────────
async function initTeams() {
  try {
    await microsoftTeams.app.initialize();
    const ctx = await microsoftTeams.app.getContext();
    document.body.dataset.theme = ctx.app?.theme || 'default';
    microsoftTeams.app.registerOnThemeChangeHandler(t => { document.body.dataset.theme = t; });
    state.inTeams = true;
    return ctx;
  } catch {
    state.inTeams = false;
    return null;
  }
}

// ─── TEAMS SSO ───────────────────────────────────────────────────────────────
// Teams passes the user's existing login silently — no popup needed.
function getTeamsToken() {
  return new Promise((resolve, reject) => {
    microsoftTeams.authentication.getAuthToken({
      successCallback: resolve,
      failureCallback: reject,
    });
  });
}

// NOTE: a browser-only app (no backend) cannot exchange a Teams SSO token
// for a Graph token itself — that exchange (On-Behalf-Of) requires a
// confidential client holding a secret on a server. Without a backend, the
// realistic path is: identify the user from the Teams token (free), then
// get a real Graph token via MSAL — silently if we already have a cached
// MSAL session, or via Teams' own auth popup (interactiveSignIn) the first
// time consent is needed.
async function getTokenViaTeamsSSO() {
  const teamsToken = await getTeamsToken();
  state.teamsToken = teamsToken;

  // Decode name/email from the Teams token (it's a standard JWT) just for display.
  try {
    const payload = JSON.parse(atob(teamsToken.split('.')[1]));
    if (!state.account) {
      state.account = {
        name:     payload.name || payload.preferred_username || '',
        username: payload.preferred_username || payload.upn || '',
      };
    }
  } catch { /* ignore decode errors */ }

  const msalRequest = { scopes: CONFIG.scopes, loginHint: state.account?.username };
  try {
    const result = await state.msalInstance.acquireTokenSilent(msalRequest);
    state.account = result.account;
    return result.accessToken;
  } catch {
    // No cached MSAL session yet — need one interactive consent via Teams' popup.
    return interactiveSignIn();
  }
}

// ─── INTERACTIVE SIGN-IN (works inside Teams, where plain window.open popups
// are blocked) ──────────────────────────────────────────────────────────────
async function interactiveSignIn() {
  if (state.inTeams) {
    await new Promise((resolve, reject) => {
      microsoftTeams.authentication.authenticate({
        url: window.location.origin + '/auth-start.html',
        width: 600,
        height: 535,
        successCallback: resolve,
        failureCallback: reject,
      });
    });
    // auth-end.html cached the token in localStorage; read it back here.
    const accounts = state.msalInstance.getAllAccounts();
    if (accounts.length === 0) throw new Error('Sign-in completed but no account was found.');
    state.account = accounts[0];
    const result = await state.msalInstance.acquireTokenSilent({ scopes: CONFIG.scopes, account: state.account });
    return result.accessToken;
  }

  // Outside Teams: a normal browser popup works fine.
  const result = await state.msalInstance.loginPopup({ scopes: CONFIG.scopes });
  state.account = result.account;
  return result.accessToken;
}

// ─── MSAL INIT ────────────────────────────────────────────────────────────────
async function initMsal() {
  const msalConfig = {
    auth: {
      clientId:    CONFIG.clientId,
      authority:   `https://login.microsoftonline.com/${CONFIG.tenantId}`,
      redirectUri: window.location.origin + window.location.pathname,
    },
    // localStorage so a token acquired in the auth-start/auth-end popup
    // (a separate window) is visible to this tab/iframe afterward.
    cache: { cacheLocation: 'localStorage' },
  };

  state.msalInstance = new msal.PublicClientApplication(msalConfig);
  await state.msalInstance.initialize();

  const response = await state.msalInstance.handleRedirectPromise();
  if (response) {
    state.account = response.account;
    return true;
  }
  const accounts = state.msalInstance.getAllAccounts();
  if (accounts.length > 0) {
    state.account = accounts[0];
    return true;
  }
  return false;
}

// ─── TOKEN ACQUISITION (unified) ──────────────────────────────────────────────
async function getToken() {
  // Inside Teams: use SSO exchange
  if (state.inTeams) {
    return getTokenViaTeamsSSO();
  }

  // Outside Teams (browser): standard MSAL silent/popup
  const request = { scopes: CONFIG.scopes, account: state.account };
  try {
    const res = await state.msalInstance.acquireTokenSilent(request);
    return res.accessToken;
  } catch {
    const res = await state.msalInstance.acquireTokenPopup(request);
    state.account = res.account;
    return res.accessToken;
  }
}

// ─── STANDARD AUTH (browser fallback) ────────────────────────────────────────
async function signIn() {
  clearError(el.signinError);
  el.signinBtn.disabled = true;
  el.signinBtn.textContent = 'Signing in…';
  try {
    await interactiveSignIn();
    await enterApp();
  } catch (e) {
    showError(el.signinError, e.message || 'Sign-in failed. Please try again.');
    el.signinBtn.disabled = false;
    el.signinBtn.textContent = 'Sign In with Microsoft';
  }
}

function signOut() {
  if (state.inTeams) {
    // Can't truly sign out of Teams SSO — just reload
    window.location.reload();
  } else {
    state.msalInstance.logoutPopup({ account: state.account });
  }
}

// ─── GRAPH HELPERS ────────────────────────────────────────────────────────────
async function graphFetch(path, options = {}) {
  const token = await getToken();
  const res = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
    ...options,
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json',
      ...(options.headers || {}),
    },
  });
  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    throw new Error(err?.error?.message || `Graph error ${res.status}`);
  }
  if (res.status === 204) return null;
  return res.json();
}

// ─── SHAREPOINT DISCOVERY ─────────────────────────────────────────────────────
async function resolveSiteAndList() {
  const siteRes = await graphFetch(`/sites/${CONFIG.sharePointHost}:${CONFIG.sitePath}`);
  state.siteId = siteRes.id;

  const listsRes = await graphFetch(`/sites/${state.siteId}/lists`);
  const list = listsRes.value.find(
    l => l.displayName.toLowerCase() === CONFIG.listName.toLowerCase()
  );
  if (!list) throw new Error(`List "${CONFIG.listName}" not found on the SharePoint site.`);
  state.listId = list.id;
}

// ─── FIELD NAME RESOLUTION ────────────────────────────────────────────────────
async function resolveFieldNames(sampleFields) {
  const keys = Object.keys(sampleFields);
  function bestMatch(candidates) {
    for (const c of candidates) {
      const found = keys.find(k => k.toLowerCase() === c.toLowerCase());
      if (found) return found;
    }
    return null;
  }
  state.workItemField    = bestMatch([CONFIG.workItemField, 'Work_x0020_Item', 'WorkItem', 'Work Item', 'Title', 'Project']);
  state.statusNotesField = bestMatch([CONFIG.statusNotesField, 'Status_x0020_Notes', 'StatusNotes', 'Status Notes', 'Notes']);
  if (!state.workItemField) state.workItemField = 'Title';
}

// ─── LIST DATA ────────────────────────────────────────────────────────────────
async function loadListItems() {
  const res = await graphFetch(
    `/sites/${state.siteId}/lists/${state.listId}/items?$expand=fields&$top=500&$orderby=fields/Title asc`
  );
  const raw = res.value || [];
  if (raw.length > 0) await resolveFieldNames(raw[0].fields);
  return raw.map(item => ({
    id:          item.id,
    workItem:    item.fields[state.workItemField]    || '(unnamed)',
    statusNotes: item.fields[state.statusNotesField] || '',
  }));
}

async function refreshItems() {
  clearError(el.loadError);
  showLoading('Refreshing projects…');
  try {
    state.items = await loadListItems();
    populateDropdown();
    if (state.selectedItem) {
      const updated = state.items.find(i => i.id === state.selectedItem.id);
      if (updated) { state.selectedItem = updated; renderStatusHistory(); }
    }
  } catch (e) {
    showError(el.loadError, 'Failed to load projects: ' + e.message);
  } finally {
    hideLoading();
  }
}

// ─── DROPDOWN ─────────────────────────────────────────────────────────────────
function populateDropdown() {
  el.projectSelect.innerHTML = '';
  if (state.items.length === 0) {
    el.projectSelect.appendChild(Object.assign(document.createElement('option'), { value: '', textContent: '— No projects found —' }));
    return;
  }
  el.projectSelect.appendChild(Object.assign(document.createElement('option'), { value: '', textContent: '— Select a project —' }));
  state.items.forEach(item => {
    el.projectSelect.appendChild(Object.assign(document.createElement('option'), { value: item.id, textContent: item.workItem }));
  });
  if (state.selectedItem) el.projectSelect.value = state.selectedItem.id;
}

function onProjectChange() {
  const id = el.projectSelect.value;
  clearError(el.submitError);
  el.submitStatus.textContent = '';
  el.updateText.value = '';
  if (!id) { state.selectedItem = null; el.detailPanel.classList.add('hidden'); return; }
  state.selectedItem = state.items.find(i => i.id === id) || null;
  if (state.selectedItem) { el.detailPanel.classList.remove('hidden'); renderStatusHistory(); }
}

// ─── STATUS HISTORY ───────────────────────────────────────────────────────────
function renderStatusHistory() {
  const notes = (state.selectedItem?.statusNotes || '').trim();
  if (!notes) { el.statusHistory.innerHTML = '<p class="placeholder-text">No status notes yet.</p>'; return; }

  const entries = notes.split(/\n?───+\n?/).filter(s => s.trim());
  if (entries.length === 0) { el.statusHistory.innerHTML = '<p class="placeholder-text">No status notes yet.</p>'; return; }

  el.statusHistory.innerHTML = entries.reverse().map(entry => {
    const lines = entry.trim().split('\n');
    const headerMatch = lines[0].match(/^\[(.+?)\]$/);
    if (headerMatch) {
      const parts   = headerMatch[1].split('|').map(s => s.trim());
      const datePart = parts[0] || '';
      const author   = parts[1] || '';
      const initials = author.split(' ').map(w => w[0]).join('').slice(0, 2).toUpperCase();
      const body     = lines.slice(1).join('\n').trim();
      return `
        <div class="status-entry">
          <div class="entry-meta">
            <div class="entry-avatar">${escHtml(initials)}</div>
            <span class="entry-author">${escHtml(author)}</span>
            <span class="entry-date">${escHtml(datePart)}</span>
          </div>
          <div class="entry-body">${escHtml(body)}</div>
        </div>`;
    }
    return `<div class="status-entry"><div class="entry-body">${escHtml(entry.trim())}</div></div>`;
  }).join('');
}

function escHtml(str) {
  return String(str)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/\n/g, '<br/>');
}

// ─── SUBMIT UPDATE ────────────────────────────────────────────────────────────
async function submitUpdate() {
  const text = el.updateText.value.trim();
  if (!text || !state.selectedItem) return;

  clearError(el.submitError);
  el.submitBtn.disabled = true;
  el.submitStatus.textContent = '';
  showLoading('Posting update…');

  try {
    const now    = new Date();
    const date   = now.toLocaleDateString('en-US', { year: 'numeric', month: 'short', day: 'numeric' });
    const time   = now.toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit' });
    const author = state.account?.name || state.account?.username || 'Unknown';
    const entry  = `[${date} ${time} | ${author}]\n${text}`;
    const divider = '\n───────────────────────────────────────\n';
    const existing = (state.selectedItem.statusNotes || '').trim();
    const updated  = existing ? existing + divider + entry : entry;

    await graphFetch(
      `/sites/${state.siteId}/lists/${state.listId}/items/${state.selectedItem.id}/fields`,
      { method: 'PATCH', body: JSON.stringify({ [state.statusNotesField]: updated }) }
    );

    state.selectedItem.statusNotes = updated;
    const idx = state.items.findIndex(i => i.id === state.selectedItem.id);
    if (idx !== -1) state.items[idx].statusNotes = updated;

    el.updateText.value = '';
    renderStatusHistory();
    el.submitStatus.textContent = 'Update posted successfully.';
    el.submitStatus.className = 'submit-status success';
  } catch (e) {
    showError(el.submitError, 'Failed to post update: ' + e.message);
  } finally {
    el.submitBtn.disabled = false;
    hideLoading();
  }
}

function onTextareaInput() {
  el.submitBtn.disabled = el.updateText.value.trim().length === 0;
  el.submitStatus.textContent = '';
}

// ─── ENTER APP ────────────────────────────────────────────────────────────────
async function enterApp() {
  el.signinScreen.classList.add('hidden');
  el.appScreen.classList.remove('hidden');

  const name = state.account?.name || state.account?.username || '';
  el.userName.textContent = name;
  if (el.userAvatar) {
    el.userAvatar.textContent = name.split(' ').map(w => w[0]).join('').slice(0, 2).toUpperCase() || '?';
  }

  showLoading('Connecting to SharePoint…');
  try {
    await resolveSiteAndList();
    state.items = await loadListItems();
    populateDropdown();
  } catch (e) {
    showError(el.loadError, 'Could not load data: ' + e.message);
  } finally {
    hideLoading();
  }
}

// ─── BOOTSTRAP ────────────────────────────────────────────────────────────────
(function bootstrap() {
  // 1. Wire up events FIRST, synchronously, before any async/init work that
  //    could throw. If init fails below, the buttons must still respond.
  el.signinBtn?.addEventListener('click', signIn);
  el.signoutBtn?.addEventListener('click', signOut);
  el.projectSelect.addEventListener('change', onProjectChange);
  el.refreshBtn.addEventListener('click', refreshItems);
  el.updateText.addEventListener('input', onTextareaInput);
  el.submitBtn.addEventListener('click', submitUpdate);

  // 2. Config guard
  if (CONFIG.clientId === 'YOUR_CLIENT_ID_HERE') {
    el.signinScreen.querySelector('p').innerHTML =
      '<strong style="color:#c00">Setup required:</strong> Fill in <code>config.js</code> with your Azure AD values.';
    el.signinBtn.disabled = true;
    return;
  }

  init();
})();

async function init() {
  try {
    await initTeams();
  } catch (e) {
    console.error('initTeams failed:', e);
  }

  try {
    await initMsal();
  } catch (e) {
    console.error('initMsal failed:', e);
    showError(el.signinError,
      'Could not initialize sign-in (' + (e?.errorCode || e?.message || e) +
      '). Click "Sign In" to retry.');
    return;
  }

  // Inside Teams → attempt silent SSO first
  if (state.inTeams) {
    showLoading('Signing you in…');
    try {
      await getTeamsToken();          // populates state.account from JWT
      await enterApp();
      return;
    } catch (e) {
      // Teams identity context itself failed (rare) → fall through to manual sign-in
      hideLoading();
      showError(el.signinError,
        'Automatic sign-in failed: ' + (e?.message || e) +
        '. Click "Sign In" to authenticate manually.');
    }
  }

  // Outside Teams or SSO failed → check for existing MSAL session
  const alreadySignedIn = state.account !== null;
  if (alreadySignedIn) {
    await enterApp();
  }
  // Otherwise show the sign-in screen (already visible by default)
}

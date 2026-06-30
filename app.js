// ─── STATE ────────────────────────────────────────────────────────────────────
const state = {
  msalInstance: null,
  account:      null,
  siteId:       null,
  listId:       null,
  items:        [],          // raw list items [{id, workItem, statusNotes}, …]
  selectedItem: null,
  workItemField:    null,    // resolved internal field name
  statusNotesField: null,
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

// ─── MSAL INIT ────────────────────────────────────────────────────────────────
async function initMsal() {
  const msalConfig = {
    auth: {
      clientId:    CONFIG.clientId,
      authority:   `https://login.microsoftonline.com/${CONFIG.tenantId}`,
      redirectUri: window.location.origin + window.location.pathname,
    },
    cache: { cacheLocation: 'sessionStorage' },
  };

  state.msalInstance = new msal.PublicClientApplication(msalConfig);
  await state.msalInstance.initialize();

  // Handle redirect response (popup not used in Teams – use redirect flow)
  const response = await state.msalInstance.handleRedirectPromise();
  if (response) {
    state.account = response.account;
  } else {
    const accounts = state.msalInstance.getAllAccounts();
    if (accounts.length > 0) state.account = accounts[0];
  }

  return !!state.account;
}

// ─── AUTH ─────────────────────────────────────────────────────────────────────
async function getToken() {
  const request = { scopes: CONFIG.scopes, account: state.account };
  try {
    const res = await state.msalInstance.acquireTokenSilent(request);
    return res.accessToken;
  } catch {
    // Silent failed → interactive
    try {
      const res = await state.msalInstance.acquireTokenPopup(request);
      state.account = res.account;
      return res.accessToken;
    } catch (e) {
      throw new Error('Authentication failed: ' + e.message);
    }
  }
}

async function signIn() {
  clearError(el.signinError);
  el.signinBtn.disabled = true;
  el.signinBtn.textContent = 'Signing in…';
  try {
    await state.msalInstance.loginPopup({ scopes: CONFIG.scopes });
    const accounts = state.msalInstance.getAllAccounts();
    if (accounts.length === 0) throw new Error('No account returned.');
    state.account = accounts[0];
    await enterApp();
  } catch (e) {
    showError(el.signinError, e.message || 'Sign-in failed. Please try again.');
    el.signinBtn.disabled = false;
    el.signinBtn.textContent = 'Sign In';
  }
}

function signOut() {
  state.msalInstance.logoutPopup({ account: state.account });
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
  // 1. Get site ID
  const siteRes = await graphFetch(
    `/sites/${CONFIG.sharePointHost}:${CONFIG.sitePath}`
  );
  state.siteId = siteRes.id;

  // 2. Get list ID by display name
  const listsRes = await graphFetch(`/sites/${state.siteId}/lists`);
  const list = listsRes.value.find(
    l => l.displayName.toLowerCase() === CONFIG.listName.toLowerCase()
  );
  if (!list) throw new Error(`List "${CONFIG.listName}" not found on the SharePoint site.`);
  state.listId = list.id;
}

// ─── FIELD NAME RESOLUTION ────────────────────────────────────────────────────
// SharePoint stores column internal names differently from display names.
// We auto-detect by fetching a single item and inspecting field keys.
async function resolveFieldNames(sampleFields) {
  const keys = Object.keys(sampleFields);

  function bestMatch(candidates) {
    for (const c of candidates) {
      if (keys.some(k => k.toLowerCase() === c.toLowerCase())) {
        return keys.find(k => k.toLowerCase() === c.toLowerCase());
      }
    }
    return null;
  }

  state.workItemField = bestMatch([
    CONFIG.workItemField,
    'Work_x0020_Item', 'WorkItem', 'Work Item', 'Title', 'Project',
  ]);

  state.statusNotesField = bestMatch([
    CONFIG.statusNotesField,
    'Status_x0020_Notes', 'StatusNotes', 'Status Notes', 'Notes',
  ]);

  if (!state.workItemField) {
    // Fall back to Title if nothing matches
    state.workItemField = 'Title';
  }
}

// ─── LIST DATA ────────────────────────────────────────────────────────────────
async function loadListItems() {
  const res = await graphFetch(
    `/sites/${state.siteId}/lists/${state.listId}/items` +
    `?$expand=fields&$top=500&$orderby=fields/Title asc`
  );

  const raw = res.value || [];
  if (raw.length === 0) return [];

  // Auto-detect field names from first item
  await resolveFieldNames(raw[0].fields);

  return raw.map(item => ({
    id:          item.id,
    workItem:    item.fields[state.workItemField] || '(unnamed)',
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
      // Keep selection and refresh its data
      const updated = state.items.find(i => i.id === state.selectedItem.id);
      if (updated) {
        state.selectedItem = updated;
        renderStatusHistory();
      }
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
    const opt = document.createElement('option');
    opt.value = '';
    opt.textContent = '— No projects found —';
    el.projectSelect.appendChild(opt);
    return;
  }

  const placeholder = document.createElement('option');
  placeholder.value = '';
  placeholder.textContent = '— Select a project —';
  el.projectSelect.appendChild(placeholder);

  // Group by unique Work Item values; if duplicates, append ID
  state.items.forEach(item => {
    const opt = document.createElement('option');
    opt.value = item.id;
    opt.textContent = item.workItem;
    el.projectSelect.appendChild(opt);
  });

  // Restore previous selection if still present
  if (state.selectedItem) {
    el.projectSelect.value = state.selectedItem.id;
  }
}

function onProjectChange() {
  const id = el.projectSelect.value;
  clearError(el.submitError);
  el.submitStatus.textContent = '';
  el.updateText.value = '';

  if (!id) {
    state.selectedItem = null;
    el.detailPanel.classList.add('hidden');
    return;
  }

  state.selectedItem = state.items.find(i => i.id === id) || null;
  if (state.selectedItem) {
    el.detailPanel.classList.remove('hidden');
    renderStatusHistory();
  }
}

// ─── STATUS HISTORY RENDERING ─────────────────────────────────────────────────
function renderStatusHistory() {
  const notes = (state.selectedItem?.statusNotes || '').trim();
  if (!notes) {
    el.statusHistory.innerHTML = '<p class="placeholder-text">No status notes yet.</p>';
    return;
  }

  // Parse entries separated by "───" dividers inserted by this app,
  // or fall back to rendering the raw text if legacy format.
  const entries = notes.split(/\n?───+\n?/).filter(s => s.trim());
  if (entries.length === 0) {
    el.statusHistory.innerHTML = '<p class="placeholder-text">No status notes yet.</p>';
    return;
  }

  el.statusHistory.innerHTML = entries.reverse().map(entry => {
    const lines = entry.trim().split('\n');
    // First line might be a header like "[2026-06-30 | Brian Rooney]"
    const headerMatch = lines[0].match(/^\[(.+?)\]$/);
    if (headerMatch) {
      const meta = headerMatch[1];
      const body = lines.slice(1).join('\n').trim();
      return `
        <div class="status-entry">
          <div class="entry-meta">${escHtml(meta)}</div>
          <div class="entry-body">${escHtml(body)}</div>
        </div>`;
    }
    // Legacy / plain text
    return `<div class="status-entry"><div class="entry-body">${escHtml(entry.trim())}</div></div>`;
  }).join('');
}

function escHtml(str) {
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/\n/g, '<br/>');
}

// ─── SUBMIT STATUS UPDATE ─────────────────────────────────────────────────────
async function submitUpdate() {
  const text = el.updateText.value.trim();
  if (!text || !state.selectedItem) return;

  clearError(el.submitError);
  el.submitBtn.disabled = true;
  el.submitStatus.textContent = '';
  showLoading('Posting update…');

  try {
    // Build timestamped entry
    const now    = new Date();
    const date   = now.toLocaleDateString('en-US', { year: 'numeric', month: 'short', day: 'numeric' });
    const time   = now.toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit' });
    const author = state.account?.name || state.account?.username || 'Unknown';
    const header = `[${date} ${time} | ${author}]`;
    const entry  = `${header}\n${text}`;
    const divider = '\n───────────────────────────────────────\n';

    const existing = (state.selectedItem.statusNotes || '').trim();
    const updated  = existing ? existing + divider + entry : entry;

    // PATCH the SharePoint list item fields
    await graphFetch(
      `/sites/${state.siteId}/lists/${state.listId}/items/${state.selectedItem.id}/fields`,
      {
        method: 'PATCH',
        body: JSON.stringify({ [state.statusNotesField]: updated }),
      }
    );

    // Update local state
    state.selectedItem.statusNotes = updated;
    const idx = state.items.findIndex(i => i.id === state.selectedItem.id);
    if (idx !== -1) state.items[idx].statusNotes = updated;

    el.updateText.value = '';
    renderStatusHistory();
    el.submitStatus.textContent = '✓ Update posted successfully.';
    el.submitStatus.className = 'submit-status success';
  } catch (e) {
    showError(el.submitError, 'Failed to post update: ' + e.message);
  } finally {
    el.submitBtn.disabled = false;
    hideLoading();
  }
}

// ─── TEXTAREA → BUTTON STATE ──────────────────────────────────────────────────
function onTextareaInput() {
  el.submitBtn.disabled = el.updateText.value.trim().length === 0;
  el.submitStatus.textContent = '';
}

// ─── ENTER APP ────────────────────────────────────────────────────────────────
async function enterApp() {
  el.signinScreen.classList.add('hidden');
  el.appScreen.classList.remove('hidden');
  el.userName.textContent = state.account?.name || state.account?.username || '';

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

// ─── TEAMS SDK INIT ───────────────────────────────────────────────────────────
async function initTeams() {
  try {
    await microsoftTeams.app.initialize();
    const ctx = await microsoftTeams.app.getContext();
    // Teams theme support
    document.body.dataset.theme = ctx.app?.theme || 'default';
    microsoftTeams.app.registerOnThemeChangeHandler(theme => {
      document.body.dataset.theme = theme;
    });
  } catch {
    // Running outside Teams (browser) — that's fine
  }
}

// ─── BOOTSTRAP ────────────────────────────────────────────────────────────────
(async function bootstrap() {
  await initTeams();

  // Guard: if client ID is still placeholder, show helpful error
  if (CONFIG.clientId === 'YOUR_CLIENT_ID_HERE') {
    el.signinScreen.querySelector('p').innerHTML =
      '<strong style="color:#c00">Setup required:</strong> Open <code>config.js</code> ' +
      'and paste your Azure AD App Registration values. See <code>README.md</code> for instructions.';
    el.signinBtn.disabled = true;
    return;
  }

  const alreadySignedIn = await initMsal();

  el.signinBtn.addEventListener('click', signIn);
  el.signoutBtn.addEventListener('click', signOut);
  el.projectSelect.addEventListener('change', onProjectChange);
  el.refreshBtn.addEventListener('click', refreshItems);
  el.updateText.addEventListener('input', onTextareaInput);
  el.submitBtn.addEventListener('click', submitUpdate);

  if (alreadySignedIn) {
    await enterApp();
  }
})();

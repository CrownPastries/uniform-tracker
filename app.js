/* ============================================================
   UniTrack — Application Logic (app.js)
   Workflow: 1-Receive from Cintas → 2-Distribute/Report Issue → 3-Return → 4-Send to Cintas
   ============================================================ */

// ============================================================
// DATA STORE
// ============================================================
let DB_MEMORY = {
  users: [],
  employees: [],
  transactions: [],
  centres: ["Main Plant","Warehouse A","Warehouse B","Office"],
  uniformTypes: ["Shirt","Pants","Jacket","Safety Vest","Cap","Gloves"]
};

// IndexedDB Helper
const IDB = {
  db: null,
  init() {
    return new Promise((resolve, reject) => {
      const request = indexedDB.open('UniTrackDB', 2);
      request.onupgradeneeded = (e) => {
        const db = e.target.result;
        if (!db.objectStoreNames.contains('store')) {
          db.createObjectStore('store');
        }
        if (!db.objectStoreNames.contains('transactions')) {
          db.createObjectStore('transactions', { keyPath: 'id' });
        }
      };
      request.onsuccess = (e) => {
        this.db = e.target.result;
        resolve();
      };
      request.onerror = (e) => reject(e);
    });
  },
  get(key) {
    return new Promise((resolve, reject) => {
      const tx = this.db.transaction('store', 'readonly');
      const req = tx.objectStore('store').get(key);
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
  },
  set(key, val) {
    return new Promise((resolve, reject) => {
      const tx = this.db.transaction('store', 'readwrite');
      const req = tx.objectStore('store').put(val, key);
      req.onsuccess = () => resolve();
      req.onerror = () => reject(req.error);
    });
  },
  // New optimized transaction methods
  getAllTransactions() {
    return new Promise((resolve, reject) => {
      const tx = this.db.transaction('transactions', 'readonly');
      const req = tx.objectStore('transactions').getAll();
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
  },
  saveTransaction(txn) {
    return new Promise((resolve, reject) => {
      const tx = this.db.transaction('transactions', 'readwrite');
      const req = tx.objectStore('transactions').put(txn);
      req.onsuccess = () => resolve();
      req.onerror = () => reject(req.error);
    });
  },
  deleteTransaction(id) {
    return new Promise((resolve, reject) => {
      const tx = this.db.transaction('transactions', 'readwrite');
      const req = tx.objectStore('transactions').delete(id);
      req.onsuccess = () => resolve();
      req.onerror = () => reject(req.error);
    });
  },
  clearTransactions() {
    return new Promise((resolve, reject) => {
      const tx = this.db.transaction('transactions', 'readwrite');
      const req = tx.objectStore('transactions').clear();
      req.onsuccess = () => resolve();
      req.onerror = () => reject(req.error);
    });
  }
};

const DB = {
  get users()        { return DB_MEMORY.users; },
  get employees()    { return DB_MEMORY.employees; },
  get transactions() { return DB_MEMORY.transactions; },
  get centres()      { return DB_MEMORY.centres; },
  get uniformTypes() { return DB_MEMORY.uniformTypes; },
  saveUsers(d)        { DB_MEMORY.users = d; IDB.set('ut_users', d); },
  saveEmployees(d)    { DB_MEMORY.employees = d; IDB.set('ut_employees', d); },
  saveCentres(d)      { DB_MEMORY.centres = d; IDB.set('ut_centres', d); },
  saveTypes(d)        { DB_MEMORY.uniformTypes = d; IDB.set('ut_types', d); },
  addUser(u)        { const list = this.users; list.push(u); this.saveUsers(list); },
  addTransaction(t) {
    this.transactions.unshift(t);
    IDB.saveTransaction(t);
  },
  removeTransaction(id) {
    DB_MEMORY.transactions = DB_MEMORY.transactions.filter(t => t.id !== id);
    IDB.deleteTransaction(id);
  },
  addEmployee(e)    { const list = this.employees; list.push(e); this.saveEmployees(list); }
};

// ============================================================
// CLOUD SYNC  (Google Sheets via Apps Script)
// ============================================================
const CloudSync = {
  get apiUrl() { return localStorage.getItem('ut_cloud_url') || ''; },
  get apiKey() { return localStorage.getItem('ut_cloud_key') || ''; },

  isReady() { return !!(this.apiUrl && this.apiKey); },

  configure(url, key) {
    localStorage.setItem('ut_cloud_url', url.trim());
    localStorage.setItem('ut_cloud_key', key.trim());
  },

  // Apps Script POST requires Content-Type text/plain (not application/json)
  async post(action, payload) {
    const resp = await fetch(this.apiUrl, {
      method: 'POST',
      redirect: 'follow',
      body: JSON.stringify({ key: this.apiKey, action, payload }),
      headers: { 'Content-Type': 'text/plain' }
    });
    return this._parseResp(resp);
  },

  async get(action) {
    const url = `${this.apiUrl}?key=${encodeURIComponent(this.apiKey)}&action=${encodeURIComponent(action)}`;
    const resp = await fetch(url, { redirect: 'follow' });
    return this._parseResp(resp);
  },

  // Parse Apps Script response safely — HTML (login redirects, errors) returns a useful message
  async _parseResp(resp) {
    const text = await resp.text();
    try {
      return JSON.parse(text);
    } catch (_) {
      // HTML means Apps Script redirected to login or the script threw an unhandled error
      if (text.includes('<!DOCTYPE') || text.includes('<html')) {
        if (text.includes('signin') || text.includes('accounts.google')) {
          throw new Error('Apps Script requires authorization. Open the Web App URL directly in your browser and sign in, then retry.');
        }
        throw new Error('Apps Script returned an HTML page instead of JSON. Make sure you deployed as "Anyone" (not "Anyone with Google Account") and created a New Version deployment.');
      }
      throw new Error('Invalid response from Apps Script: ' + text.slice(0, 120));
    }
  },

  async pushTransaction(txn) {
    if (!this.isReady()) return;
    try { await this.post('addTransaction', txn); } catch (_) {}
  },

  async pushEmployee(emp) {
    if (!this.isReady()) return;
    try { await this.post('saveEmployee', emp); } catch (_) {}
  },

  async removeEmployee(id) {
    if (!this.isReady()) return;
    try { await this.post('deleteEmployee', { id }); } catch (_) {}
  },

  async syncAll() {
    if (!this.isReady()) { showToast('Cloud sync not configured. Go to Settings → Cloud Sync.', 'error'); return; }
    const btn = document.getElementById('cloud-sync-btn');
    if (btn) { btn.disabled = true; btn.textContent = 'Syncing…'; }
    try {
      const result = await this.post('syncAll', {
        employees:    DB.employees,
        transactions: DB.transactions,
        centres:      DB.centres,
        uniformTypes: DB.uniformTypes
      });
      if (result && result.done) {
        localStorage.setItem('ut_last_sync', new Date().toISOString());
        showToast('✓ All data pushed to Google Sheets!', 'success');
      } else {
        showToast('Sync failed: ' + (result?.error || 'Unknown error'), 'error');
      }
    } catch (e) {
      showToast('Sync error: ' + e.message, 'error');
    } finally {
      if (btn) { btn.disabled = false; btn.textContent = 'Push All to Cloud'; }
      updateSyncStatus();
    }
  },

  async pullAll(silent = false) {
    if (!this.isReady()) { 
      if (!silent) showToast('Cloud sync not configured. Go to Settings → Cloud Sync.', 'error'); 
      return; 
    }
    const btn = document.getElementById('cloud-pull-btn');
    if (btn) { btn.disabled = true; btn.textContent = 'Pulling…'; }
    try {
      const data = await this.get('getAll');
      if (!data || data.error) { 
        if (!silent) showToast('Pull failed: ' + (data?.error || 'No data'), 'error'); 
        return; 
      }
      if (data.employees)    DB.saveEmployees(data.employees);
      if (data.centres && data.centres.length) DB.saveCentres(data.centres);
      if (data.uniformTypes && data.uniformTypes.length) DB.saveTypes(data.uniformTypes);
      if (data.transactions) {
        await IDB.clearTransactions();
        DB_MEMORY.transactions = data.transactions;
        for (const t of data.transactions) await IDB.saveTransaction(t);
      }
      localStorage.setItem('ut_last_sync', new Date().toISOString());
      if (!silent) showToast('✓ Data pulled from Google Sheets!', 'success');
      renderDashboard(); renderEmployeeList(); renderSettings();
    } catch (e) {
      if (!silent) showToast('Pull error: ' + e.message, 'error');
    } finally {
      if (btn) { btn.disabled = false; btn.textContent = 'Pull from Cloud'; }
      updateSyncStatus();
    }
  },

  async testConnection() {
    if (!this.isReady()) { showToast('Enter the Web App URL and API Key first.', 'error'); return; }
    const btn = document.getElementById('cloud-test-btn');
    if (btn) { btn.disabled = true; btn.textContent = 'Testing…'; }
    try {
      const data = await this.get('getAll');
      if (data && data.error) {
        showToast('Connection failed: ' + data.error, 'error');
      } else {
        const empCount = (data.employees || []).length;
        const txnCount = (data.transactions || []).length;
        showToast(`✓ Connected! Cloud has ${empCount} employees & ${txnCount} transactions.`, 'success');
        localStorage.setItem('ut_last_sync', new Date().toISOString());
      }
    } catch (e) {
      showToast('Connection error: ' + e.message, 'error');
    } finally {
      if (btn) { btn.disabled = false; btn.textContent = 'Test Connection'; }
      updateSyncStatus();
    }
  }
};

function updateSyncStatus() {
  const statusEl = document.getElementById('cloud-status-text');
  const dotEl    = document.getElementById('cloud-status-dot');
  const ts       = localStorage.getItem('ut_last_sync');
  const isAutoSync = autoSyncInterval && CloudSync.isReady();
  
  if (statusEl) {
    const baseText = ts ? 'Last sync: ' + formatDateTime(ts) : 'Never synced';
    statusEl.textContent = isAutoSync ? baseText + ' (Auto-sync active)' : baseText;
  }
  if (dotEl)    dotEl.className = 'cloud-dot ' + (CloudSync.isReady() ? 'ready' : 'off');
}

function saveCloudConfig() {
  const url = (document.getElementById('cloud-url')?.value || '').trim();
  const key = (document.getElementById('cloud-key')?.value || '').trim();
  if (!url || !key) { showToast('Please enter both the Web App URL and API Key.', 'error'); return; }
  if (!url.includes('script.google.com/macros/s/')) {
    showToast('URL looks wrong — it should contain script.google.com/macros/s/', 'error'); return;
  }
  CloudSync.configure(url, key);
  showToast('Cloud configuration saved! Click Test Connection to verify.', 'success');
  updateSyncStatus();
  startAutoSync(); // Start automatic syncing if newly configured
}
window.saveCloudConfig = saveCloudConfig;
window.cloudSyncAll    = () => CloudSync.syncAll();
window.cloudPullAll    = () => CloudSync.pullAll(false); // Manual pull, show toasts
window.cloudTestConn   = () => CloudSync.testConnection();

// Automatic sync every 5 minutes
let autoSyncInterval;
function startAutoSync() {
  // Clear any existing interval
  if (autoSyncInterval) clearInterval(autoSyncInterval);
  
  // Only start if cloud sync is configured
  if (!CloudSync.isReady()) return;
  
  // Sync every 5 minutes (300,000 ms)
  autoSyncInterval = setInterval(async () => {
    try {
      // Only auto-sync if the app is visible (not in background tab)
      if (document.hidden) return;
      
      console.log('Auto-syncing with Google Sheets...');
      await CloudSync.pullAll(true); // Silent auto-sync
      // Don't show toast for automatic syncs to avoid spam
    } catch (e) {
      console.warn('Auto-sync failed:', e.message);
      // Don't show error toast for automatic syncs
    }
  }, 5 * 60 * 1000); // 5 minutes
  
  console.log('Automatic cloud sync started (every 5 minutes)');
  updateSyncStatus(); // Update status to show auto-sync is active
}

// Stop auto sync when logging out
function stopAutoSync() {
  if (autoSyncInterval) {
    clearInterval(autoSyncInterval);
    autoSyncInterval = null;
    console.log('Automatic cloud sync stopped');
    updateSyncStatus(); // Update status to remove auto-sync indicator
  }
}

// ============================================================
// UTILITIES
// ============================================================
function uid()  { return Date.now().toString(36) + Math.random().toString(36).slice(2,7); }
function today(){ return new Date().toISOString().slice(0,10); }
function formatDate(d) { if (!d) return ''; const [y,m,day] = d.split('-'); return `${m}/${day}/${y}`; }
function formatDateTime(iso) {
  if (!iso) return '';
  const d = new Date(iso);
  return d.toLocaleDateString('en-US',{month:'short',day:'numeric'}) + ' ' +
         d.toLocaleTimeString('en-US',{hour:'2-digit',minute:'2-digit'});
}
function avatarColor(name) {
  const colors = ['#2563eb','#7c3aed','#db2777','#d97706','#059669','#0891b2','#dc2626','#65a30d'];
  let h = 0; for (let c of (name||'A')) h = (h*31 + c.charCodeAt(0)) & 0xffffffff;
  return colors[Math.abs(h) % colors.length];
}
function initials(first, last) { return ((first||'')[0]||'') + ((last||'')[0]||''); }
function actionLabel(a) {
  return {
    distributed:          'Distributed',
    returned:             'Returned',
    sent_to_cintas:       '→ Cintas',
    received_from_cintas: '← Cintas',
    reported_issue:       'Issue/Damaged'
  }[a] || a;
}
function escHtml(s) {
  return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

// ============================================================
// UI STATE
// ============================================================
let currentUser   = null;
let currentPage   = 'dashboard';
let currentAction = 'distributed';
let currentReport = 'missing';
let selectedEmployee  = null;
let sessionScans  = 0;
let sidebarOpen   = true;
let focusLockInterval = null;  // keeps barcode input focused on scan page

// Excel import state
let xlRawRows = [];
let xlHeaders  = [];
let xlMapping  = {};

// ============================================================
// INIT
// ============================================================
document.addEventListener('DOMContentLoaded', async () => {
  document.getElementById('scan-date').value = today();
  const now = new Date();
  document.getElementById('dateDisplay').textContent =
    now.toLocaleDateString('en-US',{weekday:'short',month:'long',day:'numeric',year:'numeric'});
  
  try {
    await IDB.init();
    
    // Migration logic
    const lsEmployees = localStorage.getItem('ut_employees');
    const lsTransactions = localStorage.getItem('ut_transactions');
    const lsCentres = localStorage.getItem('ut_centres');
    const lsTypes = localStorage.getItem('ut_types');
    
    let hasMigratedAny = false;
    
    // Employees
    let empData = await IDB.get('ut_employees');
    if (!empData && lsEmployees) { empData = JSON.parse(lsEmployees); hasMigratedAny = true; }
    if (empData) DB_MEMORY.employees = empData;
    
    // Transactions (Optimized Record-Based Storage)
    let txData = await IDB.getAllTransactions();
    // Migration from old array-based store if detected
    const legacyTxData = await IDB.get('ut_transactions');
    if (legacyTxData && legacyTxData.length > 0) {
      console.log('Migrating legacy transactions to record-based store...');
      for (const t of legacyTxData) {
        if (!t.id) t.id = uid();
        await IDB.saveTransaction(t);
        txData.push(t);
      }
      IDB.set('ut_transactions', []); // Clear legacy store
      hasMigratedAny = true;
    }
    // Final check for localStorage migration
    if (txData.length === 0 && lsTransactions) {
      const parsed = JSON.parse(lsTransactions);
      for (const t of parsed) {
        if (!t.id) t.id = uid();
        await IDB.saveTransaction(t);
        txData.push(t);
      }
      hasMigratedAny = true;
    }
    // Sort transactions by date (newest first)
    DB_MEMORY.transactions = txData.sort((a,b) => new Date(b.createdAt) - new Date(a.createdAt));
    
    // Centres
    let cenData = await IDB.get('ut_centres');
    if (!cenData && lsCentres) { cenData = JSON.parse(lsCentres); hasMigratedAny = true; }
    if (cenData) DB_MEMORY.centres = cenData;
    
    // Types
    let typeData = await IDB.get('ut_types');
    if (!typeData && lsTypes) { typeData = JSON.parse(lsTypes); hasMigratedAny = true; }
    if (typeData) DB_MEMORY.uniformTypes = typeData;
    
    // Users
    let userData = await IDB.get('ut_users');
    if (!userData || userData.length === 0) {
      userData = [{ username: 'admin', password: 'admin', role: 'admin' }];
      hasMigratedAny = true;
    }
    DB_MEMORY.users = userData;

    if (hasMigratedAny) {
      console.log('Migrated data from localStorage to IndexedDB');
      IDB.set('ut_employees', DB_MEMORY.employees);
      IDB.set('ut_transactions', DB_MEMORY.transactions);
      IDB.set('ut_centres', DB_MEMORY.centres);
      IDB.set('ut_types', DB_MEMORY.uniformTypes);
      IDB.set('ut_users', DB_MEMORY.users);
    }
  } catch (e) {
    console.error("IndexedDB Intialization failed, running empty defaults.", e);
  }

  // Auth Bootloader Check
  const activeSession = sessionStorage.getItem('ut_active_user');
  if (activeSession) {
    const user = DB_MEMORY.users.find(u => u.username === activeSession);
    if (user) {
      currentUser = user;
      document.body.setAttribute('data-role', currentUser.role);
      document.getElementById('login-overlay').classList.add('hidden');
      document.getElementById('app-wrapper').style.display = 'flex';
      
      showPage('dashboard');
      setAction('distributed');
    }
  } else {
    // If no auth, focus login
    setTimeout(() => {
      const lu = document.getElementById('login-user');
      if (lu) lu.focus();
    }, 100);
  }

  document.addEventListener('click', (e) => {
    const dd = document.getElementById('scan-emp-dropdown');
    const wrapper = document.querySelector('.emp-search-wrapper');
    if (dd && wrapper && !wrapper.contains(e.target)) dd.classList.add('hidden');
  });

  // Start automatic cloud sync if configured
  startAutoSync();
});

// ============================================================
// AUTH & ROLES
// ============================================================
function handleLogin() {
  const u = document.getElementById('login-user').value.trim();
  const p = document.getElementById('login-pass').value;
  const user = DB_MEMORY.users.find(x => x.username === u && x.password === p);
  if (!user) {
    showToast('Invalid username or password', 'error');
    return;
  }
  currentUser = user;
  sessionStorage.setItem('ut_active_user', user.username);
  document.body.setAttribute('data-role', currentUser.role);
  document.getElementById('login-overlay').classList.add('hidden');
  document.getElementById('app-wrapper').style.display = 'flex';
  document.getElementById('login-pass').value = '';
  
  showPage('dashboard');
  setAction('distributed');
  showToast(`Welcome, ${u}!`, 'success');
}
window.handleLogin = handleLogin;

function handleLogout() {
  console.log("handleLogout called!");
  stopAutoSync(); // Stop automatic syncing
  currentUser = null;
  sessionStorage.removeItem('ut_active_user');
  document.body.removeAttribute('data-role');
  document.getElementById('app-wrapper').style.display = 'none';
  document.getElementById('login-overlay').classList.remove('hidden');
  document.getElementById('login-user').value = '';
  document.getElementById('login-user').focus();
  showToast('Logged out successfully.', 'success');
}
window.handleLogout = handleLogout;

// ============================================================
// SIDEBAR & NAVIGATION
// ============================================================
function toggleSidebar() {
  sidebarOpen = !sidebarOpen;
  const sidebar = document.getElementById('sidebar');
  const main    = document.querySelector('.main-content');
  if (sidebarOpen) { sidebar.classList.remove('closed'); main.classList.remove('expanded'); }
  else             { sidebar.classList.add('closed');    main.classList.add('expanded');    }
}

function showPage(page) {
  currentPage = page;
  document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
  const pageEl = document.getElementById('page-' + page);
  if (pageEl) pageEl.classList.add('active');
  const navEl = document.querySelector(`.nav-item[data-page="${page}"]`);
  if (navEl) navEl.classList.add('active');
  const titles = {dashboard:'Dashboard',employees:'Employees',scan:'Scan Entry',reports:'Reports',settings:'Settings'};
  document.getElementById('pageTitle').textContent = titles[page] || page;
  // Start/stop scanner focus lock
  if (page === 'scan') startFocusLock();
  else                 stopFocusLock();

  if (page === 'dashboard') renderDashboard();
  if (page === 'employees') renderEmployeeList();
  if (page === 'scan')      setTimeout(() => { document.getElementById('barcodeInput').focus(); }, 150);
  if (page === 'reports')   renderReport();
  if (page === 'settings')  renderSettings();
}

// ============================================================
// DASHBOARD
// ============================================================
function computeStats() {
  const txns     = DB.transactions;
  const barcodes = [...new Set(txns.map(t => t.barcode))];
  let out = 0, cintas = 0, warehouse = 0, issues = 0;
  barcodes.forEach(bc => {
    const last = txns.find(t => t.barcode === bc);
    if (!last) return;
    if      (last.action === 'distributed')                                        out++;
    else if (last.action === 'sent_to_cintas')                                     cintas++;
    else if (last.action === 'reported_issue')                                     issues++;
    else if (last.action === 'returned' || last.action === 'received_from_cintas') warehouse++;
  });
  return { total: barcodes.length, out, cintas, warehouse, issues };
}

function renderDashboard() {
  const stats = computeStats();
  document.getElementById('d-total').textContent    = stats.total;
  document.getElementById('d-out').textContent      = stats.out;
  document.getElementById('d-cintas').textContent   = stats.cintas;
  document.getElementById('d-warehouse').textContent= stats.warehouse;
  document.getElementById('sb-out').textContent     = stats.out + ' out';
  document.getElementById('sb-cintas').textContent  = stats.cintas + ' at Cintas';

  const container = document.getElementById('recent-activity');
  const txns = DB.transactions.slice(0, 20);
  if (!txns.length) {
    container.innerHTML = `<div class="empty-state"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2"/><rect x="9" y="3" width="6" height="4" rx="2"/></svg><p>No activity yet.</p></div>`;
    return;
  }
  container.innerHTML = txns.map(t => {
    const emp     = t.employeeId ? DB.employees.find(e => e.id === t.employeeId) : null;
    const empName = emp ? `${emp.firstName} ${emp.lastName}` :
                    (t.action.includes('cintas') || t.action === 'reported_issue' ? 'Cintas' : 'Warehouse');
    return `<div class="activity-item">
      <span class="activity-dot ${t.action}"></span>
      <div class="activity-info">
        <div class="activity-barcode">${escHtml(t.barcode)}</div>
        <div class="activity-meta">${actionLabel(t.action)} — ${escHtml(empName)}</div>
      </div>
      <div class="activity-time">${formatDateTime(t.createdAt)}</div>
    </div>`;
  }).join('');
}

// ============================================================
// EMPLOYEES
// ============================================================
function renderEmployeeList() {
  const query = (document.getElementById('emp-search')?.value || '').toLowerCase();
  const list  = DB.employees.filter(e =>
    !query || e.firstName?.toLowerCase().includes(query) ||
    e.lastName?.toLowerCase().includes(query) ||
    e.employeeId?.toLowerCase().includes(query) ||
    e.productionCentre?.toLowerCase().includes(query)
  );
  const container = document.getElementById('employee-list');
  if (!list.length) {
    container.innerHTML = `<div class="empty-state" style="grid-column:1/-1"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M17 21v-2a4 4 0 00-4-4H5a4 4 0 00-4 4v2"/><circle cx="9" cy="7" r="4"/></svg><p>${query ? 'No employees match.' : 'No employees yet. Click "Add Employee" to start.'}</p></div>`;
    return;
  }
  const txns = DB.transactions;
  container.innerHTML = list.map(e => {
    const color = avatarColor(e.firstName + e.lastName);
    const ini   = initials(e.firstName, e.lastName);
    const myBarcodes = [...new Set(txns.filter(t => t.employeeId === e.id).map(t => t.barcode))];
    const held = myBarcodes.filter(bc => { const l = txns.find(t => t.barcode === bc); return l && l.action === 'distributed'; }).length;
    return `<div class="emp-card" onclick="showEmployeeDetail('${e.id}')">
      <div class="emp-card-header">
        <div class="emp-avatar" style="background:${color}">${escHtml(ini)}</div>
        <div><div class="emp-name">${escHtml(e.firstName)} ${escHtml(e.lastName)}</div><div class="emp-id">${escHtml(e.employeeId||'—')}</div></div>
      </div>
      <div class="emp-tags">
        ${e.productionCentre ? `<span class="emp-tag centre">${escHtml(e.productionCentre)}</span>` : ''}
        ${e.department       ? `<span class="emp-tag dept">${escHtml(e.department)}</span>` : ''}
      </div>
      <div class="emp-card-footer">
        <div class="emp-uniform-count">Holding: <strong>${held}</strong> uniform${held!==1?'s':''}</div>
        <div class="emp-card-actions" onclick="event.stopPropagation()">
          <button class="emp-action-btn" title="Edit" onclick="openEmployeeModal('${e.id}')">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M11 4H4a2 2 0 00-2 2v14a2 2 0 002 2h14a2 2 0 002-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 013 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
          </button>
          <button class="emp-action-btn del" title="Delete" onclick="deleteEmployee('${e.id}')">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="3 6 5 6 21 6"/><path d="M19 6l-1 14a2 2 0 01-2 2H8a2 2 0 01-2-2L5 6"/></svg>
          </button>
        </div>
      </div>
    </div>`;
  }).join('');
}

function openEmployeeModal(id) {
  document.getElementById('emp-modal-title').textContent = id ? 'Edit Employee' : 'Add Employee';
  document.getElementById('emp-id').value = id || '';
  const sel = document.getElementById('emp-centre');
  sel.innerHTML = '<option value="">— Select —</option>' + DB.centres.map(c => `<option value="${escHtml(c)}">${escHtml(c)}</option>`).join('');
  if (id) {
    const e = DB.employees.find(emp => emp.id === id);
    if (e) {
      document.getElementById('emp-first').value  = e.firstName || '';
      document.getElementById('emp-last').value   = e.lastName || '';
      document.getElementById('emp-empid').value  = e.employeeId || '';
      document.getElementById('emp-centre').value = e.productionCentre || '';
      document.getElementById('emp-dept').value   = e.department || '';
      document.getElementById('emp-phone').value  = e.phone || '';
      document.getElementById('emp-notes').value  = e.notes || '';
    }
  } else {
    ['emp-first','emp-last','emp-empid','emp-dept','emp-phone','emp-notes'].forEach(fid => { document.getElementById(fid).value = ''; });
    document.getElementById('emp-centre').value = '';
  }
  openModal('employee-modal');
}

function saveEmployee() {
  const first = document.getElementById('emp-first').value.trim();
  const last  = document.getElementById('emp-last').value.trim();
  if (!first || !last) { showToast('First and last name are required.', 'error'); return; }
  const id   = document.getElementById('emp-id').value;
  const data = {
    id: id || uid(),
    firstName: first, lastName: last,
    employeeId: document.getElementById('emp-empid').value.trim(),
    productionCentre: document.getElementById('emp-centre').value,
    department: document.getElementById('emp-dept').value.trim(),
    phone: document.getElementById('emp-phone').value.trim(),
    notes: document.getElementById('emp-notes').value.trim(),
    createdAt: new Date().toISOString()
  };
  let list = DB.employees;
  if (id) { const idx = list.findIndex(e => e.id === id); if (idx >= 0) { data.createdAt = list[idx].createdAt; list[idx] = data; } }
  else    { list.push(data); }
  DB.saveEmployees(list);
  CloudSync.pushEmployee(data);
  closeModal('employee-modal');
  showToast(`Employee ${first} ${last} saved!`, 'success');
  renderEmployeeList();
}

function deleteEmployee(id) {
  const e = DB.employees.find(emp => emp.id === id);
  if (!e) return;
  confirm_dialog(`Delete ${e.firstName} ${e.lastName}?`, 'This employee will be removed. Their transaction history will be kept.', () => {
    DB.saveEmployees(DB.employees.filter(emp => emp.id !== id));
    CloudSync.removeEmployee(id);
    showToast('Employee deleted.', 'success');
    renderEmployeeList();
  });
}

function showEmployeeDetail(id) {
  const e = DB.employees.find(emp => emp.id === id);
  if (!e) return;
  document.getElementById('emp-detail-name').textContent = `${e.firstName} ${e.lastName}`;
  document.getElementById('emp-detail-info').innerHTML = [
    ['Employee ID', e.employeeId||'—'],['Production Centre', e.productionCentre||'—'],
    ['Department', e.department||'—'],['Phone', e.phone||'—'],['Notes', e.notes||'—']
  ].map(([l,v]) => `<div class="emp-detail-field"><label>${l}</label><span>${escHtml(v)}</span></div>`).join('');

  const txns     = DB.transactions.filter(t => t.employeeId === id);
  const detailEl = document.getElementById('emp-detail-uniforms');
  if (!txns.length) {
    detailEl.innerHTML = '<div class="empty-state"><p>No uniform transactions found for this employee.</p></div>';
  } else {
    detailEl.innerHTML = `<div class="report-table-wrap"><table class="report-table">
      <thead><tr><th>Barcode</th><th>Action</th><th>Date</th><th>Notes</th></tr></thead>
      <tbody>${txns.map(t => `<tr>
        <td class="bc-mono">${escHtml(t.barcode)}</td>
        <td><span class="report-badge ${t.action==='distributed'?'out':t.action==='returned'?'returned':t.action==='reported_issue'?'issue':'cintas'}">${actionLabel(t.action)}</span></td>
        <td>${formatDate(t.date)}</td>
        <td>${escHtml(t.notes||'')}</td>
      </tr>`).join('')}</tbody>
    </table></div>`;
  }
  openModal('emp-detail-modal');
}

// ============================================================
// SCAN ENTRY
// ============================================================
function setAction(action) {
  currentAction = action;
  document.querySelectorAll('.action-tab').forEach(t => t.classList.remove('active'));
  const tab = document.getElementById('tab-' + action);
  if (tab) tab.classList.add('active');

  const selector      = document.getElementById('employee-selector');
  const needsEmployee = action === 'distributed' || action === 'returned';
  selector.style.display = needsEmployee ? 'block' : 'none';

  // Update scanner hint text
  const hints = {
    received_from_cintas: 'Scan each item arriving from Cintas',
    distributed:          'Select employee above — then scan each item',
    reported_issue:       'Scan damaged / dirty items to flag them',
    returned:             'Scan each item being returned to warehouse',
    sent_to_cintas:       'Scan each item going back to Cintas'
  };
  const hintEl = document.getElementById('scannerHintText');
  if (hintEl) hintEl.textContent = hints[action] || '';

  // Return focus to scanner after switching action
  setTimeout(() => {
    const inp = document.getElementById('barcodeInput');
    if (inp) inp.focus();
  }, 80);
}

// ---- Focus lock: keeps barcode input focused while on scan page ----
function startFocusLock() {
  stopFocusLock(); // clear any existing
  focusLockInterval = setInterval(() => {
    if (currentPage !== 'scan') { stopFocusLock(); return; }
    const active = document.activeElement;
    const inp    = document.getElementById('barcodeInput');
    // Do NOT steal focus from employee search or modal inputs
    const allowed = ['scan-emp-search', 'barcodeInput', 'scan-date', 'manual-barcode', 'manual-notes'];
    if (inp && active && !allowed.includes(active.id) && !active.closest('.modal-overlay')) {
      inp.focus();
    }
  }, 350);
}

function stopFocusLock() {
  if (focusLockInterval) { clearInterval(focusLockInterval); focusLockInterval = null; }
}

// ---- Scanner zone UI events ----
function onScannerFocus() {
  const dot    = document.getElementById('scannerDot');
  const status = document.getElementById('scannerStatusText');
  const zone   = document.getElementById('scannerZone');
  if (dot)    dot.classList.add('active');
  if (status) status.textContent = 'SCANNER ACTIVE';
  if (zone)   zone.classList.add('focused');
}

function onScannerBlur() {
  // Only update UI if focus didn't move to scanner-related elements
  setTimeout(() => {
    const active = document.activeElement;
    if (active && active.id === 'barcodeInput') return; // refocused immediately
    const dot    = document.getElementById('scannerDot');
    const status = document.getElementById('scannerStatusText');
    const zone   = document.getElementById('scannerZone');
    if (dot)    dot.classList.remove('active');
    if (status) status.textContent = 'SCANNER READY';
    if (zone)   zone.classList.remove('focused');
  }, 100);
}

function onBarcodeInput() {
  // Show typing feedback while barcode chars are coming in
  const val    = document.getElementById('barcodeInput').value;
  const status = document.getElementById('scannerStatusText');
  if (status && val) status.textContent = `READING… ${val}`;
}

function filterScanEmployees() { showEmpDropdown(); }

function showEmpDropdown() {
  const query     = document.getElementById('scan-emp-search').value.toLowerCase();
  const dd        = document.getElementById('scan-emp-dropdown');
  const employees = DB.employees;
  const filtered  = query ? employees.filter(e =>
    `${e.firstName} ${e.lastName}`.toLowerCase().includes(query) ||
    (e.employeeId||'').toLowerCase().includes(query) ||
    (e.productionCentre||'').toLowerCase().includes(query)
  ) : employees;

  dd.innerHTML = !filtered.length
    ? '<div class="emp-dropdown-empty">No employees found</div>'
    : filtered.slice(0,8).map(e => `
      <div class="emp-dropdown-item" onclick="selectEmployee('${e.id}')">
        <div class="emp-name-d">${escHtml(e.firstName)} ${escHtml(e.lastName)}</div>
        <div class="emp-meta-d">${escHtml(e.employeeId||'')}${e.productionCentre?' · '+escHtml(e.productionCentre):''}</div>
      </div>`).join('');
  dd.classList.remove('hidden');
}

function selectEmployee(id) {
  const e = DB.employees.find(emp => emp.id === id);
  if (!e) return;
  selectedEmployee = e;
  document.getElementById('scan-emp-search').value = `${e.firstName} ${e.lastName}`;
  document.getElementById('scan-emp-dropdown').classList.add('hidden');
  document.getElementById('clearEmpBtn').style.display = 'block';
  const badge = document.getElementById('selected-employee-badge');
  const color = avatarColor(e.firstName + e.lastName);
  badge.innerHTML = `
    <div class="sel-emp-avatar" style="background:${color}">${escHtml(initials(e.firstName,e.lastName))}</div>
    <div>
      <div class="sel-emp-name">${escHtml(e.firstName)} ${escHtml(e.lastName)}</div>
      <div class="sel-emp-meta">${escHtml(e.employeeId||'')}${e.productionCentre?' · '+escHtml(e.productionCentre):''}</div>
    </div>`;
  badge.classList.remove('hidden');
  // Immediately return focus to barcode scanner
  setTimeout(() => document.getElementById('barcodeInput').focus(), 30);
}

function clearEmployee() {
  selectedEmployee = null;
  document.getElementById('scan-emp-search').value = '';
  document.getElementById('clearEmpBtn').style.display = 'none';
  document.getElementById('selected-employee-badge').classList.add('hidden');
  document.getElementById('scan-emp-dropdown').classList.add('hidden');
}

function handleBarcodeKey(e) { if (e.key === 'Enter') { e.preventDefault(); submitScan(); } }

// Check if barcode was already logged with the same action on the same date
function checkDuplicate(barcode, action, date) {
  const txns = DB.transactions;
  return txns.find(t =>
    t.barcode === barcode &&
    t.action  === action  &&
    t.date    === date
  ) || null;
}

// Flash the scanner zone RED (duplicate warning)
function flashDuplicateWarning(barcode) {
  const zone   = document.getElementById('scannerZone');
  const dot    = document.getElementById('scannerDot');
  const status = document.getElementById('scannerStatusText');
  if (zone)   { zone.classList.add('dup-warning');   setTimeout(() => zone.classList.remove('dup-warning'),   1500); }
  if (dot)    { dot.classList.add('dup-dot');         setTimeout(() => dot.classList.remove('dup-dot'),         1500); }
  if (status) {
    status.textContent = `⚠ DUPLICATE: ${barcode}`;
    setTimeout(() => { if (status) status.textContent = 'SCANNER ACTIVE'; }, 1600);
  }
}

function submitScan() {
  const barcode = document.getElementById('barcodeInput').value.trim();
  if (!barcode) { showToast('Please scan or enter a barcode.', 'error'); return; }

  const needsEmployee = currentAction === 'distributed' || currentAction === 'returned';
  if (currentAction === 'distributed' && !selectedEmployee) {
    showToast('Please select an employee first.', 'error');
    document.getElementById('scan-emp-search').focus();
    return;
  }

  const date = document.getElementById('scan-date').value || today();

  // ---- Duplicate check ----
  const dup = checkDuplicate(barcode, currentAction, date);
  if (dup) {
    flashDuplicateWarning(barcode);
    const dupEmp = dup.employeeName ? ` (${dup.employeeName})` : '';
    confirm_dialog(
      `⚠ Duplicate Barcode Detected`,
      `Barcode "${barcode}" was already logged as "${actionLabel(currentAction)}"${dupEmp} on ${formatDate(date)}. Save it again anyway?`,
      () => doSaveScan(barcode, currentAction, date, '')   // user confirmed
    );
    // Clear input so scanner is ready for next item
    document.getElementById('barcodeInput').value = '';
    setTimeout(() => document.getElementById('barcodeInput').focus(), 20);
    return;
  }

  doSaveScan(barcode, currentAction, date, '');
}

function doSaveScan(barcode, action, date, notes) {
  const needsEmployee = action === 'distributed' || action === 'returned';
  const txn = {
    id: uid(), barcode, action,
    employeeId:   (needsEmployee && selectedEmployee) ? selectedEmployee.id   : null,
    employeeName: (needsEmployee && selectedEmployee) ? `${selectedEmployee.firstName} ${selectedEmployee.lastName}` : null,
    date, notes,
    createdAt: new Date().toISOString()
  };

  DB.addTransaction(txn);
  CloudSync.pushTransaction(txn);
  sessionScans++;

  // Flash scanner zone green
  const wrapper = document.getElementById('barcodeWrapper');
  const zone    = document.getElementById('scannerZone');
  const dot     = document.getElementById('scannerDot');
  const status  = document.getElementById('scannerStatusText');
  const lastVal = document.getElementById('scannerLastValue');

  if (wrapper) { wrapper.classList.add('flash'); setTimeout(() => wrapper.classList.remove('flash'), 600); }
  if (zone)    { zone.classList.add('scan-success'); setTimeout(() => zone.classList.remove('scan-success'), 600); }
  if (dot)     { dot.classList.add('success'); setTimeout(() => dot.classList.remove('success'), 600); }
  if (status)  { status.textContent = '✓ LOGGED'; setTimeout(() => { if (status) status.textContent = 'SCANNER ACTIVE'; }, 700); }
  if (lastVal) lastVal.textContent = barcode;

  updateScanFeed(txn);
  document.getElementById('session-count').textContent = `${sessionScans} scan${sessionScans!==1?'s':''}`;
  document.getElementById('barcodeInput').value = '';
  showToast(`✓ ${barcode} — ${actionLabel(action)}`, 'success');
  setTimeout(() => document.getElementById('barcodeInput').focus(), 20);
}

// ---- Manual entry fallback ----
function openManualEntry() {
  document.getElementById('manual-barcode').value = '';
  const mn = document.getElementById('manual-notes');
  if (mn) mn.value = '';
  openModal('manual-entry-modal');
  setTimeout(() => document.getElementById('manual-barcode').focus(), 200);
}

function manualSubmit() {
  const barcode = document.getElementById('manual-barcode').value.trim();
  if (!barcode) { showToast('Please enter a barcode.', 'error'); return; }

  if (currentAction === 'distributed' && !selectedEmployee) {
    showToast('Please select an employee first.', 'error');
    closeModal('manual-entry-modal');
    document.getElementById('scan-emp-search').focus();
    return;
  }

  const date  = document.getElementById('scan-date').value || today();
  const notes = (document.getElementById('manual-notes')?.value || '').trim();

  // ---- Duplicate check ----
  const dup = checkDuplicate(barcode, currentAction, date);
  if (dup) {
    const dupEmp = dup.employeeName ? ` (${dup.employeeName})` : '';
    confirm_dialog(
      `⚠ Duplicate Barcode Detected`,
      `Barcode "${barcode}" was already logged as "${actionLabel(currentAction)}"${dupEmp} on ${formatDate(date)}. Save it again anyway?`,
      () => {
        doSaveScan(barcode, currentAction, date, notes);
        document.getElementById('manual-barcode').value = '';
        if (document.getElementById('manual-notes')) document.getElementById('manual-notes').value = '';
        setTimeout(() => document.getElementById('manual-barcode').focus(), 80);
      }
    );
    return;
  }

  doSaveScan(barcode, currentAction, date, notes);
  document.getElementById('manual-barcode').value = '';
  if (document.getElementById('manual-notes')) document.getElementById('manual-notes').value = '';
  setTimeout(() => document.getElementById('manual-barcode').focus(), 80);
}

function updateScanFeed(txn) {
  const feed    = document.getElementById('scan-feed');
  const empStr  = txn.employeeName || (txn.action === 'reported_issue' ? 'Issue → Cintas' : txn.action.includes('cintas') ? 'Cintas' : 'Warehouse');
  const empty   = feed.querySelector('.empty-state');
  if (empty) empty.remove();
  const item = document.createElement('div');
  item.className = 'scan-feed-item new';
  item.innerHTML = `
    <span class="scan-num">${sessionScans}</span>
    <div><div class="scan-bc">${escHtml(txn.barcode)}</div><div class="scan-emp">${escHtml(empStr)}</div></div>
    <span class="scan-action-pill ${txn.action}">${actionLabel(txn.action)}</span>
    <span class="scan-time">${formatDateTime(txn.createdAt)}</span>`;
  feed.insertBefore(item, feed.firstChild);
  setTimeout(() => item.classList.remove('new'), 800);
}

// ============================================================
// REPORTS
// ============================================================
function setReport(r) {
  currentReport = r;
  document.querySelectorAll('.report-tab').forEach(t => t.classList.remove('active'));
  const tab = document.querySelector(`.report-tab[data-report="${r}"]`);
  if (tab) tab.classList.add('active');
  renderReport();
}

function renderReport() {
  const from      = document.getElementById('report-from').value;
  const to        = document.getElementById('report-to').value;
  const q         = (document.getElementById('report-search').value || '').toLowerCase();
  const container = document.getElementById('report-content');
  const txns      = DB.transactions;
  const employees = DB.employees;

  function latestTxn(barcode) { return txns.find(t => t.barcode === barcode); }
  function inRange(dateStr) {
    if (!dateStr) return true;
    if (from && dateStr < from) return false;
    if (to   && dateStr > to)   return false;
    return true;
  }

  // ---- MISSING UNIFORMS (core report) ----
  if (currentReport === 'missing') {
    const barcodes = [...new Set(txns.map(t => t.barcode))];
    const missingItems = [];

    barcodes.forEach(bc => {
      const lastTxn = latestTxn(bc);
      if (!lastTxn || lastTxn.action !== 'distributed') return;
      const emp = employees.find(e => e.id === lastTxn.employeeId);
      const row = {
        barcode: bc,
        employeeId:    lastTxn.employeeId,
        employeeName:  emp ? `${emp.firstName} ${emp.lastName}` : '— Unknown —',
        employeeEmpId: emp?.employeeId || '',
        centre:        emp?.productionCentre || '—',
        department:    emp?.department || '—',
        date:          lastTxn.date,
        notes:         lastTxn.notes || '',
        days:          Math.floor((Date.now() - new Date(lastTxn.date)) / 86400000)
      };
      if (from && lastTxn.date < from) return;
      if (to   && lastTxn.date > to)   return;
      if (q && !bc.toLowerCase().includes(q) &&
               !row.employeeName.toLowerCase().includes(q) &&
               !row.centre.toLowerCase().includes(q) &&
               !row.employeeEmpId.toLowerCase().includes(q)) return;
      missingItems.push(row);
    });

    missingItems.sort((a,b) => b.days - a.days);

    // Group by employee
    const byEmp = {};
    missingItems.forEach(r => {
      const key = r.employeeId || '__unknown__';
      if (!byEmp[key]) byEmp[key] = { name: r.employeeName, empId: r.employeeEmpId, centre: r.centre, dept: r.department, items: [] };
      byEmp[key].items.push(r);
    });

    const urgentCount = missingItems.filter(r => r.days > 7).length;

    container.innerHTML = `
      <div class="missing-banner ${missingItems.length > 0 ? 'has-missing' : 'all-clear'}">
        <div class="missing-banner-icon">${missingItems.length > 0 ? '⚠' : '✓'}</div>
        <div>
          <div class="missing-banner-title">${missingItems.length > 0 ? missingItems.length + ' Uniform' + (missingItems.length !== 1 ? 's' : '') + ' Not Returned' : 'All Uniforms Returned!'}</div>
          <div class="missing-banner-sub">${missingItems.length > 0 ? urgentCount + ' item' + (urgentCount !== 1 ? 's' : '') + ' overdue >7 days &middot; Grouped by employee below' : 'No outstanding uniforms at this time.'}</div>
        </div>
        <button class="btn-secondary" style="margin-left:auto" onclick="exportCSV()">Export CSV</button>
      </div>
      ${Object.values(byEmp).map(grp => `
        <div class="missing-emp-group">
          <div class="missing-emp-header">
            <div class="emp-avatar" style="width:36px;height:36px;font-size:0.82rem;background:${avatarColor(grp.name)}">${grp.name.split(' ').map(w=>w[0]).join('').slice(0,2)}</div>
            <div>
              <div style="font-weight:700;color:var(--text)">${escHtml(grp.name)} <span style="color:var(--text3);font-size:0.78rem;font-weight:400">${escHtml(grp.empId ? '&middot; ' + grp.empId : '')}</span></div>
              <div style="font-size:0.76rem;color:var(--text3)">${escHtml(grp.centre)}${grp.dept ? ' &middot; ' + escHtml(grp.dept) : ''}</div>
            </div>
            <span class="missing-count-badge">${grp.items.length} item${grp.items.length !== 1 ? 's' : ''}</span>
          </div>
          <div class="report-table-wrap" style="margin:0;border-top:none;border-radius:0 0 10px 10px">
            <table class="report-table">
              <thead><tr><th>Barcode</th><th>Date Out</th><th>Days Out</th><th>Notes</th></tr></thead>
              <tbody>
                ${grp.items.map(r => `<tr>
                  <td class="bc-mono">${escHtml(r.barcode)}</td>
                  <td>${formatDate(r.date)}</td>
                  <td><span class="days-pill ${r.days > 14 ? 'overdue' : r.days > 7 ? 'warning' : 'ok'}">${r.days}d</span></td>
                  <td style="color:var(--text3)">${escHtml(r.notes)}</td>
                </tr>`).join('')}
              </tbody>
            </table>
          </div>
        </div>
      `).join('')}
      ${missingItems.length === 0 ? '<div class="empty-state"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M22 11.08V12a10 10 0 11-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/></svg><p>No missing uniforms found.</p></div>' : ''}
    `;

  // ---- REPORTED ISSUES ----
  } else if (currentReport === 'issues') {
    const barcodes = [...new Set(txns.map(t => t.barcode))];
    const rows = [];
    barcodes.forEach(bc => {
      const last = latestTxn(bc);
      if (!last || last.action !== 'reported_issue') return;
      if (!inRange(last.date)) return;
      if (q && !bc.toLowerCase().includes(q)) return;
      rows.push({ barcode: bc, date: last.date, notes: last.notes || '' });
    });
    container.innerHTML = `
      <div class="report-summary">
        <div class="rs-card"><div class="rs-num" style="color:var(--yellow)">${rows.length}</div><div class="rs-label">Items Reported Damaged/Dirty</div></div>
      </div>
      <div class="report-table-wrap"><table class="report-table">
        <thead><tr><th>Barcode</th><th>Date Reported</th><th>Notes</th></tr></thead>
        <tbody>${rows.length ? rows.map(r => `<tr>
          <td class="bc-mono">${escHtml(r.barcode)}</td>
          <td>${formatDate(r.date)}</td>
          <td>${escHtml(r.notes)}</td>
        </tr>`).join('') : '<tr><td colspan="3" style="text-align:center;color:var(--text3);padding:30px">No issues reported.</td></tr>'}</tbody>
      </table></div>`;

  // ---- BY EMPLOYEE ----
  } else if (currentReport === 'employee-summary') {
    const empList = employees.filter(e => !q || `${e.firstName} ${e.lastName}`.toLowerCase().includes(q) || (e.employeeId||'').toLowerCase().includes(q));
    const rows = empList.map(e => {
      const myTxns     = txns.filter(t => t.employeeId === e.id);
      const myBarcodes = [...new Set(myTxns.map(t => t.barcode))];
      const held = myBarcodes.filter(bc => { const l = latestTxn(bc); return l && l.action === 'distributed'; }).length;
      return { e, held, total: myBarcodes.length };
    }).filter(r => r.total > 0 || !q);

    container.innerHTML = `
      <div class="report-summary">
        <div class="rs-card"><div class="rs-num">${employees.length}</div><div class="rs-label">Total Employees</div></div>
        <div class="rs-card"><div class="rs-num" style="color:var(--orange)">${rows.reduce((a,r)=>a+r.held,0)}</div><div class="rs-label">Currently Out</div></div>
      </div>
      <div class="report-table-wrap"><table class="report-table">
        <thead><tr><th>Employee</th><th>ID</th><th>Centre</th><th>Currently Holding</th><th>Total Transacted</th></tr></thead>
        <tbody>${rows.length ? rows.map(r => `<tr>
          <td>${escHtml(r.e.firstName)} ${escHtml(r.e.lastName)}</td>
          <td class="bc-mono">${escHtml(r.e.employeeId||'—')}</td>
          <td>${escHtml(r.e.productionCentre||'—')}</td>
          <td><strong style="color:var(--orange)">${r.held}</strong></td>
          <td>${r.total}</td>
        </tr>`).join('') : '<tr><td colspan="5" style="text-align:center;color:var(--text3);padding:30px">No data found.</td></tr>'}</tbody>
      </table></div>`;

  // ---- AT CINTAS ----
  } else if (currentReport === 'cintas') {
    const barcodes = [...new Set(txns.map(t => t.barcode))];
    const rows = [];
    barcodes.forEach(bc => {
      const last = latestTxn(bc);
      if (!last || (last.action !== 'sent_to_cintas' && last.action !== 'reported_issue')) return;
      if (!inRange(last.date)) return;
      if (q && !bc.toLowerCase().includes(q)) return;
      rows.push({ barcode: bc, date: last.date, how: last.action, days: Math.floor((Date.now()-new Date(last.date))/86400000), notes: last.notes||'' });
    });
    container.innerHTML = `
      <div class="report-summary">
        <div class="rs-card"><div class="rs-num" style="color:var(--purple)">${rows.length}</div><div class="rs-label">At Cintas</div></div>
        <div class="rs-card"><div class="rs-num" style="color:var(--yellow)">${rows.filter(r=>r.how==='reported_issue').length}</div><div class="rs-label">Reported Issues</div></div>
      </div>
      <div class="report-table-wrap"><table class="report-table">
        <thead><tr><th>Barcode</th><th>Reason</th><th>Date</th><th>Days at Cintas</th><th>Notes</th></tr></thead>
        <tbody>${rows.length ? rows.map(r => `<tr>
          <td class="bc-mono">${escHtml(r.barcode)}</td>
          <td><span class="report-badge ${r.how==='reported_issue'?'issue':'cintas'}">${r.how==='reported_issue'?'Issue/Damaged':'Sent for Cleaning'}</span></td>
          <td>${formatDate(r.date)}</td>
          <td><span style="color:${r.days>14?'var(--red)':'var(--text2)'}">${r.days}d</span></td>
          <td>${escHtml(r.notes)}</td>
        </tr>`).join('') : '<tr><td colspan="5" style="text-align:center;color:var(--text3);padding:30px">No uniforms currently at Cintas.</td></tr>'}</tbody>
      </table></div>`;

  // ---- WAREHOUSE ----
  } else if (currentReport === 'warehouse') {
    const barcodes = [...new Set(txns.map(t => t.barcode))];
    const rows = [];
    barcodes.forEach(bc => {
      const last = latestTxn(bc);
      if (!last || (last.action !== 'returned' && last.action !== 'received_from_cintas')) return;
      if (!inRange(last.date)) return;
      if (q && !bc.toLowerCase().includes(q)) return;
      rows.push({ barcode: bc, date: last.date, how: last.action, notes: last.notes||'' });
    });
    container.innerHTML = `
      <div class="report-summary">
        <div class="rs-card"><div class="rs-num" style="color:var(--blue)">${rows.length}</div><div class="rs-label">In Warehouse</div></div>
        <div class="rs-card"><div class="rs-num" style="color:var(--green)">${rows.filter(r=>r.how==='returned').length}</div><div class="rs-label">Employee Returns</div></div>
        <div class="rs-card"><div class="rs-num">${rows.filter(r=>r.how==='received_from_cintas').length}</div><div class="rs-label">Received from Cintas</div></div>
      </div>
      <div class="report-table-wrap"><table class="report-table">
        <thead><tr><th>Barcode</th><th>Arrived</th><th>How</th><th>Notes</th></tr></thead>
        <tbody>${rows.length ? rows.map(r => `<tr>
          <td class="bc-mono">${escHtml(r.barcode)}</td>
          <td>${formatDate(r.date)}</td>
          <td><span class="report-badge ${r.how==='returned'?'returned':'warehouse'}">${r.how==='returned'?'Employee Return':'From Cintas'}</span></td>
          <td>${escHtml(r.notes)}</td>
        </tr>`).join('') : '<tr><td colspan="4" style="text-align:center;color:var(--text3);padding:30px">No uniforms currently in warehouse.</td></tr>'}</tbody>
      </table></div>`;

  // ---- BARCODE TIMELINE ----
  } else if (currentReport === 'barcode-history') {
    let filteredTxns = txns.filter(t => inRange(t.date));
    if (q) filteredTxns = filteredTxns.filter(t => t.barcode.toLowerCase().includes(q) || (t.employeeName||'').toLowerCase().includes(q));
    const canDelete = currentUser && currentUser.role === 'admin';
    container.innerHTML = `
      <div class="report-table-wrap"><table class="report-table">
        <thead><tr><th>Date</th><th>Barcode</th><th>Action</th><th>Employee / Location</th><th>Notes</th><th class="admin-only" style="${canDelete?'':'display:none'}">Action</th></tr></thead>
        <tbody>${filteredTxns.length ? filteredTxns.map(t => `<tr>
          <td>${formatDate(t.date)}</td>
          <td class="bc-mono">${escHtml(t.barcode)}</td>
          <td><span class="report-badge ${t.action==='distributed'?'out':t.action==='returned'?'returned':t.action==='reported_issue'?'issue':t.action==='sent_to_cintas'?'cintas':'warehouse'}">${actionLabel(t.action)}</span></td>
          <td>${escHtml(t.employeeName||(t.action==='reported_issue'?'Issue → Cintas':t.action.includes('cintas')?'Cintas':'Warehouse'))}</td>
          <td>${escHtml(t.notes||'')}</td>
          <td class="admin-only" style="${canDelete?'':'display:none'}"><button class="btn-text" style="color:var(--red)" onclick="deleteTransaction('${t.id}')">Delete</button></td>
        </tr>`).join('') : '<tr><td colspan="6" style="text-align:center;color:var(--text3);padding:30px">No transactions found.</td></tr>'}</tbody>
      </table></div>`;
  }
}

function deleteTransaction(id) {
  if (!currentUser || currentUser.role !== 'admin') return;
  confirm_dialog('Delete Transaction?', 'Are you sure you want to completely remove this barcode scan record from history? This cannot be undone.', () => {
    DB.removeTransaction(id);
    showToast('Transaction deleted.', 'success');
    renderReport();
    renderDashboard();
  });
}

function clearDates() {
  document.getElementById('report-from').value = '';
  document.getElementById('report-to').value   = '';
  renderReport();
}

// ============================================================
// CSV EXPORT
// ============================================================
function exportCSV() {
  const table = document.querySelector('#report-content .report-table');
  if (!table) { showToast('No data to export.', 'error'); return; }
  const rows    = [];
  const headers = [...table.querySelectorAll('thead th')].map(th => th.textContent.trim());
  rows.push(headers.join(','));
  table.querySelectorAll('tbody tr').forEach(tr => {
    const cols = [...tr.querySelectorAll('td')].map(td => `"${td.textContent.trim().replace(/"/g,'""')}"`);
    rows.push(cols.join(','));
  });
  const blob = new Blob([rows.join('\n')], { type:'text/csv' });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a'); a.href = url;
  a.download = `unitrack_${currentReport}_${today()}.csv`; a.click();
  URL.revokeObjectURL(url);
  showToast('CSV exported!', 'success');
}

// ============================================================
// SETTINGS
// ============================================================
function renderSettings() {
  renderTags('centre-tags', DB.centres, removeCentre);
  renderTags('type-tags',   DB.uniformTypes, removeType);
  // Cloud sync fields
  const urlEl = document.getElementById('cloud-url');
  const keyEl = document.getElementById('cloud-key');
  if (urlEl && !urlEl.value) urlEl.value = CloudSync.apiUrl;
  if (keyEl && !keyEl.value) keyEl.value = CloudSync.apiKey;
  updateSyncStatus();
  
  if (currentUser && currentUser.role === 'admin') {
    const uBody = document.getElementById('users-tbody');
    if (uBody) {
      if (!DB.users.length) { uBody.innerHTML = '<tr><td colspan="3" style="text-align:center;color:var(--text3)">No users found.</td></tr>'; }
      else {
        uBody.innerHTML = DB.users.map(u => `<tr>
          <td><strong>${escHtml(u.username)}</strong></td>
          <td>${u.role==='admin'?'<span style="color:var(--red);font-weight:600">Admin</span>':'Operator'}</td>
          <td style="text-align:right">
            <button class="btn-text" style="color:var(--red)" onclick="removeUser('${u.username}')">Remove</button>
          </td>
        </tr>`).join('');
      }
    }
  }
}

function addUser() {
  console.log("addUser called with currentUser:", currentUser);
  if (!currentUser || currentUser.role !== 'admin') {
    console.warn("addUser blocked: insufficient privileges or null session.");
    return;
  }
  const u = document.getElementById('new-user-name').value.trim();
  const p = document.getElementById('new-user-pass').value;
  const r = document.getElementById('new-user-role').value;
  if (!u || !p) return showToast('Username and password required.', 'error');
  if (DB.users.find(x => x.username === u)) return showToast('Username already exists.', 'error');
  
  DB.addUser({ username: u, password: p, role: r });
  showToast(`User ${u} created.`, 'success');
  document.getElementById('new-user-name').value = '';
  document.getElementById('new-user-pass').value = '';
  renderSettings();
}
window.addUser = addUser;

function removeUser(u) {
  if (!currentUser || currentUser.role !== 'admin') return;
  if (u === currentUser.username) return showToast('Cannot delete yourself!', 'error');
  if (u === 'admin') return showToast('Cannot delete master admin account!', 'error');
  confirm_dialog('Remove user?', `Are you sure you want to remove access for user ${u}?`, () => {
    DB.saveUsers(DB.users.filter(x => x.username !== u));
    showToast(`User ${u} removed.`, 'success');
    renderSettings();
  });
}
window.removeUser = removeUser;
function renderTags(id, items, removeFn) {
  const el = document.getElementById(id);
  if (!items.length) { el.innerHTML = '<span style="color:var(--text3);font-size:0.8rem">No items.</span>'; return; }
  el.innerHTML = items.map(item => `<span class="tag">${escHtml(item)}<button onclick="${removeFn.name}('${escHtml(item)}')">✕</button></span>`).join('');
}
function addCentre() {
  const val = document.getElementById('new-centre').value.trim();
  if (!val) return;
  const list = DB.centres;
  if (list.includes(val)) { showToast('Already exists.', 'error'); return; }
  list.push(val); DB.saveCentres(list);
  document.getElementById('new-centre').value = '';
  renderSettings(); showToast('Centre added!', 'success');
}
function removeCentre(name) { DB.saveCentres(DB.centres.filter(c => c !== name)); renderSettings(); }
function addUniformType() {
  const val = document.getElementById('new-type').value.trim();
  if (!val) return;
  const list = DB.uniformTypes;
  if (list.includes(val)) { showToast('Already exists.', 'error'); return; }
  list.push(val); DB.saveTypes(list);
  document.getElementById('new-type').value = '';
  renderSettings(); showToast('Type added!', 'success');
}
function removeType(name) { DB.saveTypes(DB.uniformTypes.filter(t => t !== name)); renderSettings(); }

// ============================================================
// JSON BACKUP
// ============================================================
function exportJSON() {
  const data = { version:1, exported:new Date().toISOString(), employees:DB.employees, transactions:DB.transactions, centres:DB.centres, uniformTypes:DB.uniformTypes };
  const blob = new Blob([JSON.stringify(data,null,2)],{type:'application/json'});
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a'); a.href = url;
  a.download = `unitrack_backup_${today()}.json`; a.click();
  URL.revokeObjectURL(url);
  showToast('Backup exported!', 'success');
}
function importJSON(event) {
  const file = event.target.files[0]; if (!file) return;
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const data = JSON.parse(e.target.result);
      confirm_dialog('Import Backup?','This will REPLACE all current data. Are you sure?', async () => {
        if (data.employees)    DB.saveEmployees(data.employees);
        if (data.centres)      DB.saveCentres(data.centres);
        if (data.uniformTypes) DB.saveTypes(data.uniformTypes);
        
        if (data.transactions) {
          await IDB.clearTransactions();
          DB_MEMORY.transactions = data.transactions;
          for (const t of data.transactions) {
            await IDB.saveTransaction(t);
          }
        }
        
        showToast('Backup imported!', 'success');
        renderSettings(); renderDashboard();
      });
    } catch { showToast('Invalid backup file.', 'error'); }
  };
  reader.readAsText(file);
  event.target.value = '';
}
function clearAllData() {
  confirm_dialog('Clear ALL Data?','Permanently delete all employees, transactions, and settings. This CANNOT be undone!', async () => {
    // Clear IndexedDB
    const stores = ['store','transactions'];
    for (const s of stores) {
      const tx = IDB.db.transaction(s, 'readwrite');
      tx.objectStore(s).clear();
    }
    // Clear memory
    DB_MEMORY.employees = [];
    DB_MEMORY.transactions = [];
    // Users are NOT cleared to prevent lockout, 
    // but settings (centres/types) are reset to defaults
    DB_MEMORY.centres = ["Main Plant","Warehouse A","Warehouse B","Office"];
    DB_MEMORY.uniformTypes = ["Shirt","Pants","Jacket","Safety Vest","Cap","Gloves"];
    
    // Also clear localStorage if anything remained
    localStorage.clear();
    
    showToast('All data cleared.', 'success');
    renderDashboard(); renderEmployeeList(); renderSettings();
  });
}

// ============================================================
// EXCEL / CSV IMPORT
// ============================================================
const FIELD_ALIASES = {
  firstName:        ['first name','firstname','first','given name','prénom','prenom','name'],
  lastName:         ['last name','lastname','last','surname','family name','nom'],
  employeeId:       ['employee id','emp id','id','badge','badge number','badge no','employee no','emp no'],
  productionCentre: ['production centre','production center','centre','center','plant','site','location'],
  department:       ['department','dept','division','team'],
  phone:            ['phone','telephone','tel','mobile','cell'],
  notes:            ['notes','remarks','comments','note']
};

function openExcelImport() {
  xlRawRows = []; xlHeaders = []; xlMapping = {};
  document.getElementById('xl-step1').style.display = 'block';
  document.getElementById('xl-step2').style.display = 'none';
  document.getElementById('xl-step3').style.display = 'none';
  document.getElementById('xl-next-btn').style.display   = 'none';
  document.getElementById('xl-import-btn').style.display = 'none';
  document.getElementById('xl-file-input').value = '';
  openModal('excel-import-modal');
}

function handleExcelFile(event) {
  const file = event.target.files[0];
  if (!file) return;
  const ext = file.name.split('.').pop().toLowerCase();
  const reader = new FileReader();

  reader.onload = (e) => {
    try {
      let rows = [];
      if (ext === 'csv') {
        const text = e.target.result;
        rows = parseCSV(text);
      } else {
        const workbook = XLSX.read(e.target.result, { type: 'array' });
        const sheet    = workbook.Sheets[workbook.SheetNames[0]];
        rows           = XLSX.utils.sheet_to_json(sheet, { header:1, defval:'' });
      }

      if (rows.length < 2) { showToast('File appears to be empty.', 'error'); return; }

      xlHeaders = rows[0].map(h => String(h).trim());
      xlRawRows = rows.slice(1).filter(r => r.some(cell => String(cell).trim()));

      // Auto-detect column mapping
      xlMapping = {};
      xlHeaders.forEach((h, idx) => {
        const hl = h.toLowerCase();
        for (const [field, aliases] of Object.entries(FIELD_ALIASES)) {
          if (!xlMapping[field] && aliases.some(a => hl.includes(a))) {
            xlMapping[field] = idx;
          }
        }
      });

      document.getElementById('xl-file-info').innerHTML = `
        <div class="xl-tip" style="background:rgba(63,185,80,0.08);border-color:rgba(63,185,80,0.3)">
          <svg viewBox="0 0 24 24" fill="none" stroke="var(--green)" stroke-width="2"><polyline points="20 6 9 17 4 12"/></svg>
          <div><strong>${file.name}</strong> — ${xlRawRows.length} rows detected &middot; ${xlHeaders.length} columns</div>
        </div>`;

      renderMappingGrid();
      document.getElementById('xl-step1').style.display = 'none';
      document.getElementById('xl-step2').style.display = 'block';
      document.getElementById('xl-next-btn').style.display = 'inline-flex';
      showToast('File loaded. Review the column mapping below.', 'success');
    } catch (err) {
      showToast('Could not read file: ' + err.message, 'error');
    }
  };

  if (ext === 'csv') reader.readAsText(file);
  else               reader.readAsArrayBuffer(file);
}

function parseCSV(text) {
  return text.split('\n').map(line => {
    const row = []; let cur = ''; let inQ = false;
    for (let i = 0; i < line.length; i++) {
      const c = line[i];
      if (c === '"') { inQ = !inQ; }
      else if (c === ',' && !inQ) { row.push(cur.trim()); cur = ''; }
      else { cur += c; }
    }
    row.push(cur.trim());
    return row;
  }).filter(r => r.length > 0 && r.some(c => c));
}

function renderMappingGrid() {
  const fields = [
    { key:'firstName',        label:'First Name', required:true  },
    { key:'lastName',         label:'Last Name',  required:true  },
    { key:'employeeId',       label:'Employee ID',required:false },
    { key:'productionCentre', label:'Production Centre', required:false },
    { key:'department',       label:'Department', required:false },
    { key:'phone',            label:'Phone',      required:false },
    { key:'notes',            label:'Notes',      required:false }
  ];
  const options = ['(skip)', ...xlHeaders].map((h, i) =>
    `<option value="${i-1}" ${xlMapping[fields[0]?.key] === i-1 ? '' : ''}>${i === 0 ? '— Skip this field —' : `[Col ${i}] ${h}`}</option>`
  );

  document.getElementById('xl-mapping-grid').innerHTML = fields.map(f => {
    const sel = ['<option value="-1">— Skip this field —</option>',
      ...xlHeaders.map((h, idx) => `<option value="${idx}" ${xlMapping[f.key] === idx ? 'selected' : ''}>[Col ${idx+1}] ${h}</option>`)
    ].join('');
    return `<div class="xl-map-row">
      <div class="xl-map-label">${f.label}${f.required ? ' <span style="color:var(--red)">*</span>' : ''}</div>
      <select class="field-input xl-map-select" id="xlmap-${f.key}" onchange="xlMapping['${f.key}'] = parseInt(this.value)">
        ${sel}
      </select>
    </div>`;
  }).join('');
}

function xlNext() {
  // Validate required fields
  if (xlMapping['firstName'] === undefined || xlMapping['firstName'] < 0 ||
      xlMapping['lastName']  === undefined || xlMapping['lastName']  < 0) {
    showToast('Please map First Name and Last Name columns.', 'error');
    return;
  }
  // Read current select values
  for (const key of Object.keys(FIELD_ALIASES)) {
    const el = document.getElementById('xlmap-' + key);
    if (el) xlMapping[key] = parseInt(el.value);
  }
  renderPreview();
  document.getElementById('xl-step2').style.display = 'none';
  document.getElementById('xl-step3').style.display = 'block';
  document.getElementById('xl-next-btn').style.display   = 'none';
  document.getElementById('xl-import-btn').style.display = 'inline-flex';
}

function renderPreview() {
  const get = (row, key) => xlMapping[key] >= 0 ? String(row[xlMapping[key]] || '').trim() : '';
  const preview = xlRawRows.slice(0,20);
  document.getElementById('xl-preview-info').innerHTML = `
    <div class="xl-tip">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>
      <div>Showing first ${Math.min(20,xlRawRows.length)} of <strong>${xlRawRows.length} employees</strong> to be imported. Review and click <strong>Import Employees</strong>.</div>
    </div>`;
  document.getElementById('xl-preview-body').innerHTML = preview.map(row => `
    <tr>
      <td>${escHtml(get(row,'firstName'))}</td>
      <td>${escHtml(get(row,'lastName'))}</td>
      <td class="bc-mono">${escHtml(get(row,'employeeId'))}</td>
      <td>${escHtml(get(row,'productionCentre'))}</td>
      <td>${escHtml(get(row,'department'))}</td>
      <td>${escHtml(get(row,'phone'))}</td>
    </tr>`).join('');
}

function xlDoImport() {
  const get = (row, key) => xlMapping[key] >= 0 ? String(row[xlMapping[key]] || '').trim() : '';
  const existing = DB.employees;
  let added = 0, skipped = 0;

  xlRawRows.forEach(row => {
    const first = get(row, 'firstName');
    const last  = get(row, 'lastName');
    if (!first || !last) { skipped++; return; }
    const emp = {
      id:               uid(),
      firstName:        first,
      lastName:         last,
      employeeId:       get(row,'employeeId'),
      productionCentre: get(row,'productionCentre'),
      department:       get(row,'department'),
      phone:            get(row,'phone'),
      notes:            get(row,'notes'),
      createdAt:        new Date().toISOString()
    };
    existing.push(emp);
    added++;
  });

  DB.saveEmployees(existing);
  closeModal('excel-import-modal');
  showToast(`✓ ${added} employees imported${skipped ? `, ${skipped} skipped (missing name)` : ''}!`, 'success');
  renderEmployeeList();
}

function downloadTemplate() {
  const csv = 'First Name,Last Name,Employee ID,Production Centre,Department,Phone,Notes\nJohn,Smith,EMP-001,Main Plant,Packaging,555-1234,\nJane,Doe,EMP-002,Warehouse A,Shipping,,';
  const blob = new Blob([csv], {type:'text/csv'});
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a'); a.href = url;
  a.download = 'unitrack_employees_template.csv'; a.click();
  URL.revokeObjectURL(url);
  showToast('Template downloaded!', 'success');
}

// ============================================================
// MODAL HELPERS
// ============================================================
function openModal(id) {
  const el = document.getElementById(id);
  el.style.display = 'flex';
  setTimeout(() => el.classList.add('open'), 10);
}
function closeModal(id) {
  const el = document.getElementById(id);
  el.classList.remove('open');
  setTimeout(() => { el.style.display = 'none'; }, 200);
}
function closeModalIfBg(e, id) { if (e.target === e.currentTarget) closeModal(id); }

function confirm_dialog(title, message, onConfirm) {
  document.getElementById('confirm-title').textContent   = title;
  document.getElementById('confirm-message').textContent = message;
  const overlay   = document.getElementById('confirm-overlay');
  overlay.style.display = 'flex';
  setTimeout(() => overlay.classList.add('open'), 10);
  const cleanup = () => { overlay.classList.remove('open'); setTimeout(() => { overlay.style.display='none'; }, 200); };
  document.getElementById('confirm-ok').onclick     = () => { cleanup(); onConfirm(); };
  document.getElementById('confirm-cancel').onclick = cleanup;
}

// ============================================================
// TOAST
// ============================================================
let toastTimer;
function showToast(msg, type = 'success') {
  const toast = document.getElementById('toast');
  const icon  = type === 'success' ? '✓' : '⚠';
  toast.innerHTML = `<span class="toast-icon">${icon}</span> ${escHtml(msg)}`;
  toast.className = `toast ${type} show`;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => toast.classList.remove('show'), 3200);
}

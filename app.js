/* ============================================================
   UniTrack v3 — Application Logic (app.js)
   Intelligent Uniform Lifecycle Tracking
   State Machine + Inference Engine
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
      const request = indexedDB.open('UniTrackDB', 3);
      request.onupgradeneeded = (e) => {
        const db = e.target.result;
        if (!db.objectStoreNames.contains('store')) {
          db.createObjectStore('store');
        }
        if (!db.objectStoreNames.contains('transactions')) {
          db.createObjectStore('transactions', { keyPath: 'id' });
        }
        if (!db.objectStoreNames.contains('syncQueue')) {
          db.createObjectStore('syncQueue', { keyPath: 'queueId' });
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
      if (!txn.id) {
        txn.id = uid();
      }
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
  },
  async saveTransactionsBulk(txns) {
    // Write in chunks with a yield between each batch to prevent UI freeze
    const CHUNK = 500;
    
    // First, clear the store
    await new Promise((resolve, reject) => {
      const tx = this.db.transaction('transactions', 'readwrite');
      tx.objectStore('transactions').clear();
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
    });

    // Then write chunks, creating a new IDB transaction for each chunk
    // because IndexedDB transactions automatically close when yielding to setTimeout
    for (let i = 0; i < txns.length; i += CHUNK) {
      await new Promise((resolve, reject) => {
        const tx = this.db.transaction('transactions', 'readwrite');
        const store = tx.objectStore('transactions');
        const end = Math.min(i + CHUNK, txns.length);
        for (let j = i; j < end; j++) {
          if (!txns[j].id) txns[j].id = uid();
          store.put(txns[j]);
        }
        tx.oncomplete = () => resolve();
        tx.onerror = () => reject(tx.error);
      });
      // Yield to the browser
      await new Promise(r => setTimeout(r, 0));
    }
  },
  getSyncQueue() {
    return new Promise((resolve, reject) => {
      const tx = this.db.transaction('syncQueue', 'readonly');
      const req = tx.objectStore('syncQueue').getAll();
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
  },
  addToSyncQueue(item) {
    return new Promise((resolve, reject) => {
      const tx = this.db.transaction('syncQueue', 'readwrite');
      if (!item.queueId) item.queueId = uid();
      const req = tx.objectStore('syncQueue').put(item);
      req.onsuccess = () => resolve();
      req.onerror = () => reject(req.error);
    });
  },
  removeFromSyncQueue(queueId) {
    return new Promise((resolve, reject) => {
      const tx = this.db.transaction('syncQueue', 'readwrite');
      const req = tx.objectStore('syncQueue').delete(queueId);
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
    IDB.saveTransaction(t).catch(err => {
      console.error('Failed to save transaction to IndexedDB:', err);
      showToast('⚠ Transaction saved locally but may not persist. Check storage.', 'warning');
    });
  },
  removeTransaction(id) {
    DB_MEMORY.transactions = DB_MEMORY.transactions.filter(t => t.id !== id);
    IDB.deleteTransaction(id);
  },
  addEmployee(e)    { const list = this.employees; list.push(e); this.saveEmployees(list); }
};

// ============================================================
// STATE MACHINE ENGINE
// ============================================================

// Maps action to the resulting state after that action
const ACTION_TO_STATE = {
  received_from_cintas:   'warehouse',
  distributed:            'with_employee',
  returned:               'warehouse',
  collected_from_soil_bin:'soil_bin',
  sent_to_cintas:         'at_cintas',
  reported_issue:         'damaged'
};

// Valid transitions: from state → allowed actions
// NOTE: Some transitions are intentionally OMITTED so inference logic fires:
//   - with_employee → collected_from_soil_bin (must infer 'returned' first)
//   - damaged → received_from_cintas (must infer 'sent_to_cintas' first)
const VALID_TRANSITIONS = {
  'unknown':        ['received_from_cintas', 'distributed', 'returned', 'collected_from_soil_bin', 'sent_to_cintas', 'reported_issue'],
  'warehouse':      ['distributed', 'reported_issue', 'sent_to_cintas', 'collected_from_soil_bin'],
  'with_employee':  ['returned', 'reported_issue'],
  'soil_bin':       ['sent_to_cintas'],
  'at_cintas':      ['received_from_cintas'],
  'damaged':        ['sent_to_cintas']
};

// Resolve the current state of a barcode based on its full transaction history
function resolveState(barcode) {
  const txns = DB.transactions
    .filter(t => t.barcode === barcode)
    .sort((a, b) => {
      // Sort by date first, then by createdAt for same-date items
      const dateComp = (a.date || '').localeCompare(b.date || '');
      if (dateComp !== 0) return dateComp;
      return new Date(a.createdAt) - new Date(b.createdAt);
    });

  if (!txns.length) return { state: 'unknown', lastTxn: null, history: [] };

  const lastTxn = txns[txns.length - 1];
  const state = ACTION_TO_STATE[lastTxn.action] || 'unknown';

  return { state, lastTxn, history: txns };
}

// Check if a transition is valid and what inferences are needed
function validateTransition(barcode, newAction, newDate, employeeId) {
  const { state, lastTxn, history } = resolveState(barcode);
  const allowed = VALID_TRANSITIONS[state] || VALID_TRANSITIONS['unknown'];
  const isValid = allowed.includes(newAction);
  const warnings = [];
  const inferences = [];

  if (state === 'unknown') {
    // First time seeing this barcode, any action is OK
    return { valid: true, warnings: [], inferences: [], currentState: state };
  }

  if (isValid) {
    // Direct valid transition, no inference needed
    return { valid: true, warnings: [], inferences: [], currentState: state };
  }

  // ---- INFERENCE LOGIC: Try to bridge the gap ----

  // Case: Item is with_employee, action is collected_from_soil_bin
  // Infer: employee returned to warehouse first
  if (state === 'with_employee' && newAction === 'collected_from_soil_bin') {
    inferences.push({
      action: 'returned',
      barcode,
      employeeId: lastTxn.employeeId,
      employeeName: lastTxn.employeeName,
      date: newDate,
      notes: 'Auto-inferred: Employee returned uniform before soil bin collection',
      inferred: true
    });
    return { valid: true, warnings: ['Item was with employee — auto-inferring return step'], inferences, currentState: state };
  }

  // Case: Item is with_employee, action is sent_to_cintas
  // Infer: returned + collected from soil bin
  if (state === 'with_employee' && newAction === 'sent_to_cintas') {
    inferences.push({
      action: 'returned',
      barcode,
      employeeId: lastTxn.employeeId,
      employeeName: lastTxn.employeeName,
      date: newDate,
      notes: 'Auto-inferred: Employee returned uniform before Cintas send',
      inferred: true
    });
    inferences.push({
      action: 'collected_from_soil_bin',
      barcode,
      employeeId: null,
      employeeName: null,
      date: newDate,
      notes: 'Auto-inferred: Collected from soil bin before Cintas send',
      inferred: true
    });
    return { valid: true, warnings: ['Item was with employee — auto-inferring return + soil bin collection'], inferences, currentState: state };
  }

  // Case: Item is with_employee, action is received_from_cintas (re-receipt)
  // Infer: returned + collected from soil bin + sent to cintas
  if (state === 'with_employee' && newAction === 'received_from_cintas') {
    const lastDate = lastTxn.date || newDate;
    const midDate = estimateMidDate(lastDate, newDate);

    inferences.push({
      action: 'returned',
      barcode,
      employeeId: lastTxn.employeeId,
      employeeName: lastTxn.employeeName,
      date: midDate,
      notes: 'Auto-inferred: Employee returned uniform (estimated date)',
      inferred: true
    });
    inferences.push({
      action: 'collected_from_soil_bin',
      barcode,
      employeeId: null,
      employeeName: null,
      date: midDate,
      notes: 'Auto-inferred: Collected from soil bin (estimated date)',
      inferred: true
    });
    inferences.push({
      action: 'sent_to_cintas',
      barcode,
      employeeId: null,
      employeeName: null,
      date: midDate,
      notes: 'Auto-inferred: Sent to Cintas for cleaning (estimated date)',
      inferred: true
    });
    return { valid: true, warnings: ['Item was with employee — auto-inferring full return → soil bin → Cintas cycle'], inferences, currentState: state };
  }

  // Case: Item is in warehouse, action is received_from_cintas
  // Infer: sent to cintas first
  if (state === 'warehouse' && newAction === 'received_from_cintas') {
    const lastDate = lastTxn.date || newDate;
    const midDate = estimateMidDate(lastDate, newDate);
    inferences.push({
      action: 'collected_from_soil_bin',
      barcode,
      employeeId: null,
      employeeName: null,
      date: midDate,
      notes: 'Auto-inferred: Collected from soil bin (estimated date)',
      inferred: true
    });
    inferences.push({
      action: 'sent_to_cintas',
      barcode,
      employeeId: null,
      employeeName: null,
      date: midDate,
      notes: 'Auto-inferred: Sent to Cintas for cleaning (estimated date)',
      inferred: true
    });
    return { valid: true, warnings: ['Item was in warehouse — auto-inferring soil bin → Cintas cycle'], inferences, currentState: state };
  }

  // Case: Item is in soil_bin, action is received_from_cintas
  // Infer: sent to cintas
  if (state === 'soil_bin' && newAction === 'received_from_cintas') {
    const lastDate = lastTxn.date || newDate;
    const midDate = estimateMidDate(lastDate, newDate);
    inferences.push({
      action: 'sent_to_cintas',
      barcode,
      employeeId: null,
      employeeName: null,
      date: midDate,
      notes: 'Auto-inferred: Sent to Cintas for cleaning (estimated date)',
      inferred: true
    });
    return { valid: true, warnings: ['Item was in soil bin — auto-inferring Cintas send'], inferences, currentState: state };
  }

  // Case: Item is damaged, action is received_from_cintas
  // Infer: sent to cintas (damaged return)
  if (state === 'damaged' && newAction === 'received_from_cintas') {
    const lastDate = lastTxn.date || newDate;
    const midDate = estimateMidDate(lastDate, newDate);
    inferences.push({
      action: 'sent_to_cintas',
      barcode,
      employeeId: null,
      employeeName: null,
      date: midDate,
      notes: 'Auto-inferred: Damaged item returned to Cintas (estimated date)',
      inferred: true
    });
    return { valid: true, warnings: ['Damaged item — auto-inferring Cintas return'], inferences, currentState: state };
  }

  // Case: Item is at_cintas, action is distributed
  // Infer: received from cintas
  if (state === 'at_cintas' && newAction === 'distributed') {
    inferences.push({
      action: 'received_from_cintas',
      barcode,
      employeeId: null,
      employeeName: null,
      date: newDate,
      notes: 'Auto-inferred: Received from Cintas before distribution',
      inferred: true
    });
    return { valid: true, warnings: ['Item was at Cintas — auto-inferring receipt before distribution'], inferences, currentState: state };
  }

  // Case: Item is at_cintas, other actions
  if (state === 'at_cintas' && newAction !== 'received_from_cintas') {
    inferences.push({
      action: 'received_from_cintas',
      barcode,
      employeeId: null,
      employeeName: null,
      date: newDate,
      notes: 'Auto-inferred: Received from Cintas (estimated)',
      inferred: true
    });
    // After receiving, check if we need more inferences
    const nextState = 'warehouse';
    const nextAllowed = VALID_TRANSITIONS[nextState];
    if (nextAllowed.includes(newAction)) {
      return { valid: true, warnings: ['Item was at Cintas — auto-inferring receipt'], inferences, currentState: state };
    }
  }

  // Case: Item is with_employee, action is distributed (re-distribute)
  if (state === 'with_employee' && newAction === 'distributed') {
    inferences.push({
      action: 'returned',
      barcode,
      employeeId: lastTxn.employeeId,
      employeeName: lastTxn.employeeName,
      date: newDate,
      notes: 'Auto-inferred: Previous employee returned uniform before redistribution',
      inferred: true
    });
    return { valid: true, warnings: ['Item was with another employee — auto-inferring return'], inferences, currentState: state };
  }

  // Fallback: allow with warning
  warnings.push(`Unusual transition: ${state} → ${newAction}. Proceeding anyway.`);
  return { valid: true, warnings, inferences: [], currentState: state };
}

// Estimate a midpoint date between two dates
function estimateMidDate(dateA, dateB) {
  const a = new Date(dateA);
  const b = new Date(dateB);
  if (isNaN(a.getTime()) || isNaN(b.getTime())) return today();
  const mid = new Date((a.getTime() + b.getTime()) / 2);
  return mid.toISOString().slice(0, 10);
}

// ============================================================
// SUPABASE SYNC
// ============================================================

// Column mapping: v3 camelCase ↔ Supabase snake_case
function empToSupabase(e) {
  const row = {
    first_name: e.firstName || '',
    last_name: e.lastName || '',
    employee_id_number: e.employeeId || null,
    production_centre: e.productionCentre || null,
    department: e.department || null,
    phone: e.phone || null,
    notes: e.notes || null
  };
  if (e.id && isValidUUID(e.id)) {
    row.id = e.id;
  }
  return row;
}

function empFromSupabase(row) {
  return {
    id: row.id,
    firstName: row.first_name || '',
    lastName: row.last_name || '',
    employeeId: row.employee_id_number || '',
    productionCentre: row.production_centre || '',
    department: row.department || '',
    phone: row.phone || '',
    notes: row.notes || '',
    createdAt: row.created_at || new Date().toISOString()
  };
}

function txnToSupabase(t) {
  const row = {
    barcode: t.barcode,
    action: t.action,
    employee_id: t.employeeId && isValidUUID(t.employeeId) ? t.employeeId : null,
    uniform_type: t.uniformType || null,
    notes: t.notes ? (t.inferred ? '[AUTO] ' + t.notes : t.notes) : (t.inferred ? '[AUTO]' : null),
    date: t.date || today()
  };
  if (t.id && isValidUUID(t.id)) {
    row.id = t.id;
  }
  return row;
}

function txnFromSupabase(row, employeeMap) {
  const emp = row.employee_id ? (employeeMap[row.employee_id] || null) : null;
  const notes = row.notes || '';
  const inferred = notes.startsWith('[AUTO]');
  return {
    id: row.id,
    barcode: row.barcode,
    action: row.action,
    employeeId: row.employee_id || null,
    employeeName: emp ? `${emp.firstName} ${emp.lastName}` : null,
    uniformType: row.uniform_type || '',
    notes: inferred ? notes.replace(/^\[AUTO\]\s*/, '') : notes,
    date: row.date || today(),
    createdAt: row.created_at || new Date().toISOString(),
    inferred
  };
}

// The Supabase client instance (created after credentials are saved)
let _supabaseClient = null;
let _realtimeChannel = null;

function getSupabase() {
  if (_supabaseClient) return _supabaseClient;
  const url = 'https://ruhovbdcnnukvobbejor.supabase.co';
  const key = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InJ1aG92YmRjbm51a3ZvYmJlam9yIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODAwMTc5NzksImV4cCI6MjA5NTU5Mzk3OX0.Yx769_aDtIdQhiiYuaDCcAq05QZPaCAhBpOxMrWm7Jc';

  try {
    _supabaseClient = window.supabase.createClient(url, key);
    return _supabaseClient;
  } catch (e) {
    console.error('Failed to create Supabase client:', e);
    return null;
  }
}

const CloudSync = {
  isReady() { return !!_supabaseClient; },

  async manualSync() {
    await this.processSyncQueue();
    await this.pullAll();
  },

  async processSyncQueue() {
    if (!navigator.onLine) return;
    const sb = getSupabase();
    if (!sb) return;
    try {
      const queue = await IDB.getSyncQueue();
      if (!queue || queue.length === 0) return;
      
      console.log(`Processing sync queue with ${queue.length} items...`);
      for (const item of queue) {
        try {
          if (item.type === 'pushTransaction') {
             const row = txnToSupabase(item.payload);
             const { error } = await sb.from('transactions').upsert(row, { onConflict: 'id' });
             if (error) throw error;
          } else if (item.type === 'deleteTransaction') {
             const { error } = await sb.from('transactions').delete().eq('id', item.payload);
             if (error) throw error;
          } else if (item.type === 'pushEmployee') {
             const row = empToSupabase(item.payload);
             const { error } = await sb.from('employees').upsert(row, { onConflict: 'id' });
             if (error) throw error;
          } else if (item.type === 'removeEmployee') {
             const { error } = await sb.from('employees').delete().eq('id', item.payload);
             if (error) throw error;
          }
          await IDB.removeFromSyncQueue(item.queueId);
        } catch (itemErr) {
          console.warn('Failed to process queue item:', itemErr);
          // Stop processing if we hit an error (likely offline again)
          break;
        }
      }
      updateSyncStatus();
    } catch (e) {
      console.error('Failed to process sync queue:', e);
    }
  },

  // ── Transactions ──────────────────────────────────────────
  async pushTransaction(txn) {
    const sb = getSupabase();
    if (!sb) return;
    try {
      if (!navigator.onLine) throw new Error('Offline');
      const row = txnToSupabase(txn);
      const { error } = await sb.from('transactions').upsert(row, { onConflict: 'id' });
      if (error) throw error;
    } catch (e) {
      console.warn('pushTransaction failed, queuing:', e.message);
      IDB.addToSyncQueue({ type: 'pushTransaction', payload: txn }).then(updateSyncStatus);
    }
  },

  async pushTransactionsBulk(txns) {
    const sb = getSupabase();
    if (!sb || !txns.length) return;
    try {
      const rows = txns.map(txnToSupabase);
      // Upsert in batches of 500 to avoid payload limits
      for (let i = 0; i < rows.length; i += 500) {
        const { error } = await sb.from('transactions').upsert(rows.slice(i, i + 500), { onConflict: 'id' });
        if (error) throw new Error(error.message);
      }
    } catch (e) { throw new Error('Bulk transaction push failed: ' + e.message); }
  },

  async deleteTransaction(id) {
    const sb = getSupabase();
    if (!sb) return;
    try {
      if (!navigator.onLine) throw new Error('Offline');
      const { error } = await sb.from('transactions').delete().eq('id', id);
      if (error) throw error;
    } catch (e) {
      console.warn('deleteTransaction failed, queuing:', e.message);
      IDB.addToSyncQueue({ type: 'deleteTransaction', payload: id }).then(updateSyncStatus);
    }
  },

  // ── Employees ─────────────────────────────────────────────
  async pushEmployee(emp) {
    const sb = getSupabase();
    if (!sb) return;
    try {
      if (!navigator.onLine) throw new Error('Offline');
      const row = empToSupabase(emp);
      const { error } = await sb.from('employees').upsert(row, { onConflict: 'id' });
      if (error) throw error;
    } catch (e) {
      console.warn('pushEmployee failed, queuing:', e.message);
      IDB.addToSyncQueue({ type: 'pushEmployee', payload: emp }).then(updateSyncStatus);
    }
  },

  async removeEmployee(id) {
    const sb = getSupabase();
    if (!sb) return;
    try {
      if (!navigator.onLine) throw new Error('Offline');
      const { error } = await sb.from('employees').delete().eq('id', id);
      if (error) throw error;
    } catch (e) {
      console.warn('removeEmployee failed, queuing:', e.message);
      IDB.addToSyncQueue({ type: 'removeEmployee', payload: id }).then(updateSyncStatus);
    }
  },

  // ── Config ────────────────────────────────────────────────
  async pushConfig() {
    const sb = getSupabase();
    if (!sb) return;
    try {
      await sb.from('config').upsert(
        [{ key: 'production_centres', value: DB.centres },
         { key: 'uniform_types',      value: DB.uniformTypes }],
        { onConflict: 'key' }
      );
    } catch (e) { console.warn('pushConfig failed:', e.message); }
  },

  // ── Pull All (full sync from Supabase) ────────────────────
  async pullAll(silent = false) {
    const sb = getSupabase();
    if (!sb) {
      if (!silent) showToast('Supabase not configured. Go to Settings → Database.', 'error');
      return;
    }

    const syncNowBtn = document.getElementById('cloud-sync-now-btn');
    if (syncNowBtn) { syncNowBtn.disabled = true; syncNowBtn.textContent = 'Pulling…'; }

    // Helper: yield to browser between heavy steps
    const yield_ = () => new Promise(r => setTimeout(r, 0));

    try {
      // ── Step 1: Fetch employees ──────────────────────────────
      if (!silent) showToast('⏳ Fetching employees…', 'info');
      await yield_();

      const { data: empRows, error: empErr } = await sb.from('employees').select('*').order('created_at');
      if (empErr) throw new Error('Employees fetch failed: ' + empErr.message);

      const employees = (empRows || []).map(empFromSupabase);
      const employeeMap = {};
      employees.forEach(e => { employeeMap[e.id] = e; });

      // ── Step 2: Fetch transactions (paginated) ───────────────
      if (!silent) showToast(`⏳ Fetching transactions (${employees.length} employees loaded)…`, 'info');
      await yield_();

      let allTxnRows = [];
      let from = 0;
      const pageSize = 1000;
      while (true) {
        const { data: rows, error: txnErr } = await sb
          .from('transactions').select('*')
          .order('created_at', { ascending: false })
          .range(from, from + pageSize - 1);
        if (txnErr) throw new Error('Transactions fetch failed: ' + txnErr.message);
        if (!rows || rows.length === 0) break;
        allTxnRows = allTxnRows.concat(rows);
        if (!silent) showToast(`⏳ Fetched ${allTxnRows.length} transactions…`, 'info');
        await yield_();
        if (rows.length < pageSize) break;
        from += pageSize;
      }

      // ── Step 3: Map rows (chunked with yields) ───────────────
      if (!silent) showToast(`⏳ Processing ${allTxnRows.length} transactions…`, 'info');
      await yield_();

      const transactions = [];
      const MAP_CHUNK = 500;
      for (let i = 0; i < allTxnRows.length; i += MAP_CHUNK) {
        const chunk = allTxnRows.slice(i, i + MAP_CHUNK);
        transactions.push(...chunk.map(r => txnFromSupabase(r, employeeMap)));
        await yield_(); // yield every 500 rows
      }

      // ── Step 4: Fetch config ─────────────────────────────────
      const { data: cfgRows } = await sb.from('config').select('*');
      const centres      = cfgRows?.find(r => r.key === 'production_centres')?.value || DB.centres;
      const uniformTypes = cfgRows?.find(r => r.key === 'uniform_types')?.value    || DB.uniformTypes;

      // ── Step 5: Detect changes ───────────────────────────────
      const hasNewEmployees = employees.length    !== DB.employees.length;
      const hasNewTxns      = transactions.length !== DB.transactions.length;

      // ── Step 6: Update memory ────────────────────────────────
      DB.saveEmployees(employees);
      DB.saveCentres(centres);
      DB.saveTypes(uniformTypes);
      DB_MEMORY.transactions = transactions;
      await yield_();

      // ── Step 7: Save to IndexedDB (chunked internally) ───────
      if (!silent) showToast(`⏳ Saving to local storage…`, 'info');
      try { await IDB.saveTransactionsBulk(transactions); } catch (e) {
        console.error('Failed saving transactions to IndexedDB:', e);
      }

      localStorage.setItem('ut_last_sync', new Date().toISOString());

      // ── Step 8: Re-render ────────────────────────────────────
      if (!silent) {
        showToast(`✓ Synced — ${employees.length} employees, ${transactions.length} transactions`, 'success');
      } else if (hasNewEmployees || hasNewTxns) {
        showToast('🔄 Data updated from Supabase', 'info');
      }

      await yield_();
      renderDashboard();
      await yield_();
      renderEmployeeList();
      renderSettings();

    } catch (e) {
      if (!silent) showToast('Pull error: ' + e.message, 'error');
      else console.warn('Silent pull error:', e.message);
    } finally {
      if (syncNowBtn) { syncNowBtn.disabled = false; syncNowBtn.textContent = 'Pull from Supabase'; }
      updateSyncStatus();
    }
  },

  // ── Push All (full local data → Supabase) ─────────────────
  async syncAll() {
    const sb = getSupabase();
    if (!sb) { showToast('Supabase not configured. Go to Settings → Database.', 'error'); return; }

    const btn = document.getElementById('cloud-sync-btn');
    if (btn) { btn.disabled = true; btn.textContent = 'Pushing…'; }

    try {
      // Push employees
      const empRows = DB.employees.map(empToSupabase);
      if (empRows.length) {
        for (let i = 0; i < empRows.length; i += 500) {
          const { error } = await sb.from('employees').upsert(empRows.slice(i, i + 500), { onConflict: 'id' });
          if (error) throw new Error('Employee push failed: ' + error.message);
        }
      }

      // Push transactions
      await this.pushTransactionsBulk(DB.transactions);

      // Push config
      await this.pushConfig();

      localStorage.setItem('ut_last_sync', new Date().toISOString());
      showToast(`✓ Pushed ${DB.employees.length} employees & ${DB.transactions.length} transactions to Supabase!`, 'success');

    } catch (e) {
      showToast('Push error: ' + e.message, 'error');
    } finally {
      if (btn) { btn.disabled = false; btn.textContent = 'Push All to Supabase'; }
      updateSyncStatus();
    }
  },

  async testConnection() {
    const url = 'https://ruhovbdcnnukvobbejor.supabase.co';
    const key = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InJ1aG92YmRjbm51a3ZvYmJlam9yIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODAwMTc5NzksImV4cCI6MjA5NTU5Mzk3OX0.Yx769_aDtIdQhiiYuaDCcAq05QZPaCAhBpOxMrWm7Jc';

    const btn = document.getElementById('cloud-test-btn');
    if (btn) { btn.disabled = true; btn.innerHTML = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="22 12 18 12 15 21 9 3 6 12 2 12"/></svg> Testing…'; }

    try {
      // Recreate client
      const testClient = window.supabase.createClient(url, key, { auth: { persistSession: false } });
      const { data, error } = await testClient.from('employees').select('id', { count: 'exact', head: true });
      if (error) {
        if (error.code === 'PGRST301' || error.message.includes('JWT')) {
          showToast('Invalid Anon Key. Check your Supabase API settings.', 'error');
        } else if (error.message.includes('relation') || error.code === '42P01') {
          showToast('Connected! But tables not found — run supabase_schema.sql first.', 'error');
        } else {
          showToast('Connection error: ' + error.message, 'error');
        }
        return;
      }

      // Count transactions too
      const { count: txnCount } = await testClient
        .from('transactions').select('id', { count: 'exact', head: true });

      const { count: empCount } = await testClient
        .from('employees').select('id', { count: 'exact', head: true });

      showToast(`✓ Connected! ${empCount || 0} employees, ${txnCount || 0} transactions in Supabase.`, 'success');
      localStorage.setItem('ut_last_sync', new Date().toISOString());
      updateSyncStatus();

    } catch (e) {
      showToast('Connection failed: ' + e.message, 'error');
    } finally {
      if (btn) { btn.disabled = false; btn.innerHTML = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="22 12 18 12 15 21 9 3 6 12 2 12"/></svg> Test Connection'; }
    }
  }
};

async function updateSyncStatus() {
  const statusEl = document.getElementById('cloud-status-text');
  const dotEl    = document.getElementById('cloud-status-dot');
  const ts       = localStorage.getItem('ut_last_sync');
  const ready    = CloudSync.isReady();
  
  let queueLength = 0;
  try {
    const queue = await IDB.getSyncQueue();
    queueLength = queue ? queue.length : 0;
  } catch(e) {}

  if (statusEl) {
    let statusText = '';
    if (!ready) {
      statusText = 'Not connected';
    } else if (!navigator.onLine) {
      statusText = `Offline (${queueLength} pending)`;
    } else if (ts) {
      const minutesAgo = Math.floor((Date.now() - new Date(ts).getTime()) / 60000);
      if (minutesAgo < 1)   statusText = 'Last sync: Just now';
      else if (minutesAgo < 60) statusText = `Last sync: ${minutesAgo}m ago`;
      else                  statusText = `Last sync: ${Math.floor(minutesAgo / 60)}h ago`;
      if (queueLength > 0) {
        statusText += ` · ${queueLength} pending`;
      }
    } else {
      statusText = 'Connected (never synced)';
    }
    if (autoSyncInterval && ready && navigator.onLine) statusText += ' · Live';
    statusEl.textContent = statusText;
  }
  if (dotEl) {
    dotEl.className = 'cloud-dot ' + (ready && navigator.onLine ? 'ready' : 'off');
    if (queueLength > 0 && navigator.onLine) dotEl.classList.add('syncing');
  }
}

async function manualSync() {
  const btn = document.getElementById('cloud-sync-master-btn');
  if (btn) { btn.disabled = true; btn.innerHTML = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="23 4 23 10 17 10"/><polyline points="1 20 1 14 7 14"/><path d="M3.51 9a9 9 0 0114.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0020.49 15"/></svg> Syncing...'; }
  try {
    if (!navigator.onLine) {
      showToast('Cannot sync while offline.', 'error');
      return;
    }
    await CloudSync.processSyncQueue();
    await CloudSync.pullAll(true);
    showToast('Sync complete!', 'success');
  } catch (e) {
    showToast('Sync failed: ' + e.message, 'error');
  } finally {
    if (btn) { btn.disabled = false; btn.innerHTML = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="23 4 23 10 17 10"/><polyline points="1 20 1 14 7 14"/><path d="M3.51 9a9 9 0 0114.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0020.49 15"/></svg> Sync Data'; }
  }
}

window.manualSync      = manualSync;
window.cloudTestConn   = () => CloudSync.testConnection();

// Realtime subscription — get notified when other devices push new scans
function startRealtimeSubscription() {
  const sb = getSupabase();
  if (!sb || _realtimeChannel) return;
  try {
    _realtimeChannel = sb
      .channel('unitrack-realtime')
      .on('postgres_changes', { event: '*', schema: 'public', table: 'transactions' }, async (payload) => {
        const type = payload.eventType;
        if (type === 'DELETE') {
          const oldRow = payload.old;
          if (oldRow && oldRow.id) {
            DB_MEMORY.transactions = DB_MEMORY.transactions.filter(t => t.id !== oldRow.id);
            IDB.deleteTransaction(oldRow.id).catch(() => {});
            if (currentPage === 'dashboard') renderDashboard();
            if (currentPage === 'reports')   renderReport();
          }
          return;
        }

        const newRow = payload.new;
        if (!newRow) return;
        
        // Build employee map for name lookup
        const employeeMap = {};
        DB.employees.forEach(e => { employeeMap[e.id] = e; });
        const txn = txnFromSupabase(newRow, employeeMap);
        
        const idx = DB_MEMORY.transactions.findIndex(t => t.id === txn.id);
        if (idx >= 0) {
          DB_MEMORY.transactions[idx] = txn;
        } else {
          DB_MEMORY.transactions.unshift(txn);
        }
        
        IDB.saveTransaction(txn).catch(() => {});
        if (currentPage === 'dashboard') renderDashboard();
        if (currentPage === 'reports')   renderReport();
      })
      .on('postgres_changes', { event: '*', schema: 'public', table: 'employees' }, async (payload) => {
        const type = payload.eventType;
        if (type === 'DELETE') {
          const oldRow = payload.old;
          if (oldRow && oldRow.id) {
            DB_MEMORY.employees = DB_MEMORY.employees.filter(e => e.id !== oldRow.id);
            IDB.set('ut_employees', DB_MEMORY.employees);
            if (currentPage === 'employees') renderEmployeeList();
          }
          return;
        }

        const newRow = payload.new;
        if (!newRow) return;
        
        const emp = empFromSupabase(newRow);
        const idx = DB_MEMORY.employees.findIndex(e => e.id === emp.id);
        if (idx >= 0) {
          DB_MEMORY.employees[idx] = emp;
        } else {
          DB_MEMORY.employees.push(emp);
        }
        
        DB.saveEmployees(DB_MEMORY.employees);
        if (currentPage === 'employees') renderEmployeeList();
      })
      .subscribe((status) => {
        if (status === 'SUBSCRIBED') {
          console.log('UniTrack: Supabase realtime connected');
          updateSyncStatus();
        }
      });
  } catch (e) {
    console.warn('Realtime subscription failed:', e.message);
  }
}

// Auto-sync every 2 minutes + realtime subscription
let autoSyncInterval;
function startAutoSync() {
  if (autoSyncInterval) clearInterval(autoSyncInterval);
  const sb = getSupabase();
  if (!sb) return;

  // Start realtime for instant multi-device updates
  startRealtimeSubscription();

  autoSyncInterval = setInterval(async () => {
    try { await CloudSync.pullAll(true); } catch (e) { console.warn('Auto-sync failed:', e.message); }
  }, 120 * 1000); // 2 minutes fallback

  updateSyncStatus();
}

function stopAutoSync() {
  if (autoSyncInterval) { clearInterval(autoSyncInterval); autoSyncInterval = null; }
  if (_realtimeChannel) { try { _realtimeChannel.unsubscribe(); } catch(e){} _realtimeChannel = null; }
  updateSyncStatus();
}


// ============================================================
// UTILITIES
// ============================================================
function isValidUUID(str) {
  if (typeof str !== 'string') return false;
  return /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(str);
}

function uid() {
  if (typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function') {
    return crypto.randomUUID();
  }
  return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function(c) {
    const r = (Math.random() * 16) | 0;
    const v = c === 'x' ? r : (r & 0x3) | 0x8;
    return v.toString(16);
  });
}
function today(){ return new Date().toISOString().slice(0,10); }
function normalizeDate(d) {
  if (!d) return today();
  if (/^\d{4}-\d{2}-\d{2}$/.test(d)) return d;
  if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(d)) {
    const [m, day, y] = d.split('/');
    return `${y}-${String(m).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
  }
  return d.split('T')[0] || today();
}
function formatDate(d) {
  if (!d) return '';
  const date = normalizeDate(d);
  const [y,m,day] = date.split('-');
  return `${m}/${day}/${y}`;
}
function formatDateTime(iso) {
  if (!iso) return '';
  const d = new Date(iso);
  return d.toLocaleDateString('en-US',{month:'short',day:'numeric'}) + ' ' +
         d.toLocaleTimeString('en-US',{hour:'2-digit',minute:'2-digit'});
}
function avatarColor(name) {
  const colors = ['#2563eb','#7c3aed','#db2777','#d97706','#059669','#0891b2','#dc2626','#65a30d'];
  const text = String(name || 'A');
  let h = 0; for (let c of text) h = (h*31 + c.charCodeAt(0)) & 0xffffffff;
  return colors[Math.abs(h) % colors.length];
}
function initials(first, last) { return ((first||'')[0]||'') + ((last||'')[0]||''); }
function actionLabel(a) {
  return {
    distributed:          'Distributed',
    returned:             'Returned',
    collected_from_soil_bin: 'Soil Bin Collection',
    sent_to_cintas:       '→ Cintas',
    received_from_cintas: '← Cintas',
    reported_issue:       'Issue/Damaged'
  }[a] || a;
}

function stateLabel(s) {
  return {
    warehouse:     'In Warehouse',
    with_employee: 'With Employee',
    soil_bin:      'In Soil Bin',
    at_cintas:     'At Cintas',
    damaged:       'Damaged',
    unknown:       'Unknown'
  }[s] || s;
}

function normalizeRemoteAction(action) {
  const value = String(action || '').trim().toLowerCase();
  const map = {
    distributed: 'distributed',
    returned: 'returned',
    'soil bin collection': 'collected_from_soil_bin',
    collected_from_soil_bin: 'collected_from_soil_bin',
    'sent to cintas': 'sent_to_cintas',
    sent_to_cintas: 'sent_to_cintas',
    '→ cintas': 'sent_to_cintas',
    'received from cintas': 'received_from_cintas',
    received_from_cintas: 'received_from_cintas',
    '← cintas': 'received_from_cintas',
    'reported issue': 'reported_issue',
    reported_issue: 'reported_issue',
    'issue/damaged': 'reported_issue',
    issue: 'reported_issue',
    damaged: 'reported_issue'
  };
  return map[value] || String(action || '').trim();
}

function findEmployeeByIdOrExternal(id) {
  if (!id) return null;
  return DB.employees.find(e => e.id === id || e.employeeId === id);
}

function escHtml(s) {
  return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

// ============================================================
// iOS BLUETOOTH KEYBOARD FIX
// ============================================================
// When a Bluetooth barcode scanner is connected, iOS treats it
// as a hardware keyboard and suppresses the software keyboard.
// forceKeyboard() uses the readOnly toggle trick — the only
// cross-iOS-version reliable way to force the keyboard open.
// ============================================================
function forceKeyboard(inputId) {
  const el = typeof inputId === 'string'
    ? document.getElementById(inputId)
    : inputId;
  if (!el) return;

  // Trick: briefly set readOnly → focus → remove readOnly → focus again
  // This forces iOS to re-evaluate and show the software keyboard
  el.setAttribute('readonly', 'readonly');
  el.focus();
  requestAnimationFrame(() => {
    el.removeAttribute('readonly');
    el.focus();
    // Place cursor at end
    const len = el.value.length;
    try { el.setSelectionRange(len, len); } catch(_) {}
  });
}
window.forceKeyboard = forceKeyboard;

// On touch devices: auto-force keyboard on touchend for text inputs
// so users don't always have to tap the ⌨ button
(function initIOSKeyboardFix() {
  const isTouch = navigator.maxTouchPoints > 0 || 'ontouchstart' in window;
  if (!isTouch) return;

  document.addEventListener('touchend', (e) => {
    const target = e.target;
    if (!target) return;
    const tag  = target.tagName;
    const type = (target.type || '').toLowerCase();
    const im   = (target.getAttribute('inputmode') || '').toLowerCase();

    // Apply to text inputs but NOT the barcode input (inputmode=none)
    if ((tag === 'INPUT' || tag === 'TEXTAREA') && im !== 'none' && type !== 'date') {
      // Small delay so the tap registers first
      setTimeout(() => forceKeyboard(target), 80);
    }
  }, { passive: true });
})();


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
let focusLockInterval = null;

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
    if (empData) {
      empData = empData.map(e => {
        if (!e.firstName && e['Column 2']) {
          return {
            id: String(e.id || e['Column 1'] || uid()).trim(),
            firstName: String(e['Column 2'] || '').trim(),
            lastName: String(e['Column 3'] || '').trim(),
            employeeId: String(e['Column 4'] || '').trim(),
            productionCentre: String(e['Column 5'] || '').trim(),
            department: String(e['Column 6'] || '').trim(),
            phone: String(e['Column 7'] || '').trim(),
            notes: String(e['Column 8'] || '').trim(),
            createdAt: e['Column 9'] || e.createdAt || new Date().toISOString()
          };
        }
        return e;
      });
      DB_MEMORY.employees = empData;
      await IDB.set('ut_employees', empData);
    }

    // Transactions
    let txData = await IDB.getAllTransactions();
    txData = txData.map(t => {
      if (!t.barcode && t['Column 2']) {
        return {
          id: String(t.id || t['Column 1'] || uid()).trim(),
          action: normalizeRemoteAction(t['Column 3'] || ''),
          barcode: String(t['Column 2'] || '').trim(),
          employeeId: String(t['Column 4'] || '').trim(),
          employeeName: String(t['Column 5'] || '').trim(),
          uniformType: String(t['Column 6'] || '').trim(),
          notes: String(t['Column 10'] || '').trim(),
          date: normalizeDate(t['Column 8'] || t['Column 7'] || ''),
          createdAt: t['Column 9'] || t.createdAt || new Date().toISOString(),
          inferred: t.inferred || false
        };
      }
      return t;
    });

    // Migration from legacy store
    const legacyTxData = await IDB.get('ut_transactions');
    if (legacyTxData && legacyTxData.length > 0) {
      IDB.set('ut_transactions', []);
      hasMigratedAny = true;
    }
    if (txData.length === 0 && lsTransactions) {
      const parsed = JSON.parse(lsTransactions);
      for (const t of parsed) {
        if (!t.id) t.id = uid();
        await IDB.saveTransaction(t);
        txData.push(t);
      }
      hasMigratedAny = true;
    }

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
      userData = [
        { username: 'Admin', password: 'Admin1980', role: 'admin' },
        { username: 'Operator', password: 'Oper1234', role: 'operator' },
        { username: 'Warehouse', password: 'Wh1234', role: 'warehouse' },
        { username: 'Manager', password: 'Manager123', role: 'manager' }
      ];
      hasMigratedAny = true;
    }
    DB_MEMORY.users = userData;

    if (hasMigratedAny) {
      IDB.set('ut_employees', DB_MEMORY.employees);
      IDB.set('ut_centres', DB_MEMORY.centres);
      IDB.set('ut_types', DB_MEMORY.uniformTypes);
      IDB.set('ut_users', DB_MEMORY.users);
    }
  } catch (e) {
    console.error("IndexedDB init failed, using defaults.", e);
    DB_MEMORY.users = [
      { username: 'Admin', password: 'Admin1980', role: 'admin' },
      { username: 'Operator', password: 'Oper1234', role: 'operator' },
      { username: 'Warehouse', password: 'Wh1234', role: 'warehouse' },
      { username: 'Manager', password: 'Manager123', role: 'manager' }
    ];
  }

  // Auth check
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
    setTimeout(() => {
      const lu = document.getElementById('login-user');
      if (lu) lu.focus();
    }, 100);
  }

  // Close employee dropdown on outside click
  document.addEventListener('click', (e) => {
    const dd = document.getElementById('scan-emp-dropdown');
    const wrapper = document.querySelector('.emp-search-wrapper');
    if (dd && wrapper && !wrapper.contains(e.target)) dd.classList.add('hidden');
  });

  // Pre-configure Supabase if not already set up
  // (credentials from v2 project — user can override in Settings → Database)
  if (!localStorage.getItem('ut_cloud_url') || !localStorage.getItem('ut_cloud_key')) {
    localStorage.setItem('ut_cloud_url', 'https://ruhovbdcnnukvobbejor.supabase.co');
    localStorage.setItem('ut_cloud_key', 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InJ1aG92YmRjbm51a3ZvYmJlam9yIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODAwMTc5NzksImV4cCI6MjA5NTU5Mzk3OX0.Yx769_aDtIdQhiiYuaDCcAq05QZPaCAhBpOxMrWm7Jc');
  }
  // Populate settings fields if on settings page
  const urlEl = document.getElementById('cloud-url');
  const keyEl = document.getElementById('cloud-key');
  if (urlEl) urlEl.value = localStorage.getItem('ut_cloud_url') || '';
  if (keyEl) keyEl.value = localStorage.getItem('ut_cloud_key') || '';

  startAutoSync();

  // On first load, pull from Supabase to get all cloud data
  const isFirstLoad = !localStorage.getItem('ut_last_sync');
  if (isFirstLoad && getSupabase()) {
    setTimeout(() => CloudSync.pullAll(false), 1500);
  }

  // Responsive sidebar
  window.addEventListener('resize', () => {
    const sidebar = document.getElementById('sidebar');
    const main = document.querySelector('.main-content');
    const backdrop = document.getElementById('sidebar-backdrop');
    if (window.innerWidth > 768) {
      if (backdrop) backdrop.style.display = 'none';
      sidebar.classList.remove('open');
      if (sidebarOpen) { sidebar.classList.remove('closed'); main.classList.remove('expanded'); }
      else { sidebar.classList.add('closed'); main.classList.add('expanded'); }
    } else {
      sidebar.classList.remove('closed');
      main.classList.remove('expanded');
      if (sidebarOpen) { sidebar.classList.add('open'); if (backdrop) backdrop.style.display = 'block'; }
      else { sidebar.classList.remove('open'); if (backdrop) backdrop.style.display = 'none'; }
    }
  });
});

// ============================================================
// AUTH & ROLES
// ============================================================
function handleLogin() {
  const u = document.getElementById('login-user').value.trim();
  const p = document.getElementById('login-pass').value;
  let user = DB_MEMORY.users.find(x => x.username.toLowerCase() === u.toLowerCase() && x.password === p);

  if (!user && u.toLowerCase() === 'admin' && (p === 'admin' || p === 'admin123')) {
    const adminUser = DB_MEMORY.users.find(x => x.role === 'admin');
    if (adminUser) { user = adminUser; showToast('Legacy admin login accepted.', 'info'); }
  }
  if (!user) { showToast('Invalid username or password', 'error'); return; }
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
  stopAutoSync();
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
// ROLE-BASED PERMISSIONS
// ============================================================
let ROLE_PERMISSIONS = {
  admin: {
    actions: ['received_from_cintas', 'distributed', 'reported_issue', 'returned', 'collected_from_soil_bin', 'sent_to_cintas'],
    canManageUsers: true, canManageEmployees: true, canDeleteTransactions: true, canExportData: true
  },
  operator: {
    actions: ['received_from_cintas', 'distributed', 'reported_issue', 'returned', 'collected_from_soil_bin', 'sent_to_cintas'],
    canManageUsers: false, canManageEmployees: true, canDeleteTransactions: true, canExportData: false
  },
  warehouse: {
    actions: ['received_from_cintas', 'returned', 'collected_from_soil_bin', 'sent_to_cintas'],
    canManageUsers: false, canManageEmployees: false, canDeleteTransactions: false, canExportData: true
  },
  manager: {
    actions: [],
    canManageUsers: false, canManageEmployees: false, canDeleteTransactions: false, canExportData: false
  }
};

function loadRolePermissions() {
  const saved = localStorage.getItem('ut_role_permissions');
  if (saved) {
    try {
      ROLE_PERMISSIONS = JSON.parse(saved);
    } catch (e) {
      console.error('Failed to parse saved role permissions', e);
    }
  }
}
loadRolePermissions();

function saveRolePermissions() {
  localStorage.setItem('ut_role_permissions', JSON.stringify(ROLE_PERMISSIONS));
  showToast('Role permissions saved!', 'success');
}
window.saveRolePermissions = saveRolePermissions;

function hasPermission(action) {
  if (!currentUser) return false;
  return (ROLE_PERMISSIONS[currentUser.role] || ROLE_PERMISSIONS.operator).actions.includes(action);
}
function canManageUsers() {
  if (!currentUser) return false;
  return (ROLE_PERMISSIONS[currentUser.role] || ROLE_PERMISSIONS.operator).canManageUsers;
}
function canManageEmployees() {
  if (!currentUser) return false;
  return (ROLE_PERMISSIONS[currentUser.role] || ROLE_PERMISSIONS.operator).canManageEmployees;
}
function canDeleteTransactions() {
  if (!currentUser) return false;
  return (ROLE_PERMISSIONS[currentUser.role] || ROLE_PERMISSIONS.operator).canDeleteTransactions;
}
function canExportData() {
  if (!currentUser) return false;
  return (ROLE_PERMISSIONS[currentUser.role] || ROLE_PERMISSIONS.operator).canExportData;
}

function updateActionTabsVisibility() {
  if (!currentUser) return;
  const allowedActions = (ROLE_PERMISSIONS[currentUser.role] || ROLE_PERMISSIONS.operator).actions;
  document.querySelectorAll('.action-tab').forEach(tab => {
    const action = tab.id.replace('tab-', '');
    tab.style.display = allowedActions.includes(action) ? 'flex' : 'none';
  });
  if (!hasPermission(currentAction)) {
    const firstAllowed = allowedActions[0];
    if (firstAllowed) setAction(firstAllowed);
  }
}

// ============================================================
// NAVIGATION
// ============================================================
function toggleSidebar() {
  sidebarOpen = !sidebarOpen;
  const sidebar = document.getElementById('sidebar');
  const main    = document.querySelector('.main-content');
  const backdrop = document.getElementById('sidebar-backdrop');
  if (window.innerWidth <= 768) {
    if (sidebarOpen) { sidebar.classList.add('open'); if (backdrop) backdrop.style.display = 'block'; }
    else { sidebar.classList.remove('open'); if (backdrop) backdrop.style.display = 'none'; }
  } else {
    if (sidebarOpen) { sidebar.classList.remove('closed'); main.classList.remove('expanded'); }
    else { sidebar.classList.add('closed'); main.classList.add('expanded'); }
  }
}

function showPage(page) {
  currentPage = page;
  document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
  const pageEl = document.getElementById('page-' + page);
  if (pageEl) pageEl.classList.add('active');
  const navEl = document.querySelector(`.nav-item[data-page="${page}"]`);
  if (navEl) navEl.classList.add('active');
  const titles = {dashboard:'Dashboard',employees:'Employees',scan:'Scan Entry',reports:'Reports',settings:'Settings',users:'User Management'};
  document.getElementById('pageTitle').textContent = titles[page] || page;

  if (page === 'scan') { updateActionTabsVisibility(); startFocusLock(); }
  else stopFocusLock();

  if (page === 'dashboard') renderDashboard();
  if (page === 'employees') renderEmployeeList();
  if (page === 'scan')      setTimeout(() => { document.getElementById('barcodeInput').focus(); }, 150);
  if (page === 'reports')   renderReport();
  if (page === 'settings')  renderSettings();
  if (page === 'users')     renderUsers();

  // Close sidebar on mobile when navigating
  if (window.innerWidth <= 768 && sidebarOpen) {
    toggleSidebar();
  }
}

// ============================================================
// DASHBOARD
// ============================================================
function computeStats() {
  const txns = DB.transactions;
  // Build latest state for each barcode using the state machine
  const latestByBarcode = {};
  txns.forEach(t => {
    if (!latestByBarcode[t.barcode]) latestByBarcode[t.barcode] = t;
  });

  let out = 0, cintas = 0, warehouse = 0, issues = 0;
  Object.values(latestByBarcode).forEach(last => {
    const state = ACTION_TO_STATE[last.action] || 'unknown';
    if (state === 'with_employee') out++;
    else if (state === 'at_cintas') cintas++;
    else if (state === 'damaged') issues++;
    else if (state === 'warehouse' || state === 'soil_bin') warehouse++;
  });

  return { total: Object.keys(latestByBarcode).length, out, cintas, warehouse, issues };
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
    const emp = findEmployeeByIdOrExternal(t.employeeId);
    const actionKey = t.action || '';
    const empName = emp ? `${emp.firstName} ${emp.lastName}` :
                    (actionKey.includes('cintas') || actionKey === 'reported_issue' ? 'Cintas' : 'Warehouse');
    const inferredMark = t.inferred ? ' <span class="inferred-indicator"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 15V9m0 0l-3 3m3-3l3 3"/><circle cx="12" cy="12" r="10"/></svg>inferred</span>' : '';
    return `<div class="activity-item${t.inferred ? ' inferred' : ''}">
      <span class="activity-dot ${t.action}"></span>
      <div class="activity-info">
        <div class="activity-barcode">${escHtml(t.barcode)}${inferredMark}</div>
        <div class="activity-meta">${actionLabel(t.action)} — ${escHtml(empName)}</div>
      </div>
      <div class="activity-time">${formatDateTime(t.createdAt)}</div>
    </div>`;
  }).join('');
}

// ============================================================
// EMPLOYEES
// ============================================================
function getLatestByBarcode() {
  const latest = {};
  DB.transactions.forEach(t => {
    if (!latest[t.barcode]) latest[t.barcode] = t;
  });
  return latest;
}

function renderEmployeeList() {
  const query = (document.getElementById('emp-search')?.value || '').toLowerCase();
  const list  = DB.employees.filter(e => {
    const searchText = `${e.firstName || ''} ${e.lastName || ''} ${e.name || ''}`.toLowerCase();
    return !query || searchText.includes(query) ||
    (e.employeeId && String(e.employeeId).toLowerCase().includes(query)) ||
    (e.productionCentre && String(e.productionCentre).toLowerCase().includes(query));
  });

  const container = document.getElementById('employee-list');
  const canEdit = canManageEmployees();
  const addButton = document.querySelector('#page-employees .btn-primary');
  if (addButton) addButton.style.display = canEdit ? 'inline-flex' : 'none';

  if (!list.length) {
    container.innerHTML = `<div class="empty-state"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M17 21v-2a4 4 0 00-4-4H5a4 4 0 00-4 4v2"/><circle cx="9" cy="7" r="4"/></svg><p>${query ? 'No employees match.' : 'No employees yet. Click "Add Employee" to start.'}</p></div>`;
    return;
  }

  // FIXED: Use proper latest-by-barcode for held count
  const latestByBarcode = getLatestByBarcode();

  let tableHtml = `<div class="employee-grid-table">
    <div class="employee-grid-header">
      <div>Employee</div>
      <div>Department</div>
      <div>Production Centre</div>
      <div>Actions</div>
    </div>`;

  tableHtml += list.map(e => {
    const firstName = e.firstName || (e.name ? e.name.split(' ')[0] : '');
    const lastName = e.lastName || (e.name ? e.name.split(' ').slice(1).join(' ') : '');
    const fullName = firstName && lastName ? `${firstName} ${lastName}` : (e.name || 'Unknown Employee');
    const color = avatarColor(fullName);
    const ini   = initials(firstName, lastName) || initials(e.name || '', '');

    // FIXED: Get barcodes associated with this employee, check GLOBAL latest state
    const myBarcodes = [...new Set(
      DB.transactions.filter(t => t.employeeId === e.id || t.employeeId === e.employeeId).map(t => t.barcode)
    )];
    const held = myBarcodes.filter(bc => {
      const latest = latestByBarcode[bc];
      return latest && latest.action === 'distributed' && (latest.employeeId === e.id || latest.employeeId === e.employeeId);
    }).length;

    // Mobile metadata pills
    const pills = [];
    if (e.productionCentre) pills.push(`<span class="emp-mobile-pill">${escHtml(e.productionCentre)}</span>`);
    if (e.department) pills.push(`<span class="emp-mobile-pill">${escHtml(e.department)}</span>`);
    const pillsHtml = pills.length ? `<div class="emp-mobile-meta-pills">${pills.join('')}</div>` : '';

    return `<div class="employee-grid-row" data-id="${e.id}">
      <div class="emp-row-info" onclick="showEmployeeDetail('${e.id}')">
        <div>
          <div class="emp-avatar-name-cell">
            <div class="emp-avatar" style="background:${color}">${escHtml(ini)}</div>
            <div class="emp-name-held-wrap">
              <div class="emp-table-name">${escHtml(fullName)}</div>
              <div class="emp-table-held">Holding: <strong>${held}</strong> uniform${held!==1?'s':''}</div>
              ${pillsHtml}
            </div>
          </div>
        </div>
        <div class="emp-table-dept">${escHtml(e.department || '—')}</div>
        <div class="emp-table-centre">${escHtml(e.productionCentre || '—')}</div>
      </div>
      <div class="emp-row-actions" onclick="event.stopPropagation()">
        <div class="emp-actions-cell">
          <div class="emp-actions-flex">
            <!-- Mobile Close Button -->
            <button class="emp-action-btn mobile-close-btn" title="Back" onclick="toggleRowActions('${e.id}', false)">
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M19 12H5M12 19l-7-7 7-7"/></svg>
            </button>
            
            <button class="emp-action-btn" title="View Detail" onclick="showEmployeeDetail('${e.id}')">
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/><circle cx="12" cy="12" r="3"/></svg>
            </button>
            
            ${canEdit ? `
              <button class="emp-action-btn" title="Edit" onclick="openEmployeeModal('${e.id}')">
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M11 4H4a2 2 0 00-2 2v14a2 2 0 002 2h14a2 2 0 002-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 013 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
              </button>
              <button class="emp-action-btn del" title="Delete" onclick="deleteEmployee('${e.id}')">
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="3 6 5 6 21 6"/><path d="M19 6l-1 14a2 2 0 01-2 2H8a2 2 0 01-2-2L5 6"/></svg>
              </button>
            ` : ''}
            
            <!-- Mobile Trigger Button -->
            <button class="emp-action-btn mobile-trigger-btn" title="Actions" onclick="toggleRowActions('${e.id}', true)">
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="1"/><circle cx="12" cy="5" r="1"/><circle cx="12" cy="19" r="1"/></svg>
            </button>
          </div>
        </div>
      </div>
    </div>`;
  }).join('');

  tableHtml += `</div>`;
  container.innerHTML = tableHtml;
  
  // Add mobile touch swipe left/right detection to each row
  setTimeout(() => {
    document.querySelectorAll('.employee-grid-row').forEach(row => {
      let touchStartX = 0;
      let touchEndX = 0;
      
      row.addEventListener('touchstart', (e) => {
        touchStartX = e.changedTouches[0].screenX;
      }, { passive: true });
      
      row.addEventListener('touchend', (e) => {
        touchEndX = e.changedTouches[0].screenX;
        const diffX = touchStartX - touchEndX;
        const id = row.getAttribute('data-id');
        if (diffX > 50) {
          toggleRowActions(id, true);
        } else if (diffX < -50) {
          toggleRowActions(id, false);
        }
      }, { passive: true });
    });
  }, 100);
}

function toggleRowActions(id, show) {
  const row = document.querySelector(`.employee-grid-row[data-id="${id}"]`);
  if (row) {
    if (show) {
      row.classList.add('swipe-active');
    } else {
      row.classList.remove('swipe-active');
    }
  }
}
window.toggleRowActions = toggleRowActions;

function openEmployeeModal(id) {
  if (!canManageEmployees()) { showToast('Access denied.', 'error'); return; }
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
    ['emp-first','emp-last','emp-empid','emp-dept','emp-phone','emp-notes'].forEach(fid => document.getElementById(fid).value = '');
    document.getElementById('emp-centre').value = '';
  }
  openModal('employee-modal');
}

function saveEmployee() {
  if (!canManageEmployees()) { showToast('Access denied.', 'error'); return; }
  const first = document.getElementById('emp-first').value.trim();
  const last  = document.getElementById('emp-last').value.trim();
  if (!first || !last) { showToast('First and last name required.', 'error'); return; }
  const id   = document.getElementById('emp-id').value;
  const data = {
    id: id || uid(), firstName: first, lastName: last,
    employeeId: document.getElementById('emp-empid').value.trim(),
    productionCentre: document.getElementById('emp-centre').value,
    department: document.getElementById('emp-dept').value.trim(),
    phone: document.getElementById('emp-phone').value.trim(),
    notes: document.getElementById('emp-notes').value.trim(),
    createdAt: new Date().toISOString()
  };
  let list = DB.employees;
  if (id) {
    const idx = list.findIndex(e => e.id === id);
    if (idx >= 0) { data.createdAt = list[idx].createdAt; list[idx] = data; }
  } else {
    list.push(data);
  }
  DB.saveEmployees(list);
  CloudSync.pushEmployee(data);
  closeModal('employee-modal');
  showToast(`Employee ${first} ${last} saved!`, 'success');
  renderEmployeeList();
}

function deleteEmployee(id) {
  if (!canManageEmployees()) { showToast('Access denied.', 'error'); return; }
  const e = DB.employees.find(emp => emp.id === id);
  if (!e) return;
  confirm_dialog(`Delete ${e.firstName} ${e.lastName}?`, 'This employee will be removed. Transaction history is kept.', () => {
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

  // FIXED: Use global latest-by-barcode for accurate held count
  const latestByBarcode = getLatestByBarcode();
  const myBarcodes = [...new Set(
    DB.transactions.filter(t => t.employeeId === id).map(t => t.barcode)
  )];
  const held = myBarcodes.filter(bc => {
    const latest = latestByBarcode[bc];
    return latest && latest.action === 'distributed' && latest.employeeId === id;
  }).length;

  document.getElementById('emp-detail-info').innerHTML = `
    <div style="display:grid; grid-template-columns: repeat(auto-fit, minmax(150px, 1fr)); gap:12px; margin-bottom:20px">
      <div class="emp-detail-field" style="background:rgba(59,130,246,0.05); padding:12px; border-radius:8px">
        <label style="font-weight:600; color:var(--blue)">Uniforms On Hold</label>
        <span style="font-size:1.8rem; font-weight:bold; color:var(--blue)">${held}</span>
      </div>
      <div class="emp-detail-field" style="background:rgba(107,114,128,0.05); padding:12px; border-radius:8px">
        <label style="font-weight:600; color:var(--text2)">Total Uniforms</label>
        <span style="font-size:1.8rem; font-weight:bold; color:var(--text2)">${myBarcodes.length}</span>
      </div>
    </div>
    ${[
      ['Employee ID', e.employeeId||'—'],
      ['Production Centre', e.productionCentre||'—'],
      ['Department', e.department||'—'],
      ['Phone', e.phone||'—'],
      ['Notes', e.notes||'—']
    ].map(([l,v]) => `<div class="emp-detail-field"><label>${l}</label><span>${escHtml(v)}</span></div>`).join('')}
  `;

  const txns = DB.transactions.filter(t => t.employeeId === id);
  const detailEl = document.getElementById('emp-detail-uniforms');
  if (!txns.length) {
    detailEl.innerHTML = '<div class="empty-state"><p>No uniform transactions found.</p></div>';
  } else {
    detailEl.innerHTML = `<div class="report-table-wrap"><table class="report-table">
      <thead><tr><th>Barcode</th><th>Action</th><th>Date</th><th>Notes</th></tr></thead>
      <tbody>${txns.sort((a,b) => new Date(b.date) - new Date(a.date)).map(t => `<tr class="${t.inferred ? 'inferred-row' : ''}">
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
  if (!hasPermission(action)) {
    showToast(`Access denied for "${actionLabel(action)}".`, 'error');
    return;
  }
  currentAction = action;
  document.querySelectorAll('.action-tab').forEach(t => t.classList.remove('active'));
  const tab = document.getElementById('tab-' + action);
  if (tab) tab.classList.add('active');

  const selector = document.getElementById('employee-selector');
  const needsEmployee = action === 'distributed' || action === 'returned' || action === 'collected_from_soil_bin';
  selector.style.display = needsEmployee ? 'block' : 'none';

  const hints = {
    received_from_cintas: 'Scan each item arriving from Cintas',
    distributed:          'Select employee above — then scan each item',
    reported_issue:       'Scan damaged / dirty items to flag them',
    returned:             'Scan each item being returned to warehouse',
    collected_from_soil_bin: 'Select employee — then scan items collected from soil bin',
    sent_to_cintas:       'Scan each item going back to Cintas'
  };
  const hintEl = document.getElementById('scannerHintText');
  if (hintEl) hintEl.textContent = hints[action] || '';

  // Clear state info display
  const stateInfoEl = document.getElementById('scan-state-info');
  if (stateInfoEl) stateInfoEl.innerHTML = '';

  setTimeout(() => {
    const inp = document.getElementById('barcodeInput');
    if (inp) inp.focus();
  }, 80);
}

// Focus lock
function startFocusLock() {
  stopFocusLock();
  focusLockInterval = setInterval(() => {
    if (currentPage !== 'scan') { stopFocusLock(); return; }
    const active = document.activeElement;
    const inp    = document.getElementById('barcodeInput');
    const allowed = ['scan-emp-search', 'barcodeInput', 'scan-date', 'manual-barcode', 'manual-notes'];
    if (inp && active && !allowed.includes(active.id) && !active.closest('.modal-overlay')) {
      inp.focus();
    }
  }, 350);
}
function stopFocusLock() {
  if (focusLockInterval) { clearInterval(focusLockInterval); focusLockInterval = null; }
}

// Scanner zone UI
function onScannerFocus() {
  const dot = document.getElementById('scannerDot');
  const status = document.getElementById('scannerStatusText');
  const zone = document.getElementById('scannerZone');
  if (dot) dot.classList.add('active');
  if (status) status.textContent = 'SCANNER ACTIVE';
  if (zone) zone.classList.add('focused');
}

function onScannerBlur() {
  setTimeout(() => {
    if (document.activeElement && document.activeElement.id === 'barcodeInput') return;
    const dot = document.getElementById('scannerDot');
    const status = document.getElementById('scannerStatusText');
    const zone = document.getElementById('scannerZone');
    if (dot) dot.classList.remove('active');
    if (status) status.textContent = 'SCANNER READY';
    if (zone) zone.classList.remove('focused');
  }, 100);
}

function onBarcodeInput() {
  const val    = document.getElementById('barcodeInput').value;
  const status = document.getElementById('scannerStatusText');
  if (status && val) {
    const validation = validateBarcode(val);
    status.textContent = validation.valid ? `✓ READY (${val.length} chars)` : `⚠ INCOMPLETE (${val.length}/${BARCODE_CONFIG.minLength} chars)`;
  }
}

function filterScanEmployees() { showEmpDropdown(); }

function showEmpDropdown() {
  const query     = document.getElementById('scan-emp-search').value.toLowerCase();
  const dd        = document.getElementById('scan-emp-dropdown');
  const employees = DB.employees;
  const filtered  = query ? employees.filter(e =>
    `${e.firstName} ${e.lastName}`.toLowerCase().includes(query) ||
    String(e.employeeId||'').toLowerCase().includes(query) ||
    String(e.productionCentre||'').toLowerCase().includes(query)
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
  setTimeout(() => document.getElementById('barcodeInput').focus(), 30);
}

function clearEmployee() {
  selectedEmployee = null;
  document.getElementById('scan-emp-search').value = '';
  document.getElementById('clearEmpBtn').style.display = 'none';
  document.getElementById('selected-employee-badge').classList.add('hidden');
  document.getElementById('scan-emp-dropdown').classList.add('hidden');
}

// ============================================================
// BARCODE VALIDATION
// ============================================================
const BARCODE_CONFIG = { minLength: 13, maxLength: 15 };

function playErrorSound() {
  try {
    const ctx = new (window.AudioContext || window.webkitAudioContext)();
    const osc = ctx.createOscillator();
    const gain = ctx.createGain();
    osc.connect(gain);
    gain.connect(ctx.destination);
    osc.frequency.setValueAtTime(200, ctx.currentTime);
    osc.frequency.setValueAtTime(150, ctx.currentTime + 0.1);
    gain.gain.setValueAtTime(0.3, ctx.currentTime);
    gain.gain.setValueAtTime(0, ctx.currentTime + 0.15);
    osc.start(ctx.currentTime);
    osc.stop(ctx.currentTime + 0.15);
  } catch (e) {}
}

function validateBarcode(barcode) {
  barcode = (barcode || '').trim();
  if (!barcode) return { valid: false, reason: 'Barcode is empty' };
  if (!/^\d+$/.test(barcode)) return { valid: false, reason: 'Barcode must contain only digits (0-9)' };
  if (barcode.length < BARCODE_CONFIG.minLength) return { valid: false, reason: `Too short (${barcode.length}). Must be 13-15 digits.` };
  if (barcode.length > BARCODE_CONFIG.maxLength) return { valid: false, reason: `Too long (${barcode.length}). Must be 13-15 digits.` };
  return { valid: true, reason: '' };
}

function handleBarcodeKey(e) {
  if (e.key === 'Enter') {
    e.preventDefault();
    const barcode = document.getElementById('barcodeInput').value.trim();
    const validation = validateBarcode(barcode);
    if (!validation.valid) {
      playErrorSound();
      showToast(`⚠ Invalid Barcode: ${validation.reason}`, 'error');
      document.getElementById('barcodeInput').value = '';
      return;
    }
    submitScan();
  }
}

function checkDuplicate(barcode, action, date) {
  return DB.transactions.find(t => t.barcode === barcode && t.action === action && t.date === date) || null;
}

function flashDuplicateWarning(barcode) {
  const zone   = document.getElementById('scannerZone');
  const dot    = document.getElementById('scannerDot');
  const status = document.getElementById('scannerStatusText');
  if (zone) { zone.classList.add('dup-warning'); setTimeout(() => zone.classList.remove('dup-warning'), 1500); }
  if (dot)  { dot.classList.add('dup-dot'); setTimeout(() => dot.classList.remove('dup-dot'), 1500); }
  if (status) {
    status.textContent = `⚠ DUPLICATE: ${barcode}`;
    setTimeout(() => { if (status) status.textContent = 'SCANNER ACTIVE'; }, 1600);
  }
}

// ============================================================
// SMART SCAN SUBMISSION (with State Machine + Inference)
// ============================================================
function submitScan() {
  const barcode = document.getElementById('barcodeInput').value.trim();
  const validation = validateBarcode(barcode);
  if (!validation.valid) {
    playErrorSound();
    showToast(`⚠ Invalid: ${validation.reason}`, 'error');
    return;
  }

  const needsEmployee = currentAction === 'distributed' || currentAction === 'returned' || currentAction === 'collected_from_soil_bin';
  if ((currentAction === 'distributed' || currentAction === 'collected_from_soil_bin') && !selectedEmployee) {
    showToast('Please select an employee first.', 'error');
    document.getElementById('barcodeInput').value = '';
    document.getElementById('scan-emp-search').focus();
    return;
  }

  const date = normalizeDate(document.getElementById('scan-date').value || today());

  // Duplicate check
  const dup = checkDuplicate(barcode, currentAction, date);
  if (dup) {
    flashDuplicateWarning(barcode);
    const dupEmp = dup.employeeName ? ` (${dup.employeeName})` : '';
    confirm_dialog(
      `⚠ Duplicate Barcode Detected`,
      `Barcode "${barcode}" was already logged as "${actionLabel(currentAction)}"${dupEmp} on ${formatDate(date)}. Save again?`,
      () => doSaveScanWithInference(barcode, currentAction, date, '')
    );
    document.getElementById('barcodeInput').value = '';
    setTimeout(() => document.getElementById('barcodeInput').focus(), 20);
    return;
  }

  // STATE MACHINE VALIDATION + INFERENCE
  const result = validateTransition(barcode, currentAction, date, selectedEmployee?.id);

  if (result.inferences.length > 0) {
    // Show inference confirmation
    const inferenceHtml = result.inferences.map(inf =>
      `<li><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 15V9m0 0l-3 3m3-3l3 3"/><circle cx="12" cy="12" r="10"/></svg>${actionLabel(inf.action)}${inf.employeeName ? ' — ' + escHtml(inf.employeeName) : ''} (${formatDate(inf.date)})</li>`
    ).join('');

    confirm_dialog(
      '⚙ Auto-Inferred Steps Required',
      `<div style="margin-bottom:12px">${result.warnings.join('<br>')}</div>
       <div style="font-size:0.85rem;color:var(--text2);margin-bottom:8px">The following steps will be auto-created:</div>
       <ul class="inference-list">${inferenceHtml}</ul>
       <div style="font-size:0.82rem;color:var(--text3)">These inferred steps are marked with ⚙ and can be edited later.</div>`,
      () => doSaveScanWithInference(barcode, currentAction, date, ''),
      true // use innerHTML for message
    );
    document.getElementById('barcodeInput').value = '';
    setTimeout(() => document.getElementById('barcodeInput').focus(), 20);
    return;
  }

  // Show warnings if any
  if (result.warnings.length > 0) {
    result.warnings.forEach(w => showToast(`ℹ ${w}`, 'info'));
  }

  doSaveScanWithInference(barcode, currentAction, date, '');
}

function doSaveScanWithInference(barcode, action, date, notes) {
  // Re-validate and get inferences
  const result = validateTransition(barcode, action, date, selectedEmployee?.id);

  // Save inferred transactions first
  for (const inf of result.inferences) {
    const inferTxn = {
      id: uid(),
      barcode: inf.barcode,
      action: inf.action,
      employeeId: inf.employeeId || null,
      employeeName: inf.employeeName || null,
      date: inf.date,
      notes: inf.notes,
      inferred: true,
      createdAt: new Date().toISOString()
    };
    DB.addTransaction(inferTxn);
    CloudSync.pushTransaction(inferTxn);
  }

  // Save actual scan
  doSaveScan(barcode, action, date, notes);
}

function doSaveScan(barcode, action, date, notes) {
  const needsEmployee = action === 'distributed' || action === 'returned' || action === 'collected_from_soil_bin';

  let finalNotes = notes;
  if (action === 'collected_from_soil_bin') {
    finalNotes = (notes ? notes + ' • ' : '') + 'Collected from soil bin';
  }

  const txn = {
    id: uid(), barcode, action,
    employeeId:   (needsEmployee && selectedEmployee) ? selectedEmployee.id   : null,
    employeeName: (needsEmployee && selectedEmployee) ? `${selectedEmployee.firstName} ${selectedEmployee.lastName}` : null,
    date, notes: finalNotes,
    inferred: false,
    createdAt: new Date().toISOString()
  };

  DB.addTransaction(txn);
  CloudSync.pushTransaction(txn);
  sessionScans++;

  // Visual feedback
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

// Manual entry
function openManualEntry() {
  document.getElementById('manual-barcode').value = '';
  const mn = document.getElementById('manual-notes');
  if (mn) mn.value = '';
  openModal('manual-entry-modal');
  setTimeout(() => document.getElementById('manual-barcode').focus(), 200);
}

function manualSubmit() {
  const barcode = document.getElementById('manual-barcode').value.trim();
  const validation = validateBarcode(barcode);
  if (!validation.valid) { playErrorSound(); showToast(validation.reason, 'error'); return; }

  if ((currentAction === 'distributed' || currentAction === 'collected_from_soil_bin') && !selectedEmployee) {
    showToast('Please select an employee first.', 'error');
    closeModal('manual-entry-modal');
    document.getElementById('scan-emp-search').focus();
    return;
  }

  const date  = normalizeDate(document.getElementById('scan-date').value || today());
  const notes = (document.getElementById('manual-notes')?.value || '').trim();

  const dup = checkDuplicate(barcode, currentAction, date);
  if (dup) {
    confirm_dialog(
      `⚠ Duplicate Barcode`,
      `"${barcode}" already logged as "${actionLabel(currentAction)}" on ${formatDate(date)}. Save again?`,
      () => {
        doSaveScanWithInference(barcode, currentAction, date, notes);
        document.getElementById('manual-barcode').value = '';
        if (document.getElementById('manual-notes')) document.getElementById('manual-notes').value = '';
        setTimeout(() => document.getElementById('manual-barcode').focus(), 80);
      }
    );
    return;
  }

  doSaveScanWithInference(barcode, currentAction, date, notes);
  document.getElementById('manual-barcode').value = '';
  if (document.getElementById('manual-notes')) document.getElementById('manual-notes').value = '';
  setTimeout(() => document.getElementById('manual-barcode').focus(), 80);
}

function updateScanFeed(txn) {
  const feed    = document.getElementById('scan-feed');
  const actionKey = txn.action || '';
  const empStr  = txn.employeeName || (actionKey === 'reported_issue' ? 'Issue → Cintas' : actionKey.includes('cintas') ? 'Cintas' : 'Warehouse');
  const empty   = feed.querySelector('.empty-state');
  if (empty) empty.remove();
  const item = document.createElement('div');
  item.className = 'scan-feed-item new' + (txn.inferred ? ' inferred' : '');
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
  barcodeHistoryPage = 0; // reset pagination on tab change
  document.querySelectorAll('.report-tab').forEach(t => t.classList.remove('active'));
  const tab = document.querySelector(`.report-tab[data-report="${r}"]`);
  if (tab) tab.classList.add('active');
  renderReport();
}

// ============================================================
// BARCODE HISTORY PAGINATION STATE
// ============================================================
let barcodeHistoryPage = 0;
const BARCODE_HISTORY_PAGE_SIZE = 100;

// FIX 2 helper: split a barcode's transactions (ascending) into labelled cycles.
// A new cycle starts whenever action is 'distributed' or 'received_from_cintas'
// and there is already content in the current accumulator.
function splitIntoCycles(txns) {
  const cycles = [];
  let current = [];
  txns.forEach(t => {
    if ((t.action === 'distributed' || t.action === 'received_from_cintas') && current.length > 0) {
      cycles.push(current);
      current = [];
    }
    current.push(t);
  });
  if (current.length) cycles.push(current);
  return cycles.reverse(); // newest first
}

function renderReport() {
  const from      = document.getElementById('report-from').value;
  const to        = document.getElementById('report-to').value;
  const q         = (document.getElementById('report-search').value || '').toLowerCase();
  const container = document.getElementById('report-content');
  const txns      = DB.transactions;
  const employees = DB.employees;

  const latestByBarcode = getLatestByBarcode();

  // FIX 4: Pre-compute barcode → current state map once (avoids O(n²) resolveState per row)
  const barcodeStateMap = {};
  Object.entries(latestByBarcode).forEach(([bc, txn]) => {
    barcodeStateMap[bc] = ACTION_TO_STATE[txn.action] || 'unknown';
  });

  function inRange(dateStr) {
    if (!dateStr) return true;
    if (from && dateStr < from) return false;
    if (to   && dateStr > to)   return false;
    return true;
  }

  // ---- MISSING UNIFORMS (FIX 9: always show ALL currently missing, no date gate) ----
  if (currentReport === 'missing') {
    // Read extra filters specific to missing report
    const centreFilter = (document.getElementById('missing-centre-filter')?.value || '');
    const overdueOnly  = document.getElementById('missing-overdue-toggle')?.checked || false;

    const missingItems = [];
    Object.entries(latestByBarcode).forEach(([bc, lastTxn]) => {
      if (lastTxn.action !== 'distributed') return; // only items currently out with employee
      const emp = findEmployeeByIdOrExternal(lastTxn.employeeId);
      const row = {
        barcode: bc,
        employeeId: lastTxn.employeeId,
        employeeName: emp ? `${emp.firstName} ${emp.lastName}` : '— Unknown —',
        employeeEmpId: emp?.employeeId || '',
        centre: emp?.productionCentre || '—',
        department: emp?.department || '—',
        date: lastTxn.date,
        notes: lastTxn.notes || '',
        days: Math.floor((Date.now() - new Date(lastTxn.date)) / 86400000)
      };
      // FIX 9: date filter repurposed as "Distributed After" secondary filter
      if (from && lastTxn.date < from) return;
      if (centreFilter && row.centre !== centreFilter) return;
      if (overdueOnly && row.days <= 7) return;
      if (q && !String(bc).toLowerCase().includes(q) &&
               !String(row.employeeName).toLowerCase().includes(q) &&
               !String(row.centre).toLowerCase().includes(q) &&
               !String(row.employeeEmpId).toLowerCase().includes(q)) return;
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
    const avgDays = missingItems.length ? Math.round(missingItems.reduce((s,r)=>s+r.days,0)/missingItems.length) : 0;
    const allCentres = [...new Set(DB.employees.map(e=>e.productionCentre).filter(Boolean))].sort();

    container.innerHTML = `
      <div style="display:flex;gap:10px;align-items:center;flex-wrap:wrap;margin-bottom:12px">
        <select id="missing-centre-filter" class="form-input" style="flex:1;min-width:160px;max-width:220px" onchange="renderReport()">
          <option value="">All Centres</option>
          ${allCentres.map(c=>`<option value="${escHtml(c)}" ${centreFilter===c?'selected':''}>${escHtml(c)}</option>`).join('')}
        </select>
        <label style="display:flex;align-items:center;gap:6px;font-size:0.88rem;color:var(--text2);cursor:pointer">
          <input type="checkbox" id="missing-overdue-toggle" ${overdueOnly?'checked':''} onchange="renderReport()" style="width:16px;height:16px">
          Overdue only (&gt;7 days)
        </label>
        <span style="margin-left:auto;font-size:0.82rem;color:var(--text3)">Avg days out: <strong>${avgDays}d</strong></span>
        <button class="btn-secondary" onclick="exportCSV()">Export CSV</button>
      </div>
      <div class="missing-banner ${missingItems.length > 0 ? 'has-missing' : 'all-clear'}">
        <div class="missing-banner-icon">${missingItems.length > 0 ? '⚠' : '✓'}</div>
        <div>
          <div class="missing-banner-title">${missingItems.length > 0 ? missingItems.length + ' Uniform' + (missingItems.length !== 1 ? 's' : '') + ' Not Returned' : 'All Uniforms Returned!'}</div>
          <div class="missing-banner-sub">${missingItems.length > 0 ? urgentCount + ' item' + (urgentCount !== 1 ? 's' : '') + ' overdue >7 days' : 'No outstanding uniforms.'}</div>
        </div>
      </div>
      ${Object.values(byEmp).map(grp => `
        <div class="missing-emp-group">
          <div class="missing-emp-header">
            <div class="emp-avatar" style="width:36px;height:36px;font-size:0.82rem;background:${avatarColor(grp.name)}">${grp.name.split(' ').map(w=>w[0]).join('').slice(0,2)}</div>
            <div>
              <div style="font-weight:700;color:var(--text)">${escHtml(grp.name)} <span style="color:var(--text3);font-size:0.78rem;font-weight:400">${escHtml(grp.empId ? '· ' + grp.empId : '')}</span></div>
              <div style="font-size:0.76rem;color:var(--text3)">${escHtml(grp.centre)}${grp.dept ? ' · ' + escHtml(grp.dept) : ''}</div>
            </div>
            <span class="missing-count-badge">${grp.items.length} item${grp.items.length !== 1 ? 's' : ''}</span>
          </div>
          <div class="report-table-wrap" style="margin:0;border-top:none;border-radius:0 0 10px 10px">
            <table class="report-table">
              <thead><tr><th>Barcode</th><th>Distributed After</th><th>Days Out</th><th>Notes</th></tr></thead>
              <tbody>${grp.items.map(r => `<tr>
                <td class="bc-mono">${escHtml(r.barcode)}</td>
                <td>${formatDate(r.date)}</td>
                <td><span class="days-pill ${r.days > 14 ? 'overdue' : r.days > 7 ? 'warning' : 'ok'}">${r.days}d</span></td>
                <td style="color:var(--text3)">${escHtml(r.notes)}</td>
              </tr>`).join('')}</tbody>
            </table>
          </div>
        </div>
      `).join('')}
      ${missingItems.length === 0 ? '<div class="empty-state"><p>No missing uniforms found.</p></div>' : ''}
    `;

  // ---- ISSUES / DAMAGED (FIX 5: show last employee who held each item) ----
  } else if (currentReport === 'issues') {
    // Build barcode → all transactions map for lookups
    const txnsByBarcode = {};
    txns.forEach(t => { if(!txnsByBarcode[t.barcode]) txnsByBarcode[t.barcode] = []; txnsByBarcode[t.barcode].push(t); });

    const rows = [];
    Object.entries(latestByBarcode).forEach(([bc, last]) => {
      if (last.action !== 'reported_issue') return;
      if (!inRange(last.date)) return;
      if (q && !String(bc).toLowerCase().includes(q)) return;
      // FIX 5: find last employee who held this item (most recent distributed txn)
      const bcTxns = (txnsByBarcode[bc] || []).slice().sort((a,b)=>new Date(b.date)-new Date(a.date));
      const lastDist = bcTxns.find(t => t.action === 'distributed');
      const lastEmpName = lastDist?.employeeName || '—';
      rows.push({ barcode: bc, date: last.date, notes: last.notes || '', lastEmployee: lastEmpName });
    });
    container.innerHTML = `
      <div class="report-summary">
        <div class="rs-card"><div class="rs-num" style="color:var(--yellow)">${rows.length}</div><div class="rs-label">Items Damaged/Dirty</div></div>
      </div>
      <div class="report-table-wrap"><table class="report-table">
        <thead><tr><th>Barcode</th><th>Date Reported</th><th>Last Held By</th><th>Notes</th></tr></thead>
        <tbody>${rows.length ? rows.map(r => `<tr>
          <td class="bc-mono">${escHtml(r.barcode)}</td>
          <td>${formatDate(r.date)}</td>
          <td style="color:var(--text2)">${escHtml(r.lastEmployee)}</td>
          <td>${escHtml(r.notes)}</td>
        </tr>`).join('') : '<tr><td colspan="4" style="text-align:center;color:var(--text3);padding:30px">No issues reported.</td></tr>'}</tbody>
      </table></div>`;

  // ---- CINTAS WORKFLOW (FIX 2+10: all cycles stacked, newest first; correct date filter per cycle) ----
  } else if (currentReport === 'cintas-workflow') {
    const barcodes = [...new Set(txns.map(t => t.barcode))];
    // Each entry: { barcode, cycleIndex, cycleLabel, ...cycle data }
    const allCycleRows = [];

    barcodes.forEach(bc => {
      const allTxns = txns.filter(t => t.barcode === bc).sort((a, b) => new Date(a.date) - new Date(b.date));
      if (!allTxns.length) return;

      if (q && !String(bc).toLowerCase().includes(q)) {
        // Also check employee names across all txns
        const hasEmpMatch = allTxns.some(t => t.employeeName && t.employeeName.toLowerCase().includes(q));
        if (!hasEmpMatch) return;
      }

      const cycles = splitIntoCycles(allTxns); // newest first
      const totalCycles = cycles.length;

      cycles.forEach((cycleTxns, idx) => {
        const cycleNum = totalCycles - idx; // e.g. totalCycles=2 → idx0=Cycle2, idx1=Cycle1
        const isCurrent = idx === 0;
        const cycleLabel = totalCycles > 1
          ? `Cycle ${cycleNum}${isCurrent ? ' — Current' : ' — Completed'}`
          : null; // single cycle: no label needed

        const distributed   = cycleTxns.find(t => t.action === 'distributed');
        const returned      = cycleTxns.find(t => t.action === 'returned');
        const collected     = cycleTxns.find(t => t.action === 'collected_from_soil_bin');
        const sentToCintas  = cycleTxns.find(t => t.action === 'sent_to_cintas');
        const receivedBack  = cycleTxns.find(t => t.action === 'received_from_cintas');

        if (!distributed && !sentToCintas && !receivedBack) return;

        // FIX 10: date filter applied to THIS cycle's earliest date, not first-ever
        const cycleDate = distributed?.date || sentToCintas?.date || receivedBack?.date || '';
        if (!inRange(cycleDate)) return;

        const emp = distributed ? findEmployeeByIdOrExternal(distributed.employeeId) : null;
        const hasInferred = cycleTxns.some(t => t.inferred);

        // FIX 14: If returned/collected step was auto-inferred, show Auto-Processed, not Skipped Return
        const returnWasInferred = (returned?.inferred || collected?.inferred);

        let status = '⚠ Incomplete';
        let statusColor = 'var(--orange)';
        if (distributed && (returned || collected) && sentToCintas) {
          status = '✓ Complete'; statusColor = 'var(--green)';
        } else if (distributed && sentToCintas && !returned && !collected) {
          status = returnWasInferred ? '✓ Auto-Processed' : '⚠ Skipped Return';
          statusColor = returnWasInferred ? 'var(--green)' : 'var(--purple)';
        } else if (distributed && (returned || collected) && !sentToCintas) {
          status = '→ Ready for Cintas'; statusColor = 'var(--blue)';
        } else if (distributed && !returned && !collected && !sentToCintas) {
          status = '⏳ With Employee'; statusColor = 'var(--yellow)';
        } else if (sentToCintas && receivedBack) {
          status = '✓ Received Back'; statusColor = 'var(--green)';
        } else if (sentToCintas) {
          status = '🔄 At Cintas'; statusColor = 'var(--purple)';
        }

        const isComplete   = !!(distributed && (returned || collected) && sentToCintas);
        const isReadyForCintas = !!(distributed && (returned || collected) && !sentToCintas);

        allCycleRows.push({
          barcode: bc, cycleLabel, isCurrent, cycleNum,
          status, statusColor, hasInferred,
          distributedDate: distributed?.date || '—',
          distributedEmp: emp ? `${emp.firstName} ${emp.lastName}` : (distributed?.employeeName || '—'),
          returnedDate: returned?.date || collected?.date || '—',
          sentToCintasDate: sentToCintas?.date || '—',
          receivedBackDate: receivedBack?.date || '—',
          isComplete, isReadyForCintas
        });
      });
    });

    // Sort: ready for Cintas first, then incomplete, then complete; within group by barcode
    allCycleRows.sort((a, b) => {
      if (a.isReadyForCintas !== b.isReadyForCintas) return a.isReadyForCintas ? -1 : 1;
      if (a.isComplete !== b.isComplete) return a.isComplete ? 1 : -1;
      const bc = a.barcode.localeCompare(b.barcode);
      if (bc !== 0) return bc;
      return b.cycleNum - a.cycleNum; // newest cycle first within same barcode
    });

    const totalBarcodes = new Set(allCycleRows.map(r => r.barcode)).size;
    const complete       = allCycleRows.filter(i => i.isComplete).length;
    const readyForCintas = allCycleRows.filter(i => i.isReadyForCintas).length;
    const incomplete     = allCycleRows.filter(i => !i.isComplete && !i.isReadyForCintas).length;

    container.innerHTML = `
      <div class="report-summary">
        <div class="rs-card"><div class="rs-num" style="color:var(--green)">${complete}</div><div class="rs-label">Complete Cycles</div></div>
        <div class="rs-card"><div class="rs-num" style="color:var(--blue)">${readyForCintas}</div><div class="rs-label">Ready for Cintas</div></div>
        <div class="rs-card"><div class="rs-num" style="color:var(--orange)">${incomplete}</div><div class="rs-label">Incomplete</div></div>
        <div class="rs-card"><div class="rs-num" style="color:var(--text2)">${totalBarcodes}</div><div class="rs-label">Barcodes</div></div>
      </div>
      <div class="report-table-wrap"><table class="report-table" style="font-size:0.88rem">
        <thead><tr><th>Barcode</th><th>Cycle</th><th>Status</th><th>Distributed</th><th>Returned</th><th>→ Cintas</th><th>← Received</th></tr></thead>
        <tbody>${allCycleRows.length ? allCycleRows.map(item => `<tr class="${item.hasInferred ? 'inferred-row' : ''}" style="background:${item.isCurrent ? (item.status.includes('✓') ? 'rgba(34,197,94,0.06)' : item.status.includes('→') ? 'rgba(59,130,246,0.06)' : 'rgba(249,115,22,0.06)') : 'transparent'}">
          <td class="bc-mono"><strong>${escHtml(item.barcode)}</strong></td>
          <td style="white-space:nowrap;font-size:0.78rem;color:${item.isCurrent?'var(--text2)':'var(--text3)'}">${item.cycleLabel ? escHtml(item.cycleLabel) : '<em style="color:var(--text3)">—</em>'}</td>
          <td><span class="report-badge" style="background:${item.statusColor}; color:white; padding:3px 7px; border-radius:4px; font-size:0.82rem;white-space:nowrap">${item.status}</span></td>
          <td><small>${item.distributedDate === '—' ? '<em style="color:var(--text3)">—</em>' : formatDate(item.distributedDate)}<br><em style="color:var(--text3);font-size:0.75rem">${escHtml(item.distributedEmp)}</em></small></td>
          <td><small>${item.returnedDate === '—' ? '<em style="color:var(--text3)">Pending…</em>' : formatDate(item.returnedDate)}</small></td>
          <td><small>${item.sentToCintasDate === '—' ? '<em style="color:var(--text3)">Pending…</em>' : formatDate(item.sentToCintasDate)}</small></td>
          <td><small>${item.receivedBackDate === '—' ? '<em style="color:var(--text3)">—</em>' : formatDate(item.receivedBackDate)}</small></td>
        </tr>`).join('') : '<tr><td colspan="7" style="text-align:center;color:var(--text3);padding:30px">No workflow data.</td></tr>'}</tbody>
      </table></div>`;

  // ---- BY EMPLOYEE (FIX 6: renamed "Total"→"Ever Handled"; sort by holding desc; dept column) ----
  } else if (currentReport === 'employee-summary') {
    const empList = employees.filter(e => !q || `${e.firstName} ${e.lastName}`.toLowerCase().includes(q) || String(e.employeeId||'').toLowerCase().includes(q));
    const rows = empList.map(e => {
      const myTxns = txns.filter(t => t.employeeId === e.id || t.employeeId === e.employeeId);
      const myBarcodes = [...new Set(myTxns.map(t => t.barcode))];
      const held = myBarcodes.filter(bc => {
        const l = latestByBarcode[bc];
        return l && l.action === 'distributed' && (l.employeeId === e.id || l.employeeId === e.employeeId);
      }).length;
      return { e, held, total: myBarcodes.length };
    }).filter(r => r.total > 0 || !q);

    // FIX 6: sort by currently holding (most burdened employees first)
    rows.sort((a, b) => b.held - a.held || b.total - a.total);

    container.innerHTML = `
      <div class="report-summary">
        <div class="rs-card"><div class="rs-num">${employees.length}</div><div class="rs-label">Total Employees</div></div>
        <div class="rs-card"><div class="rs-num" style="color:var(--orange)">${rows.reduce((a,r)=>a+r.held,0)}</div><div class="rs-label">Currently Out</div></div>
      </div>
      <div class="report-table-wrap"><table class="report-table">
        <thead><tr><th>Employee</th><th>ID</th><th>Centre</th><th>Dept</th><th>Holding ↓</th><th>Ever Handled</th></tr></thead>
        <tbody>${rows.length ? rows.map(r => `<tr>
          <td>${escHtml(r.e.firstName)} ${escHtml(r.e.lastName)}</td>
          <td class="bc-mono">${escHtml(r.e.employeeId||'—')}</td>
          <td>${escHtml(r.e.productionCentre||'—')}</td>
          <td style="color:var(--text3)">${escHtml(r.e.department||'—')}</td>
          <td><strong style="color:${r.held>0?'var(--orange)':'var(--text3)'}">${r.held}</strong></td>
          <td style="color:var(--text3)">${r.total}</td>
        </tr>`).join('') : '<tr><td colspan="6" style="text-align:center;color:var(--text3);padding:30px">No data.</td></tr>'}</tbody>
      </table></div>`;

  // ---- AT CINTAS (FIX 3: only sent_to_cintas; removed incorrect reported_issue inclusion) ----
  } else if (currentReport === 'cintas') {
    const rows = [];
    Object.entries(latestByBarcode).forEach(([bc, last]) => {
      if (last.action !== 'sent_to_cintas') return; // FIX 3: damaged items excluded
      if (!inRange(last.date)) return;
      if (q && !String(bc).toLowerCase().includes(q)) return;
      rows.push({ barcode: bc, date: last.date, days: Math.floor((Date.now()-new Date(last.date))/86400000), notes: last.notes||'' });
    });
    const longStay = rows.filter(r=>r.days>14).length;
    container.innerHTML = `
      <div class="report-summary">
        <div class="rs-card"><div class="rs-num" style="color:var(--purple)">${rows.length}</div><div class="rs-label">At Cintas</div></div>
        <div class="rs-card"><div class="rs-num" style="color:var(--red)">${longStay}</div><div class="rs-label">&gt;14 Days at Cintas</div></div>
      </div>
      <div class="report-table-wrap"><table class="report-table">
        <thead><tr><th>Barcode</th><th>Date Sent</th><th>Days at Cintas</th><th>Notes</th></tr></thead>
        <tbody>${rows.length ? rows.map(r => `<tr>
          <td class="bc-mono">${escHtml(r.barcode)}</td>
          <td>${formatDate(r.date)}</td>
          <td><span style="color:${r.days>14?'var(--red)':r.days>7?'var(--orange)':'var(--text2)'}">${r.days}d</span></td>
          <td>${escHtml(r.notes)}</td>
        </tr>`).join('') : '<tr><td colspan="4" style="text-align:center;color:var(--text3);padding:30px">No uniforms at Cintas.</td></tr>'}</tbody>
      </table></div>`;

  // ---- WAREHOUSE ----
  } else if (currentReport === 'warehouse') {
    const rows = [];
    Object.entries(latestByBarcode).forEach(([bc, last]) => {
      const state = ACTION_TO_STATE[last.action];
      if (state !== 'warehouse' && state !== 'soil_bin') return;
      if (!inRange(last.date)) return;
      if (q && !String(bc).toLowerCase().includes(q)) return;
      const howDisplay = last.action === 'collected_from_soil_bin' ? 'Collected from Soil Bin' :
                         last.action === 'returned' ? 'Returned' : 'Received from Cintas';
      rows.push({ barcode: bc, date: last.date, how: last.action, howDisplay, notes: last.notes||'' });
    });
    container.innerHTML = `
      <div class="report-summary">
        <div class="rs-card"><div class="rs-num" style="color:var(--blue)">${rows.length}</div><div class="rs-label">In Warehouse</div></div>
        <div class="rs-card"><div class="rs-num" style="color:var(--green)">${rows.filter(r=>r.how==='returned').length}</div><div class="rs-label">Employee Returns</div></div>
        <div class="rs-card"><div class="rs-num">${rows.filter(r=>r.how==='received_from_cintas').length}</div><div class="rs-label">From Cintas</div></div>
      </div>
      <div class="report-table-wrap"><table class="report-table">
        <thead><tr><th>Barcode</th><th>Date</th><th>Source</th><th>Notes</th></tr></thead>
        <tbody>${rows.length ? rows.map(r => `<tr>
          <td class="bc-mono">${escHtml(r.barcode)}</td>
          <td>${formatDate(r.date)}</td>
          <td><span class="report-badge ${r.how==='collected_from_soil_bin'?'issue':r.how==='returned'?'returned':'warehouse'}">${r.howDisplay}</span></td>
          <td>${escHtml(r.notes)}</td>
        </tr>`).join('') : '<tr><td colspan="4" style="text-align:center;color:var(--text3);padding:30px">No uniforms in warehouse.</td></tr>'}</tbody>
      </table></div>`;

  // ---- ACTIVITY SUMMARY (FIX 7+8: real vs inferred split; daily table respects date range) ----
  } else if (currentReport === 'activity-summary') {
    const filteredTxns = txns.filter(t => inRange(t.date));
    const realTxns     = filteredTxns.filter(t => !t.inferred);
    const inferredCount = filteredTxns.length - realTxns.length;
    const stats = {
      totalScans: realTxns.length, // FIX 8: real scans only
      uniqueBarcodes: new Set(filteredTxns.map(t => t.barcode)).size,
      uniqueEmployees: new Set(filteredTxns.filter(t => t.employeeId).map(t => t.employeeId)).size,
      byAction: {},
      byActionInferred: {},
      byDate: {},
      byEmployee: {},
      byCentre: {}
    };

    realTxns.forEach(t => { stats.byAction[t.action] = (stats.byAction[t.action] || 0) + 1; });
    filteredTxns.filter(t => t.inferred).forEach(t => { stats.byActionInferred[t.action] = (stats.byActionInferred[t.action] || 0) + 1; });
    // FIX 7: daily table uses ALL filtered txns (respects date range)
    filteredTxns.forEach(t => { stats.byDate[t.date] = (stats.byDate[t.date] || 0) + 1; });
    filteredTxns.filter(t => t.employeeId).forEach(t => {
      const emp = findEmployeeByIdOrExternal(t.employeeId);
      const name = emp ? `${emp.firstName} ${emp.lastName}` : 'Unknown';
      stats.byEmployee[name] = (stats.byEmployee[name] || 0) + 1;
    });
    filteredTxns.filter(t => t.employeeId).forEach(t => {
      const emp = findEmployeeByIdOrExternal(t.employeeId);
      const centre = emp?.productionCentre || 'Unknown';
      stats.byCentre[centre] = (stats.byCentre[centre] || 0) + 1;
    });

    const topEmployees  = Object.entries(stats.byEmployee).sort((a,b) => b[1] - a[1]).slice(0, 10);
    const topCentres    = Object.entries(stats.byCentre).sort((a,b) => b[1] - a[1]).slice(0, 10);
    // FIX 7: show all days in date range (up to 60), sorted ascending
    const dailyActivity = Object.entries(stats.byDate).sort((a,b) => a[0].localeCompare(b[0])).slice(-60);
    const allActions    = [...new Set([...Object.keys(stats.byAction), ...Object.keys(stats.byActionInferred)])];
    const totalForPct   = realTxns.length || 1;

    container.innerHTML = `
      <div class="report-summary" style="grid-template-columns: repeat(auto-fit, minmax(180px, 1fr))">
        <div class="rs-card"><div class="rs-num" style="color:var(--blue)">${stats.totalScans}</div><div class="rs-label">Real Scans</div><div style="font-size:0.72rem;color:var(--text3);margin-top:3px">+${inferredCount} auto-inferred</div></div>
        <div class="rs-card"><div class="rs-num" style="color:var(--green)">${stats.uniqueBarcodes}</div><div class="rs-label">Unique Barcodes</div></div>
        <div class="rs-card"><div class="rs-num" style="color:var(--orange)">${stats.uniqueEmployees}</div><div class="rs-label">Active Employees</div></div>
        <div class="rs-card"><div class="rs-num" style="color:var(--purple)">${inferredCount}</div><div class="rs-label">Auto-Inferred Steps</div></div>
      </div>
      <div style="display:grid; grid-template-columns:1fr 1fr; gap:20px; margin:20px 0; flex-wrap:wrap">
        <div class="card">
          <div class="card-header"><h3 class="card-title">Scans by Action</h3></div>
          <div class="report-table-wrap" style="max-height:300px"><table class="report-table">
            <thead><tr><th>Action</th><th>Real</th><th>Auto</th><th>%</th></tr></thead>
            <tbody>${allActions.sort((a,b)=>(stats.byAction[b]||0)-(stats.byAction[a]||0)).map(action => {
              const real = stats.byAction[action] || 0;
              const auto = stats.byActionInferred[action] || 0;
              return `<tr>
                <td><span class="report-badge ${action==='distributed'?'out':action==='returned'?'returned':action==='reported_issue'?'issue':action==='sent_to_cintas'?'cintas':'warehouse'}">${actionLabel(action)}</span></td>
                <td>${real}</td>
                <td style="color:var(--text3);font-size:0.82rem">${auto > 0 ? auto : '—'}</td>
                <td>${((real/totalForPct)*100).toFixed(1)}%</td>
              </tr>`;
            }).join('')}</tbody>
          </table></div>
        </div>
        <div class="card">
          <div class="card-header"><h3 class="card-title">Top Employees</h3></div>
          <div class="report-table-wrap" style="max-height:300px"><table class="report-table">
            <thead><tr><th>Employee</th><th>Scans</th></tr></thead>
            <tbody>${topEmployees.map(([name, count]) => `<tr><td>${escHtml(name)}</td><td>${count}</td></tr>`).join('')}</tbody>
          </table></div>
        </div>
      </div>
      <div style="display:grid; grid-template-columns:1fr 1fr; gap:20px; margin:20px 0">
        <div class="card">
          <div class="card-header"><h3 class="card-title">By Production Centre</h3></div>
          <div class="report-table-wrap" style="max-height:300px"><table class="report-table">
            <thead><tr><th>Centre</th><th>Scans</th></tr></thead>
            <tbody>${topCentres.map(([centre, count]) => `<tr><td>${escHtml(centre)}</td><td>${count}</td></tr>`).join('')}</tbody>
          </table></div>
        </div>
        <div class="card">
          <div class="card-header"><h3 class="card-title">Daily Activity${dailyActivity.length === 60 ? ' (last 60 days)' : ''}</h3></div>
          <div class="report-table-wrap" style="max-height:300px"><table class="report-table">
            <thead><tr><th>Date</th><th>Scans</th></tr></thead>
            <tbody>${dailyActivity.map(([date, count]) => `<tr><td>${formatDate(date)}</td><td>${count}</td></tr>`).join('')}</tbody>
          </table></div>
        </div>
      </div>
    `;

  // ---- BARCODE HISTORY (FIX 4: pre-computed state map + 100-row pagination) ----
  } else if (currentReport === 'barcode-history') {
    let filteredTxns = txns.filter(t => inRange(t.date));
    if (q) filteredTxns = filteredTxns.filter(t =>
      String(t.barcode).toLowerCase().includes(q) ||
      String(t.employeeName||'').toLowerCase().includes(q));

    const totalRows  = filteredTxns.length;
    const totalPages = Math.max(1, Math.ceil(totalRows / BARCODE_HISTORY_PAGE_SIZE));
    // clamp page in case filters changed
    if (barcodeHistoryPage >= totalPages) barcodeHistoryPage = totalPages - 1;
    const pageStart  = barcodeHistoryPage * BARCODE_HISTORY_PAGE_SIZE;
    const pageEnd    = Math.min(pageStart + BARCODE_HISTORY_PAGE_SIZE, totalRows);
    const pageTxns   = filteredTxns.slice(pageStart, pageEnd);

    const canDelete  = canDeleteTransactions();
    const cols       = canDelete ? 7 : 6;

    const paginationHtml = totalPages > 1 ? `
      <div style="display:flex;align-items:center;justify-content:center;gap:10px;padding:12px;border-top:1px solid var(--border)">
        <button class="btn-secondary" style="padding:4px 12px" ${barcodeHistoryPage===0?'disabled':''}
          onclick="barcodeHistoryPage=Math.max(0,barcodeHistoryPage-1);renderReport()">← Prev</button>
        <span style="font-size:0.85rem;color:var(--text2)">Page ${barcodeHistoryPage+1} of ${totalPages} &nbsp;·&nbsp; ${totalRows} rows</span>
        <button class="btn-secondary" style="padding:4px 12px" ${barcodeHistoryPage>=totalPages-1?'disabled':''}
          onclick="barcodeHistoryPage=Math.min(totalPages-1,barcodeHistoryPage+1);renderReport()">Next →</button>
      </div>` : '';

    container.innerHTML = `
      <div class="report-table-wrap">
        <table class="report-table">
          <thead><tr><th>Date</th><th>Barcode</th><th>Action</th><th>Employee / Location</th><th>Current State</th><th>Notes</th>${canDelete ? '<th>Del</th>' : ''}</tr></thead>
          <tbody>${pageTxns.length ? pageTxns.map(t => {
            // FIX 4: use pre-computed map instead of calling resolveState() per row
            const curState = barcodeStateMap[t.barcode] || 'unknown';
            return `<tr class="${t.inferred ? 'inferred-row' : ''}">
              <td>${formatDate(t.date)}</td>
              <td class="bc-mono">${escHtml(t.barcode)}</td>
              <td><span class="report-badge ${t.action==='distributed'?'out':t.action==='returned'?'returned':t.action==='reported_issue'?'issue':t.action==='sent_to_cintas'?'cintas':'warehouse'}">${actionLabel(t.action)}</span></td>
              <td>${escHtml(t.employeeName||(t.action==='reported_issue'?'Issue → Cintas':(t.action||'').includes('cintas')?'Cintas':'Warehouse'))}</td>
              <td><span class="state-badge ${curState}">${stateLabel(curState)}</span></td>
              <td>${escHtml(t.notes||'')}${t.inferred ? ' <span class="inferred-indicator">⚙ inferred</span>' : ''}</td>
              ${canDelete ? `<td><button class="btn-text" style="color:var(--red);font-size:0.8rem" onclick="deleteTransaction('${t.id}')">Del</button></td>` : ''}
            </tr>`;
          }).join('') : `<tr><td colspan="${cols}" style="text-align:center;color:var(--text3);padding:30px">No transactions found.</td></tr>`}
          </tbody>
        </table>
        ${paginationHtml}
      </div>`;
  }
}

function deleteTransaction(id) {
  if (!canDeleteTransactions()) { showToast('Access denied.', 'error'); return; }
  confirm_dialog('Delete Transaction?', 'Remove this scan record? This cannot be undone.', () => {
    DB.removeTransaction(id);
    CloudSync.deleteTransaction(id); // FIX 1: sync deletion to Supabase (queued if offline)
    showToast('Transaction deleted.', 'success');
    renderReport();
    renderDashboard();
  });
}

function clearDates() {
  document.getElementById('report-from').value = '';
  document.getElementById('report-to').value = '';
  renderReport();
}

// ============================================================
// CSV EXPORT
// ============================================================
function exportCSV() {
  if (!canExportData()) {
    showToast('Access denied. Your role cannot export data.', 'error');
    return;
  }

  const txns = DB.transactions;
  if (!txns.length) { showToast('No data to export.', 'error'); return; }

  const from = document.getElementById('report-from').value;
  const to   = document.getElementById('report-to').value;
  let filtered = txns;
  if (from) filtered = filtered.filter(t => t.date >= from);
  if (to)   filtered = filtered.filter(t => t.date <= to);

  const headers = ['Date','Barcode','Action','Employee ID','Employee Name','Uniform Type','Notes','State','Inferred','Created At'];
  const rows = filtered.map(t => {
    const { state } = resolveState(t.barcode);
    return [
      t.date, t.barcode, actionLabel(t.action),
      t.employeeId || '', t.employeeName || '', t.uniformType || '',
      t.notes || '', stateLabel(state), t.inferred ? 'Yes' : 'No', t.createdAt
    ];
  });

  let csv = headers.join(',') + '\n';
  rows.forEach(row => {
    csv += row.map(v => `"${String(v).replace(/"/g,'""')}"`).join(',') + '\n';
  });

  const blob = new Blob([csv], { type: 'text/csv' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `unitrack_export_${today()}.csv`;
  a.click();
  URL.revokeObjectURL(url);
  showToast('✓ CSV exported!', 'success');
}

// ============================================================
// SETTINGS
// ============================================================
function renderSettings() {
  // Centres
  const centreContainer = document.getElementById('centre-tags');
  centreContainer.innerHTML = DB.centres.map(c =>
    `<span class="tag">${escHtml(c)}<button onclick="removeCentre('${escHtml(c)}')">✕</button></span>`
  ).join('');
  // Types
  const typeContainer = document.getElementById('type-tags');
  typeContainer.innerHTML = DB.uniformTypes.map(t =>
    `<span class="tag">${escHtml(t)}<button onclick="removeType('${escHtml(t)}')">✕</button></span>`
  ).join('');
  // Supabase connection status (FIX 15: removed dead CloudSync.apiUrl/apiKey references)
  updateSyncStatus();

}

function addCentre() {
  const input = document.getElementById('new-centre');
  const val = input.value.trim();
  if (!val) return;
  if (DB.centres.includes(val)) { showToast('Already exists.', 'error'); return; }
  DB.saveCentres([...DB.centres, val]);
  input.value = '';
  renderSettings();
}

function removeCentre(name) {
  DB.saveCentres(DB.centres.filter(c => c !== name));
  renderSettings();
}

function addUniformType() {
  const input = document.getElementById('new-type');
  const val = input.value.trim();
  if (!val) return;
  if (DB.uniformTypes.includes(val)) { showToast('Already exists.', 'error'); return; }
  DB.saveTypes([...DB.uniformTypes, val]);
  input.value = '';
  renderSettings();
}

function removeType(name) {
  DB.saveTypes(DB.uniformTypes.filter(t => t !== name));
  renderSettings();
}

function exportJSON() {
  const data = {
    employees: DB.employees,
    transactions: DB.transactions,
    centres: DB.centres,
    uniformTypes: DB.uniformTypes,
    exportDate: new Date().toISOString(),
    version: 'v3'
  };
  const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `unitrack_backup_${today()}.json`;
  a.click();
  URL.revokeObjectURL(url);
  showToast('✓ Backup exported!', 'success');
}

function importJSON(event) {
  if (!canManageEmployees()) { showToast('Access denied.', 'error'); return; }
  const file = event.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = async (e) => {
    try {
      const data = JSON.parse(e.target.result);
      if (data.employees) DB.saveEmployees(data.employees);
      if (data.transactions) {
        DB_MEMORY.transactions = data.transactions;
        await IDB.saveTransactionsBulk(data.transactions);
      }
      if (data.centres) DB.saveCentres(data.centres);
      if (data.uniformTypes) DB.saveTypes(data.uniformTypes);
      showToast('✓ Data imported!', 'success');
      renderDashboard(); renderEmployeeList(); renderSettings();
    } catch (err) {
      showToast('Import failed: ' + err.message, 'error');
    }
  };
  reader.readAsText(file);
  event.target.value = '';
}

// ============================================================
// USERS PAGE (NEW - complete implementation)
// ============================================================
function renderUsers() {
  if (!canManageUsers()) {
    const tbody = document.getElementById('users-tbody');
    if (tbody) tbody.innerHTML = '<tr><td colspan="3" style="text-align:center;padding:30px;color:var(--text3)">Access denied. Admin only.</td></tr>';
    return;
  }
  
  renderRolePermissions();

  const q = (document.getElementById('user-search')?.value || '').toLowerCase();
  const tbody = document.getElementById('users-tbody');
  if (!tbody) return;

  const users = DB.users.filter(u => !q || u.username.toLowerCase().includes(q) || u.role.toLowerCase().includes(q));

  tbody.innerHTML = users.length ? users.map((u, idx) => `<tr>
    <td><strong>${escHtml(u.username)}</strong></td>
    <td><span class="role-badge ${u.role}">${u.role.charAt(0).toUpperCase() + u.role.slice(1)}</span></td>
    <td>
      <button class="btn-text" onclick="openUserModal(${idx})" style="margin-right:8px">Edit</button>
      <button class="btn-text" style="color:var(--red)" onclick="deleteUser(${idx})">Delete</button>
    </td>
  </tr>`).join('') : '<tr><td colspan="3" style="text-align:center;padding:30px;color:var(--text3)">No users found.</td></tr>';
}

function renderRolePermissions() {
  const tbody = document.getElementById('roles-tbody');
  if (!tbody) return;
  
  const roles = Object.keys(ROLE_PERMISSIONS);
  
  tbody.innerHTML = roles.map(role => {
    const p = ROLE_PERMISSIONS[role];
    const roleName = role.charAt(0).toUpperCase() + role.slice(1);
    
    // Admins can't remove their own core user-management permission (safeguard)
    const isCoreAdmin = (role === 'admin');
    
    return `<tr>
      <td><strong>${roleName}</strong></td>
      <td><input type="checkbox" onchange="updateRolePermission('${role}', 'canManageUsers', this.checked)" ${p.canManageUsers ? 'checked' : ''} ${isCoreAdmin ? 'disabled' : ''}></td>
      <td><input type="checkbox" onchange="updateRolePermission('${role}', 'canManageEmployees', this.checked)" ${p.canManageEmployees ? 'checked' : ''}></td>
      <td><input type="checkbox" onchange="updateRolePermission('${role}', 'canDeleteTransactions', this.checked)" ${p.canDeleteTransactions ? 'checked' : ''}></td>
      <td><input type="checkbox" onchange="updateRolePermission('${role}', 'canExportData', this.checked)" ${p.canExportData ? 'checked' : ''}></td>
    </tr>`;
  }).join('');
}

window.updateRolePermission = (role, perm, val) => {
  if (ROLE_PERMISSIONS[role]) {
    ROLE_PERMISSIONS[role][perm] = val;
  }
};

function openUserModal(idx) {
  if (!canManageUsers()) { showToast('Access denied.', 'error'); return; }
  const isEdit = idx !== undefined && idx !== null;
  document.getElementById('user-modal-title').textContent = isEdit ? 'Edit User' : 'Add User';
  document.getElementById('user-edit-idx').value = isEdit ? idx : '';

  if (isEdit) {
    const u = DB.users[idx];
    document.getElementById('user-username').value = u.username;
    document.getElementById('user-password').value = u.password;
    document.getElementById('user-role').value = u.role;
  } else {
    document.getElementById('user-username').value = '';
    document.getElementById('user-password').value = '';
    document.getElementById('user-role').value = 'operator';
  }
  openModal('user-modal');
}

function saveUser() {
  if (!canManageUsers()) { showToast('Access denied.', 'error'); return; }
  const username = document.getElementById('user-username').value.trim();
  const password = document.getElementById('user-password').value.trim();
  const role     = document.getElementById('user-role').value;
  if (!username || !password) { showToast('Username and password required.', 'error'); return; }

  const idx = document.getElementById('user-edit-idx').value;
  const list = DB.users;

  if (idx !== '') {
    // Edit
    list[parseInt(idx)] = { username, password, role };
  } else {
    // Check duplicate
    if (list.some(u => u.username.toLowerCase() === username.toLowerCase())) {
      showToast('Username already exists.', 'error'); return;
    }
    list.push({ username, password, role });
  }
  DB.saveUsers(list);
  closeModal('user-modal');
  showToast(`User ${username} saved!`, 'success');
  renderUsers();
}

function deleteUser(idx) {
  if (!canManageUsers()) { showToast('Access denied.', 'error'); return; }
  const user = DB.users[idx];
  if (!user) return;
  if (user.username === currentUser?.username) {
    showToast('Cannot delete the currently logged-in user.', 'error'); return;
  }
  confirm_dialog(`Delete user "${user.username}"?`, 'This cannot be undone.', () => {
    const list = DB.users;
    list.splice(idx, 1);
    DB.saveUsers(list);
    showToast('User deleted.', 'success');
    renderUsers();
  });
}

// ============================================================
// DANGER ZONE
// ============================================================
// ============================================================
// SMART DATABASE CLEANUP ENGINE
// ============================================================

// What intermediate steps are needed to go from state → new action
const INFERENCE_STEPS_NEEDED = {
  'with_employee': {
    'collected_from_soil_bin': ['returned'],
    'sent_to_cintas':          ['returned', 'collected_from_soil_bin'],
    'received_from_cintas':    ['returned', 'collected_from_soil_bin', 'sent_to_cintas'],
    'distributed':             ['returned'],   // re-distribute to another employee
    'reported_issue':          []              // valid direct
  },
  'warehouse': {
    'received_from_cintas':    ['collected_from_soil_bin', 'sent_to_cintas']
  },
  'soil_bin': {
    'received_from_cintas':    ['sent_to_cintas']
  },
  'at_cintas': {
    'distributed':             ['received_from_cintas'],
    'reported_issue':          ['received_from_cintas']
  },
  'damaged': {
    'received_from_cintas':    ['sent_to_cintas'],
    'distributed':             ['sent_to_cintas', 'received_from_cintas']
  }
};

// Replay state machine for one barcode's sorted history.
// Returns { finalState, stepsToInject[] }
function replayBarcodeHistory(sortedTxns) {
  const ACTION_TO_ST = {
    received_from_cintas:    'warehouse',
    distributed:             'with_employee',
    returned:                'warehouse',
    collected_from_soil_bin: 'soil_bin',
    sent_to_cintas:          'at_cintas',
    reported_issue:          'damaged'
  };

  let state = 'unknown';
  let lastTxn = null;
  const stepsToInject = [];

  for (const txn of sortedTxns) {
    if (txn.inferred) {
      // Already an inferred step — update state and continue
      state = ACTION_TO_ST[txn.action] || state;
      lastTxn = txn;
      continue;
    }

    const neededMap = INFERENCE_STEPS_NEEDED[state] || {};
    const needed    = neededMap[txn.action] || [];

    if (needed.length > 0) {
      // Calculate estimated date between last txn and this txn
      const prevDate = lastTxn ? lastTxn.date : txn.date;
      const midDate  = estimateMidDate(prevDate, txn.date);

      for (const missingAction of needed) {
        const inferredTxn = {
          id:           uid(),
          barcode:      txn.barcode,
          action:       missingAction,
          date:         midDate,
          createdAt:    new Date(midDate + 'T12:00:00Z').toISOString(),
          employeeId:   missingAction === 'returned' ? (lastTxn?.employeeId || null) : null,
          employeeName: missingAction === 'returned' ? (lastTxn?.employeeName || null) : null,
          uniformType:  txn.uniformType || lastTxn?.uniformType || '',
          notes:        'Auto-corrected by Smart Cleanup',
          inferred:     true
        };
        stepsToInject.push(inferredTxn);
        state = ACTION_TO_ST[missingAction] || state;
        lastTxn = inferredTxn;
      }
    }

    state   = ACTION_TO_ST[txn.action] || state;
    lastTxn = txn;
  }

  return { finalState: state, stepsToInject };
}

// ============================================================
// SMART CLEANUP ORCHESTRATOR
// ============================================================
// Uses replayBarcodeHistory() to find every barcode that has missing
// intermediate steps (e.g. distributed → received_from_cintas with no
// return/soil-bin/sent step in between) and injects the corrected records.
//
// Date rule (per business logic):
//   If the gap starts from a 'with_employee' state, the inferred steps are
//   dated 7 days after the last distribution — matching the physical cycle
//   where uniforms are collected from employees after ~1 week.
//   For all other gaps, we use the midpoint between the two surrounding dates.

async function runSmartCleanup() {
  if (!canDeleteTransactions()) {
    showToast('Access denied. Admins and operators only.', 'error');
    return;
  }

  // ── Pass 1: Dry-run to count and preview what will be fixed ──────────────
  const barcodes = [...new Set(DB.transactions.map(t => t.barcode))];
  let previewRows = [];

  for (const bc of barcodes) {
    const sorted = DB.transactions
      .filter(t => t.barcode === bc)
      .sort((a, b) => {
        const d = (a.date || '').localeCompare(b.date || '');
        return d !== 0 ? d : new Date(a.createdAt) - new Date(b.createdAt);
      });

    const { stepsToInject } = replayBarcodeHistory(sorted);
    if (stepsToInject.length > 0) {
      previewRows.push({ bc, count: stepsToInject.length, actions: stepsToInject.map(s => s.action) });
    }
  }

  if (previewRows.length === 0) {
    confirm_dialog(
      '✓ Database Clean',
      'All uniform histories are complete — no missing steps detected. The dashboard numbers are accurate.',
      () => {},
      { hideConfirm: true, cancelLabel: 'Close' }
    );
    return;
  }

  const totalFixes = previewRows.reduce((s, r) => s + r.count, 0);
  const previewHtml = `
    <p style="margin-bottom:12px">Found <strong>${previewRows.length} barcodes</strong> with missing flow steps.
    This will insert <strong>${totalFixes} inferred transactions</strong> to fill the gaps,
    using a 7-day rule for uniforms with employees.</p>
    <div style="max-height:220px;overflow-y:auto;font-size:0.82rem;border:1px solid var(--border);border-radius:8px;padding:10px">
      ${previewRows.slice(0, 30).map(r =>
        `<div style="padding:4px 0;border-bottom:1px solid var(--border);display:flex;gap:8px;align-items:center">
          <code style="font-size:0.8rem;color:var(--text2)">${escHtml(r.bc)}</code>
          <span style="color:var(--text3)">+${r.count} step${r.count > 1 ? 's' : ''}</span>
          <span style="color:var(--purple);font-size:0.75rem">${r.actions.map(a => actionLabel(a)).join(' → ')}</span>
        </div>`
      ).join('')}
      ${previewRows.length > 30 ? `<div style="padding:4px 0;color:var(--text3)">…and ${previewRows.length - 30} more barcodes</div>` : ''}
    </div>
    <p style="margin-top:10px;font-size:0.8rem;color:var(--text3)">⚙ All inserted steps are marked as <em>auto-inferred</em> and will sync to Supabase.</p>`;

  confirm_dialog('🔧 Smart Database Cleanup', previewHtml, async () => {
    // ── Pass 2: Actually apply the fixes ─────────────────────────────────
    const btn = document.getElementById('smart-cleanup-btn');
    if (btn) { btn.disabled = true; btn.textContent = '⏳ Running…'; }

    let totalInjected = 0;
    const toSync = [];

    for (const bc of barcodes) {
      const sorted = DB.transactions
        .filter(t => t.barcode === bc)
        .sort((a, b) => {
          const d = (a.date || '').localeCompare(b.date || '');
          return d !== 0 ? d : new Date(a.createdAt) - new Date(b.createdAt);
        });

      // Custom replay with business date rule:
      // For gaps starting from 'with_employee', date = distributionDate + 7 days
      // For all other gaps, use midpoint
      const ACTION_TO_ST = {
        received_from_cintas: 'warehouse', distributed: 'with_employee',
        returned: 'warehouse', collected_from_soil_bin: 'soil_bin',
        sent_to_cintas: 'at_cintas', reported_issue: 'damaged'
      };

      let state = 'unknown';
      let lastTxn = null;
      const toInject = [];

      for (const txn of sorted) {
        if (txn.inferred) {
          state = ACTION_TO_ST[txn.action] || state;
          lastTxn = txn;
          continue;
        }

        const neededMap = INFERENCE_STEPS_NEEDED[state] || {};
        const needed = neededMap[txn.action] || [];

        if (needed.length > 0) {
          // Business date rule: if coming from with_employee state,
          // use distribution_date + 7 days; otherwise midpoint.
          let inferDate;
          if (state === 'with_employee' && lastTxn) {
            const distDate = new Date(lastTxn.date);
            distDate.setDate(distDate.getDate() + 7);
            // Don't exceed the next real txn's date
            const nextDate = new Date(txn.date);
            inferDate = distDate <= nextDate
              ? distDate.toISOString().slice(0, 10)
              : estimateMidDate(lastTxn.date, txn.date);
          } else {
            inferDate = estimateMidDate(lastTxn ? lastTxn.date : txn.date, txn.date);
          }

          for (const missingAction of needed) {
            const inferredTxn = {
              id:           uid(),
              barcode:      txn.barcode,
              action:       missingAction,
              date:         inferDate,
              createdAt:    new Date(inferDate + 'T12:00:00Z').toISOString(),
              employeeId:   missingAction === 'returned' ? (lastTxn?.employeeId || null) : null,
              employeeName: missingAction === 'returned' ? (lastTxn?.employeeName || null) : null,
              uniformType:  txn.uniformType || lastTxn?.uniformType || '',
              notes:        'Auto-corrected by Smart Cleanup (7-day rule)',
              inferred:     true
            };
            toInject.push(inferredTxn);
            toSync.push(inferredTxn);
            state = ACTION_TO_ST[missingAction] || state;
            lastTxn = inferredTxn;
            totalInjected++;
          }
        }

        state = ACTION_TO_ST[txn.action] || state;
        lastTxn = txn;
      }

      // Insert all injected steps into memory + IndexedDB
      for (const t of toInject) {
        DB.addTransaction(t);
      }
    }

    // Re-sort in-memory transactions (newest first by createdAt)
    DB_MEMORY.transactions.sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));

    // Push all to Supabase in one batch
    if (toSync.length > 0) {
      showToast(`⏳ Syncing ${toSync.length} corrected steps to Supabase…`, 'info');
      try {
        await CloudSync.pushTransactionsBulk(toSync);
      } catch(e) {
        // Queue individually if bulk fails
        for (const t of toSync) CloudSync.pushTransaction(t);
      }
    }

    if (btn) { btn.disabled = false; btn.innerHTML = '🔧 Fix Database'; }

    // Refresh all views
    renderDashboard();
    if (currentPage === 'reports') renderReport();

    showToast(
      `✓ Smart Cleanup complete — ${totalInjected} missing steps added across ${previewRows.length} barcodes. Dashboard updated.`,
      'success'
    );
  }, { useInnerHTML: true, confirmLabel: 'Apply Fixes', cancelLabel: 'Cancel' });
}

// ============================================================
// MODALS & TOAST
// ============================================================
function openModal(id) {
  const el = document.getElementById(id);
  if (el) {
    el.style.display = 'flex';
    setTimeout(() => {
      el.classList.add('open');
    }, 10);
  }
}

function closeModal(id) {
  const el = document.getElementById(id);
  if (el) {
    el.classList.remove('open');
    setTimeout(() => {
      if (!el.classList.contains('open')) el.style.display = 'none';
    }, 300);
  }
}

function closeModalIfBg(event, id) {
  if (event.target === event.currentTarget) closeModal(id);
}

function showToast(msg, type = 'info') {
  const toast = document.getElementById('toast');
  toast.textContent = msg;
  toast.className = 'toast ' + type + ' show';
  clearTimeout(toast._timeout);
  toast._timeout = setTimeout(() => { toast.classList.remove('show'); }, 3500);
}

function confirm_dialog(title, message, onOk, optsOrBool = false) {
  // Backwards compat: 4th arg can be boolean (legacy) or options object
  const opts = (typeof optsOrBool === 'object' && optsOrBool !== null) ? optsOrBool : { useInnerHTML: optsOrBool };
  const useInnerHTML  = opts.useInnerHTML  || false;
  const hideConfirm   = opts.hideConfirm   || false;
  const cancelLabel   = opts.cancelLabel   || 'Cancel';
  const confirmLabel  = opts.confirmLabel  || 'Confirm';

  const overlay = document.getElementById('confirm-overlay');
  document.getElementById('confirm-title').textContent = title;
  const msgEl = document.getElementById('confirm-message');
  if (useInnerHTML || typeof message === 'string' && message.includes('<')) {
    msgEl.innerHTML = message;
  } else {
    msgEl.textContent = message;
  }
  openModal('confirm-overlay');

  const okBtn = document.getElementById('confirm-ok');
  const cancelBtn = document.getElementById('confirm-cancel');
  const newOk = okBtn.cloneNode(true);
  const newCancel = cancelBtn.cloneNode(true);
  okBtn.parentNode.replaceChild(newOk, okBtn);
  cancelBtn.parentNode.replaceChild(newCancel, cancelBtn);

  newOk.textContent  = confirmLabel;
  newOk.style.display = hideConfirm ? 'none' : '';
  newCancel.textContent = cancelLabel;

  if (!hideConfirm && onOk) {
    newOk.addEventListener('click', () => { closeModal('confirm-overlay'); onOk(); });
  }
  newCancel.addEventListener('click', () => { closeModal('confirm-overlay'); });
}


// ============================================================
// EXCEL / CSV IMPORT
// ============================================================
const FIELD_ALIASES = {
  firstName:        ['first name','first_name','firstname','given name','fname','prenom'],
  lastName:         ['last name','last_name','lastname','surname','family name','lname'],
  employeeId:       ['employee id','employee_id','emp id','emp_id','id number','id_number','badge','badge id','badge_id','employeeid'],
  productionCentre: ['production centre','production_centre','centre','center','plant','location','site','production center'],
  department:       ['department','dept','dept.','division','section','area'],
  phone:            ['phone','phone number','phone_number','tel','telephone','mobile','cell','contact'],
  notes:            ['notes','note','comments','comment','remarks','memo','description']
};

function openExcelImport() {
  xlRawRows = []; xlHeaders = []; xlMapping = {};
  document.getElementById('xl-step1').style.display = 'block';
  document.getElementById('xl-step2').style.display = 'none';
  document.getElementById('xl-step3').style.display = 'none';
  document.getElementById('xl-next-btn').style.display = 'none';
  document.getElementById('xl-import-btn').style.display = 'none';
  const dropzone = document.getElementById('xl-dropzone');
  if (dropzone) {
    dropzone.classList.remove('has-file');
    const sub = dropzone.querySelector('.xl-drop-text');
    if (sub) sub.textContent = 'Click to browse or drag & drop your file here';
  }
  openModal('excel-import-modal');
}

function handleExcelFile(event) {
  const file = event.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const workbook = XLSX.read(e.target.result, { type: 'array' });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const json = XLSX.utils.sheet_to_json(sheet, { header: 1 });
      if (json.length < 2) { showToast('File has no data rows.', 'error'); return; }
      xlHeaders = json[0].map(h => String(h || '').trim());
      xlRawRows = json.slice(1).filter(r => r.some(c => c !== undefined && c !== null && c !== ''));

      // Auto-detect column mapping
      xlMapping = {};
      Object.entries(FIELD_ALIASES).forEach(([field, aliases]) => {
        const match = xlHeaders.findIndex(h =>
          aliases.some(a => h.toLowerCase().includes(a))
        );
        if (match >= 0) xlMapping[field] = match;
      });

      // Check for single "Name" or "Full Name" column
      if (xlMapping.firstName === undefined && xlMapping.lastName === undefined) {
        const nameIdx = xlHeaders.findIndex(h => /^(name|full\s*name|employee\s*name)$/i.test(h.trim()));
        if (nameIdx >= 0) { xlMapping.firstName = nameIdx; xlMapping._singleName = true; }
      }

      showStep2();
    } catch (err) {
      showToast('Error reading file: ' + err.message, 'error');
    }
  };
  reader.readAsArrayBuffer(file);
}

function showStep2() {
  document.getElementById('xl-step1').style.display = 'none';
  document.getElementById('xl-step2').style.display = 'block';
  document.getElementById('xl-next-btn').style.display = 'inline-flex';
  document.getElementById('xl-file-info').innerHTML = `<strong>${xlRawRows.length}</strong> rows detected, <strong>${xlHeaders.length}</strong> columns`;

  const grid = document.getElementById('xl-mapping-grid');
  const fields = [
    { key: 'firstName', label: 'First Name', required: true },
    { key: 'lastName', label: 'Last Name', required: !xlMapping._singleName },
    { key: 'employeeId', label: 'Employee ID' },
    { key: 'productionCentre', label: 'Production Centre' },
    { key: 'department', label: 'Department' },
    { key: 'phone', label: 'Phone' },
    { key: 'notes', label: 'Notes' }
  ];

  grid.innerHTML = fields.map(f => `
    <div class="xl-map-row">
      <label class="xl-map-label">${f.label}${f.required ? ' <span style="color:var(--red)">*</span>' : ''}</label>
      <select class="field-input xl-map-select" data-field="${f.key}" onchange="xlMapping['${f.key}'] = this.value === '' ? undefined : parseInt(this.value)">
        <option value="">— Skip —</option>
        ${xlHeaders.map((h, i) => `<option value="${i}" ${xlMapping[f.key] === i ? 'selected' : ''}>${escHtml(h)}</option>`).join('')}
      </select>
    </div>
  `).join('');
}

function xlNext() {
  // Validate required fields
  if (xlMapping.firstName === undefined) {
    showToast('First Name column is required.', 'error'); return;
  }
  if (!xlMapping._singleName && xlMapping.lastName === undefined) {
    showToast('Last Name column is required (or map a single Name column).', 'error'); return;
  }

  document.getElementById('xl-step2').style.display = 'none';
  document.getElementById('xl-step3').style.display = 'block';
  document.getElementById('xl-next-btn').style.display = 'none';
  document.getElementById('xl-import-btn').style.display = 'inline-flex';

  const previewBody = document.getElementById('xl-preview-body');
  const previewRows = xlRawRows.slice(0, 10);

  previewBody.innerHTML = previewRows.map(row => {
    let firstName = '', lastName = '';
    if (xlMapping._singleName) {
      const parts = String(row[xlMapping.firstName] || '').trim().split(' ');
      firstName = parts[0] || '';
      lastName = parts.slice(1).join(' ') || '';
    } else {
      firstName = String(row[xlMapping.firstName] || '').trim();
      lastName = xlMapping.lastName !== undefined ? String(row[xlMapping.lastName] || '').trim() : '';
    }
    return `<tr>
      <td>${escHtml(firstName)}</td>
      <td>${escHtml(lastName)}</td>
      <td>${xlMapping.employeeId !== undefined ? escHtml(row[xlMapping.employeeId] || '') : '—'}</td>
      <td>${xlMapping.productionCentre !== undefined ? escHtml(row[xlMapping.productionCentre] || '') : '—'}</td>
      <td>${xlMapping.department !== undefined ? escHtml(row[xlMapping.department] || '') : '—'}</td>
      <td>${xlMapping.phone !== undefined ? escHtml(row[xlMapping.phone] || '') : '—'}</td>
    </tr>`;
  }).join('');

  document.getElementById('xl-preview-info').innerHTML = `Showing first <strong>${previewRows.length}</strong> of <strong>${xlRawRows.length}</strong> employees`;
}

function xlDoImport() {
  let imported = 0, skipped = 0;
  xlRawRows.forEach(row => {
    let firstName = '', lastName = '';
    if (xlMapping._singleName) {
      const parts = String(row[xlMapping.firstName] || '').trim().split(' ');
      firstName = parts[0] || '';
      lastName = parts.slice(1).join(' ') || '';
    } else {
      firstName = String(row[xlMapping.firstName] || '').trim();
      lastName = xlMapping.lastName !== undefined ? String(row[xlMapping.lastName] || '').trim() : '';
    }
    if (!firstName) return;

    const newEmpId = xlMapping.employeeId !== undefined ? String(row[xlMapping.employeeId] || '').trim() : '';
    // FIX 11: Skip duplicates — check by employeeId if available, else by full name
    const isDuplicate = newEmpId
      ? DB.employees.some(e => e.employeeId && e.employeeId === newEmpId)
      : DB.employees.some(e =>
          (e.firstName||'').toLowerCase() === firstName.toLowerCase() &&
          (e.lastName||'').toLowerCase() === (lastName||'').toLowerCase()
        );
    if (isDuplicate) { skipped++; return; }

    const emp = {
      id: uid(), firstName, lastName,
      employeeId: newEmpId,
      productionCentre: xlMapping.productionCentre !== undefined ? String(row[xlMapping.productionCentre] || '').trim() : '',
      department: xlMapping.department !== undefined ? String(row[xlMapping.department] || '').trim() : '',
      phone: xlMapping.phone !== undefined ? String(row[xlMapping.phone] || '').trim() : '',
      notes: xlMapping.notes !== undefined ? String(row[xlMapping.notes] || '').trim() : '',
      createdAt: new Date().toISOString()
    };
    DB.addEmployee(emp);
    CloudSync.pushEmployee(emp);
    imported++;
  });

  closeModal('excel-import-modal');
  const msg = skipped > 0
    ? `✓ Imported ${imported} employees. ${skipped} skipped (duplicates).`
    : `✓ Imported ${imported} employees!`;
  showToast(msg, 'success');
  renderEmployeeList();
}

function downloadTemplate() {
  const csv = 'First Name,Last Name,Employee ID,Production Centre,Department,Phone,Notes\nJohn,Smith,EMP-001,Main Plant,Packaging,+1 555-1234,\nJane,Doe,EMP-002,Warehouse A,Shipping,+1 555-5678,Night shift\n';
  const blob = new Blob([csv], { type: 'text/csv' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'unitrack_employee_template.csv';
  a.click();
  URL.revokeObjectURL(url);
}

// ============================================================
// GLOBAL EXPORTS (for onclick handlers in HTML)
// ============================================================
window.toggleSidebar = toggleSidebar;
window.showPage = showPage;
window.setAction = setAction;
window.setReport = setReport;
window.onScannerFocus = onScannerFocus;
window.onScannerBlur = onScannerBlur;
window.onBarcodeInput = onBarcodeInput;
window.handleBarcodeKey = handleBarcodeKey;
window.filterScanEmployees = filterScanEmployees;
window.showEmpDropdown = showEmpDropdown;
window.selectEmployee = selectEmployee;
window.clearEmployee = clearEmployee;
window.openManualEntry = openManualEntry;
window.manualSubmit = manualSubmit;
window.openEmployeeModal = openEmployeeModal;
window.saveEmployee = saveEmployee;
window.deleteEmployee = deleteEmployee;
window.showEmployeeDetail = showEmployeeDetail;
window.renderEmployeeList = renderEmployeeList;
window.renderReport = renderReport;
window.clearDates = clearDates;
window.exportCSV = exportCSV;
// Expose pagination state to inline onclick handlers
Object.defineProperty(window, 'barcodeHistoryPage', {
  get() { return barcodeHistoryPage; },
  set(v) { barcodeHistoryPage = v; }
});
Object.defineProperty(window, 'totalPages', {
  get() { return Math.max(1, Math.ceil(
    (() => {
      const from = document.getElementById('report-from')?.value || '';
      const to   = document.getElementById('report-to')?.value || '';
      const q    = (document.getElementById('report-search')?.value || '').toLowerCase();
      let f = DB.transactions.filter(t => (!from || t.date >= from) && (!to || t.date <= to));
      if (q) f = f.filter(t => String(t.barcode).toLowerCase().includes(q) || String(t.employeeName||'').toLowerCase().includes(q));
      return f.length;
    })()
  ) / BARCODE_HISTORY_PAGE_SIZE); }
});
window.exportJSON = exportJSON;
window.importJSON = importJSON;
window.addCentre = addCentre;
window.removeCentre = removeCentre;
window.addUniformType = addUniformType;
window.removeType = removeType;
// window.cleanupDatabase = cleanupDatabase; (Replaced by Smart Cleanup)

window.openExcelImport = openExcelImport;
window.handleExcelFile = handleExcelFile;
window.xlNext = xlNext;
window.xlDoImport = xlDoImport;
window.downloadTemplate = downloadTemplate;
window.openUserModal = openUserModal;
window.saveUser = saveUser;
window.deleteUser = deleteUser;
window.renderUsers = renderUsers;
window.deleteTransaction = deleteTransaction;
window.closeModal = closeModal;
window.closeModalIfBg = closeModalIfBg;
window.openModal = openModal;
window.confirm_dialog = confirm_dialog;
window.runSmartCleanup = runSmartCleanup;

// Network listeners
window.addEventListener('online', () => {
  updateSyncStatus();
  CloudSync.processSyncQueue().then(() => CloudSync.pullAll(true));
});
window.addEventListener('offline', () => {
  updateSyncStatus();
});

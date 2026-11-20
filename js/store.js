// Data layer. Two interchangeable backends with the same methods:
//   firebase : Google sign-in + Cloud Firestore (production)
//   demo     : this browser's localStorage, no login (used until Firebase is configured, or with ?demo=1)

const FB = 'https://www.gstatic.com/firebasejs/10.14.1/';

export async function createStore(firebaseConfig) {
  const forceDemo = new URLSearchParams(location.search).has('demo');
  const configured = firebaseConfig && firebaseConfig.apiKey && !String(firebaseConfig.apiKey).startsWith('PASTE');
  return configured && !forceDemo ? firebaseStore(firebaseConfig) : demoStore();
}

// ============================================================ FIREBASE
async function firebaseStore(cfg) {
  const [{ initializeApp }, A, F] = await Promise.all([
    import(FB + 'firebase-app.js'), import(FB + 'firebase-auth.js'), import(FB + 'firebase-firestore.js'),
  ]);
  const app = initializeApp(cfg);
  const auth = A.getAuth(app);
  const db = F.getFirestore(app);
  let me = null;
  let unsubs = [];
  let allUnsub = null;
  let billMaps = { fy: new Map(), open: new Map() };

  const stamp = () => ({ updatedAt: new Date().toISOString(), updatedBy: me?.email || '' });
  const clean = (o) => JSON.parse(JSON.stringify(o)); // strips undefined values Firestore rejects

  const store = {
    mode: 'firebase',
    onAuth(cb) {
      A.onAuthStateChanged(auth, (u) => {
        me = u ? { email: u.email.toLowerCase(), name: u.displayName || u.email, photo: u.photoURL } : null;
        cb(me);
      });
    },
    async signIn() {
      const p = new A.GoogleAuthProvider();
      p.setCustomParameters({ prompt: 'select_account' });
      try { await A.signInWithPopup(auth, p); }
      catch (e) {
        if (e.code === 'auth/popup-blocked' || e.code === 'auth/operation-not-supported-in-this-environment') {
          await A.signInWithRedirect(auth, p);
        } else throw e;
      }
    },
    async signOut() { store.unsubscribe(); await A.signOut(auth); },
    async getProfile(email) {
      const s = await F.getDoc(F.doc(db, 'users', email));
      return s.exists() ? { id: s.id, ...s.data() } : null;
    },
    async setOwnPin(pinHash) {
      await F.updateDoc(F.doc(db, 'users', me.email), { pinHash, lastSeen: new Date().toISOString() });
    },
    subscribe(fy, h, isAdmin) {
      unsubs.forEach((u) => u()); unsubs = [];
      billMaps = { fy: new Map(), open: new Map() };
      const emitBills = () => {
        const m = new Map([...billMaps.open, ...billMaps.fy]);
        h.bills([...m.values()]);
      };
      const err = (e) => h.error && h.error(e);
      const bills = F.collection(db, 'bills');
      const billQueries = fy === 'ALL'
        ? [['fy', bills]]
        : [['fy', F.query(bills, F.where('fy', '==', fy))], ['open', F.query(bills, F.where('open', '==', true))]];
      for (const [k, q] of billQueries) {
        unsubs.push(F.onSnapshot(q, (snap) => {
          billMaps[k] = new Map(snap.docs.map((d) => [d.id, { id: d.id, ...d.data() }]));
          emitBills();
        }, err));
      }
      unsubs.push(F.onSnapshot(F.collection(db, 'vendors'), (snap) => h.vendors(snap.docs.map((d) => ({ id: d.id, ...d.data() }))), err));
      unsubs.push(F.onSnapshot(F.doc(db, 'settings', 'app'), (s) => h.settings(s.exists() ? s.data() : {}), err));
      unsubs.push(F.onSnapshot(F.doc(db, 'settings', 'sync'), (s) => h.sync && h.sync(s.exists() ? s.data() : {}), err));
      if (isAdmin) unsubs.push(F.onSnapshot(F.collection(db, 'users'), (snap) => h.users(snap.docs.map((d) => ({ id: d.id, ...d.data() }))), err));
    },
    unsubscribe() { unsubs.forEach((u) => u()); unsubs = []; allUnsub && allUnsub(); allUnsub = null; },
    // Live list of every bill, used only on the device that syncs from OneDrive.
    subscribeAllBills(cb) {
      if (allUnsub) return;
      allUnsub = F.onSnapshot(F.collection(db, 'bills'), (snap) => cb(snap.docs.map((d) => ({ id: d.id, ...d.data() }))));
    },
    async saveSync(o) { await F.setDoc(F.doc(db, 'settings', 'sync'), clean(o), { merge: true }); },
    async saveVendor(v) { await F.setDoc(F.doc(db, 'vendors', v.id), clean({ ...v, ...stamp() }), { merge: true }); },
    async deleteVendor(id) { await F.deleteDoc(F.doc(db, 'vendors', id)); },
    async saveBill(b) { await F.setDoc(F.doc(db, 'bills', b.id), clean({ ...b, ...stamp() })); },
    async deleteBill(id) { await F.deleteDoc(F.doc(db, 'bills', id)); },
    async writeMany(ops, onProgress) {
      // ops: [{col, id, data, del}]. Written as individual requests, 25 at a time, because the
      // security rules look up the user's role on every write and batched writes cap those lookups.
      for (let i = 0; i < ops.length; i += 25) {
        await Promise.all(ops.slice(i, i + 25).map((o) => {
          const ref = F.doc(db, o.col, o.id);
          return o.del ? F.deleteDoc(ref) : F.setDoc(ref, clean({ ...o.data, ...stamp() }), { merge: !!o.merge });
        }));
        onProgress && onProgress(Math.min(i + 25, ops.length), ops.length);
      }
    },
    async saveSettings(s) { await F.setDoc(F.doc(db, 'settings', 'app'), clean({ ...s, ...stamp() }), { merge: true }); },
    async saveUser(u) { await F.setDoc(F.doc(db, 'users', u.id), clean({ ...u, ...stamp() }), { merge: true }); },
    async deleteUser(id) { await F.deleteDoc(F.doc(db, 'users', id)); },
    async fetchVendorBills(vendorId) {
      const s = await F.getDocs(F.query(F.collection(db, 'bills'), F.where('vendorId', '==', vendorId)));
      return s.docs.map((d) => ({ id: d.id, ...d.data() }));
    },
    async fetchAllBills() {
      const s = await F.getDocs(F.collection(db, 'bills'));
      return s.docs.map((d) => ({ id: d.id, ...d.data() }));
    },
    async addLog(action, detail) {
      try {
        await F.addDoc(F.collection(db, 'logs'), { at: new Date().toISOString(), by: me?.email || '', action, detail: String(detail || '').slice(0, 500) });
      } catch (e) { console.warn('log failed', e); }
    },
    async fetchLogs() {
      const s = await F.getDocs(F.query(F.collection(db, 'logs'), F.orderBy('at', 'desc'), F.limit(200)));
      return s.docs.map((d) => ({ id: d.id, ...d.data() }));
    },
  };
  return store;
}

// ============================================================ DEMO (localStorage)
function demoStore() {
  const KEY = 'mkf_payables_demo_v1';
  const load = () => {
    try { return JSON.parse(localStorage.getItem(KEY)) || null; } catch { return null; }
  };
  let db = load() || { vendors: {}, bills: {}, settings: {}, users: {}, logs: [] };
  db.sync = db.sync || {};
  let allCb = null;
  const me = { email: 'demo@local', name: 'Demo User', photo: '' };
  if (!db.users[me.email]) db.users[me.email] = { id: me.email, name: 'Demo User', role: 'admin', active: true };
  let h = null, fy = null;
  const persist = () => { try { localStorage.setItem(KEY, JSON.stringify(db)); } catch (e) { console.warn(e); } };
  const stamp = () => ({ updatedAt: new Date().toISOString(), updatedBy: me.email });
  const emit = () => {
    persist();
    if (!h) return;
    const bills = Object.values(db.bills).filter((b) => fy === 'ALL' || b.fy === fy || b.open);
    h.vendors(Object.values(db.vendors));
    h.bills(bills);
    h.settings(db.settings || {});
    h.users && h.users(Object.values(db.users));
    h.sync && h.sync({ ...db.sync });
    allCb && allCb(Object.values(db.bills));
  };
  const async = (fn) => async (...a) => { const r = fn(...a); emit(); return r; };
  return {
    mode: 'demo',
    onAuth(cb) { setTimeout(() => cb(me), 0); },
    async signIn() {}, async signOut() { location.reload(); },
    async getProfile(email) { return db.users[email] || null; },
    setOwnPin: async (pinHash) => { db.users[me.email].pinHash = pinHash; persist(); },
    subscribe(f, handlers) { fy = f; h = handlers; setTimeout(emit, 0); },
    unsubscribe() { h = null; },
    subscribeAllBills(cb) { allCb = cb; setTimeout(emit, 0); },
    saveSync: async((o) => { db.sync = { ...db.sync, ...o }; }),
    saveVendor: async((v) => { db.vendors[v.id] = { ...(db.vendors[v.id] || {}), ...v, ...stamp() }; }),
    deleteVendor: async((id) => { delete db.vendors[id]; }),
    saveBill: async((b) => { db.bills[b.id] = { ...b, ...stamp() }; }),
    deleteBill: async((id) => { delete db.bills[id]; }),
    writeMany: async((ops, onProgress) => {
      for (const o of ops) {
        if (o.del) { delete db[o.col][o.id]; continue; }
        const base = o.merge ? db[o.col][o.id] || {} : {};
        if (o.col === 'settings') { db.settings = { ...db.settings, ...o.data }; continue; }
        else db[o.col][o.id] = { ...base, ...o.data, ...stamp() };
      }
      onProgress && onProgress(ops.length, ops.length);
    }),
    saveSettings: async((s) => { db.settings = { ...db.settings, ...s }; }),
    saveUser: async((u) => { db.users[u.id] = { ...(db.users[u.id] || {}), ...u }; }),
    deleteUser: async((id) => { delete db.users[id]; }),
    async fetchVendorBills(vendorId) { return Object.values(db.bills).filter((b) => b.vendorId === vendorId); },
    async fetchAllBills() { return Object.values(db.bills); },
    async addLog(action, detail) { db.logs.unshift({ at: new Date().toISOString(), by: me.email, action, detail }); db.logs = db.logs.slice(0, 200); persist(); },
    async fetchLogs() { return db.logs; },
    resetDemo() { localStorage.removeItem(KEY); location.reload(); },
  };
}

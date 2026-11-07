// ===== שכבת Firebase לממשק הניהול =====
// מחקה בדיוק את הפעולות של שרת Apps Script (אותם שמות פעולה ואותן תשובות),
// כך שממשק הניהול עובד אותו דבר – רק מהר יותר ועם עדכונים חיים.
import { initializeApp } from "https://www.gstatic.com/firebasejs/13.2.0/firebase-app.js";
import { getAuth, onAuthStateChanged, signInWithPopup, GoogleAuthProvider, signInWithEmailAndPassword,
  createUserWithEmailAndPassword, signOut, reauthenticateWithPopup, deleteUser, setPersistence, inMemoryPersistence
} from "https://www.gstatic.com/firebasejs/13.2.0/firebase-auth.js";
import { initializeFirestore, persistentLocalCache, persistentMultipleTabManager, collection, doc, getDoc, getDocs,
  getDocsFromServer, setDoc, updateDoc, deleteDoc, addDoc, onSnapshot, writeBatch, serverTimestamp, Timestamp
} from "https://www.gstatic.com/firebasejs/13.2.0/firebase-firestore.js";

export const firebaseConfig = {
  apiKey: "AIzaSyB9Xv5GMEx_-wCfcsAW8HkIRS8TY0BoL5I",
  authDomain: "jony-thai.firebaseapp.com",
  projectId: "jony-thai",
  storageBucket: "jony-thai.firebasestorage.app",
  messagingSenderId: "1033410505499",
  appId: "1:1033410505499:web:872dfafbe56d91677ec1be"
};
// המנהל נכנס עם חשבון Google הזה. חייב להיות זהה לכתובת שב-firestore.rules
const ADMIN_EMAIL = "jonynemes@gmail.com";
const STATUSES = ["חדשה", "דווח תשלום", "שולם", "בהכנה", "מוכנה", "נמסרה", "בוטלה"];
const LISTS = ["shop", "prep", "todo"];
const PRI = ["דחוף", "רגיל", "יכול לחכות"];

const app = initializeApp(firebaseConfig);
const auth = getAuth(app);
const db = initializeFirestore(app, { localCache: persistentLocalCache({ tabManager: persistentMultipleTabManager() }) });
// אפליקציה שנייה, רק ליצירה/מחיקה של חשבונות צוות בלי לנתק את המנהל
let secAuth = null;
async function sec() {
  if (!secAuth) { secAuth = getAuth(initializeApp(firebaseConfig, "staff-admin")); await setPersistence(secAuth, inMemoryPersistence); }
  return secAuth;
}

// ===== תפקיד המשתמש המחובר =====
let ROLE = null, NAME = "";
async function roleOf(u) {
  if (!u) return null;
  if (u.email === ADMIN_EMAIL && u.emailVerified) return "admin";
  try { const s = await getDoc(doc(db, "staff", u.uid)); if (s.exists()) { NAME = s.data().name || ""; return "staff"; } } catch (e) {}
  return null;
}
let authResolve; const authReady = new Promise(r => authResolve = r);
onAuthStateChanged(auth, async u => {
  const r = await roleOf(u);
  if (r !== ROLE) { ROLE = r; r ? listen() : unlisten(); }
  authResolve(r);
});

// ===== האזנה חיה להזמנות ולמצב החנות =====
let ORDERS = [], OPEN = true, unsubs = [], firstOrders = null, changeCb = null;
const toIso = t => t instanceof Timestamp ? t.toDate().toISOString() : (t ? String(t) : "");
function listen() {
  unlisten();
  let res; firstOrders = new Promise(r => res = r);
  unsubs.push(onSnapshot(collection(db, "orders"), snap => {
    ORDERS = snap.docs.map(d => { const o = d.data({ serverTimestamps: "estimate" });
      return { orderNo: d.id, name: o.name || "", phone: String(o.phone || ""), items: o.items || "", note: o.note || "", total: o.total || 0, status: o.status || "חדשה", time: toIso(o.time) }; })
      .sort((a, b) => b.time.localeCompare(a.time));
    res(); changeCb && changeCb("orders");
  }, () => res()));
  unsubs.push(onSnapshot(doc(db, "settings", "shop"), s => { OPEN = !s.exists() || s.data().open !== false; changeCb && changeCb("shop"); }, () => {}));
}
function unlisten() { unsubs.forEach(f => f()); unsubs = []; ORDERS = []; firstOrders = null; }

// ===== כניסה / יציאה =====
async function staffEmail(pass) {
  const b = await crypto.subtle.digest("SHA-256", new TextEncoder().encode("jony-thai-staff:" + pass));
  return [...new Uint8Array(b)].slice(0, 12).map(x => x.toString(16).padStart(2, "0")).join("") + "@staff.jony-thai.app";
}
async function loginGoogle() {
  try {
    const p = new GoogleAuthProvider(); p.setCustomParameters({ prompt: "select_account" });
    const c = await signInWithPopup(auth, p);
    const r = await roleOf(c.user);
    if (r !== "admin") { await signOut(auth); return "not admin"; }
    ROLE = r; listen(); return "ok";
  } catch (e) { return e.code === "auth/popup-closed-by-user" || e.code === "auth/cancelled-popup-request" ? "cancel" : "network"; }
}
async function loginStaff(pass) {
  if (!pass || pass.length < 8) return "bad pass";
  try {
    const c = await signInWithEmailAndPassword(auth, await staffEmail(pass), pass);
    const r = await roleOf(c.user);
    if (r !== "staff") { await signOut(auth); return "bad pass"; }
    ROLE = r; listen(); return "ok";
  } catch (e) { return /network/.test(e.code || "") ? "network" : "bad pass"; }
}
async function logout() { unlisten(); ROLE = null; await signOut(auth); }
async function reauth() {
  try { await reauthenticateWithPopup(auth.currentUser, new GoogleAuthProvider()); return auth.currentUser.email === ADMIN_EMAIL; }
  catch (e) { return false; }
}

// ===== קריאה – אותן תשובות כמו doGet =====
async function get(action) {
  if (!ROLE) return { error: "unauthorized" };
  if (action === "orders") {
    if (firstOrders) await firstOrders;
    return { orders: ORDERS.map(o => ({ ...o })), statuses: STATUSES, open: OPEN, role: ROLE };
  }
  if (action === "lists") {
    const [ls, rs] = await Promise.all([getDocs(collection(db, "lists")), getDocs(collection(db, "recipes"))]);
    const lists = { shop: [], prep: [], todo: [] };
    ls.docs.forEach(d => { const x = d.data(); if (lists[x.list]) lists[x.list].push({ id: d.id, text: x.text || "", qty: String(x.qty || ""), note: x.note || "", group: x.group || "", pri: x.pri || "רגיל", done: x.done === true, due: x.due || "", c: x.created || 0 }); });
    Object.values(lists).forEach(a => a.sort((p, q) => p.c - q.c));
    const recipes = [];
    rs.docs.forEach(d => (d.data().rows || []).forEach(r => recipes.push({ dish: d.data().dish, ing: r.ing, qty: Number(r.qty) || 0, unit: r.unit || "" })));
    return { lists, recipes, role: ROLE };
  }
  if (action === "menuall") {
    const s = await getDocs(collection(db, "menu"));
    const menu = s.docs.map(d => ({ row: d.id, ...d.data() })).sort((a, b) => (a.order || 0) - (b.order || 0));
    return { menu, role: ROLE };
  }
  if (action === "staff") {
    if (ROLE !== "admin") return { error: "unauthorized" };
    const s = await getDocs(collection(db, "staff"));
    return { staff: s.docs.map(d => d.data().name).sort((a, b) => a.localeCompare(b, "he")) };
  }
  return { error: "unknown" };
}

// ===== כתיבה – אותן תשובות כמו doPost =====
const errCode = e => {
  const c = (e && e.code) || "";
  return c === "permission-denied" ? (ROLE === "staff" ? "forbidden" : "unauthorized") : c === "not-found" ? "not found" : /unavailable|network/.test(c) ? "network" : (c || "error");
};
const clean = v => String(v == null ? "" : v).slice(0, 500);
const listItem = (it, list) => ({ list, text: clean(it.text), qty: clean(it.qty), note: clean(it.note), group: clean(it.group),
  pri: PRI.includes(it.pri) ? it.pri : "רגיל", done: it.done === true, due: clean(it.due) });

async function post(body) {
  const d = body || {};
  if (!ROLE) return "unauthorized";
  try {
    if (d.action === "setStatus") {
      if (!STATUSES.includes(d.status)) return "bad status";
      if (ROLE === "staff") {
        if (d.status === "בוטלה") return "forbidden";
        const o = ORDERS.find(x => x.orderNo === d.orderNo);
        if (o && o.status === "בוטלה") return "forbidden";
      }
      await updateDoc(doc(db, "orders", String(d.orderNo)), { status: d.status });
      return "ok";
    }
    if (ROLE !== "admin") return "unauthorized";

    switch (d.action) {
      case "deleteOrder": await deleteDoc(doc(db, "orders", String(d.orderNo))); return "ok";
      case "setOpen": await setDoc(doc(db, "settings", "shop"), { open: !!d.open }, { merge: true }); return "ok";
      case "closeDay": return await closeDay();
      case "saveMenuRow": case "addMenuRow": {
        const m = d.item || {}, price = Number(m.price);
        if (!m.name || isNaN(price) || price < 0) return "bad item";
        const row = { cat: clean(m.cat), name: clean(m.name), desc: clean(m.desc), price, active: m.active === "לא" ? "לא" : "כן", img: clean(m.img), spicy: Math.max(0, Math.min(5, Math.round(Number(m.spicy) || 0))), vegan: m.vegan === "כן" ? "כן" : "לא" };
        if (d.action === "saveMenuRow" && m.row) { await setDoc(doc(db, "menu", String(m.row)), row, { merge: true }); return "ok"; }
        const all = await getDocs(collection(db, "menu"));
        row.order = Math.max(0, ...all.docs.map(x => x.data().order || 0)) + 1;
        await addDoc(collection(db, "menu"), row); return "ok";
      }
      case "recipeSave": {
        const dish = clean(d.dish); if (!dish) return "bad item";
        const rows = (d.rows || []).filter(r => r.ing).map(r => ({ ing: clean(r.ing), qty: Number(r.qty) || 0, unit: clean(r.unit) }));
        const id = dish.replace(/\//g, "∕");
        rows.length ? await setDoc(doc(db, "recipes", id), { dish, rows }) : await deleteDoc(doc(db, "recipes", id));
        return "ok";
      }
      case "listSave": {
        if (!LISTS.includes(d.list)) return "bad list";
        const it = d.item || {}; if (!it.text) return "bad item";
        if (it.id) { await updateDoc(doc(db, "lists", it.id), listItem(it, d.list)); return "ok"; }
        await addDoc(collection(db, "lists"), { ...listItem(it, d.list), created: Date.now() }); return "ok";
      }
      case "listAdd": {
        if (!LISTS.includes(d.list)) return "bad list";
        const b = writeBatch(db); let t = Date.now();
        (d.items || []).filter(it => it.text).forEach(it => b.set(doc(collection(db, "lists")), { ...listItem(it, d.list), created: t++ }));
        await b.commit(); return "ok";
      }
      case "listDone": await updateDoc(doc(db, "lists", String(d.id)), { done: d.done === true }); return "ok";
      case "listDelete": await deleteDoc(doc(db, "lists", String(d.id))); return "ok";
      case "listClearDone": case "listReset": {
        const s = await getDocs(collection(db, "lists")), b = writeBatch(db);
        s.docs.filter(x => x.data().list === d.list && x.data().done === true)
          .forEach(x => d.action === "listClearDone" ? b.delete(x.ref) : b.update(x.ref, { done: false }));
        await b.commit(); return "ok";
      }
      case "adminCheck": return "ok"; // האימות כבר נעשה עם Google (reauth)
      case "staffAdd": return await staffAdd(d.name, d.pass);
      case "staffReveal": { const s = await staffByName(d.name); return s ? "pass:" + (s.data().pass || "") : "not found"; }
      case "staffChange": return await staffChange(d.name, d.pass);
      case "staffDelete": return await staffDelete(d.name);
      case "importSheets": return await importSheets(d.url, d.oldKey);
      case "imageAdd": {
        if (!/^data:image\/(webp|jpeg|png);base64,/.test(d.data || "") || d.data.length > 900000) return "bad image";
        const r = await addDoc(collection(db, "images"), { name: clean(d.name).slice(0, 80), data: d.data, thumb: String(d.thumb || "").slice(0, 60000), created: Date.now() });
        return "ok:" + r.id;
      }
      case "imageDelete": await deleteDoc(doc(db, "images", String(d.id))); return "ok";
    }
    return "unknown action";
  } catch (e) { console.error(e); return errCode(e); }
}

// ===== סגירת יום: כל ההזמנות עוברות לארכיון =====
async function closeDay() {
  const s = await getDocsFromServer(collection(db, "orders"));
  if (s.empty) return "nothing";
  const now = new Date(), p = n => String(n).padStart(2, "0");
  const name = "ארכיון " + p(now.getDate()) + "." + p(now.getMonth() + 1) + "." + now.getFullYear();
  const docs = s.docs;
  for (let i = 0; i < docs.length; i += 200) {
    const b = writeBatch(db);
    docs.slice(i, i + 200).forEach(x => { b.set(doc(collection(db, "archive")), { ...x.data(), orderNo: x.id, day: name, archivedAt: serverTimestamp() }); b.delete(x.ref); });
    await b.commit();
  }
  return "ok:" + docs.length + ":" + name;
}

// ===== צוות =====
async function staffByName(n) {
  const s = await getDocs(collection(db, "staff"));
  return s.docs.find(x => x.data().name === n) || null;
}
function passErr(p) { return String(p || "").length < 8 ? "short pass" : ""; }
async function makeStaffUser(pass) {
  const a = await sec();
  try { const c = await createUserWithEmailAndPassword(a, await staffEmail(pass), pass); const uid = c.user.uid; await signOut(a); return { uid }; }
  catch (e) { return { err: e.code === "auth/email-already-in-use" ? "pass exists" : e.code === "auth/weak-password" ? "short pass" : errCode(e) }; }
}
async function removeStaffUser(pass) { // לא חובה: הגישה נחסמת ממילא כשמסמך הצוות נמחק
  try { const a = await sec(); const c = await signInWithEmailAndPassword(a, await staffEmail(pass), pass); await deleteUser(c.user); } catch (e) {}
}
async function staffAdd(name, pass) {
  name = clean(name).trim().slice(0, 40);
  if (!name) return "bad name";
  if (await staffByName(name)) return "name exists";
  const bad = passErr(pass); if (bad) return bad;
  const u = await makeStaffUser(pass); if (u.err) return u.err;
  await setDoc(doc(db, "staff", u.uid), { name, pass, created: Date.now() });
  return "ok";
}
async function staffChange(name, pass) {
  const old = await staffByName(name); if (!old) return "not found";
  const bad = passErr(pass); if (bad) return bad;
  const u = await makeStaffUser(pass); if (u.err) return u.err;
  await setDoc(doc(db, "staff", u.uid), { name, pass, created: Date.now() });
  await deleteDoc(old.ref);
  await removeStaffUser(old.data().pass);
  return "ok";
}
async function staffDelete(name) {
  const s = await staffByName(name); if (!s) return "ok";
  await deleteDoc(s.ref);
  await removeStaffUser(s.data().pass);
  return "ok";
}

// ===== ייבוא חד-פעמי מהגיליון (דרך שרת Apps Script הקיים) =====
async function importSheets(url, oldKey) {
  const g = async a => (await fetch(url + "?action=" + a + "&key=" + encodeURIComponent(oldKey))).json();
  const [o, l, m, op] = await Promise.all([g("orders"), g("lists"), g("menuall"), g("open")]);
  if (o.error || l.error || m.error) return "bad admin";
  const ops = [];
  (m.menu || []).filter(x => x.name).forEach((x, i) => ops.push(b => b.set(doc(db, "menu", "m" + String(i + 1).padStart(3, "0")),
    { cat: clean(x.cat), name: clean(x.name), desc: clean(x.desc), price: Number(x.price) || 0, active: String(x.active).trim() === "לא" ? "לא" : "כן", img: clean(x.img), spicy: Math.max(0, Math.min(5, Math.round(Number(x.spicy) || 0))), ...(x.vegan === "כן" || x.vegan === "לא" ? { vegan: x.vegan } : {}), order: i + 1 })));
  (o.orders || []).filter(x => x.orderNo).forEach(x => { const t = new Date(x.time);
    ops.push(b => b.set(doc(db, "orders", String(x.orderNo)), { orderNo: String(x.orderNo), name: clean(x.name), phone: clean(x.phone), items: String(x.items || ""), note: clean(x.note), total: Number(x.total) || 0, status: STATUSES.includes(x.status) ? x.status : "חדשה", time: isNaN(t) ? serverTimestamp() : Timestamp.fromDate(t) })); });
  let t = Date.now();
  LISTS.forEach(k => ((l.lists || {})[k] || []).forEach(it => { const c = t++; ops.push(b => b.set(doc(db, "lists", "i_" + k + "_" + String(it.id).replace(/[^\w-]/g, "")),{ ...listItem(it, k), created: c })); }));
  const byDish = {};
  (l.recipes || []).forEach(r => { (byDish[r.dish] = byDish[r.dish] || []).push({ ing: clean(r.ing), qty: Number(r.qty) || 0, unit: clean(r.unit) }); });
  Object.entries(byDish).forEach(([dish, rows]) => ops.push(b => b.set(doc(db, "recipes", String(dish).replace(/\//g, "∕")), { dish: String(dish), rows })));
  ops.push(b => b.set(doc(db, "settings", "shop"), { open: op.open !== false }));
  for (let i = 0; i < ops.length; i += 400) { const b = writeBatch(db); ops.slice(i, i + 400).forEach(f => f(b)); await b.commit(); }
  return "ok:" + (m.menu || []).length + ":" + (o.orders || []).length + ":" + Object.keys(byDish).length;
}

window.FB = { authReady, get, post, loginGoogle, loginStaff, logout, reauth, onChange: cb => { changeCb = cb; }, role: () => ROLE, name: () => NAME };
window.dispatchEvent(new Event("fbready"));

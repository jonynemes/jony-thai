// ===== הגדרות =====
// השאירו כאן את המפתח הסודי שלכם (12 תווים ומעלה). אל תשתפו אותו ואל תעלו אותו ל-GitHub.
const ADMIN_KEY = "CHANGE_ME";
// סיסמאות הצוות מנוהלות מתוך ממשק האדמין (מסך "צוות") ונשמרות מוצפנות בהגדרות הסקריפט – לא כאן ולא בגיליון.

const ORDERS = "הזמנות";
const MENU = "תפריט";
const STATUSES = ["חדשה", "דווח תשלום", "שולם", "בהכנה", "מוכנה", "נמסרה", "בוטלה"];

// מבנה גיליון ההזמנות (עמודות A עד I)
const HEADERS = ["תאריך", "שעה", "מספר הזמנה", "שם", "טלפון", "מנות", "הערה", "סכום", "סטטוס"];
const C = { date: 0, time: 1, no: 2, name: 3, phone: 4, items: 5, note: 6, total: 7, status: 8 };

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function text_(t) { return ContentService.createTextOutput(t); }
function sheet_(n) { return SpreadsheetApp.getActive().getSheetByName(n); }
// ===== צוות =====
// כל איש צוות: { n: שם, h: גיבוב SHA-256 של הסיסמה, e: הסיסמה מוצפנת בסיסמת המנהל }.
function hash_(k) {
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(k), Utilities.Charset.UTF_8));
}
function staffList_() {
  try { return JSON.parse(PropertiesService.getScriptProperties().getProperty("staff") || "[]"); } catch (e) { return []; }
}
function staffSave_(list) { PropertiesService.getScriptProperties().setProperty("staff", JSON.stringify(list)); }
// הצפנת הסיסמה עם מפתח המנהל (HMAC-SHA256 כזרם מפתח), כדי שאפשר יהיה להציג אותה – רק למי שיודע את סיסמת המנהל
function keystream_(iv, len) {
  const out = [];
  for (let i = 0; out.length < len; i++) {
    const blk = Utilities.computeHmacSha256Signature(iv + ":" + i, ADMIN_KEY);
    for (let j = 0; j < blk.length && out.length < len; j++) out.push(blk[j]);
  }
  return out;
}
function enc_(plain) {
  const iv = Utilities.getUuid();
  const b = Utilities.newBlob(String(plain)).getBytes(), ks = keystream_(iv, b.length);
  const x = b.map(function (v, i) { const r = (v ^ ks[i]) & 255; return r > 127 ? r - 256 : r; });
  return { iv: iv, c: Utilities.base64Encode(x) };
}
function dec_(e) {
  if (!e || !e.iv) return null;
  const b = Utilities.base64Decode(e.c), ks = keystream_(e.iv, b.length);
  const x = b.map(function (v, i) { const r = (v ^ ks[i]) & 255; return r > 127 ? r - 256 : r; });
  return Utilities.newBlob(x).getDataAsString("UTF-8");
}
// אימות חוזר של סיסמת המנהל לפני הצגה או שינוי של סיסמת צוות
function reauth_(d) { return isAdmin_(d.confirm); }
function checkPass_(list, pass, exceptName) {
  if (pass.length < 8) return "short pass";
  if (pass === ADMIN_KEY) return "same as admin";
  const h = hash_(pass);
  if (list.some(function (s) { return s.h === h && s.n !== exceptName; })) return "pass exists";
  return "";
}
function staffOf_(k) {
  if (!k || String(k).length < 8) return null;
  const h = hash_(k);
  return staffList_().filter(function (s) { return s.h === h; })[0] || null;
}
function isViewer_(k) { return isAdmin_(k) || !!staffOf_(k); }
function role_(k) { return isAdmin_(k) ? "admin" : "staff"; }
function isAdmin_(k) {
  return ADMIN_KEY !== "CHANGE_ME" && ADMIN_KEY.length >= 12 && String(k || "") === ADMIN_KEY;
}
// מונע הפיכת טקסט מהלקוחות לנוסחה בגיליון, וחותך אורך
function clean_(v) {
  let s = String(v == null ? "" : v).slice(0, 500);
  return /^[=+\-@]/.test(s) ? " " + s : s;
}
function menuRows_() {
  return sheet_(MENU).getDataRange().getValues().slice(1).map(function (r, i) {
    return { row: i + 2, cat: r[0], name: r[1], desc: r[2], price: r[3], active: r[4], img: r[5], spicy: Number(r[6]) || 0, vegan: String(r[7] || "") };
  });
}
// ===== רשימות (קניות / הכנות / משימות) ומתכונים =====
// כל רשימה בלשונית משלה בגיליון. הלשוניות נוצרות אוטומטית בפעם הראשונה.
const LISTS = { shop: "קניות", prep: "הכנות", todo: "משימות" };
const LIST_HEADERS = ["מזהה", "פריט", "כמות", "הערה", "קבוצה", "דחיפות", "בוצע", "תאריך יעד"];
const RECIPES = "מתכונים";
const RECIPE_HEADERS = ["מנה", "מצרך", "כמות למנה", "יחידה"];

function ensureSheet_(name, headers, checkboxCol) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (sh) return sh;
  sh = ss.insertSheet(name);
  sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
  sh.setFrozenRows(1);
  sh.setRightToLeft(true);
  if (checkboxCol) sh.getRange(2, checkboxCol, sh.getMaxRows() - 1).insertCheckboxes();
  return sh;
}
function listSheet_(key) {
  if (!LISTS[key]) return null;
  return ensureSheet_(LISTS[key], LIST_HEADERS, 7);
}
function listItems_(key) {
  const sh = listSheet_(key);
  return sh.getDataRange().getValues().slice(1).filter(function (r) { return r[0] && r[1]; }).map(function (r) {
    const due = r[7] instanceof Date ? Utilities.formatDate(r[7], Session.getScriptTimeZone(), "yyyy-MM-dd") : String(r[7] || "");
    return { id: String(r[0]), text: r[1], qty: String(r[2]), note: r[3], group: r[4], pri: r[5] || "רגיל", done: r[6] === true, due: due };
  });
}
function listRowOf_(sh, id) {
  const ids = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
  for (let i = 1; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 1;
  return 0;
}
function listRowValues_(id, it) {
  const pri = ["דחוף", "רגיל", "יכול לחכות"].indexOf(it.pri) >= 0 ? it.pri : "רגיל";
  return [id, clean_(it.text), clean_(it.qty), clean_(it.note), clean_(it.group), pri, it.done === true, clean_(it.due)];
}
// כותב שורה אחרי השורה האחרונה עם תוכן (appendRow היה מדלג על שורות תיבות הסימון הריקות)
function listAppend_(sh, rows) {
  const ids = sh.getRange(1, 1, sh.getMaxRows(), 1).getValues();
  let last = 1;
  for (let i = ids.length - 1; i > 0; i--) if (ids[i][0] !== "") { last = i + 1; break; }
  if (last + rows.length > sh.getMaxRows()) {
    sh.insertRowsAfter(sh.getMaxRows(), rows.length + 50);
    sh.getRange(2, 7, sh.getMaxRows() - 1).insertCheckboxes();
  }
  sh.getRange(last + 1, 1, rows.length, LIST_HEADERS.length).setValues(rows);
}
function recipeSheet_() { return ensureSheet_(RECIPES, RECIPE_HEADERS); }
function recipeRows_() {
  return recipeSheet_().getDataRange().getValues().slice(1).filter(function (r) { return r[0] && r[1]; }).map(function (r) {
    return { dish: r[0], ing: r[1], qty: Number(r[2]) || 0, unit: r[3] };
  });
}

function listOp_(d) {
  if (d.action === "recipeSave") {
    const sh = recipeSheet_();
    const dish = clean_(d.dish);
    if (!dish) return text_("bad item");
    const v = sh.getDataRange().getValues();
    for (let i = v.length - 1; i >= 1; i--) if (String(v[i][0]) === dish) sh.deleteRow(i + 1);
    const rows = (d.rows || []).filter(function (r) { return r.ing; }).map(function (r) {
      return [dish, clean_(r.ing), Number(r.qty) || 0, clean_(r.unit)];
    });
    if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, 4).setValues(rows);
    return text_("ok");
  }

  const sh = listSheet_(d.list);
  if (!sh) return text_("bad list");

  if (d.action === "listSave") {
    const it = d.item || {};
    if (!it.text) return text_("bad item");
    if (it.id) {
      const r = listRowOf_(sh, it.id);
      if (!r) return text_("not found");
      sh.getRange(r, 1, 1, LIST_HEADERS.length).setValues([listRowValues_(it.id, it)]);
    } else {
      listAppend_(sh, [listRowValues_(Utilities.getUuid().slice(0, 8), it)]);
    }
    return text_("ok");
  }
  if (d.action === "listAdd") {
    const rows = (d.items || []).filter(function (it) { return it.text; }).map(function (it) {
      return listRowValues_(Utilities.getUuid().slice(0, 8), it);
    });
    if (rows.length) listAppend_(sh, rows);
    return text_("ok");
  }
  if (d.action === "listDone") {
    const r = listRowOf_(sh, d.id);
    if (!r) return text_("not found");
    sh.getRange(r, 7).setValue(d.done === true);
    return text_("ok");
  }
  if (d.action === "listDelete") {
    const r = listRowOf_(sh, d.id);
    if (!r) return text_("not found");
    sh.deleteRow(r);
    return text_("ok");
  }
  if (d.action === "listClearDone" || d.action === "listReset") {
    const v = sh.getDataRange().getValues();
    for (let i = v.length - 1; i >= 1; i--) {
      if (!v[i][0]) continue;
      if (d.action === "listClearDone" && v[i][6] === true) sh.deleteRow(i + 1);
      if (d.action === "listReset" && v[i][6] === true) sh.getRange(i + 1, 7).setValue(false);
    }
    return text_("ok");
  }
  return text_("unknown action");
}

// ===== סגירת יום =====
// מעביר את כל ההזמנות ללשונית ארכיון לפי תאריך, ולשונית ההזמנות מתחילה מחדש ריקה.
function closeDay_() {
  const ss = SpreadsheetApp.getActive();
  const sh = sheet_(ORDERS);
  const data = sh.getDataRange().getValues();
  const move = [];
  for (let i = 1; i < data.length; i++) if (data[i][C.no]) move.push(i + 1);
  if (!move.length) return text_("nothing");

  const name = "ארכיון " + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd.MM.yyyy");
  let ar = ss.getSheetByName(name);
  if (!ar) {
    // העתק של לשונית ההזמנות שומר על העיצוב, הרוחבים והכיוון
    ar = sh.copyTo(ss).setName(name);
    if (ar.getLastRow() > 1) ar.getRange(2, 1, ar.getMaxRows() - 1, ar.getMaxColumns()).clearContent().clearDataValidations();
    ss.setActiveSheet(ar); ss.moveActiveSheet(ss.getNumSheets());
    ss.setActiveSheet(sh);
  } else {
    // אם כבר נסגר יום היום – מוחקים את שורת הסיכום הקודמת ומוסיפים אחריה
    const last = ar.getLastRow();
    if (last > 1 && String(ar.getRange(last, 1).getValue()).indexOf("סיכום") === 0) ar.deleteRow(last);
  }

  const rows = move.map(function (r) { return data[r - 1].slice(0, HEADERS.length); });
  const start = Math.max(ar.getLastRow(), 1) + 1;
  ar.getRange(start, 1, rows.length, HEADERS.length).setValues(rows);

  // שורת סיכום: כמה הזמנות נמסרו ובאיזה סכום
  const all = ar.getRange(2, 1, ar.getLastRow() - 1, HEADERS.length).getValues();
  let n = 0, sum = 0;
  all.forEach(function (r) { if (r[C.status] === "נמסרה") { n++; sum += Number(r[C.total]) || 0; } });
  const summary = new Array(HEADERS.length).fill("");
  summary[0] = "סיכום"; summary[C.name] = n + " הזמנות נמסרו"; summary[C.total] = sum;
  ar.getRange(ar.getLastRow() + 1, 1, 1, HEADERS.length).setValues([summary]).setFontWeight("bold").setBackground("#fff4d6");

  // מרוקנים את לשונית ההזמנות (העיצוב והכותרות נשארים)
  sh.deleteRows(2, data.length - 1);
  return text_("ok:" + move.length + ":" + name);
}

// האם האתר מקבל הזמנות (נשמר בהגדרות הסקריפט; ברירת מחדל: פתוח)
function isOpen_() { return PropertiesService.getScriptProperties().getProperty("open") !== "no"; }

// מחזיר את מספר השורה בגיליון של הזמנה, או 0
function findOrderRow_(data, orderNo) {
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][C.no]) === String(orderNo)) return i + 1;
  }
  return 0;
}

// ===== קריאה =====
function doGet(e) {
  const p = e.parameter || {};

  if (p.action === "open") return json_({ open: isOpen_() });

  if (p.action === "orders") {
    if (!isViewer_(p.key)) return json_({ error: "unauthorized" });
    const v = sheet_(ORDERS).getDataRange().getValues().slice(1);
    const orders = v.map(function (r, i) {
      const t = r[C.date];
      return {
        row: i + 2, name: r[C.name], phone: String(r[C.phone]), items: r[C.items], note: r[C.note],
        total: r[C.total], status: r[C.status],
        time: t instanceof Date ? t.toISOString() : String(t), orderNo: String(r[C.no])
      };
    }).filter(function (o) { return o.orderNo; }).reverse();
    return json_({ orders: orders, statuses: STATUSES, open: isOpen_(), role: role_(p.key) });
  }

  if (p.action === "lists") {
    if (!isViewer_(p.key)) return json_({ error: "unauthorized" });
    const lists = {};
    Object.keys(LISTS).forEach(function (k) { lists[k] = listItems_(k); });
    return json_({ lists: lists, recipes: recipeRows_(), role: role_(p.key) });
  }

  if (p.action === "staff") {
    if (!isAdmin_(p.key)) return json_({ error: "unauthorized" });
    return json_({ staff: staffList_().map(function (s) { return s.n; }) });
  }

  if (p.action === "menuall") {
    if (!isViewer_(p.key)) return json_({ error: "unauthorized" });
    return json_({ menu: menuRows_(), role: role_(p.key) });
  }

  // תפריט ציבורי (רק מנות פעילות)
  return json_(menuRows_().filter(function (m) { return m.name && String(m.active).trim() !== "לא"; }));
}

// ===== כתיבה =====
function doPost(e) {
  const d = JSON.parse(e.postData.contents);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (d.action) {
      if (isAdmin_(d.key)) return adminOp_(d);
      // צוות: רק שינוי סטטוס של הזמנה, בלי ביטול
      if (d.action === "setStatus" && staffOf_(d.key)) return staffStatus_(d);
      return text_("unauthorized");
    }
    return customerOrder_(d);
  } finally {
    lock.releaseLock();
  }
}

// הזמנה מהאתר הציבורי
function customerOrder_(d) {
  const sh = sheet_(ORDERS);
  const orderNo = clean_(d.orderNo);
  const data = sh.getDataRange().getValues();
  const r = findOrderRow_(data, orderNo);
  if (r) {
    // לקוח יכול רק לדווח תשלום על הזמנה חדשה
    if (d.status === "דווח תשלום" && data[r - 1][C.status] === "חדשה") sh.getRange(r, C.status + 1).setValue("דווח תשלום");
    return text_("updated");
  }
  if (!isOpen_()) return text_("closed"); // הזמנות סגורות – לא נרשמות הזמנות חדשות
  const now = new Date();
  // אותו רגע בשתי העמודות: A מוצגת כתאריך, B כשעה
  sh.appendRow([now, now, orderNo, clean_(d.name), "'" + clean_(d.phone), clean_(String(d.items || "").replace(/;\s*/g, "\n")), clean_(d.note),
    Number(d.total) || 0, "חדשה"]);
  return text_("ok");
}

// שינוי סטטוס על ידי איש צוות: אסור לבטל הזמנה, ואסור לגעת בהזמנה שכבר בוטלה
function staffStatus_(d) {
  if (STATUSES.indexOf(d.status) < 0) return text_("bad status");
  if (d.status === "בוטלה") return text_("forbidden");
  const sh = sheet_(ORDERS);
  const data = sh.getDataRange().getValues();
  const r = findOrderRow_(data, d.orderNo);
  if (!r) return text_("not found");
  if (data[r - 1][C.status] === "בוטלה") return text_("forbidden");
  sh.getRange(r, C.status + 1).setValue(d.status);
  return text_("ok");
}

// פעולות ניהול (דורשות מפתח)
function adminOp_(d) {
  if (/^list|^recipe/.test(d.action)) return listOp_(d);

  if (d.action === "staffAdd") {
    const name = clean_(d.name).trim().slice(0, 40), pass = String(d.pass || "");
    if (!name) return text_("bad name");
    const list = staffList_();
    if (list.some(function (s) { return s.n === name; })) return text_("name exists");
    const bad = checkPass_(list, pass, null);
    if (bad) return text_(bad);
    list.push({ n: name, h: hash_(pass), e: enc_(pass) });
    staffSave_(list);
    return text_("ok");
  }
  if (d.action === "adminCheck") return text_(reauth_(d) ? "ok" : "bad admin");
  if (d.action === "staffReveal") {
    if (!reauth_(d)) return text_("bad admin");
    const s = staffList_().filter(function (x) { return x.n === d.name; })[0];
    if (!s) return text_("not found");
    const p = dec_(s.e);
    // אם הסיסמה נשמרה לפני האפשרות הזאת, או שסיסמת המנהל הוחלפה מאז – אי אפשר לשחזר
    if (p === null || hash_(p) !== s.h) return text_("no reveal");
    return text_("pass:" + p);
  }
  if (d.action === "staffChange") {
    if (!reauth_(d)) return text_("bad admin");
    const list = staffList_(), pass = String(d.pass || "");
    const s = list.filter(function (x) { return x.n === d.name; })[0];
    if (!s) return text_("not found");
    const bad = checkPass_(list, pass, s.n);
    if (bad) return text_(bad);
    s.h = hash_(pass); s.e = enc_(pass);
    staffSave_(list);
    return text_("ok");
  }
  if (d.action === "staffDelete") {
    if (!reauth_(d)) return text_("bad admin");
    staffSave_(staffList_().filter(function (s) { return s.n !== d.name; }));
    return text_("ok");
  }

  if (d.action === "setStatus") {
    if (STATUSES.indexOf(d.status) < 0) return text_("bad status");
    const sh = sheet_(ORDERS);
    const r = findOrderRow_(sh.getDataRange().getValues(), d.orderNo);
    if (!r) return text_("not found");
    sh.getRange(r, C.status + 1).setValue(d.status);
    return text_("ok");
  }

  if (d.action === "setOpen") {
    PropertiesService.getScriptProperties().setProperty("open", d.open ? "yes" : "no");
    return text_("ok");
  }

  if (d.action === "closeDay") return closeDay_();

  if (d.action === "deleteOrder") {
    const sh = sheet_(ORDERS);
    const r = findOrderRow_(sh.getDataRange().getValues(), d.orderNo);
    if (!r) return text_("not found");
    sh.deleteRow(r);
    return text_("ok");
  }

  if (d.action === "saveMenuRow" || d.action === "addMenuRow") {
    const m = d.item || {};
    const price = Number(m.price);
    if (!m.name || isNaN(price) || price < 0) return text_("bad item");
    const active = m.active === "לא" ? "לא" : "כן";
    const spicy = Math.max(0, Math.min(5, Math.round(Number(m.spicy) || 0)));
    const vegan = m.vegan === "כן" ? "כן" : "לא";
    const row = [clean_(m.cat), clean_(m.name), clean_(m.desc), price, active, clean_(m.img), spicy, vegan]; // G: רמת חריפות 0-5, H: טבעונית
    const sh = sheet_(MENU);
    if (d.action === "addMenuRow") {
      sh.appendRow(row);
    } else {
      const r = Number(m.row);
      if (!(r >= 2 && r <= sh.getLastRow())) return text_("bad row");
      sh.getRange(r, 1, 1, 8).setValues([row]);
    }
    return text_("ok");
  }
  return text_("unknown action");
}

// ===== הרצה חד-פעמית =====
// ממיר את גיליון ההזמנות למבנה החדש (תאריך ושעה בעמודות נפרדות).
// מזהה גם שורות במבנה הישן וגם במבנה הקודם. בטוח להריץ שוב – אם כבר הומר, לא עושה כלום.
function migrateLayout() {
  const sh = sheet_(ORDERS);
  const data = sh.getDataRange().getValues();
  if (data[0][0] === HEADERS[0] && data[0][1] === HEADERS[1]) { Logger.log("כבר במבנה החדש"); return; }

  const isNo = function (v) { return /^J\d/.test(String(v)); };
  const fixPhone = function (p) { p = String(p); return /^5\d{8}$/.test(p) ? "0" + p : p; };
  const rows = [];
  data.slice(1).forEach(function (r) {
    let o;
    if (isNo(r[1])) {        // ישן: זמן, מס׳ הזמנה, שם, טלפון, מנות, הערה, סכום, סטטוס
      o = { t: r[0], no: r[1], name: r[2], phone: r[3], items: r[4], note: r[5], total: r[6], status: r[7] };
    } else if (isNo(r[7])) { // קודם: שם, טלפון, מנות, הערה, סכום, סטטוס, זמן, מס׳ הזמנה
      o = { t: r[6], no: r[7], name: r[0], phone: r[1], items: r[2], note: r[3], total: r[4], status: r[5] };
    } else return;           // שורה ריקה או לא מזוהה
    rows.push([o.t, o.t, o.no, o.name, fixPhone(o.phone), o.items, o.note, o.total, o.status || "חדשה"]);
  });

  if (sh.getMaxColumns() < HEADERS.length) sh.insertColumnsAfter(sh.getMaxColumns(), HEADERS.length - sh.getMaxColumns());
  const maxRows = sh.getMaxRows();
  sh.getRange(1, 1, maxRows, sh.getMaxColumns()).clearContent().clearDataValidations();
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);

  // פורמטים לכל העמודה, כך שגם הזמנות חדשות יוצגו נכון
  sh.getRange(2, C.date + 1, maxRows - 1).setNumberFormat("dd/MM/yyyy");
  sh.getRange(2, C.time + 1, maxRows - 1).setNumberFormat("HH:mm");
  sh.getRange(2, C.no + 1, maxRows - 1).setNumberFormat("@");
  sh.getRange(2, C.phone + 1, maxRows - 1).setNumberFormat("@");
  if (rows.length) sh.getRange(2, 1, rows.length, HEADERS.length).setValues(rows);

  const rule = SpreadsheetApp.newDataValidation().requireValueInList(STATUSES, true).setAllowInvalid(true).build();
  sh.getRange(2, C.status + 1, maxRows - 1).setDataValidation(rule);

  sh.setRightToLeft(true);
  // הכל ממורכז, חוץ ממנות והערה שמיושרות לימין
  sh.getRange(1, 1, maxRows, HEADERS.length).setHorizontalAlignment("center").setVerticalAlignment("top");
  sh.getRange(1, C.items + 1, maxRows, 2).setHorizontalAlignment("right");
  sh.getRange(1, C.items + 1, maxRows, 2).setWrap(true);
  [110, 70, 125, 150, 120, 620, 420, 90, 130].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  sh.setFrozenRows(1);
  Logger.log("הומרו " + rows.length + " הזמנות");
}

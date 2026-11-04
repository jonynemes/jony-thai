// ===== הגדרות =====
// השאירו כאן את המפתח הסודי שלכם (12 תווים ומעלה). אל תשתפו אותו ואל תעלו אותו ל-GitHub.
const ADMIN_KEY = "CHANGE_ME";

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
    return { row: i + 2, cat: r[0], name: r[1], desc: r[2], price: r[3], active: r[4], img: r[5] };
  });
}
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

  if (p.action === "orders") {
    if (!isAdmin_(p.key)) return json_({ error: "unauthorized" });
    const v = sheet_(ORDERS).getDataRange().getValues().slice(1);
    const orders = v.map(function (r, i) {
      const t = r[C.date];
      return {
        row: i + 2, name: r[C.name], phone: String(r[C.phone]), items: r[C.items], note: r[C.note],
        total: r[C.total], status: r[C.status],
        time: t instanceof Date ? t.toISOString() : String(t), orderNo: String(r[C.no])
      };
    }).filter(function (o) { return o.orderNo; }).reverse();
    return json_({ orders: orders, statuses: STATUSES });
  }

  if (p.action === "menuall") {
    if (!isAdmin_(p.key)) return json_({ error: "unauthorized" });
    return json_({ menu: menuRows_() });
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
      if (!isAdmin_(d.key)) return text_("unauthorized");
      return adminOp_(d);
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
  const now = new Date();
  // אותו רגע בשתי העמודות: A מוצגת כתאריך, B כשעה
  sh.appendRow([now, now, orderNo, clean_(d.name), "'" + clean_(d.phone), clean_(String(d.items || "").replace(/;\s*/g, "\n")), clean_(d.note),
    Number(d.total) || 0, "חדשה"]);
  return text_("ok");
}

// פעולות ניהול (דורשות מפתח)
function adminOp_(d) {
  if (d.action === "setStatus") {
    if (STATUSES.indexOf(d.status) < 0) return text_("bad status");
    const sh = sheet_(ORDERS);
    const r = findOrderRow_(sh.getDataRange().getValues(), d.orderNo);
    if (!r) return text_("not found");
    sh.getRange(r, C.status + 1).setValue(d.status);
    return text_("ok");
  }

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
    const row = [clean_(m.cat), clean_(m.name), clean_(m.desc), price, active, clean_(m.img)];
    const sh = sheet_(MENU);
    if (d.action === "addMenuRow") {
      sh.appendRow(row);
    } else {
      const r = Number(m.row);
      if (!(r >= 2 && r <= sh.getLastRow())) return text_("bad row");
      sh.getRange(r, 1, 1, 6).setValues([row]);
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

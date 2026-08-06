import React, { useState, useEffect, useMemo, useRef, useCallback } from "react";
import * as XLSX from "xlsx";
import {
  authApi, teamApi, projectsApi, boqApi, saveProject, snapshotOf, newId,
} from "./supabaseClient";

/* ============================================================
   SEGEV BOQ — גרסת ווב
   מבנה זהה לגרסת הדסקטופ:
   פרויקטים · כתב כמויות · תקציב · לוח זמנים · רכש · דוחות
   ============================================================ */

const VAT_RATE = 0.18;
const uid = () => newId();

const UNITS = ["מ\"ר", "מ\"א", "יח'", "קומפ'", "מ\"ק", "טון", "ק\"ג", "יום"];

/* ---------- שירותים: זהים ללוגיקה של הדסקטופ ---------- */

/* schedule_service.py */
const ST_DONE_ON_TIME = "הושלם בזמן";
const ST_DONE_LATE = "הושלם באיחור";
const ST_DELAYED = "באיחור";
const ST_UPCOMING = "מתוכנן";
const ST_UNKNOWN = "לא ידוע";

const parseDate = (v) => {
  if (!v) return null;
  const d = new Date(String(v).trim());
  return isNaN(d.getTime()) ? null : d;
};

const milestoneStatus = (m, today = new Date()) => {
  const planned = parseDate(m.planned_date);
  const actual = parseDate(m.actual_date);
  if (m.done) {
    if (planned && actual) return actual <= planned ? ST_DONE_ON_TIME : ST_DONE_LATE;
    return ST_DONE_ON_TIME;
  }
  if (planned && planned < today) return ST_DELAYED;
  if (planned) return ST_UPCOMING;
  return ST_UNKNOWN;
};

const scheduleSummary = (ms, today = new Date()) => ({
  total: ms.length,
  done: ms.filter((m) => m.done).length,
  delayed: ms.filter((m) => milestoneStatus(m, today) === ST_DELAYED).length,
});

/* procurement_service.py */
const PO_ORDERED = "הוזמן";
const PO_RECEIVED = "התקבל";
const PO_CANCELLED = "בוטל";
const PO_STATUSES = [PO_ORDERED, PO_RECEIVED, PO_CANCELLED];

const orderTotal = (o) => (+o.quantity || 0) * (+o.unit_price || 0);

const procurementSummary = (orders) => ({
  ordered: orders
    .filter((o) => o.status !== PO_CANCELLED)
    .reduce((s, o) => s + orderTotal(o), 0),
  received: orders
    .filter((o) => o.status === PO_RECEIVED)
    .reduce((s, o) => s + orderTotal(o), 0),
});

/* budget.py */
const emptyBudget = () => ({
  planned_materials: 0,
  planned_labor: 0,
  planned_subcontractors: 0,
  actual_materials: 0,
  actual_labor: 0,
  actual_subcontractors: 0,
});

const budgetTotals = (b) => {
  const planned =
    (+b.planned_materials || 0) +
    (+b.planned_labor || 0) +
    (+b.planned_subcontractors || 0);
  const actual =
    (+b.actual_materials || 0) +
    (+b.actual_labor || 0) +
    (+b.actual_subcontractors || 0);
  return { planned, actual, remaining: planned - actual };
};

/* ---------- תבניות פרקים ---------- */

const TEMPLATES = {
  renovation: {
    label: "שיפוץ",
    hint: "דירה, משרד, שיפוץ כללי",
    chapters: [
      "עבודות הריסה ופינוי",
      "עבודות שלד ובנייה",
      "אינסטלציה",
      "חשמל",
      "עבודות גמר",
      "נגרות ואלומיניום",
    ],
  },
  construction: {
    label: "בנייה",
    hint: "מבנה חדש, תוספת בנייה",
    chapters: [
      "עבודות עפר",
      "עבודות בטון יצוק באתר",
      "עבודות בנייה",
      "עבודות איטום",
      "עבודות טיח",
      "עבודות ריצוף וחיפוי",
      "עבודות אלומיניום",
      "עבודות מסגרות ונגרות",
      "מתקני תברואה",
      "מתקני חשמל",
      "עבודות פיתוח",
    ],
  },
  custom: { label: "ריק", hint: "אבנה את הפרקים לבד", chapters: [] },
};

const makeChapters = (kind) =>
  (TEMPLATES[kind]?.chapters || []).map((title) => ({
    id: uid(),
    title,
    items: [],
  }));

/* ---------- מחירון ---------- */

const CATALOG = [
  { name: "הריסת קיר בלוקים", cat: "הריסה", unit: "מ\"ר", price: 45 },
  { name: "הריסת ריצוף קיים", cat: "הריסה", unit: "מ\"ר", price: 38 },
  { name: "פינוי פסולת בניין", cat: "הריסה", unit: "טון", price: 35 },
  { name: "יציקת רצפת בטון", cat: "שלד", unit: "מ\"ק", price: 120 },
  { name: "בניית קיר בלוקים", cat: "שלד", unit: "מ\"ר", price: 90 },
  { name: "יציקת תקרה", cat: "שלד", unit: "מ\"ר", price: 210 },
  { name: "התקנת צנרת מים", cat: "אינסטלציה", unit: "מ\"א", price: 60 },
  { name: "התקנת אסלה וכיור", cat: "אינסטלציה", unit: "יח'", price: 450 },
  { name: "נקודת ביוב", cat: "אינסטלציה", unit: "יח'", price: 320 },
  { name: "נקודת חשמל", cat: "חשמל", unit: "יח'", price: 150 },
  { name: "לוח חשמל ראשי", cat: "חשמל", unit: "קומפ'", price: 1800 },
  { name: "נקודת תאורה", cat: "חשמל", unit: "יח'", price: 130 },
  { name: "ריצוף גרניט פורצלן", cat: "גמר", unit: "מ\"ר", price: 150 },
  { name: "חיפוי קרמיקה", cat: "גמר", unit: "מ\"ר", price: 140 },
  { name: "צביעת קירות", cat: "גמר", unit: "מ\"ר", price: 25 },
  { name: "איטום גג", cat: "גמר", unit: "מ\"ר", price: 80 },
  { name: "התקנת דלת פנים", cat: "נגרות", unit: "יח'", price: 1450 },
  { name: "ארון מטבח תחתון", cat: "נגרות", unit: "מ\"א", price: 2200 },
  { name: "חפירה כללית בכלים מכניים", cat: "עבודות עפר", unit: "מ\"ק", price: 42 },
  { name: "מילוי מובא והידוק בשכבות", cat: "עבודות עפר", unit: "מ\"ק", price: 68 },
  { name: "חציבה בסלע", cat: "עבודות עפר", unit: "מ\"ק", price: 165 },
  { name: "בטון רזה מתחת ליסודות", cat: "עבודות בטון", unit: "מ\"ק", price: 520 },
  { name: "בטון ב-30 ליסודות", cat: "עבודות בטון", unit: "מ\"ק", price: 720 },
  { name: "בטון ב-30 לעמודים וקורות", cat: "עבודות בטון", unit: "מ\"ק", price: 890 },
  { name: "תבניות לעמודים", cat: "עבודות בטון", unit: "מ\"ר", price: 105 },
  { name: "תבניות לתקרות", cat: "עבודות בטון", unit: "מ\"ר", price: 95 },
  { name: "ברזל זיון מצולע", cat: "עבודות בטון", unit: "טון", price: 4900 },
  { name: "רשת ברזל מרותכת", cat: "עבודות בטון", unit: "מ\"ר", price: 28 },
  { name: "בנייה בבלוקי בטון 20 ס\"מ", cat: "עבודות בנייה", unit: "מ\"ר", price: 118 },
  { name: "בנייה בבלוקי איטונג", cat: "עבודות בנייה", unit: "מ\"ר", price: 145 },
  { name: "מחיצות גבס דו-שכבתי", cat: "עבודות בנייה", unit: "מ\"ר", price: 175 },
  { name: "איטום יריעות ביטומניות", cat: "איטום", unit: "מ\"ר", price: 92 },
  { name: "איטום ממברנה לגגות", cat: "איטום", unit: "מ\"ר", price: 110 },
  { name: "איטום חדרים רטובים", cat: "איטום", unit: "מ\"ר", price: 78 },
  { name: "טיח פנים שתי שכבות", cat: "טיח", unit: "מ\"ר", price: 52 },
  { name: "טיח חוץ מיישר", cat: "טיח", unit: "מ\"ר", price: 78 },
  { name: "חלון אלומיניום כנף על כנף", cat: "אלומיניום", unit: "מ\"ר", price: 1250 },
  { name: "ויטרינה אלומיניום", cat: "אלומיניום", unit: "מ\"ר", price: 1680 },
  { name: "תריס גלילה חשמלי", cat: "אלומיניום", unit: "מ\"ר", price: 780 },
  { name: "מעקה מסגרות מגולוון", cat: "מסגרות", unit: "מ\"א", price: 620 },
  { name: "דלת פלדה רב-בריח", cat: "מסגרות", unit: "יח'", price: 4200 },
  { name: "ריצוף אבן משתלבת", cat: "פיתוח", unit: "מ\"ר", price: 135 },
  { name: "קיר תומך מבטון", cat: "פיתוח", unit: "מ\"ק", price: 980 },
  { name: "גינון והשקיה", cat: "פיתוח", unit: "מ\"ר", price: 95 },
];

/* ---------- ייבוא מאקסל ---------- */

/* הסדר קובע: "מחיר יחידה" חייב להיבדק כמחיר לפני שהמילה
   "יחידה" תגרור אותו לעמודת היחידות. */
const COL_RULES = [
  { key: "price", words: ["מחיר"] },
  { key: "total", words: ["סה\"כ", "סהכ", "סכום", "סך"] },
  { key: "qty", words: ["כמות"] },
  { key: "unit", words: ["יחידה", "יח", "מידה"] },
  { key: "desc", words: ["תיאור", "תאור", "פירוט", "נושא", "העבודה", "מהות", "פריט"] },
  { key: "code", words: ["סעיף", "מספר", "מס'", "קוד", "סידורי", "קטלוג"] },
];

/* קוד כמו 01.20.3100 מקודד היררכיה:
   01 = פרק ראשי, 01.20 = תת-פרק, 01.20.3100 = סעיף. */
const codeDepth = (code) =>
  code ? String(code).trim().split(".").filter(Boolean).length : 0;

const guessColumn = (header) => {
  const h = String(header || "").trim();
  if (!h) return null;
  for (const rule of COL_RULES) {
    if (rule.words.some((w) => h.includes(w))) return rule.key;
  }
  return null;
};

const findHeaderRow = (rows) => {
  let best = { row: 0, score: 0 };
  rows.slice(0, 25).forEach((row, i) => {
    const keys = new Set(row.map(guessColumn).filter(Boolean));
    if (keys.size > best.score) best = { row: i, score: keys.size };
  });
  return best.score >= 2 ? best.row : 0;
};

const autoMap = (headerCells) => {
  const map = {};
  headerCells.forEach((cell, idx) => {
    const key = guessColumn(cell);
    if (key && map[key] === undefined) map[key] = idx;
  });
  return map;
};

const toNumber = (v) => {
  if (typeof v === "number") return v;
  const s = String(v ?? "").replace(/[^\d.,-]/g, "").replace(/,/g, "");
  const n = parseFloat(s);
  return Number.isFinite(n) ? n : 0;
};

const looksLikeChapter = (rec) => rec.desc && !rec.qty && !rec.price && !rec.unit;

/* כותרת פרק בתא ממוזג נוחתת בעמודה הראשונה, כלומר בעמודת
   הקוד. אם שם יושב טקסט ולא מספר סעיף — זו הכותרת. */
const rescueMergedTitle = (rec) => {
  if (rec.desc || !rec.code) return rec;
  if (/^[\d.\-/\s]+$/.test(rec.code)) return rec;
  return { ...rec, desc: rec.code, code: "" };
};

/* יחידות מוכרות. כל דבר אחר מסומן לבדיקה במקום להיות מנוחש —
   יחידה שגויה מכפילה מחיר בסדר גודל שלם. */
const KNOWN_UNITS = [
  "מ\"ר", "מ\"ק", "מ\"א", "מ'", "ס\"מ",
  "יח'", "יח", "יחידה", "קומפ'", "קומפלט", "קומפ",
  "טון", "ק\"ג", "ליטר", "יום", "שעה", "ש\"ע", "נק'", "נקודה", "זוג", "סט",
];

const JUNK_UNITS = ["יח' רשום", "רשום", "הערה", "-", "—"];

const unitOk = (u) => {
  const s = String(u || "").trim();
  if (!s) return false;
  if (JUNK_UNITS.includes(s)) return false;
  return KNOWN_UNITS.some((k) => s === k || s.startsWith(k));
};

/* ---------- עיצוב מספרים ---------- */

const shekel = (n) =>
  new Intl.NumberFormat("he-IL", {
    style: "currency",
    currency: "ILS",
    maximumFractionDigits: 0,
  }).format(Number.isFinite(+n) ? +n : 0);

const num = (n, d = 2) =>
  new Intl.NumberFormat("he-IL", { maximumFractionDigits: d }).format(
    Number.isFinite(+n) ? +n : 0
  );

const pct = (n) => `${(Number.isFinite(n) ? n : 0).toFixed(1)}%`;
const chapterNo = (i) => String(i + 1).padStart(2, "0");
const today = () => new Date().toISOString().slice(0, 10);

/* ---------- אחסון ---------- */

const ROLE_ADMIN = "מנהל מערכת";
const ROLE_DEVELOPER = "יזם";
const ROLE_CONTRACTOR = "קבלן";
const ROLE_PM = "מנהל פרויקט";
const ROLE_ENGINEER = "מהנדס ביצוע";
const ROLE_FOREMAN = "מנהל עבודה";

const ROLES = [
  ROLE_ADMIN,
  ROLE_DEVELOPER,
  ROLE_CONTRACTOR,
  ROLE_PM,
  ROLE_ENGINEER,
  ROLE_FOREMAN,
];

const NONE = "none";
const VIEW = "view";
const EDIT = "edit";
const LEVELS = [
  { key: NONE, label: "אין" },
  { key: VIEW, label: "צפייה" },
  { key: EDIT, label: "עריכה" },
];

const MODULES = [
  { key: "boq", label: "כתב כמויות" },
  { key: "execution", label: "ביצוע" },
  { key: "changes", label: "שינויים" },
  { key: "site", label: "פיקוח" },
  { key: "budget", label: "תקציב" },
  { key: "schedule", label: "לוח זמנים" },
  { key: "procurement", label: "רכש" },
  { key: "report", label: "דוחות" },
];

const FLAGS = [
  { key: "seePrices", label: "רואה מחירים", note: "בלי זה יראה כמויות בלבד" },
  { key: "seeProfit", label: "רואה רווחיות", note: "עלויות ושולי רווח" },
  { key: "approveChanges", label: "מאשר שינויים" },
  { key: "issueDocs", label: "מפיק הצעות וחשבונות" },
  { key: "deleteProject", label: "מוחק פרויקטים" },
  { key: "exportData", label: "מייצא נתונים" },
  { key: "manageTeam", label: "מנהל משתמשים" },
];

const mods = (spec) => {
  const out = {};
  MODULES.forEach((m) => (out[m.key] = spec[m.key] ?? NONE));
  return out;
};

const flags = (spec) => {
  const out = {};
  FLAGS.forEach((f) => (out[f.key] = !!spec[f.key]));
  return out;
};

const ROLE_PRESETS = {
  [ROLE_ADMIN]: {
    modules: mods(Object.fromEntries(MODULES.map((m) => [m.key, EDIT]))),
    flags: flags(Object.fromEntries(FLAGS.map((f) => [f.key, true]))),
  },

  /* היזם רואה התקדמות ומה שאושר. לא רואה את העלויות של הקבלן. */
  [ROLE_DEVELOPER]: {
    modules: mods({
      boq: VIEW, execution: VIEW, changes: VIEW,
      site: VIEW, schedule: VIEW, report: VIEW,
    }),
    flags: flags({ seePrices: true, approveChanges: true }),
  },

  /* הקבלן מנהל את העבודה והרכש שלו. */
  [ROLE_CONTRACTOR]: {
    modules: mods({
      boq: EDIT, execution: EDIT, changes: EDIT,
      site: VIEW, schedule: VIEW, procurement: EDIT,
    }),
    flags: flags({ seePrices: true, issueDocs: true, exportData: true }),
  },

  /* מנהל הפרויקט רואה הכל חוץ מניהול משתמשים. */
  [ROLE_PM]: {
    modules: mods(Object.fromEntries(MODULES.map((m) => [m.key, EDIT]))),
    flags: flags({
      seePrices: true, seeProfit: true, approveChanges: true,
      issueDocs: true, exportData: true,
    }),
  },

  /* מהנדס ביצוע — כמויות וטכניקה, בלי כסף. */
  [ROLE_ENGINEER]: {
    modules: mods({
      boq: EDIT, execution: EDIT, site: EDIT,
      schedule: EDIT, procurement: VIEW, changes: VIEW,
    }),
    flags: flags({ seePrices: true, exportData: true }),
  },

  /* מנהל עבודה — מה שקורה באתר היום. */
  [ROLE_FOREMAN]: {
    modules: mods({ execution: EDIT, site: EDIT, schedule: VIEW, boq: VIEW }),
    flags: flags({}),
  },
};

const permsFor = (user) => {
  const preset = ROLE_PRESETS[user?.role] || ROLE_PRESETS[ROLE_FOREMAN];
  return {
    modules: { ...preset.modules, ...(user?.perms?.modules || {}) },
    flags: { ...preset.flags, ...(user?.perms?.flags || {}) },
  };
};

const canSee = (perms, key) => perms.modules[key] !== NONE;
const canEdit = (perms, key) => perms.modules[key] === EDIT;

/* השרת מחזיר snake_case; האפליקציה עובדת ב-camelCase */
const toUser = (member, email) =>
  member && {
    id: member.id,
    email: email || "",
    name: member.full_name,
    role: HE_ROLE[member.role] || ROLE_FOREMAN,
    perms: member.perms || { modules: {}, flags: {} },
  };

/* התפקידים נשמרים באנגלית במסד ומוצגים בעברית */
const ROLE_DB = {
  [ROLE_ADMIN]: "admin",
  [ROLE_DEVELOPER]: "developer",
  [ROLE_CONTRACTOR]: "contractor",
  [ROLE_PM]: "pm",
  [ROLE_ENGINEER]: "engineer",
  [ROLE_FOREMAN]: "foreman",
};
const HE_ROLE = Object.fromEntries(
  Object.entries(ROLE_DB).map(([he, en]) => [en, he])
);

/* שלד ריק, כדי שקומפוננטות לא יקרסו על שדה חסר */
const blankProject = () => ({
  id: null, name: "", description: "", client_name: "", address: "",
  start_date: "", end_date: "", site_manager_name: "", site_manager_phone: "",
  kind: "renovation", overhead: 0,
  chapters: [], budget: null, milestones: [], orders: [],
  changes: [], visits: [], defects: [],
});

const boqTotal = (chapters) =>
  chapters.reduce(
    (s, c) => s + c.items.reduce((a, it) => a + (+it.qty || 0) * (+it.price || 0), 0),
    0
  );

/* ---------- ביצוע ----------
   done = הכמות שבוצעה בפועל בסעיף. ההפרש בינה לבין הכמות
   בחוזה הוא מה שנשאר, והמכפלה במחיר היא מה שמגיע לתשלום. */

const itemContract = (it) => (+it.qty || 0) * (+it.price || 0);
const itemDone = (it) => (+it.done || 0) * (+it.price || 0);

const executionTotals = (chapters) => {
  let contract = 0;
  let done = 0;
  let items = 0;
  let started = 0;
  let complete = 0;

  chapters.forEach((c) =>
    c.items.forEach((it) => {
      const q = +it.qty || 0;
      const d = +it.done || 0;
      contract += itemContract(it);
      done += itemDone(it);
      items += 1;
      if (d > 0) started += 1;
      if (q > 0 && d >= q) complete += 1;
    })
  );

  return {
    contract,
    done,
    remaining: contract - done,
    percent: contract > 0 ? (done / contract) * 100 : 0,
    items,
    started,
    complete,
  };
};

/* חריגה: בוצע יותר ממה שבחוזה */
const itemOver = (it) => (+it.done || 0) > (+it.qty || 0);

/* ---------- שינויים ותוספות ----------
   סעיף שלא היה בחוזה המקורי. מוחזק בנפרד כדי שתמיד יהיה
   אפשר להראות ללקוח מה היה בחוזה ומה נוסף אחריו, ובאיזה
   סטטוס אישור. זה מה שמונע ויכוח בסוף הפרויקט. */

const CO_DRAFT = "טיוטה";
const CO_SUBMITTED = "הוגש לאישור";
const CO_APPROVED = "אושר";
const CO_REJECTED = "נדחה";
const CO_STATUSES = [CO_DRAFT, CO_SUBMITTED, CO_APPROVED, CO_REJECTED];

const coTotal = (co) =>
  (co.items || []).reduce(
    (s, it) => s + (+it.qty || 0) * (+it.price || 0),
    0
  );

const changeSummary = (orders) => {
  const list = orders || [];
  const by = (st) => list.filter((c) => c.status === st);
  return {
    count: list.length,
    approved: by(CO_APPROVED).reduce((s, c) => s + coTotal(c), 0),
    approvedCount: by(CO_APPROVED).length,
    pending: by(CO_SUBMITTED).reduce((s, c) => s + coTotal(c), 0),
    pendingCount: by(CO_SUBMITTED).length,
    draft: by(CO_DRAFT).reduce((s, c) => s + coTotal(c), 0),
    draftCount: by(CO_DRAFT).length,
    rejectedCount: by(CO_REJECTED).length,
  };
};

/* ---------- פיקוח באתר ----------
   יומן ביקורים וליקויים. ליקוי הוא הדבר שנסגר או לא נסגר,
   ולכן הוא מנוהל עם אחראי, יעד וסטטוס — לא כרשימת הערות. */

const DEF_OPEN = "פתוח";
const DEF_FIXED = "תוקן";
const DEF_VERIFIED = "אומת";
const DEF_WAIVED = "בוטל";
const DEF_STATUSES = [DEF_OPEN, DEF_FIXED, DEF_VERIFIED, DEF_WAIVED];

const SEVERITIES = ["נמוכה", "בינונית", "גבוהה"];

const defectOpen = (d) => d.status === DEF_OPEN || d.status === DEF_FIXED;

const defectOverdue = (d, ref = new Date()) => {
  if (!defectOpen(d) || !d.due) return false;
  const due = parseDate(d.due);
  return !!due && due < ref;
};

const defectSummary = (list, ref = new Date()) => {
  const all = list || [];
  return {
    total: all.length,
    open: all.filter((d) => d.status === DEF_OPEN).length,
    fixed: all.filter((d) => d.status === DEF_FIXED).length,
    verified: all.filter((d) => d.status === DEF_VERIFIED).length,
    overdue: all.filter((d) => defectOverdue(d, ref)).length,
    high: all.filter((d) => d.severity === "גבוהה" && defectOpen(d)).length,
  };
};

const newVisit = () => ({
  id: uid(),
  date: today(),
  weather: "",
  workers: "",
  subcontractors: "",
  notes: "",
});

const newDefect = () => ({
  id: uid(),
  title: "",
  location: "",
  responsible: "",
  severity: "בינונית",
  status: DEF_OPEN,
  opened: today(),
  due: "",
  closed: "",
});

const newChangeOrder = (n) => ({
  id: uid(),
  number: n,
  title: "",
  reason: "",
  requested_by: "",
  date: today(),
  status: CO_DRAFT,
  items: [],
});

/* ============================================================
   APP
   ============================================================ */

const TABS = [
  { key: "boq", label: "כתב כמויות" },
  { key: "execution", label: "ביצוע" },
  { key: "changes", label: "שינויים" },
  { key: "site", label: "פיקוח" },
  { key: "budget", label: "תקציב" },
  { key: "schedule", label: "לוח זמנים" },
{ key: "procurement", label: "רכש" },
{ key: "crm", label: "CRM" },
{ key: "report", label: "דוחות" },
 ]; 
export default function SegevBoq() {
  const [index, setIndex] = useState([]);
  const [project, setProject] = useState(null);
  const [company, setCompany] = useState({ name: "", phone: "", vat: "" });
  const [tab, setTab] = useState("boq");
  const [booting, setBooting] = useState(true);
  const [saveState, setSaveState] = useState("idle");
  const [loadErr, setLoadErr] = useState("");
  const [user, setUser] = useState(null);
  const [orgId, setOrgId] = useState(null);
  const [showTeam, setShowTeam] = useState(false);

  const snapshot = useRef(null);
  const skipSave = useRef(true);

  const perms = useMemo(() => permsFor(user), [user]);
  const allowed = useMemo(
    () => MODULES.filter((m) => canSee(perms, m.key)).map((m) => m.key),
    [perms]
  );

  /* ---------- מי אני ---------- */
  const loadMe = useCallback(async () => {
    const session = await authApi.currentSession();
    if (!session) {
      setUser(null);
      setBooting(false);
      return;
    }
    const member = await authApi.me();
    if (!member) {
      /* משתמש קיים בלי שיוך לארגון — קורה אם ההרשמה נקטעה */
      setLoadErr("החשבון קיים אך אינו משויך לארגון. פנה למנהל המערכת.");
      setUser(null);
      setBooting(false);
      return;
    }
    setUser(toUser(member, session.user.email));
    setOrgId(member.org_id);

    const org = await authApi.myOrg();
    if (org) {
      setCompany({
        name: org.name || "",
        phone: org.phone || "",
        vat: org.vat_id || "",
      });
    }

    try {
      setIndex(await projectsApi.list());
    } catch (ex) {
      setLoadErr(ex.message);
    }
    setBooting(false);
  }, []);

  useEffect(() => {
    loadMe();
    return authApi.onChange((session) => {
      if (!session) {
        setUser(null);
        setProject(null);
        setIndex([]);
      }
    });
  }, [loadMe]);

  /* ---------- שמירה אוטומטית לענן ---------- */
  useEffect(() => {
    if (!project || booting || !user) return;
    if (skipSave.current) {
      skipSave.current = false;
      return;
    }
    const timer = setTimeout(async () => {
      setSaveState("saving");
      try {
        await saveProject(project, snapshot.current || {}, perms);
        snapshot.current = snapshotOf(project);
        setIndex((prev) =>
          prev.map((p) =>
            p.id === project.id
              ? { ...p, name: project.name, client_name: project.client_name,
                  address: project.address, updated_at: new Date().toISOString() }
              : p
          )
        );
        setSaveState("saved");
        setTimeout(() => setSaveState("idle"), 1400);
      } catch (ex) {
        setSaveState("error");
        setLoadErr(ex.message);
      }
    }, 900);
    return () => clearTimeout(timer);
  }, [project, booting, user, perms]);

  /* ---------- פעולות ---------- */
  const openProject = async (id) => {
    setLoadErr("");
    try {
      const full = await projectsApi.load(id);
      skipSave.current = true;
      snapshot.current = snapshotOf(full);
      setProject({ ...blankProject(), ...full });
      setTab(allowed[0] || "boq");
    } catch (ex) {
      setLoadErr(ex.message);
    }
  };

  const createProject = async (kind) => {
    setLoadErr("");
    try {
      const row = await projectsApi.create(orgId, { name: "", kind });
      const chapters = makeChapters(kind);
      if (chapters.length) {
        await boqApi.importChapters(row.id, chapters, false);
      }
      setIndex((prev) => [row, ...prev]);
      await openProject(row.id);
    } catch (ex) {
      setLoadErr(ex.message);
    }
  };

  const deleteProject = async (id, name) => {
    if (!confirm(`למחוק את "${name || "פרויקט ללא שם"}"?\n\nלא ניתן לשחזר.`)) return;
    try {
      await projectsApi.remove(id);
      setIndex((prev) => prev.filter((p) => p.id !== id));
      if (project?.id === id) setProject(null);
    } catch (ex) {
      setLoadErr(ex.message);
    }
  };

  const duplicateProject = async (id) => {
    try {
      const src = await projectsApi.load(id);
      const row = await projectsApi.create(orgId, {
        name: `${src.name || "פרויקט"} — עותק`,
        client_name: src.client_name,
        address: src.address,
        kind: src.kind,
        overhead: src.overhead,
      });
      const copy = src.chapters.map((c) => ({
        title: c.title,
        items: c.items.map((it) => ({ ...it, done: 0 })),
      }));
      if (copy.length) await boqApi.importChapters(row.id, copy, false);
      setIndex((prev) => [row, ...prev]);
    } catch (ex) {
      setLoadErr(ex.message);
    }
  };

  const signOut = async () => {
    await authApi.signOut();
    setUser(null);
    setProject(null);
    setIndex([]);
  };

  const patch = useCallback(
    (fields) => setProject((p) => (p ? { ...p, ...fields } : p)),
    []
  );

  /* שמירת פרטי החברה */
  const saveCompany = useCallback(async (next) => {
    setCompany(next);
    try {
      await authApi.updateOrg({
        name: next.name || "",
        phone: next.phone || null,
        vat_id: next.vat || null,
      });
    } catch {
      /* אין הרשאה לעדכן — נשאר מקומי לתצוגה */
    }
  }, []);

  useEffect(() => {
    if (user && tab && !allowed.includes(tab)) setTab(allowed[0] || "boq");
  }, [user, tab, allowed]);

  if (booting) {
    return (
      <div className="boq" dir="rtl">
        <style>{CSS}</style>
        <div className="boot">טוען…</div>
      </div>
    );
  }

  if (!user) {
    return (
      <div className="boq" dir="rtl">
        <style>{CSS}</style>
        <AuthScreen onSignedIn={() => { setBooting(true); loadMe(); }} />
      </div>
    );
  }

  return (
    <div className="boq" dir="rtl">
      <style>{CSS}</style>

      <header className="topbar">
        <div className="brand">
          <span className="mark" aria-hidden="true" />
          <div>
            <h1>{company.name || "SEGEV BOQ"}</h1>
            <p>{project ? project.name || "פרויקט ללא שם" : "ניהול פרויקטים"}</p>
          </div>
        </div>

        <div className="topActions">
          {project && (
            <>
              <span className={`saveDot ${saveState}`}>
                {saveState === "saving"
                  ? "שומר…"
                  : saveState === "saved"
                  ? "נשמר"
                  : saveState === "error"
                  ? "שמירה נכשלה"
                  : "נשמר אוטומטית"}
              </span>
              <button className="ghost" onClick={() => setProject(null)}>
                כל הפרויקטים
              </button>
            </>
          )}

          <div className="userChip">
            <span className="uname">{user.name}</span>
            <span className="urole">{user.role}</span>
          </div>
          {perms.flags.manageTeam && (
            <button className="ghost" onClick={() => setShowTeam(true)}>
              הרשאות
            </button>
          )}
          <button className="ghost" onClick={signOut}>
            יציאה
          </button>
        </div>
      </header>

      {loadErr && (
        <div className="page">
          <p className="warnBox">{loadErr}</p>
        </div>
      )}

      {!project ? (
        <ProjectsList
          index={index}
          onOpen={openProject}
          onCreate={createProject}
          onDelete={perms.flags.deleteProject ? deleteProject : null}
          onDuplicate={duplicateProject}
        />
      ) : (
        <>
          <nav className="tabs">
            {TABS.filter((x) => allowed.includes(x.key)).map((x) => (
              <button
                key={x.key}
                className={tab === x.key ? "on" : ""}
                onClick={() => setTab(x.key)}
              >
                {x.label}
              </button>
            ))}
          </nav>

          {tab === "boq" && (
            <BoqTab project={project} patch={patch} company={company} setCompany={saveCompany} perms={perms} />
          )}
          {tab === "execution" && (
            <ExecutionTab project={project} patch={patch} company={company} perms={perms} />
          )}
          {tab === "changes" && (
            <ChangesTab project={project} patch={patch} company={company} perms={perms} />
          )}
          {tab === "site" && <SiteTab project={project} patch={patch} />}
          {tab === "budget" && <BudgetTab project={project} patch={patch} />}
          {tab === "schedule" && <ScheduleTab project={project} patch={patch} />}
          {tab === "procurement" && <ProcurementTab project={project} patch={patch} />}
          {tab === "report" && <ReportTab project={project} />}
        </>
      )}

      {showTeam && <TeamDialog user={user} onClose={() => setShowTeam(false)} />}
    </div>
  );
}


/* ============================================================
   מסך כניסה
   ============================================================ */

function AuthScreen({ onSignedIn }) {
  const [mode, setMode] = useState("in");
  const [form, setForm] = useState({ email: "", password: "", name: "", company: "" });
  const [err, setErr] = useState("");
  const [info, setInfo] = useState("");
  const [busy, setBusy] = useState(false);

  const set = (k, v) => setForm((f) => ({ ...f, [k]: v }));

  const submit = async () => {
    setErr("");
    setInfo("");
    if (!form.email.trim() || !form.password) {
      setErr("יש למלא מייל וסיסמה.");
      return;
    }
    if (mode === "up" && !form.name.trim()) {
      setErr("יש למלא שם.");
      return;
    }
    setBusy(true);
    try {
      if (mode === "in") {
        await authApi.signIn(form.email, form.password);
      } else {
        const res = await authApi.signUpOwner({
          email: form.email,
          password: form.password,
          fullName: form.name,
          orgName: form.company,
        });
        if (res.needsEmailConfirm) {
          setInfo("נשלח אליך מייל לאישור. אשר אותו וחזור לכאן להתחבר.");
          setMode("in");
          setBusy(false);
          return;
        }
      }
      onSignedIn();
    } catch (ex) {
      setErr(ex.message || "משהו השתבש. נסה שוב.");
    } finally {
      setBusy(false);
    }
  };

  return (
    <div className="authWrap">
      <div className="authCard">
        <div className="authBrand">
          <span className="mark" aria-hidden="true" />
          <div>
            <h1>SEGEV BOQ</h1>
            <p>כתבי כמויות, תמחור ופיקוח</p>
          </div>
        </div>

        <div className="authTabs">
          <button className={mode === "in" ? "on" : ""} onClick={() => { setMode("in"); setErr(""); }}>
            כניסה
          </button>
          <button className={mode === "up" ? "on" : ""} onClick={() => { setMode("up"); setErr(""); }}>
            פתיחת חשבון
          </button>
        </div>

        <div className="authForm">
          {mode === "up" && (
            <>
              <label className="authField">
                <span>שם מלא</span>
                <input value={form.name} onChange={(e) => set("name", e.target.value)} placeholder="ישראל ישראלי" />
              </label>
              <label className="authField">
                <span>שם החברה</span>
                <input value={form.company} onChange={(e) => set("company", e.target.value)} placeholder="יופיע בהצעות המחיר" />
              </label>
            </>
          )}

          <label className="authField">
            <span>מייל</span>
            <input
              type="email" dir="ltr" autoComplete="username"
              value={form.email}
              onChange={(e) => set("email", e.target.value)}
              placeholder="you@example.com"
            />
          </label>

          <label className="authField">
            <span>סיסמה</span>
            <input
              type="password" dir="ltr"
              autoComplete={mode === "in" ? "current-password" : "new-password"}
              value={form.password}
              onChange={(e) => set("password", e.target.value)}
              onKeyDown={(e) => e.key === "Enter" && submit()}
              placeholder={mode === "up" ? "8 תווים לפחות" : ""}
            />
          </label>

          {err && <p className="authErr">{err}</p>}
          {info && <p className="authInfo">{info}</p>}

          <button className="authBtn" onClick={submit} disabled={busy}>
            {busy ? "רגע…" : mode === "in" ? "כניסה" : "פתח חשבון"}
          </button>

          <p className="authNote">
            הנתונים נשמרים בענן. אפשר להיכנס מכל מחשב, והצוות רואה את
            אותם פרויקטים.
          </p>
        </div>
      </div>
    </div>
  );
}

/* ---------- ניהול צוות ---------- */

function TeamDialog({ user, onClose }) {
  const [members, setMembers] = useState([]);
  const [editing, setEditing] = useState(null);
  const [err, setErr] = useState("");
  const [busy, setBusy] = useState(true);

  const load = useCallback(async () => {
    setBusy(true);
    try {
      const rows = await teamApi.list();
      setMembers(rows.filter((m) => m.active).map((m) => toUser(m)));
    } catch (ex) {
      setErr(ex.message);
    }
    setBusy(false);
  }, []);

  useEffect(() => { load(); }, [load]);

  const remove = async (m) => {
    if (m.id === user.id) return;
    if (!confirm(`להסיר את ${m.name}?`)) return;
    try {
      await teamApi.deactivate(m.id);
      load();
    } catch (ex) {
      setErr(ex.message);
    }
  };

  const savePerms = async (id, fields) => {
    try {
      await teamApi.update(id, {
        role: ROLE_DB[fields.role],
        perms: fields.perms,
      });
      setEditing(null);
      load();
    } catch (ex) {
      setErr(ex.message);
    }
  };

  return (
    <div className="scrim" onClick={onClose}>
      <div className="drawer wide" onClick={(e) => e.stopPropagation()}>
        <div className="drawerHead">
          <h2>{editing ? `הרשאות · ${editing.name}` : "משתמשים והרשאות"}</h2>
          <button onClick={() => (editing ? setEditing(null) : onClose())}>×</button>
        </div>

        {err && <p className="authErr" style={{ margin: "14px 18px" }}>{err}</p>}

        {editing ? (
          <PermsEditor
            member={editing}
            isSelf={editing.id === user.id}
            onSave={(fields) => savePerms(editing.id, fields)}
            onCancel={() => setEditing(null)}
          />
        ) : busy ? (
          <p className="empty">טוען…</p>
        ) : (
          <>
            <table className="grid">
              <thead>
                <tr>
                  <th>שם</th>
                  <th className="gsev">תפקיד</th>
                  <th className="gnum">מודולים</th>
                  <th />
                </tr>
              </thead>
              <tbody>
                {members.map((m) => {
                  const pm = permsFor(m);
                  const open = MODULES.filter((x) => canSee(pm, x.key)).length;
                  return (
                    <tr key={m.id}>
                      <td>{m.name}{m.id === user.id && <em className="youTag">אתה</em>}</td>
                      <td className="gsev"><span className="pill">{m.role}</span></td>
                      <td className="gnum">{open} / {MODULES.length}</td>
                      <td className="cX rowTools">
                        <button onClick={() => setEditing(m)}>הרשאות</button>
                        {m.id !== user.id && <button className="del" onClick={() => remove(m)}>×</button>}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>

            <div className="teamAdd">
              <h3>הוספת משתמש</h3>
              <p className="teamNote">
                משתמש חדש נרשם בעצמו במסך הכניסה, ואז מופיע כאן ואפשר
                לקבוע לו תפקיד והרשאות. כך הסיסמה שלו נשארת רק אצלו.
              </p>
            </div>
          </>
        )}
      </div>
    </div>
  );
}

/* ---------- עורך ההרשאות ---------- */

function PermsEditor({ member, isSelf, onSave, onCancel }) {
  const start = permsFor(member);
  const [role, setRole] = useState(member.role);
  const [modules, setModules] = useState(start.modules);
  const [flg, setFlg] = useState(start.flags);

  const applyRole = (r) => {
    setRole(r);
    const preset = ROLE_PRESETS[r];
    setModules({ ...preset.modules });
    setFlg({ ...preset.flags });
  };

  const dirty =
    JSON.stringify(ROLE_PRESETS[role]?.modules) !== JSON.stringify(modules) ||
    JSON.stringify(ROLE_PRESETS[role]?.flags) !== JSON.stringify(flg);

  return (
    <div className="permsBody">
      <div className="permsTop">
        <label className="field">
          <span>תפקיד</span>
          <select value={role} onChange={(e) => applyRole(e.target.value)}>
            {ROLES.map((r) => <option key={r}>{r}</option>)}
          </select>
        </label>
        <p className="permsHint">
          בחירת תפקיד טוענת ברירות מחדל. כל שינוי מתחתיה נשמר למשתמש הזה בלבד.
          {dirty && <b> · מותאם אישית</b>}
        </p>
      </div>

      <table className="permsGrid">
        <thead>
          <tr>
            <th>מודול</th>
            {LEVELS.map((l) => <th key={l.key} className="plev">{l.label}</th>)}
          </tr>
        </thead>
        <tbody>
          {MODULES.map((m) => (
            <tr key={m.key}>
              <td>{m.label}</td>
              {LEVELS.map((l) => (
                <td key={l.key} className="plev">
                  <input
                    type="radio"
                    name={`mod-${m.key}`}
                    checked={modules[m.key] === l.key}
                    onChange={() => setModules({ ...modules, [m.key]: l.key })}
                  />
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>

      <div className="flagList">
        {FLAGS.map((f) => (
          <label key={f.key} className="flagRow">
            <input
              type="checkbox"
              checked={!!flg[f.key]}
              disabled={isSelf && f.key === "manageTeam"}
              onChange={(e) => setFlg({ ...flg, [f.key]: e.target.checked })}
            />
            <span>
              {f.label}
              {f.note && <i>{f.note}</i>}
            </span>
          </label>
        ))}
      </div>

      {isSelf && (
        <p className="teamNote">
          אינך יכול לשלול מעצמך את ניהול המשתמשים — אחרת אף אחד לא יוכל
          להחזיר הרשאות.
        </p>
      )}

      <div className="rowBtns permsFoot">
        <button className="primaryBtn" onClick={() => onSave({ role, perms: { modules, flags: flg } })}>
          שמור
        </button>
        <button className="toolBtn" onClick={onCancel}>ביטול</button>
      </div>
    </div>
  );
}

/* ============================================================
   רשימת פרויקטים
   ============================================================ */

function ProjectsList({ index, onOpen, onCreate, onDelete, onDuplicate }) {
  const [q, setQ] = useState("");
  const [picking, setPicking] = useState(false);

  const rows = useMemo(() => {
    const t = q.trim();
    if (!t) return index;
    return index.filter(
      (p) =>
        (p.name || "").includes(t) ||
        (p.client_name || "").includes(t) ||
        (p.address || "").includes(t)
    );
  }, [index, q]);

  return (
    <div className="page">
      <div className="listHead">
        <h2>הפרויקטים שלי</h2>
        <div className="listTools">
          <input
            className="searchBox"
            placeholder="חפש לפי שם, לקוח או כתובת…"
            value={q}
            onChange={(e) => setQ(e.target.value)}
          />
          <button className="primaryBtn" onClick={() => setPicking(true)}>
            פרויקט חדש
          </button>
        </div>
      </div>

      {index.length === 0 ? (
        <div className="blank">
          <h3>עדיין אין פרויקטים</h3>
          <p>
            צור פרויקט חדש, או ייבא כתב כמויות מאקסל שקיבלת מלקוח.
          </p>
          <button className="primaryBtn" onClick={() => setPicking(true)}>
            צור פרויקט ראשון
          </button>
        </div>
      ) : rows.length === 0 ? (
        <p className="empty">אין פרויקט שתואם לחיפוש.</p>
      ) : (
        <div className="cards">
          {rows.map((p) => (
            <article className="pcard" key={p.id}>
              <button className="pcardMain" onClick={() => onOpen(p.id)}>
                <span className={`ptag ${p.kind}`}>
                  {TEMPLATES[p.kind]?.label || "פרויקט"}
                </span>
                <h3>{p.name || "פרויקט ללא שם"}</h3>
                <p className="pmeta">
                  {[p.client_name, p.address].filter(Boolean).join(" · ") || "—"}
                </p>
                
                <p className="pdate">עודכן {(p.updated_at || "").slice(0, 10)}</p>
              </button>
              <div className="pcardTools">
                <button onClick={() => onDuplicate(p.id)} title="שכפל">
                  שכפל
                </button>
                {onDelete && (
                  <button className="del" onClick={() => onDelete(p.id, p.name)} title="מחק">
                    מחק
                  </button>
                )}
              </div>
            </article>
          ))}
        </div>
      )}

      {picking && (
        <div className="scrim" onClick={() => setPicking(false)}>
          <div className="drawer" onClick={(e) => e.stopPropagation()}>
            <div className="drawerHead">
              <h2>סוג הפרויקט</h2>
              <button onClick={() => setPicking(false)}>×</button>
            </div>
            <div className="kindGrid">
              {Object.entries(TEMPLATES).map(([key, tpl]) => (
                <button
                  key={key}
                  className="kindCard"
                  onClick={() => {
                    setPicking(false);
                    onCreate(key);
                  }}
                >
                  <b>{tpl.label}</b>
                  <i>{tpl.hint}</i>
                  <em>
                    {tpl.chapters.length
                      ? `${tpl.chapters.length} פרקים מוכנים`
                      : "בלי פרקים"}
                  </em>
                </button>
              ))}
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

/* ============================================================
   כתב כמויות
   ============================================================ */

function BoqTab({ project, patch, company, setCompany, perms }) {
  const ro = !canEdit(perms, "boq");
  const money = perms.flags.seePrices;
  const [picker, setPicker] = useState(null);
  const [query, setQuery] = useState("");
  const [imp, setImp] = useState(null);
  const [showQuote, setShowQuote] = useState(false);
  const fileInput = useRef(null);

  const chapters = project.chapters;
  const overhead = project.overhead || 0;

  const chapterTotals = useMemo(
    () =>
      chapters.map((c) =>
        c.items.reduce((s, it) => s + (+it.qty || 0) * (+it.price || 0), 0)
      ),
    [chapters]
  );
  const net = chapterTotals.reduce((a, b) => a + b, 0);
  const oh = net * (overhead / 100);
  const beforeVat = net + oh;
  const vat = beforeVat * VAT_RATE;
  const total = beforeVat + vat;
  const itemCount = chapters.reduce((s, c) => s + c.items.length, 0);
  const missingUnits = useMemo(
    () =>
      chapters.reduce(
        (s, c) => s + c.items.filter((it) => it.qty && !unitOk(it.unit)).length,
        0
      ),
    [chapters]
  );

  const setChapters = (fn) =>
    patch({ chapters: typeof fn === "function" ? fn(chapters) : fn });

  const patchChapter = (id, fn) =>
    setChapters((cs) => cs.map((c) => (c.id === id ? fn(c) : c)));

  const addItem = (chId, seed = {}) =>
    patchChapter(chId, (c) => ({
      ...c,
      items: [
        ...c.items,
        {
          id: uid(),
          code: "",
          desc: seed.name || "",
          unit: seed.unit || "יח'",
          qty: seed.qty ?? 1,
          price: seed.price ?? 0,
        },
      ],
    }));

  const patchItem = (chId, itId, field, value) =>
    patchChapter(chId, (c) => ({
      ...c,
      items: c.items.map((it) => (it.id === itId ? { ...it, [field]: value } : it)),
    }));

  const removeItem = (chId, itId) =>
    patchChapter(chId, (c) => ({
      ...c,
      items: c.items.filter((it) => it.id !== itId),
    }));

  const moveChapter = (i, dir) =>
    setChapters((cs) => {
      const j = i + dir;
      if (j < 0 || j >= cs.length) return cs;
      const copy = [...cs];
      [copy[i], copy[j]] = [copy[j], copy[i]];
      return copy;
    });

  const switchKind = (kind) => {
    if (kind === project.kind) return;
    const filled = chapters.filter((c) => c.items.length);
    if (
      filled.length &&
      !confirm(
        `להחליף ל"${TEMPLATES[kind].label}"?\n\nפרקים שכבר יש בהם סעיפים יישמרו.`
      )
    )
      return;
    const fresh = makeChapters(kind).filter(
      (n) => !filled.some((f) => f.title === n.title)
    );
    patch({ kind, chapters: [...filled, ...fresh] });
  };

  /* ---------- import ---------- */
  const readFile = async (file) => {
    if (!file) return;
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      const sheets = wb.SheetNames.map((name) => ({
        name,
        rows: XLSX.utils.sheet_to_json(wb.Sheets[name], {
          header: 1,
          blankrows: false,
          defval: "",
        }),
      })).filter((s) => s.rows.length);
      if (!sheets.length) {
        setImp({ error: "הקובץ ריק או לא נקרא." });
        return;
      }
      selectSheet(sheets, 0, file.name);
    } catch {
      setImp({ error: "לא הצלחתי לקרוא את הקובץ. נסה xlsx או csv." });
    }
  };

  const selectSheet = (sheets, idx, fileName) => {
    const rows = sheets[idx].rows;
    const headerRow = findHeaderRow(rows);
    setImp({ fileName, sheets, idx, rows, headerRow, map: autoMap(rows[headerRow] || []) });
  };

  const impRecords = useMemo(() => {
    if (!imp || imp.error || !imp.rows) return [];
    const { rows, headerRow, map } = imp;
    const at = (row, key) => (map[key] === undefined ? "" : row[map[key]]);
    return rows
      .slice(headerRow + 1)
      .map((row) => {
        const rec = {
          code: String(at(row, "code") ?? "").trim(),
          desc: String(at(row, "desc") ?? "").trim(),
          unit: String(at(row, "unit") ?? "").trim(),
          qty: toNumber(at(row, "qty")),
          price: toNumber(at(row, "price")),
        };
        const fixed = rescueMergedTitle(rec);
        return {
          ...fixed,
          depth: codeDepth(fixed.code),
          isChapter: looksLikeChapter(fixed),
          unitBad: !!(fixed.qty && !unitOk(fixed.unit)),
        };
      })
      .filter((r) => r.desc || r.qty);
  }, [imp]);

  /* קובץ עם קודים היררכיים (01 / 01.20 / 01.20.3100) מאפשר
     לבחור באיזו רמה לפתוח פרקים. */
  const hier = useMemo(() => {
    const depths = impRecords.map((r) => r.depth);
    const maxDepth = Math.max(0, ...depths);
    if (maxDepth < 3) return null;
    return {
      maxDepth,
      main: impRecords.filter((r) => r.depth === 1 && !r.qty).length,
      sub: impRecords.filter((r) => r.depth === 2 && !r.qty).length,
      items: impRecords.filter((r) => r.depth >= 3 || r.qty).length,
    };
  }, [impRecords]);

  const level = imp?.level ?? 1;

  const impStats = useMemo(() => {
    if (hier) {
      const items = impRecords.filter((r) => r.qty || r.depth >= 3);
      return {
        chapters: impRecords.filter((r) => r.depth === level && !r.qty).length,
        items: items.length,
        priced: items.filter((r) => r.price > 0).length,
        unitBad: items.filter((r) => r.unitBad).length,
      };
    }
    const ch = impRecords.filter((r) => r.isChapter).length;
    return {
      chapters: ch,
      items: impRecords.length - ch,
      priced: impRecords.filter((r) => !r.isChapter && r.price > 0).length,
      unitBad: impRecords.filter((r) => r.unitBad).length,
    };
  }, [impRecords, hier, level]);

  const applyImport = (mode) => {
    const built = [];
    let current = null;

    const push = (r) => {
      if (!current) {
        current = { id: uid(), title: "כללי", items: [] };
        built.push(current);
      }
      current.items.push({
        id: uid(),
        code: r.code || "",
        desc: r.desc || "—",
        unit: unitOk(r.unit) ? r.unit : "",
        qty: r.qty || 0,
        price: r.price || 0,
      });
    };

    impRecords.forEach((r) => {
      if (hier) {
        const isItem = r.qty || r.depth >= 3;
        if (!isItem && r.depth === level) {
          current = {
            id: uid(),
            title: [r.code, r.desc].filter(Boolean).join(" "),
            items: [],
          };
          built.push(current);
          return;
        }
        if (!isItem) return; /* רמת ביניים — מדלגים, הקוד בסעיף שומר אותה */
        push(r);
        return;
      }

      if (r.isChapter) {
        current = { id: uid(), title: r.desc, items: [] };
        built.push(current);
        return;
      }
      push(r);
    });
    const clean = built.filter((c) => c.items.length);
    if (!clean.length) {
      alert("לא נמצאו סעיפים. בדוק את התאמת העמודות.");
      return;
    }
    setChapters((cs) =>
      mode === "replace" ? clean : [...cs.filter((c) => c.items.length), ...clean]
    );
    setImp(null);
    if (fileInput.current) fileInput.current.value = "";
  };

  const exportCsv = () => {
    const rows = [["פרק", "סעיף", "תיאור", "יחידה", "כמות", "מחיר יחידה", 'סה"כ']];
    chapters.forEach((c, i) =>
      c.items.forEach((it) =>
        rows.push([
          `${chapterNo(i)} ${c.title}`,
          it.code || "",
          it.desc,
          it.unit,
          it.qty,
          it.price,
          (+it.qty || 0) * (+it.price || 0),
        ])
      )
    );
    rows.push([], ['סה"כ לפני מע"מ', "", "", "", "", "", beforeVat.toFixed(2)]);
    rows.push([`מע"מ ${VAT_RATE * 100}%`, "", "", "", "", "", vat.toFixed(2)]);
    rows.push(['סה"כ כולל מע"מ', "", "", "", "", "", total.toFixed(2)]);
    downloadCsv(rows, `${project.name || "כתב_כמויות"}.csv`);
  };

  const filtered = useMemo(() => {
    const q = query.trim();
    return q ? CATALOG.filter((c) => c.name.includes(q) || c.cat.includes(q)) : CATALOG;
  }, [query]);

  return (
    <div className="layout">
      {money && (
      <aside className="tape">
        <div className="tapeHead">
          <span>סיכום</span>
          <em>{itemCount} סעיפים</em>
        </div>
        <div className="tapeBody">
          {chapters.map((c, i) => (
            <div className="tapeRow" key={c.id}>
              <span className="tno">{chapterNo(i)}</span>
              <span className="tname">{c.title}</span>
              <span className="tval">{shekel(chapterTotals[i])}</span>
            </div>
          ))}
        </div>
        <div className="tapeFoot">
          <div className="tapeRow sum">
            <span className="tname">סכום עבודות</span>
            <span className="tval">{shekel(net)}</span>
          </div>
          <label className="ohRow">
            <span>רווח וקבלנות</span>
            <span className="ohInput">
              <input
                type="number"
                min="0"
                max="100"
                value={overhead}
                onChange={(e) => patch({ overhead: +e.target.value || 0 })}
              />
              <i>%</i>
            </span>
            <span className="tval">{shekel(oh)}</span>
          </label>
          <div className="tapeRow">
            <span className="tname">לפני מע"מ</span>
            <span className="tval">{shekel(beforeVat)}</span>
          </div>
          <div className="tapeRow">
            <span className="tname">מע"מ {VAT_RATE * 100}%</span>
            <span className="tval">{shekel(vat)}</span>
          </div>
          <div className="grand">
            <span>סה"כ לתשלום</span>
            <strong>{shekel(total)}</strong>
          </div>
        </div>
      </aside>
      )}

      <main className="sheet">
        {missingUnits > 0 && (
          <p className="warnBox">
            {missingUnits} סעיפים ללא יחידת מדידה. השלם אותם לפני תמחור —
            אותו מחיר במ"א ובטון נותן סכומים שונים לחלוטין.
          </p>
        )}

        <section className="toolRow">
          <input
            ref={fileInput}
            type="file"
            accept=".xlsx,.xls,.csv"
            style={{ display: "none" }}
            onChange={(e) => readFile(e.target.files?.[0])}
          />
          {!ro && (
            <button className="toolBtn accent" onClick={() => fileInput.current?.click()}>
              ייבוא מאקסל
            </button>
          )}
          {perms.flags.exportData && (
            <button className="toolBtn" onClick={exportCsv}>
              ייצוא CSV
            </button>
          )}
          {perms.flags.issueDocs && money && (
            <button className="toolBtn" onClick={() => setShowQuote(true)}>
              הצעת מחיר
            </button>
          )}
          <span className="spacer" />
          <div className="kindPicker">
            {Object.entries(TEMPLATES).map(([key, tpl]) => (
              <button
                key={key}
                className={`kindBtn ${project.kind === key ? "on" : ""}`}
                onClick={() => switchKind(key)}
              >
                {tpl.label}
              </button>
            ))}
          </div>
        </section>

        <section className="projectCard">
          <Field label="שם הפרויקט" value={project.name} onChange={(v) => patch({ name: v })} placeholder="תוספת בנייה, הרצל 12" />
          <Field label="לקוח" value={project.client_name} onChange={(v) => patch({ client_name: v })} placeholder="שם הלקוח" />
          <Field label="כתובת" value={project.address} onChange={(v) => patch({ address: v })} placeholder="עיר ורחוב" />
          <Field label="תאריך התחלה" type="date" value={project.start_date} onChange={(v) => patch({ start_date: v })} />
          <Field label="תאריך סיום" type="date" value={project.end_date} onChange={(v) => patch({ end_date: v })} />
          <Field label="מנהל עבודה" value={project.site_manager_name} onChange={(v) => patch({ site_manager_name: v })} placeholder="שם" />
          <Field label="טלפון מנהל עבודה" value={project.site_manager_phone} onChange={(v) => patch({ site_manager_phone: v })} placeholder="050-0000000" />
          <Field label="תיאור" value={project.description} onChange={(v) => patch({ description: v })} placeholder="הערות על הפרויקט" />
        </section>

        {chapters.map((c, i) => (
          <section className="chapter" key={c.id}>
            <div className="chHead">
              <span className="chNo">{chapterNo(i)}</span>
              <input
                className="chTitle"
                value={c.title}
                onChange={(e) => patchChapter(c.id, (x) => ({ ...x, title: e.target.value }))}
              />
              {money && <span className="chSum">{shekel(chapterTotals[i])}</span>}
              <span className="chTools">
                <button onClick={() => moveChapter(i, -1)} disabled={i === 0}>↑</button>
                <button onClick={() => moveChapter(i, 1)} disabled={i === chapters.length - 1}>↓</button>
                <button className="x" onClick={() => setChapters((cs) => cs.filter((x) => x.id !== c.id))}>×</button>
              </span>
            </div>

            {c.items.length > 0 && (
              <table className="items">
                <thead>
                  <tr>
                    <th className="cDesc">תיאור</th>
                    <th className="cUnit">יחידה</th>
                    <th className="cNum">כמות</th>
                    {money && <th className="cNum">מחיר יח'</th>}
                    {money && <th className="cNum">סה"כ</th>}
                    {!ro && <th />}
                  </tr>
                </thead>
                <tbody>
                  {c.items.map((it, k) => (
                    <tr key={it.id}>
                      <td className="cDesc">
                        <span className={`itNo ${it.code ? "code" : ""}`}>
                          {it.code || `${chapterNo(i)}.${k + 1}`}
                        </span>
                        <input value={it.desc} readOnly={ro} placeholder="תיאור העבודה" onChange={(e) => patchItem(c.id, it.id, "desc", e.target.value)} />
                      </td>
                      <td className="cUnit">
                        <input
                          list="unitList"
                          className={!unitOk(it.unit) ? "needsUnit" : ""}
                          placeholder="יחידה?"
                          title={!unitOk(it.unit) ? "היחידה חסרה — בדוק לפני תמחור" : ""}
                          value={it.unit}
                          onChange={(e) => patchItem(c.id, it.id, "unit", e.target.value)}
                        />
                      </td>
                      <td className="cNum">
                        <input type="number" step="any" readOnly={ro} value={it.qty} onChange={(e) => patchItem(c.id, it.id, "qty", e.target.value)} />
                      </td>
                      {money && (
                        <td className="cNum">
                          <input type="number" step="any" readOnly={ro} value={it.price} onChange={(e) => patchItem(c.id, it.id, "price", e.target.value)} />
                        </td>
                      )}
                      {money && (
                        <td className="cNum lineTotal">{num((+it.qty || 0) * (+it.price || 0), 0)}</td>
                      )}
                      {!ro && (
                        <td className="cX">
                          <button onClick={() => removeItem(c.id, it.id)}>×</button>
                        </td>
                      )}
                    </tr>
                  ))}
                </tbody>
              </table>
            )}

            {!ro && (
              <div className="chFoot">
                <button className="add" onClick={() => addItem(c.id)}>+ סעיף ריק</button>
                <button className="add primary" onClick={() => setPicker(c.id)}>+ מהמחירון</button>
              </div>
            )}
          </section>
        ))}

        <datalist id="unitList">
          {UNITS.map((u) => <option key={u} value={u} />)}
        </datalist>

        {!ro && (
          <button className="addChapter" onClick={() => setChapters((cs) => [...cs, { id: uid(), title: "פרק חדש", items: [] }])}>
            + הוסף פרק
          </button>
        )}
      </main>

      {picker && (
        <div className="scrim" onClick={() => setPicker(null)}>
          <div className="drawer" onClick={(e) => e.stopPropagation()}>
            <div className="drawerHead">
              <h2>מחירון</h2>
              <button onClick={() => setPicker(null)}>×</button>
            </div>
            <input className="search" autoFocus placeholder="חפש עבודה או קטגוריה…" value={query} onChange={(e) => setQuery(e.target.value)} />
            <div className="catList">
              {filtered.map((row, idx) => (
                <button key={idx} className="catRow" onClick={() => { addItem(picker, row); setPicker(null); setQuery(""); }}>
                  <span className="catName">{row.name}</span>
                  <span className="catTag">{row.cat}</span>
                  <span className="catPrice">{shekel(row.price)} / {row.unit}</span>
                </button>
              ))}
              {filtered.length === 0 && <p className="empty">אין התאמה. אפשר להוסיף סעיף ריק ולכתוב ידנית.</p>}
            </div>
          </div>
        </div>
      )}

      {imp && (
        <ImportDialog
          imp={imp}
          setImp={setImp}
          records={impRecords}
          stats={impStats}
          hier={hier}
          level={level}
          selectSheet={selectSheet}
          onApply={applyImport}
        />
      )}

      {showQuote && (
        <QuoteDialog
          project={project}
          chapters={chapters}
          chapterTotals={chapterTotals}
          totals={{ net, oh, beforeVat, vat, total, overhead }}
          company={company}
          setCompany={setCompany}
          onClose={() => setShowQuote(false)}
        />
      )}
    </div>
  );
}


/* ============================================================
   ביצוע וחשבון חלקי
   ============================================================ */

function ExecutionTab({ project, patch, company, perms }) {
  const ro = !canEdit(perms, "execution");
  const money = perms.flags.seePrices;
  const [showBill, setShowBill] = useState(false);
  const [onlyOpen, setOnlyOpen] = useState(false);
  const chapters = project.chapters;
  const ex = useMemo(() => executionTotals(chapters), [chapters]);

  const setDone = (chId, itId, value) =>
    patch({
      chapters: chapters.map((c) =>
        c.id !== chId
          ? c
          : {
              ...c,
              items: c.items.map((it) =>
                it.id === itId ? { ...it, done: value } : it
              ),
            }
      ),
    });

  const fillChapter = (chId, mode) =>
    patch({
      chapters: chapters.map((c) =>
        c.id !== chId
          ? c
          : {
              ...c,
              items: c.items.map((it) => ({
                ...it,
                done: mode === "all" ? it.qty : 0,
              })),
            }
      ),
    });

  const overCount = chapters.reduce(
    (s, c) => s + c.items.filter(itemOver).length,
    0
  );

  const exportCsv = () => {
    const rows = [
      ["פרק", "סעיף", "תיאור", "יחידה", "כמות בחוזה", "בוצע", "נותר", "מחיר יח'", "סכום שבוצע"],
    ];
    chapters.forEach((c, i) =>
      c.items.forEach((it) =>
        rows.push([
          `${chapterNo(i)} ${c.title}`,
          it.code || "",
          it.desc,
          it.unit,
          it.qty,
          it.done || 0,
          (+it.qty || 0) - (+it.done || 0),
          it.price,
          itemDone(it).toFixed(2),
        ])
      )
    );
    rows.push([], ["סה\"כ חוזה", "", "", "", "", "", "", "", ex.contract.toFixed(2)]);
    rows.push(["סה\"כ בוצע", "", "", "", "", "", "", "", ex.done.toFixed(2)]);
    downloadCsv(rows, `${project.name || "ביצוע"}_ביצוע.csv`);
  };

  return (
    <div className="page">
      <h2 className="pageTitle">מעקב ביצוע</h2>

      <div className="statRow">
        {money && <Stat label="שווי החוזה" value={shekel(ex.contract)} />}
        {money && <Stat label="בוצע עד כה" value={shekel(ex.done)} tone="good" />}
        {money && <Stat label="נותר לביצוע" value={shekel(ex.remaining)} />}
        <Stat label="התקדמות" value={pct(ex.percent)} />
        <Stat label="סעיפים שהושלמו" value={`${ex.complete} / ${ex.items}`} />
      </div>

      <div className="progressWide">
        <i style={{ width: `${Math.min(ex.percent, 100)}%` }} />
        <em>
          {ex.complete} מתוך {ex.items} סעיפים הושלמו · {ex.started} התחילו
        </em>
      </div>

      {overCount > 0 && (
        <p className="warnBox">
          {overCount} סעיפים חורגים מהכמות שבחוזה. חריגה דורשת אישור בכתב מהלקוח
          לפני שהיא נכנסת לחשבון.
        </p>
      )}

      <div className="toolRow">
        {perms.flags.issueDocs && money && (
          <button className="toolBtn accent" onClick={() => setShowBill(true)}>
            חשבון חלקי
          </button>
        )}
        {perms.flags.exportData && (
          <button className="toolBtn" onClick={exportCsv}>
            ייצוא CSV
          </button>
        )}
        <span className="spacer" />
        <label className="checkRow">
          <input
            type="checkbox"
            checked={onlyOpen}
            onChange={(e) => setOnlyOpen(e.target.checked)}
          />
          הצג רק סעיפים שלא הושלמו
        </label>
      </div>

      {chapters.map((c, i) => {
        const items = onlyOpen
          ? c.items.filter((it) => (+it.done || 0) < (+it.qty || 0))
          : c.items;
        if (!c.items.length) return null;

        const cContract = c.items.reduce((s, it) => s + itemContract(it), 0);
        const cDone = c.items.reduce((s, it) => s + itemDone(it), 0);
        const cPct = cContract > 0 ? (cDone / cContract) * 100 : 0;

        return (
          <section className="chapter" key={c.id}>
            <div className="chHead exec">
              <span className="chNo">{chapterNo(i)}</span>
              <span className="chTitleText">{c.title}</span>
              <span className="chProgress">
                <i style={{ width: `${Math.min(cPct, 100)}%` }} />
                <em>{pct(cPct)}</em>
              </span>
              {money && (
                <span className="chSum">
                  {shekel(cDone)} / {shekel(cContract)}
                </span>
              )}
              {!ro && (
                <span className="chTools">
                  <button onClick={() => fillChapter(c.id, "all")} title="סמן הכל כבוצע">הכל</button>
                  <button onClick={() => fillChapter(c.id, "none")} title="אפס">אפס</button>
                </span>
              )}
            </div>

            {items.length > 0 ? (
              <table className="items exec">
                <thead>
                  <tr>
                    <th className="cDesc">תיאור</th>
                    <th className="cUnit">יחידה</th>
                    <th className="cNum">בחוזה</th>
                    <th className="cNum">בוצע</th>
                    <th className="cNum">נותר</th>
                    <th className="cBar">התקדמות</th>
                    {money && <th className="cNum">סכום</th>}
                  </tr>
                </thead>
                <tbody>
                  {items.map((it, k) => {
                    const q = +it.qty || 0;
                    const d = +it.done || 0;
                    const ip = q > 0 ? (d / q) * 100 : 0;
                    const over = itemOver(it);
                    return (
                      <tr key={it.id} className={over ? "over" : ""}>
                        <td className="cDesc">
                          <span className={`itNo ${it.code ? "code" : ""}`}>
                            {it.code || `${chapterNo(i)}.${k + 1}`}
                          </span>
                          <span className="descText">{it.desc || "—"}</span>
                        </td>
                        <td className="cUnit muted">{it.unit || "—"}</td>
                        <td className="cNum muted">{num(q)}</td>
                        <td className="cNum">
                          <input
                            type="number"
                            step="any"
                            min="0"
                            value={it.done ?? ""}
                            placeholder="0"
                            readOnly={ro}
                            onChange={(e) => setDone(c.id, it.id, e.target.value)}
                          />
                        </td>
                        <td className="cNum muted">{num(q - d)}</td>
                        <td className="cBar">
                          <span className="miniBar">
                            <i
                              className={over ? "over" : ""}
                              style={{ width: `${Math.min(ip, 100)}%` }}
                            />
                          </span>
                        </td>
                        {money && <td className="cNum lineTotal">{num(itemDone(it), 0)}</td>}
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            ) : (
              <p className="allDone">כל הסעיפים בפרק הושלמו.</p>
            )}
          </section>
        );
      })}

      {showBill && (
        <BillDialog
          project={project}
          chapters={chapters}
          ex={ex}
          company={company}
          onClose={() => setShowBill(false)}
        />
      )}
    </div>
  );
}

/* ---------- חשבון חלקי ---------- */

function BillDialog({ project, chapters, ex, company, onClose }) {
  const [prev, setPrev] = useState(0);
  const vat = ex.done * VAT_RATE;
  const thisBill = ex.done - (+prev || 0);

  return (
    <div className="scrim" onClick={onClose}>
      <div className="quoteWrap" onClick={(e) => e.stopPropagation()}>
        <div className="quoteBar">
          <label className="barField">
            <span>שולם עד כה</span>
            <input
              type="number"
              step="any"
              value={prev}
              onChange={(e) => setPrev(e.target.value)}
            />
          </label>
          <button onClick={() => window.print()}>הדפס / PDF</button>
          <button onClick={onClose}>סגור</button>
        </div>

        <div className="quote">
          <header className="qHead">
            <div>
              <h2>{company.name || "———"}</h2>
              <p>
                {[company.phone, company.vat && `ע.מ ${company.vat}`]
                  .filter(Boolean)
                  .join(" · ")}
              </p>
            </div>
            <div className="qMeta">
              <strong>חשבון חלקי</strong>
              <span>{new Date().toLocaleDateString("he-IL")}</span>
            </div>
          </header>

          <div className="qProject">
            <p><b>פרויקט:</b> {project.name || "—"}</p>
            <p><b>לקוח:</b> {project.client_name || "—"}</p>
            <p><b>כתובת:</b> {project.address || "—"}</p>
          </div>

          {chapters.map((c, i) => {
            const done = c.items.filter((it) => (+it.done || 0) > 0);
            if (!done.length) return null;
            const cDone = done.reduce((s, it) => s + itemDone(it), 0);
            return (
              <div className="qChapter" key={c.id}>
                <h3><span>{chapterNo(i)}</span> {c.title}</h3>
                <table>
                  <tbody>
                    {done.map((it, k) => (
                      <tr key={it.id}>
                        <td className="qn">{it.code || `${chapterNo(i)}.${k + 1}`}</td>
                        <td>{it.desc || "—"}</td>
                        <td className="qu">{it.unit}</td>
                        <td className="qnum">{num(+it.done || 0)}</td>
                        <td className="qnum">{num(+it.price || 0, 0)}</td>
                        <td className="qnum b">{num(itemDone(it), 0)}</td>
                      </tr>
                    ))}
                    <tr className="qSub">
                      <td colSpan={5}>סה"כ פרק {chapterNo(i)}</td>
                      <td className="qnum">{num(cDone, 0)}</td>
                    </tr>
                  </tbody>
                </table>
              </div>
            );
          })}

          <div className="qTotals">
            <div><span>שווי החוזה</span><b>{shekel(ex.contract)}</b></div>
            <div><span>בוצע עד כה ({pct(ex.percent)})</span><b>{shekel(ex.done)}</b></div>
            {+prev > 0 && (
              <>
                <div><span>שולם בחשבונות קודמים</span><b>{shekel(+prev)}</b></div>
                <div><span>לתשלום בחשבון זה</span><b>{shekel(thisBill)}</b></div>
              </>
            )}
            <div><span>מע"מ {VAT_RATE * 100}%</span><b>{shekel((+prev > 0 ? thisBill : ex.done) * VAT_RATE)}</b></div>
            <div className="qGrand">
              <span>סה"כ לתשלום</span>
              <b>{shekel((+prev > 0 ? thisBill : ex.done) * (1 + VAT_RATE))}</b>
            </div>
          </div>

          <p className="qNote">
            החשבון מבוסס על מדידת הכמויות שבוצעו בפועל נכון לתאריך שלעיל.
            סעיפים החורגים מהכמות שבחוזה טעונים אישור בכתב.
          </p>
        </div>
      </div>
    </div>
  );
}


/* ============================================================
   שינויים ותוספות
   ============================================================ */

function ChangesTab({ project, patch, company, perms }) {
  const ro = !canEdit(perms, "changes");
  const mayApprove = perms.flags.approveChanges;
  const changes = project.changes || [];
  const [openId, setOpenId] = useState(null);
  const [printId, setPrintId] = useState(null);
  const sum = changeSummary(changes);
  const contract = boqTotal(project.chapters);

  const setChanges = (fn) =>
    patch({ changes: typeof fn === "function" ? fn(changes) : fn });

  const add = () => {
    const next = newChangeOrder(changes.length + 1);
    setChanges((l) => [...l, next]);
    setOpenId(next.id);
  };

  const upd = (id, field, value) =>
    setChanges((l) => l.map((c) => (c.id === id ? { ...c, [field]: value } : c)));

  const updItem = (coId, itId, field, value) =>
    setChanges((l) =>
      l.map((c) =>
        c.id !== coId
          ? c
          : {
              ...c,
              items: c.items.map((it) =>
                it.id === itId ? { ...it, [field]: value } : it
              ),
            }
      )
    );

  const addItem = (coId) =>
    setChanges((l) =>
      l.map((c) =>
        c.id !== coId
          ? c
          : {
              ...c,
              items: [
                ...c.items,
                { id: uid(), desc: "", unit: "יח'", qty: 1, price: 0 },
              ],
            }
      )
    );

  const removeItem = (coId, itId) =>
    setChanges((l) =>
      l.map((c) =>
        c.id !== coId ? c : { ...c, items: c.items.filter((i) => i.id !== itId) }
      )
    );

  const remove = (id, title) => {
    if (!confirm(`למחוק את "${title || "השינוי"}"?`)) return;
    setChanges((l) => l.filter((c) => c.id !== id));
  };

  const printed = changes.find((c) => c.id === printId);

  return (
    <div className="page">
      <h2 className="pageTitle">שינויים ותוספות</h2>

      <div className="statRow">
        <Stat label="חוזה מקורי" value={shekel(contract)} />
        <Stat
          label={`אושרו (${sum.approvedCount})`}
          value={shekel(sum.approved)}
          tone="good"
        />
        <Stat
          label={`ממתינים (${sum.pendingCount})`}
          value={shekel(sum.pending)}
          tone={sum.pendingCount ? "warn" : ""}
        />
        <Stat label="חוזה מעודכן" value={shekel(contract + sum.approved)} />
      </div>

      {sum.pendingCount > 0 && (
        <p className="warnBox">
          {sum.pendingCount} שינויים בסך {shekel(sum.pending)} הוגשו ועדיין לא
          אושרו. עבודה שמתבצעת לפני אישור בכתב היא על אחריותך.
        </p>
      )}

      {changes.length === 0 ? (
        <div className="blank small">
          <h3>אין שינויים</h3>
          <p>
            כל בקשה של הלקוח שאינה בחוזה המקורי נרשמת כאן בנפרד, עם סטטוס
            אישור. כך תמיד ברור מה היה בחוזה ומה נוסף.
          </p>
          <button className="primaryBtn" onClick={add}>
            הוסף שינוי ראשון
          </button>
        </div>
      ) : (
        <>
          {changes.map((co) => {
            const open = openId === co.id;
            const tot = coTotal(co);
            return (
              <section className={`chapter co ${coClass(co.status)}`} key={co.id}>
                <div className="coHead">
                  <button
                    className="coToggle"
                    onClick={() => setOpenId(open ? null : co.id)}
                  >
                    {open ? "−" : "+"}
                  </button>
                  <span className="coNo">ש{co.number}</span>
                  <input
                    className="chTitle"
                    value={co.title}
                    placeholder="נושא השינוי"
                    onChange={(e) => upd(co.id, "title", e.target.value)}
                  />
                  <select
                    className={`coStatus ${coClass(co.status)}`}
                    value={co.status}
                    disabled={ro}
                    onChange={(e) => upd(co.id, "status", e.target.value)}
                  >
                    {CO_STATUSES.filter(
                      (s) =>
                        mayApprove ||
                        s === CO_DRAFT ||
                        s === CO_SUBMITTED ||
                        s === co.status
                    ).map((s) => (
                      <option key={s}>{s}</option>
                    ))}
                  </select>
                  <span className="chSum">{shekel(tot)}</span>
                  <span className="chTools">
                    <button onClick={() => setPrintId(co.id)} title="מסמך">
                      מסמך
                    </button>
                    <button className="x" onClick={() => remove(co.id, co.title)}>
                      ×
                    </button>
                  </span>
                </div>

                {open && (
                  <>
                    <div className="coMeta">
                      <Field
                        label="מי ביקש"
                        value={co.requested_by}
                        onChange={(v) => upd(co.id, "requested_by", v)}
                        placeholder="הלקוח / האדריכל"
                      />
                      <Field
                        label="תאריך"
                        type="date"
                        value={co.date}
                        onChange={(v) => upd(co.id, "date", v)}
                      />
                      <Field
                        label="סיבה"
                        value={co.reason}
                        onChange={(v) => upd(co.id, "reason", v)}
                        placeholder="למה נדרש השינוי"
                      />
                    </div>

                    {co.items.length > 0 && (
                      <table className="items">
                        <thead>
                          <tr>
                            <th className="cDesc">תיאור</th>
                            <th className="cUnit">יחידה</th>
                            <th className="cNum">כמות</th>
                            <th className="cNum">מחיר יח'</th>
                            <th className="cNum">סה"כ</th>
                            <th />
                          </tr>
                        </thead>
                        <tbody>
                          {co.items.map((it, k) => (
                            <tr key={it.id}>
                              <td className="cDesc">
                                <span className="itNo">
                                  ש{co.number}.{k + 1}
                                </span>
                                <input
                                  value={it.desc}
                                  placeholder="תיאור העבודה"
                                  onChange={(e) =>
                                    updItem(co.id, it.id, "desc", e.target.value)
                                  }
                                />
                              </td>
                              <td className="cUnit">
                                <input
                                  list="unitList"
                                  value={it.unit}
                                  onChange={(e) =>
                                    updItem(co.id, it.id, "unit", e.target.value)
                                  }
                                />
                              </td>
                              <td className="cNum">
                                <input
                                  type="number"
                                  step="any"
                                  value={it.qty}
                                  onChange={(e) =>
                                    updItem(co.id, it.id, "qty", e.target.value)
                                  }
                                />
                              </td>
                              <td className="cNum">
                                <input
                                  type="number"
                                  step="any"
                                  value={it.price}
                                  onChange={(e) =>
                                    updItem(co.id, it.id, "price", e.target.value)
                                  }
                                />
                              </td>
                              <td className="cNum lineTotal">
                                {num((+it.qty || 0) * (+it.price || 0), 0)}
                              </td>
                              <td className="cX">
                                <button onClick={() => removeItem(co.id, it.id)}>
                                  ×
                                </button>
                              </td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    )}

                    <div className="chFoot">
                      <button className="add" onClick={() => addItem(co.id)}>
                        + סעיף
                      </button>
                    </div>
                  </>
                )}
              </section>
            );
          })}

          <button className="addChapter" onClick={add}>
            + שינוי
          </button>
        </>
      )}

      {printed && (
        <ChangeDoc
          co={printed}
          project={project}
          company={company}
          onClose={() => setPrintId(null)}
        />
      )}
    </div>
  );
}

const coClass = (st) =>
  st === CO_APPROVED ? "ok" : st === CO_SUBMITTED ? "warn" : st === CO_REJECTED ? "bad" : "";

/* ---------- מסמך שינוי לחתימה ---------- */

function ChangeDoc({ co, project, company, onClose }) {
  const tot = coTotal(co);
  const vat = tot * VAT_RATE;

  return (
    <div className="scrim" onClick={onClose}>
      <div className="quoteWrap" onClick={(e) => e.stopPropagation()}>
        <div className="quoteBar">
          <button onClick={() => window.print()}>הדפס / PDF</button>
          <button onClick={onClose}>סגור</button>
        </div>

        <div className="quote">
          <header className="qHead">
            <div>
              <h2>{company.name || "———"}</h2>
              <p>
                {[company.phone, company.vat && `ע.מ ${company.vat}`]
                  .filter(Boolean)
                  .join(" · ")}
              </p>
            </div>
            <div className="qMeta">
              <strong>בקשת שינוי ש{co.number}</strong>
              <span>{co.date}</span>
            </div>
          </header>

          <div className="qProject">
            <p><b>פרויקט:</b> {project.name || "—"}</p>
            <p><b>לקוח:</b> {project.client_name || "—"}</p>
            <p><b>כתובת:</b> {project.address || "—"}</p>
          </div>

          <div className="coDocHead">
            <h3>{co.title || "שינוי ללא נושא"}</h3>
            {co.reason && <p><b>סיבה:</b> {co.reason}</p>}
            {co.requested_by && <p><b>לבקשת:</b> {co.requested_by}</p>}
          </div>

          <div className="qChapter">
            <table>
              <tbody>
                {co.items.map((it, k) => (
                  <tr key={it.id}>
                    <td className="qn">ש{co.number}.{k + 1}</td>
                    <td>{it.desc || "—"}</td>
                    <td className="qu">{it.unit}</td>
                    <td className="qnum">{num(+it.qty || 0)}</td>
                    <td className="qnum">{num(+it.price || 0, 0)}</td>
                    <td className="qnum b">
                      {num((+it.qty || 0) * (+it.price || 0), 0)}
                    </td>
                  </tr>
                ))}
                {co.items.length === 0 && (
                  <tr><td className="empty">אין סעיפים בשינוי זה.</td></tr>
                )}
              </tbody>
            </table>
          </div>

          <div className="qTotals">
            <div><span>סכום השינוי</span><b>{shekel(tot)}</b></div>
            <div><span>מע"מ {VAT_RATE * 100}%</span><b>{shekel(vat)}</b></div>
            <div className="qGrand">
              <span>תוספת לחוזה</span>
              <b>{shekel(tot + vat)}</b>
            </div>
          </div>

          <div className="signRow">
            <div className="signBox">
              <span>חתימת הלקוח</span>
              <i />
              <em>תאריך</em>
            </div>
            <div className="signBox">
              <span>חתימת הקבלן</span>
              <i />
              <em>תאריך</em>
            </div>
          </div>

          <p className="qNote">
            השינוי ייכנס לתוקף רק לאחר חתימת שני הצדדים. עבודה שתתבצע לפני
            החתימה אינה מחייבת את הלקוח בתשלום.
          </p>
        </div>
      </div>
    </div>
  );
}


/* ============================================================
   פיקוח באתר
   ============================================================ */

function SiteTab({ project, patch }) {
  const [view, setView] = useState("defects");
  const [filter, setFilter] = useState("open");
  const defects = project.defects || [];
  const visits = project.visits || [];
  const ds = defectSummary(defects);

  const setDefects = (fn) =>
    patch({ defects: typeof fn === "function" ? fn(defects) : fn });
  const setVisits = (fn) =>
    patch({ visits: typeof fn === "function" ? fn(visits) : fn });

  const updDefect = (id, field, value) =>
    setDefects((l) => l.map((d) => (d.id === id ? { ...d, [field]: value } : d)));

  const shown = useMemo(() => {
    const sorted = [...defects].sort((a, b) => {
      const ao = defectOverdue(a) ? 0 : 1;
      const bo = defectOverdue(b) ? 0 : 1;
      if (ao !== bo) return ao - bo;
      return (a.due || "9999").localeCompare(b.due || "9999");
    });
    if (filter === "open") return sorted.filter(defectOpen);
    if (filter === "overdue") return sorted.filter((d) => defectOverdue(d));
    return sorted;
  }, [defects, filter]);

  const exportDefects = () => {
    const rows = [["ליקוי", "מיקום", "אחראי", "חומרה", "סטטוס", "נפתח", "יעד", "נסגר"]];
    defects.forEach((d) =>
      rows.push([d.title, d.location, d.responsible, d.severity, d.status, d.opened, d.due, d.closed])
    );
    downloadCsv(rows, `${project.name || "פיקוח"}_ליקויים.csv`);
  };

  return (
    <div className="page">
      <h2 className="pageTitle">פיקוח באתר</h2>

      <div className="statRow">
        <Stat label="ליקויים פתוחים" value={ds.open} tone={ds.open ? "warn" : "good"} />
        <Stat label="ממתינים לאימות" value={ds.fixed} />
        <Stat label="באיחור" value={ds.overdue} tone={ds.overdue ? "bad" : ""} />
        <Stat label="ביקורים" value={visits.length} />
      </div>

      {ds.high > 0 && (
        <p className="warnBox">
          {ds.high} ליקויים בחומרה גבוהה עדיין לא אומתו כסגורים.
        </p>
      )}

      <div className="toolRow">
        <div className="segmented">
          <button className={view === "defects" ? "on" : ""} onClick={() => setView("defects")}>
            ליקויים
          </button>
          <button className={view === "log" ? "on" : ""} onClick={() => setView("log")}>
            יומן ביקורים
          </button>
        </div>
        <span className="spacer" />
        {view === "defects" && (
          <>
            <div className="segmented">
              {[
                ["open", "פתוחים"],
                ["overdue", "באיחור"],
                ["all", "הכל"],
              ].map(([k, l]) => (
                <button key={k} className={filter === k ? "on" : ""} onClick={() => setFilter(k)}>
                  {l}
                </button>
              ))}
            </div>
            <button className="toolBtn" onClick={exportDefects}>ייצוא CSV</button>
          </>
        )}
      </div>

      {view === "defects" ? (
        defects.length === 0 ? (
          <div className="blank small">
            <h3>אין ליקויים</h3>
            <p>רשום כאן כל ליקוי שמצאת בביקור, עם אחראי ותאריך יעד לתיקון.</p>
            <button className="primaryBtn" onClick={() => setDefects((l) => [...l, newDefect()])}>
              הוסף ליקוי
            </button>
          </div>
        ) : (
          <>
            <table className="grid">
              <thead>
                <tr>
                  <th>הליקוי</th>
                  <th className="gloc">מיקום</th>
                  <th className="gloc">אחראי</th>
                  <th className="gsev">חומרה</th>
                  <th className="gstat">סטטוס</th>
                  <th className="gdate">יעד</th>
                  <th />
                </tr>
              </thead>
              <tbody>
                {shown.map((d) => (
                  <tr key={d.id} className={defectOverdue(d) ? "overdue" : ""}>
                    <td>
                      <input value={d.title} placeholder="מה לא תקין" onChange={(e) => updDefect(d.id, "title", e.target.value)} />
                    </td>
                    <td className="gloc">
                      <input value={d.location} placeholder="קומה / חדר" onChange={(e) => updDefect(d.id, "location", e.target.value)} />
                    </td>
                    <td className="gloc">
                      <input value={d.responsible} placeholder="קבלן משנה" onChange={(e) => updDefect(d.id, "responsible", e.target.value)} />
                    </td>
                    <td className="gsev">
                      <select className={`sev ${sevClass(d.severity)}`} value={d.severity} onChange={(e) => updDefect(d.id, "severity", e.target.value)}>
                        {SEVERITIES.map((s) => <option key={s}>{s}</option>)}
                      </select>
                    </td>
                    <td className="gstat">
                      <select
                        className={`sev ${defClass(d.status)}`}
                        value={d.status}
                        onChange={(e) => {
                          updDefect(d.id, "status", e.target.value);
                          if ((e.target.value === DEF_VERIFIED || e.target.value === DEF_WAIVED) && !d.closed)
                            updDefect(d.id, "closed", today());
                        }}
                      >
                        {DEF_STATUSES.map((s) => <option key={s}>{s}</option>)}
                      </select>
                    </td>
                    <td className="gdate">
                      <input type="date" value={d.due || ""} onChange={(e) => updDefect(d.id, "due", e.target.value)} />
                    </td>
                    <td className="cX">
                      <button onClick={() => setDefects((l) => l.filter((x) => x.id !== d.id))}>×</button>
                    </td>
                  </tr>
                ))}
                {shown.length === 0 && (
                  <tr><td colSpan={7} className="empty">אין ליקויים בסינון הזה.</td></tr>
                )}
              </tbody>
            </table>
            <button className="addChapter" onClick={() => setDefects((l) => [...l, newDefect()])}>
              + ליקוי
            </button>
          </>
        )
      ) : visits.length === 0 ? (
        <div className="blank small">
          <h3>היומן ריק</h3>
          <p>תעד כל ביקור באתר: מי עבד, מה התקדם, ומה דורש טיפול.</p>
          <button className="primaryBtn" onClick={() => setVisits((l) => [newVisit(), ...l])}>
            רשום ביקור
          </button>
        </div>
      ) : (
        <>
          {visits.map((v) => (
            <section className="visit" key={v.id}>
              <div className="visitHead">
                <input type="date" className="visitDate" value={v.date}
                  onChange={(e) => setVisits((l) => l.map((x) => x.id === v.id ? { ...x, date: e.target.value } : x))} />
                <input placeholder="מזג אוויר" value={v.weather}
                  onChange={(e) => setVisits((l) => l.map((x) => x.id === v.id ? { ...x, weather: e.target.value } : x))} />
                <input placeholder="מספר עובדים" value={v.workers}
                  onChange={(e) => setVisits((l) => l.map((x) => x.id === v.id ? { ...x, workers: e.target.value } : x))} />
                <input placeholder="קבלני משנה באתר" value={v.subcontractors}
                  onChange={(e) => setVisits((l) => l.map((x) => x.id === v.id ? { ...x, subcontractors: e.target.value } : x))} />
                <button className="x" onClick={() => setVisits((l) => l.filter((x) => x.id !== v.id))}>×</button>
              </div>
              <textarea
                placeholder="מה בוצע היום, מה עוכב, ומה דורש החלטה"
                value={v.notes}
                onChange={(e) => setVisits((l) => l.map((x) => x.id === v.id ? { ...x, notes: e.target.value } : x))}
              />
            </section>
          ))}
          <button className="addChapter" onClick={() => setVisits((l) => [newVisit(), ...l])}>
            + ביקור
          </button>
        </>
      )}
    </div>
  );
}

const sevClass = (s) => (s === "גבוהה" ? "bad" : s === "בינונית" ? "warn" : "");
const defClass = (s) =>
  s === DEF_VERIFIED ? "ok" : s === DEF_FIXED ? "warn" : s === DEF_WAIVED ? "" : "bad";

/* ============================================================
   תקציב
   ============================================================ */

const BUDGET_ROWS = [
  { key: "materials", label: "חומרים" },
  { key: "labor", label: "עבודה" },
  { key: "subcontractors", label: "קבלני משנה" },
];

function BudgetTab({ project, patch }) {
  const b = project.budget || emptyBudget();
  const t = budgetTotals(b);
  const revenue = boqTotal(project.chapters);

  const set = (field, value) =>
    patch({ budget: { ...b, [field]: +value || 0 } });

  return (
    <div className="page narrow">
      <h2 className="pageTitle">תקציב ומעקב ביצוע</h2>

      <div className="budgetGrid">
        <div className="bHead">
          <span>סעיף</span>
          <span className="bnum">מתוכנן</span>
          <span className="bnum">בפועל</span>
          <span className="bnum">נותר</span>
          <span className="bbar" />
        </div>

        {BUDGET_ROWS.map((row) => {
          const planned = +b[`planned_${row.key}`] || 0;
          const actual = +b[`actual_${row.key}`] || 0;
          const used = planned > 0 ? (actual / planned) * 100 : actual > 0 ? 100 : 0;
          const over = planned > 0 && actual > planned;
          return (
            <div className="bRow" key={row.key}>
              <span className="blabel">{row.label}</span>
              <span className="bnum">
                <input
                  type="number"
                  step="any"
                  value={b[`planned_${row.key}`]}
                  onChange={(e) => set(`planned_${row.key}`, e.target.value)}
                />
              </span>
              <span className="bnum">
                <input
                  type="number"
                  step="any"
                  value={b[`actual_${row.key}`]}
                  onChange={(e) => set(`actual_${row.key}`, e.target.value)}
                />
              </span>
              <span className={`bnum bres ${planned - actual < 0 ? "neg" : ""}`}>
                {shekel(planned - actual)}
              </span>
              <span className="bbar">
                <i
                  className={over ? "over" : ""}
                  style={{ width: `${Math.min(used, 100)}%` }}
                />
                <em>{planned > 0 ? pct(used) : "—"}</em>
              </span>
            </div>
          );
        })}

        <div className="bRow bTotal">
          <span className="blabel">סה"כ</span>
          <span className="bnum">{shekel(t.planned)}</span>
          <span className="bnum">{shekel(t.actual)}</span>
          <span className={`bnum bres ${t.remaining < 0 ? "neg" : ""}`}>
            {shekel(t.remaining)}
          </span>
          <span className="bbar">
            <i
              className={t.actual > t.planned && t.planned > 0 ? "over" : ""}
              style={{
                width: `${t.planned > 0 ? Math.min((t.actual / t.planned) * 100, 100) : 0}%`,
              }}
            />
            <em>{t.planned > 0 ? pct((t.actual / t.planned) * 100) : "—"}</em>
          </span>
        </div>
      </div>

      <div className="statRow">
        <Stat label="הכנסה לפי כתב הכמויות" value={shekel(revenue)} />
        <Stat label="עלות מתוכננת" value={shekel(t.planned)} />
        <Stat label="עלות בפועל" value={shekel(t.actual)} />
        <Stat
          label="רווח צפוי"
          value={shekel(revenue - t.planned)}
          tone={revenue - t.planned < 0 ? "bad" : "good"}
        />
      </div>

      {t.remaining < 0 && (
        <p className="warnBox">
          חריגה מהתקציב של {shekel(Math.abs(t.remaining))}. בדוק את סעיפי
          ההוצאה מול המתוכנן.
        </p>
      )}
    </div>
  );
}

/* ============================================================
   לוח זמנים
   ============================================================ */

function ScheduleTab({ project, patch }) {
  const ms = project.milestones || [];
  const sum = scheduleSummary(ms);

  const setMs = (fn) =>
    patch({ milestones: typeof fn === "function" ? fn(ms) : fn });

  const add = () =>
    setMs((list) => [
      ...list,
      { id: uid(), name: "", planned_date: "", actual_date: "", done: false },
    ]);

  const upd = (id, field, value) =>
    setMs((list) => list.map((m) => (m.id === id ? { ...m, [field]: value } : m)));

  const sorted = useMemo(
    () =>
      [...ms].sort((a, b) =>
        (a.planned_date || "9999").localeCompare(b.planned_date || "9999")
      ),
    [ms]
  );

  return (
    <div className="page narrow">
      <h2 className="pageTitle">אבני דרך ולוח זמנים</h2>

      <div className="statRow">
        <Stat label="אבני דרך" value={sum.total} />
        <Stat label="הושלמו" value={sum.done} tone="good" />
        <Stat label="באיחור" value={sum.delayed} tone={sum.delayed ? "bad" : ""} />
        <Stat
          label="התקדמות"
          value={sum.total ? pct((sum.done / sum.total) * 100) : "—"}
        />
      </div>

      {ms.length === 0 ? (
        <div className="blank small">
          <h3>אין אבני דרך</h3>
          <p>הוסף שלבים כמו "סיום שלד" או "מסירה ללקוח" כדי לעקוב אחרי הלו״ז.</p>
          <button className="primaryBtn" onClick={add}>הוסף אבן דרך</button>
        </div>
      ) : (
        <>
          <table className="grid">
            <thead>
              <tr>
                <th>אבן דרך</th>
                <th className="gdate">מתוכנן</th>
                <th className="gdate">בפועל</th>
                <th className="gstat">סטטוס</th>
                <th className="gdone">בוצע</th>
                <th />
              </tr>
            </thead>
            <tbody>
              {sorted.map((m) => {
                const st = milestoneStatus(m);
                return (
                  <tr key={m.id}>
                    <td>
                      <input
                        value={m.name}
                        placeholder="שם השלב"
                        onChange={(e) => upd(m.id, "name", e.target.value)}
                      />
                    </td>
                    <td className="gdate">
                      <input
                        type="date"
                        value={m.planned_date || ""}
                        onChange={(e) => upd(m.id, "planned_date", e.target.value)}
                      />
                    </td>
                    <td className="gdate">
                      <input
                        type="date"
                        value={m.actual_date || ""}
                        onChange={(e) => upd(m.id, "actual_date", e.target.value)}
                      />
                    </td>
                    <td className="gstat">
                      <span className={`pill ${statusClass(st)}`}>{st}</span>
                    </td>
                    <td className="gdone">
                      <input
                        type="checkbox"
                        checked={!!m.done}
                        onChange={(e) => {
                          upd(m.id, "done", e.target.checked);
                          if (e.target.checked && !m.actual_date)
                            upd(m.id, "actual_date", today());
                        }}
                      />
                    </td>
                    <td className="cX">
                      <button onClick={() => setMs((l) => l.filter((x) => x.id !== m.id))}>×</button>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
          <button className="addChapter" onClick={add}>+ אבן דרך</button>
        </>
      )}
    </div>
  );
}

const statusClass = (st) =>
  st === ST_DONE_ON_TIME ? "ok" : st === ST_DONE_LATE ? "warn" : st === ST_DELAYED ? "bad" : "";

/* ============================================================
   רכש
   ============================================================ */

function ProcurementTab({ project, patch }) {
  const orders = project.orders || [];
  const sum = procurementSummary(orders);

  const setOrders = (fn) =>
    patch({ orders: typeof fn === "function" ? fn(orders) : fn });

  const add = () =>
    setOrders((l) => [
      ...l,
      {
        id: uid(),
        supplier: "",
        description: "",
        quantity: 1,
        unit_price: 0,
        status: PO_ORDERED,
        order_date: today(),
        received_date: "",
      },
    ]);

  const upd = (id, field, value) =>
    setOrders((l) => l.map((o) => (o.id === id ? { ...o, [field]: value } : o)));

  const exportCsv = () => {
    const rows = [["ספק", "תיאור", "כמות", "מחיר יחידה", 'סה"כ', "סטטוס", "הוזמן", "התקבל"]];
    orders.forEach((o) =>
      rows.push([
        o.supplier, o.description, o.quantity, o.unit_price,
        orderTotal(o), o.status, o.order_date, o.received_date,
      ])
    );
    downloadCsv(rows, `${project.name || "רכש"}_רכש.csv`);
  };

  return (
    <div className="page">
      <h2 className="pageTitle">הזמנות רכש</h2>

      <div className="statRow">
        <Stat label="הוזמן" value={shekel(sum.ordered)} />
        <Stat label="התקבל" value={shekel(sum.received)} tone="good" />
        <Stat label="בדרך" value={shekel(sum.ordered - sum.received)} />
        <Stat label="הזמנות" value={orders.length} />
      </div>

      {orders.length === 0 ? (
        <div className="blank small">
          <h3>אין הזמנות</h3>
          <p>רשום כאן הזמנות מספקים כדי לעקוב אחרי מה שהוזמן ומה שהגיע.</p>
          <button className="primaryBtn" onClick={add}>הוסף הזמנה</button>
        </div>
      ) : (
        <>
          <table className="grid">
            <thead>
              <tr>
                <th className="gsup">ספק</th>
                <th>תיאור</th>
                <th className="gnum">כמות</th>
                <th className="gnum">מחיר יח'</th>
                <th className="gnum">סה"כ</th>
                <th className="gstat">סטטוס</th>
                <th className="gdate">הוזמן</th>
                <th className="gdate">התקבל</th>
                <th />
              </tr>
            </thead>
            <tbody>
              {orders.map((o) => (
                <tr key={o.id} className={o.status === PO_CANCELLED ? "cancelled" : ""}>
                  <td className="gsup">
                    <input value={o.supplier} placeholder="שם הספק" onChange={(e) => upd(o.id, "supplier", e.target.value)} />
                  </td>
                  <td>
                    <input value={o.description} placeholder="מה הוזמן" onChange={(e) => upd(o.id, "description", e.target.value)} />
                  </td>
                  <td className="gnum">
                    <input type="number" step="any" value={o.quantity} onChange={(e) => upd(o.id, "quantity", e.target.value)} />
                  </td>
                  <td className="gnum">
                    <input type="number" step="any" value={o.unit_price} onChange={(e) => upd(o.id, "unit_price", e.target.value)} />
                  </td>
                  <td className="gnum lineTotal">{num(orderTotal(o), 0)}</td>
                  <td className="gstat">
                    <select
                      value={o.status}
                      onChange={(e) => {
                        upd(o.id, "status", e.target.value);
                        if (e.target.value === PO_RECEIVED && !o.received_date)
                          upd(o.id, "received_date", today());
                      }}
                    >
                      {PO_STATUSES.map((s) => <option key={s}>{s}</option>)}
                    </select>
                  </td>
                  <td className="gdate">
                    <input type="date" value={o.order_date || ""} onChange={(e) => upd(o.id, "order_date", e.target.value)} />
                  </td>
                  <td className="gdate">
                    <input type="date" value={o.received_date || ""} onChange={(e) => upd(o.id, "received_date", e.target.value)} />
                  </td>
                  <td className="cX">
                    <button onClick={() => setOrders((l) => l.filter((x) => x.id !== o.id))}>×</button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
          <div className="rowBtns">
            <button className="addChapter" onClick={add}>+ הזמנה</button>
            <button className="toolBtn" onClick={exportCsv}>ייצוא CSV</button>
          </div>
        </>
      )}
    </div>
  );
}

/* ============================================================
   דוחות
   ============================================================ */

function ReportTab({ project }) {
  const base = boqTotal(project.chapters);
  const oh = base * ((project.overhead || 0) / 100);
  const original = base + oh;

  const ch = changeSummary(project.changes);
  const revised = original + ch.approved;

  const ex = executionTotals(project.chapters);
  const b = project.budget || emptyBudget();
  const bt = budgetTotals(b);
  const proc = procurementSummary(project.orders || []);
  const sched = scheduleSummary(project.milestones || []);
  const ds = defectSummary(project.defects);

  /* עלות בפועל אם נרשמה, אחרת המתוכננת */
  const cost = bt.actual || bt.planned;
  const profit = revised - cost;
  const margin = revised > 0 ? (profit / revised) * 100 : 0;

  /* רווח לפי מה שבוצע עד כה, ולא לפי החוזה כולו */
  const earned = ex.done;
  const profitToDate = earned - bt.actual;

  const chapterRows = project.chapters
    .map((c, i) => ({
      no: chapterNo(i),
      title: c.title,
      items: c.items.length,
      total: c.items.reduce((s, it) => s + (+it.qty || 0) * (+it.price || 0), 0),
    }))
    .filter((r) => r.items > 0);

  const exportAll = () => {
    const rows = [
      ["דוח רווחיות", project.name || ""],
      ["לקוח", project.client_name || ""],
      ["כתובת", project.address || ""],
      [],
      ["כתב כמויות", base.toFixed(2)],
      ["רווח וקבלנות", oh.toFixed(2)],
      ["שינויים שאושרו", ch.approved.toFixed(2)],
      ["חוזה מעודכן", revised.toFixed(2)],
      ["שינויים ממתינים (לא נספרו)", ch.pending.toFixed(2)],
      [],
      ["בוצע עד כה", earned.toFixed(2)],
      ["התקדמות %", ex.percent.toFixed(1)],
      ["נותר לביצוע", ex.remaining.toFixed(2)],
      ["עלות בפועל", bt.actual.toFixed(2)],
      ["עלות מתוכננת", bt.planned.toFixed(2)],
      ["רווח", profit.toFixed(2)],
      ["שיעור רווח %", margin.toFixed(1)],
      [],
      ["ליקויים פתוחים", ds.open + ds.fixed],
      ["ליקויים באיחור", ds.overdue],
      ["אבני דרך", `${sched.done} / ${sched.total}`],
      [],
      ["פרק", "סעיפים", 'סה"כ'],
      ...chapterRows.map((r) => [`${r.no} ${r.title}`, r.items, r.total.toFixed(2)]),
    ];
    downloadCsv(rows, `${project.name || "דוח"}_רווחיות.csv`);
  };

  return (
    <div className="page narrow">
      <h2 className="pageTitle">דוח רווחיות</h2>

      <div className="statRow big">
        <Stat label="חוזה מעודכן" value={shekel(revised)} />
        <Stat label="עלות" value={shekel(cost)} />
        <Stat label="רווח צפוי" value={shekel(profit)} tone={profit < 0 ? "bad" : "good"} />
        <Stat
          label="שיעור רווח"
          value={pct(margin)}
          tone={margin < 0 ? "bad" : margin < 10 ? "warn" : "good"}
        />
      </div>

      {margin < 10 && revised > 0 && (
        <p className="warnBox">
          שיעור הרווח {pct(margin)} נמוך. בדוק את התמחור או את העלויות בפועל.
        </p>
      )}

      <h3 className="secTitle">הרכב שווי החוזה</h3>
      <table className="grid readonly">
        <tbody>
          <tr>
            <td>כתב כמויות</td>
            <td className="gnum lineTotal">{num(base, 0)}</td>
          </tr>
          {oh > 0 && (
            <tr>
              <td>רווח וקבלנות {project.overhead}%</td>
              <td className="gnum lineTotal">{num(oh, 0)}</td>
            </tr>
          )}
          <tr>
            <td>שינויים שאושרו ({ch.approvedCount})</td>
            <td className="gnum lineTotal">{num(ch.approved, 0)}</td>
          </tr>
          <tr className="sumRow">
            <td>חוזה מעודכן</td>
            <td className="gnum lineTotal">{num(revised, 0)}</td>
          </tr>
          {ch.pending > 0 && (
            <tr className="muteRow">
              <td>ממתינים לאישור ({ch.pendingCount}) — לא נספרו</td>
              <td className="gnum">{num(ch.pending, 0)}</td>
            </tr>
          )}
        </tbody>
      </table>

      <h3 className="secTitle">מצב ביצוע ותזרים</h3>
      <div className="statRow">
        <Stat label="בוצע עד כה" value={shekel(earned)} />
        <Stat label="התקדמות" value={pct(ex.percent)} />
        <Stat label="נותר לביצוע" value={shekel(ex.remaining)} />
        <Stat
          label="רווח על מה שבוצע"
          value={shekel(profitToDate)}
          tone={profitToDate < 0 ? "bad" : "good"}
        />
      </div>

      <h3 className="secTitle">פירוט לפי פרקים</h3>
      <table className="grid readonly">
        <thead>
          <tr>
            <th className="gno">#</th>
            <th>פרק</th>
            <th className="gnum">סעיפים</th>
            <th className="gnum">סה"כ</th>
            <th className="gnum">חלק</th>
          </tr>
        </thead>
        <tbody>
          {chapterRows.map((r) => (
            <tr key={r.no}>
              <td className="gno">{r.no}</td>
              <td>{r.title}</td>
              <td className="gnum">{r.items}</td>
              <td className="gnum lineTotal">{num(r.total, 0)}</td>
              <td className="gnum">{base > 0 ? pct((r.total / base) * 100) : "—"}</td>
            </tr>
          ))}
          {chapterRows.length === 0 && (
            <tr><td colSpan={5} className="empty">אין סעיפים בכתב הכמויות.</td></tr>
          )}
        </tbody>
      </table>

      <h3 className="secTitle">מצב הפרויקט</h3>
      <div className="miniGrid">
        <MiniStat label="אבני דרך שהושלמו" value={`${sched.done} / ${sched.total}`} />
        <MiniStat label="אבני דרך באיחור" value={sched.delayed} tone={sched.delayed ? "bad" : ""} />
        <MiniStat label="רכש שהוזמן" value={shekel(proc.ordered)} />
        <MiniStat label="רכש שהתקבל" value={shekel(proc.received)} />
        <MiniStat label="ליקויים פתוחים" value={ds.open + ds.fixed} tone={ds.open ? "bad" : ""} />
        <MiniStat label="ליקויים באיחור" value={ds.overdue} tone={ds.overdue ? "bad" : ""} />
        <MiniStat label="סעיפים שהושלמו" value={`${ex.complete} / ${ex.items}`} />
        <MiniStat label="שינויים ממתינים" value={ch.pendingCount} tone={ch.pendingCount ? "bad" : ""} />
      </div>

      <button className="toolBtn accent" onClick={exportAll}>ייצוא הדוח ל-CSV</button>
    </div>
  );
}

/* ============================================================
   רכיבים משותפים
   ============================================================ */

function Field({ label, value, onChange, placeholder, type = "text" }) {
  return (
    <label className="field">
      <span>{label}</span>
      <input
        type={type}
        value={value || ""}
        placeholder={placeholder}
        onChange={(e) => onChange(e.target.value)}
      />
    </label>
  );
}

function Stat({ label, value, tone = "" }) {
  return (
    <div className={`stat ${tone}`}>
      <span>{label}</span>
      <strong>{value}</strong>
    </div>
  );
}

function MiniStat({ label, value, tone = "" }) {
  return (
    <div className={`mini ${tone}`}>
      <span>{label}</span>
      <b>{value}</b>
    </div>
  );
}

function downloadCsv(rows, filename) {
  const csv =
    "\uFEFF" +
    rows
      .map((r) => (r || []).map((v) => `"${String(v ?? "").replace(/"/g, '""')}"`).join(","))
      .join("\n");
  const url = URL.createObjectURL(new Blob([csv], { type: "text/csv;charset=utf-8;" }));
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

/* ============================================================
   דיאלוג ייבוא מאקסל
   ============================================================ */

function ImportDialog({ imp, setImp, records, stats, hier, level, selectSheet, onApply }) {
  return (
    <div className="scrim" onClick={() => setImp(null)}>
      <div className="importWrap" onClick={(e) => e.stopPropagation()}>
        <div className="drawerHead">
          <h2>ייבוא כתב כמויות</h2>
          <button onClick={() => setImp(null)}>×</button>
        </div>

        {imp.error ? (
          <p className="empty">{imp.error}</p>
        ) : (
          <>
            <div className="impTop">
              <span className="impFile">{imp.fileName}</span>

              {imp.sheets.length > 1 && (
                <label className="impField">
                  <span>גיליון</span>
                  <select
                    value={imp.idx}
                    onChange={(e) => selectSheet(imp.sheets, +e.target.value, imp.fileName)}
                  >
                    {imp.sheets.map((s, i) => (
                      <option key={s.name} value={i}>{s.name}</option>
                    ))}
                  </select>
                </label>
              )}

              <label className="impField">
                <span>שורת הכותרות</span>
                <select
                  value={imp.headerRow}
                  onChange={(e) => {
                    const hr = +e.target.value;
                    setImp({ ...imp, headerRow: hr, map: autoMap(imp.rows[hr] || []) });
                  }}
                >
                  {imp.rows.slice(0, 25).map((r, i) => (
                    <option key={i} value={i}>
                      {i + 1}: {r.filter(Boolean).slice(0, 4).join(" | ").slice(0, 46) || "(ריקה)"}
                    </option>
                  ))}
                </select>
              </label>
            </div>

            {hier && (
              <div className="hierBar">
                <span className="hierLabel">רמת הפרקים</span>
                <div className="hierPick">
                  <button
                    className={level === 1 ? "on" : ""}
                    onClick={() => setImp({ ...imp, level: 1 })}
                  >
                    ראשית · {hier.main} פרקים
                  </button>
                  <button
                    className={level === 2 ? "on" : ""}
                    onClick={() => setImp({ ...imp, level: 2 })}
                  >
                    משנית · {hier.sub} פרקים
                  </button>
                </div>
                <span className="hierNote">
                  הקובץ בנוי בשלוש רמות. מספר הקטלוג נשמר בכל סעיף.
                </span>
              </div>
            )}

            <div className="mapGrid">
              {[
                ["desc", "תיאור", true],
                ["qty", "כמות", true],
                ["unit", "יחידה", false],
                ["price", "מחיר יחידה", false],
                ["code", "מספר סעיף", false],
              ].map(([key, label, required]) => (
                <label key={key} className="mapField">
                  <span>{label}{required && <b> *</b>}</span>
                  <select
                    value={imp.map[key] ?? ""}
                    onChange={(e) =>
                      setImp({
                        ...imp,
                        map: {
                          ...imp.map,
                          [key]: e.target.value === "" ? undefined : +e.target.value,
                        },
                      })
                    }
                  >
                    <option value="">— אין —</option>
                    {(imp.rows[imp.headerRow] || []).map((h, i) => (
                      <option key={i} value={i}>
                        {String(h).slice(0, 30) || `עמודה ${i + 1}`}
                      </option>
                    ))}
                  </select>
                </label>
              ))}
            </div>

            <div className="impStats">
              <span><b>{stats.items}</b> סעיפים</span>
              <span><b>{stats.chapters}</b> פרקים</span>
              <span className={stats.priced ? "" : "warn"}><b>{stats.priced}</b> מתומחרים</span>
              {stats.unitBad > 0 && (
                <span className="warn"><b>{stats.unitBad}</b> ללא יחידה תקינה</span>
              )}
            </div>

            {stats.unitBad > 0 && (
              <p className="unitWarn">
                ב-{stats.unitBad} סעיפים עמודת היחידה ריקה או מכילה ערך שאינו
                יחידת מדידה. הם ייובאו בלי יחידה ויסומנו בכתב הכמויות — יחידה
                שגויה משנה את המחיר בסדר גודל שלם, ולכן אני לא מנחש אותה.
              </p>
            )}

            <div className="preview">
              <table>
                <thead>
                  <tr>
                    <th className="pc">סעיף</th>
                    <th>תיאור</th>
                    <th className="pu">יח'</th>
                    <th className="pn">כמות</th>
                    <th className="pn">מחיר</th>
                  </tr>
                </thead>
                <tbody>
                  {records.slice(0, 40).map((r, i) => {
                    const isChap = hier
                      ? !(r.qty || r.depth >= 3) && r.depth === level
                      : r.isChapter;
                    const skipped = hier && !(r.qty || r.depth >= 3) && !isChap;
                    return (
                    <tr key={i} className={isChap ? "pchap" : skipped ? "pskip" : ""}>
                      <td className="pc">{r.code}</td>
                      <td>{r.desc || "—"}</td>
                      <td className={`pu ${r.unitBad ? "bad" : ""}`}>
                        {isChap ? "" : r.unitBad ? (r.unit || "—") : r.unit}
                      </td>
                      <td className="pn">{isChap ? "" : num(r.qty)}</td>
                      <td className="pn">{isChap ? "" : num(r.price, 0)}</td>
                    </tr>
                    );
                  })}
                </tbody>
              </table>
              {records.length > 40 && <p className="more">ועוד {records.length - 40} שורות…</p>}
              {records.length === 0 && (
                <p className="empty">לא זוהו שורות. שנה את שורת הכותרות או את התאמת העמודות.</p>
              )}
            </div>

            <div className="impFoot">
              <p className="impHint">
                שורה עם טקסט בלי כמות מזוהה כפרק. אפשר לתקן הכל אחרי הייבוא.
              </p>
              <div className="impBtns">
                <button className="add" disabled={!records.length} onClick={() => onApply("append")}>
                  הוסף לקיים
                </button>
                <button className="add primary" disabled={!records.length} onClick={() => onApply("replace")}>
                  החלף הכל
                </button>
              </div>
            </div>
          </>
        )}
      </div>
    </div>
  );
}

/* ============================================================
   הצעת מחיר
   ============================================================ */

function QuoteDialog({ project, chapters, chapterTotals, totals, company, setCompany, onClose }) {
  const { net, oh, beforeVat, vat, total, overhead } = totals;

  return (
    <div className="scrim" onClick={onClose}>
      <div className="quoteWrap" onClick={(e) => e.stopPropagation()}>
        <div className="quoteBar">
          <input placeholder="שם החברה שלך" value={company.name}
            onChange={(e) => setCompany({ ...company, name: e.target.value })} />
          <input placeholder="טלפון" value={company.phone}
            onChange={(e) => setCompany({ ...company, phone: e.target.value })} />
          <input placeholder="ע.מ / ח.פ" value={company.vat}
            onChange={(e) => setCompany({ ...company, vat: e.target.value })} />
          <button onClick={() => window.print()}>הדפס / PDF</button>
          <button onClick={onClose}>סגור</button>
        </div>

        <div className="quote">
          <header className="qHead">
            <div>
              <h2>{company.name || "———"}</h2>
              <p>
                {[company.phone, company.vat && `ע.מ ${company.vat}`]
                  .filter(Boolean)
                  .join(" · ")}
              </p>
            </div>
            <div className="qMeta">
              <strong>הצעת מחיר</strong>
              <span>{new Date().toLocaleDateString("he-IL")}</span>
            </div>
          </header>

          <div className="qProject">
            <p><b>פרויקט:</b> {project.name || "—"}</p>
            <p><b>לקוח:</b> {project.client_name || "—"}</p>
            <p><b>כתובת:</b> {project.address || "—"}</p>
          </div>

          {chapters.map((c, i) =>
            c.items.length ? (
              <div className="qChapter" key={c.id}>
                <h3><span>{chapterNo(i)}</span> {c.title}</h3>
                <table>
                  <tbody>
                    {c.items.map((it, k) => (
                      <tr key={it.id}>
                        <td className="qn">{it.code || `${chapterNo(i)}.${k + 1}`}</td>
                        <td>{it.desc || "—"}</td>
                        <td className="qu">{it.unit}</td>
                        <td className="qnum">{num(+it.qty || 0)}</td>
                        <td className="qnum">{num(+it.price || 0, 0)}</td>
                        <td className="qnum b">{num((+it.qty || 0) * (+it.price || 0), 0)}</td>
                      </tr>
                    ))}
                    <tr className="qSub">
                      <td colSpan={5}>סה"כ פרק {chapterNo(i)}</td>
                      <td className="qnum">{num(chapterTotals[i], 0)}</td>
                    </tr>
                  </tbody>
                </table>
              </div>
            ) : null
          )}

          <div className="qTotals">
            <div><span>סכום עבודות</span><b>{shekel(net)}</b></div>
            {overhead > 0 && (
              <div><span>רווח וקבלנות {overhead}%</span><b>{shekel(oh)}</b></div>
            )}
            <div><span>לפני מע"מ</span><b>{shekel(beforeVat)}</b></div>
            <div><span>מע"מ {VAT_RATE * 100}%</span><b>{shekel(vat)}</b></div>
            <div className="qGrand"><span>סה"כ לתשלום</span><b>{shekel(total)}</b></div>
          </div>

          <p className="qNote">
            ההצעה תקפה ל־30 יום. המחירים אינם כוללים שינויים ותוספות שלא פורטו לעיל.
          </p>
        </div>
      </div>
    </div>
  );
}

const CSS = `
@import url('https://fonts.googleapis.com/css2?family=Frank+Ruhl+Libre:wght@500;700&family=Heebo:wght@400;500;700&display=swap');

.boq{
  --ink:#131a20;
  --slate:#4b5a63;
  --mute:#8c9aa1;
  --concrete:#e7eae6;
  --paper:#ffffff;
  --rule:#d3d9d3;
  --ochre:#b3760a;
  --ochre-soft:#f6ecd9;
  --green:#0f6b52;
  --red:#a33b2a;

  font-family:'Heebo',system-ui,sans-serif;
  color:var(--ink);
  background:var(--concrete);
  min-height:100vh;
  font-variant-numeric:tabular-nums;
}
.boq *{box-sizing:border-box}
.boq button{font-family:inherit;cursor:pointer}
.boq input,.boq select{font-family:inherit;font-size:inherit;color:inherit}
.boq input:focus-visible,.boq select:focus-visible,.boq button:focus-visible{
  outline:2px solid var(--ochre);outline-offset:1px;
}

/* ---------- topbar ---------- */
.topbar{
  display:flex;align-items:center;justify-content:space-between;gap:16px;
  padding:14px 22px;background:var(--ink);color:#eef1ee;flex-wrap:wrap;
}
.brand{display:flex;align-items:center;gap:12px}
.mark{
  width:26px;height:26px;flex:0 0 auto;border-radius:3px;
  background:
    linear-gradient(var(--ochre),var(--ochre)) 0 0/100% 5px no-repeat,
    linear-gradient(#eef1ee,#eef1ee) 0 9px/70% 3px no-repeat,
    linear-gradient(#eef1ee,#eef1ee) 0 16px/85% 3px no-repeat,
    linear-gradient(#eef1ee,#eef1ee) 0 23px/55% 3px no-repeat;
}
.brand h1{font-family:'Frank Ruhl Libre',serif;font-size:19px;margin:0;letter-spacing:.2px}
.brand p{margin:1px 0 0;font-size:12px;color:#9daaa5}
.topActions{display:flex;align-items:center;gap:8px;flex-wrap:wrap}
.saveDot{font-size:12px;color:#9daaa5;margin-left:6px}
.saveDot.saved{color:#7fc9a8}
.saveDot.error{color:#e08a7a}
.ghost{
  background:transparent;color:#eef1ee;border:1px solid #3a464d;
  border-radius:4px;padding:6px 12px;font-size:13px;
}
.ghost:hover{background:#222c33}
.ghost.danger:hover{background:var(--red);border-color:var(--red)}

/* ---------- layout ---------- */
.layout{display:flex;align-items:flex-start;gap:22px;padding:22px;max-width:1400px;margin:0 auto}
.sheet{flex:1;min-width:0;display:flex;flex-direction:column;gap:16px}

/* ---------- tape (signature) ---------- */
.tape{
  flex:0 0 290px;position:sticky;top:22px;
  background:var(--paper);border:1px solid var(--rule);border-radius:6px;
  box-shadow:0 1px 0 rgba(0,0,0,.04);overflow:hidden;
}
.tapeHead{
  display:flex;justify-content:space-between;align-items:baseline;
  padding:12px 14px;border-bottom:2px solid var(--ink);
  font-family:'Frank Ruhl Libre',serif;font-size:16px;font-weight:700;
}
.tapeHead em{font-style:normal;font-size:12px;color:var(--mute);font-family:'Heebo'}
.tapeBody{padding:6px 0;max-height:34vh;overflow:auto}
.tapeRow{
  display:grid;grid-template-columns:22px 1fr auto;gap:8px;align-items:baseline;
  padding:5px 14px;font-size:13px;
}
.tno{color:var(--ochre);font-weight:700;font-size:11px}
.tname{color:var(--slate);overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.tval{font-weight:500}
.tapeFoot{border-top:1px dashed var(--rule);padding:8px 0 0;background:#fbfbfa}
.tapeRow.sum{grid-template-columns:1fr auto;font-weight:600}
.ohRow{
  display:grid;grid-template-columns:1fr auto auto;gap:8px;align-items:center;
  padding:6px 14px;font-size:13px;color:var(--slate);
}
.ohInput{display:flex;align-items:center;gap:2px}
.ohInput input{
  width:46px;text-align:center;border:1px solid var(--rule);border-radius:3px;
  padding:2px 4px;background:#fff;
}
.ohInput i{font-style:normal;color:var(--mute);font-size:12px}
.grand{
  display:flex;justify-content:space-between;align-items:baseline;
  margin-top:8px;padding:14px;background:var(--ink);color:#fff;
}
.grand span{font-size:13px;color:#a9b6b0}
.grand strong{font-family:'Frank Ruhl Libre',serif;font-size:22px;color:#f0c374}

/* ---------- project ---------- */
.projectCard{
  display:grid;grid-template-columns:repeat(auto-fit,minmax(190px,1fr));gap:14px;
  background:var(--paper);border:1px solid var(--rule);border-radius:6px;padding:16px;
}
.field{display:flex;flex-direction:column;gap:5px}
.field span{font-size:11px;letter-spacing:.4px;color:var(--mute)}
.field input{
  border:0;border-bottom:1px solid var(--rule);background:transparent;
  padding:5px 0;font-size:15px;
}
.field input:focus{border-bottom-color:var(--ochre);outline:none}

/* ---------- chapters ---------- */
.chapter{background:var(--paper);border:1px solid var(--rule);border-radius:6px;overflow:hidden}
.chHead{
  display:grid;grid-template-columns:auto 1fr auto auto;gap:12px;align-items:center;
  padding:11px 16px;border-bottom:1px solid var(--rule);background:#fafbfa;
}
.chNo{
  font-family:'Frank Ruhl Libre',serif;font-size:15px;font-weight:700;
  color:#fff;background:var(--ink);border-radius:3px;padding:2px 7px;
}
.chTitle{
  border:0;background:transparent;font-size:16px;font-weight:600;
  font-family:'Frank Ruhl Libre',serif;padding:2px 0;min-width:0;
}
.chTitle:focus{outline:none;border-bottom:1px solid var(--ochre)}
.chSum{font-weight:600;color:var(--ochre)}
.chTools{display:flex;gap:3px}
.chTools button{
  border:1px solid var(--rule);background:#fff;border-radius:3px;
  width:24px;height:24px;line-height:1;color:var(--slate);font-size:13px;
}
.chTools button:disabled{opacity:.3;cursor:default}
.chTools .x:hover{background:var(--red);color:#fff;border-color:var(--red)}

.items{width:100%;border-collapse:collapse;font-size:14px}
.items th{
  text-align:right;font-size:11px;font-weight:500;color:var(--mute);
  padding:7px 10px;border-bottom:1px solid var(--rule);letter-spacing:.3px;
}
.items td{padding:3px 10px;border-bottom:1px solid #eef0ed;vertical-align:middle}
.items tr:last-child td{border-bottom:0}
.items input,.items select{
  width:100%;border:1px solid transparent;background:transparent;
  border-radius:3px;padding:5px 6px;
}
.items input:hover,.items select:hover{border-color:var(--rule)}
.items input.needsUnit{border-color:#e0a99a;background:#fdf3ef}
.items input.needsUnit::placeholder{color:#b5705c}
.items input:focus,.items select:focus{border-color:var(--ochre);outline:none;background:#fff}
.cDesc{width:auto}
.cDesc{display:flex;align-items:center;gap:8px}
.itNo{font-size:11px;color:var(--mute);flex:0 0 auto;min-width:34px}
.itNo.code{min-width:78px;color:var(--ochre);font-weight:600;letter-spacing:.2px}
.cUnit{width:96px}
.cNum{width:110px;text-align:left}
.cNum input{text-align:left}
th.cNum{text-align:left}
.lineTotal{font-weight:600;padding-left:16px !important}
.cX{width:30px}
.cX button{
  border:0;background:transparent;color:var(--mute);font-size:16px;
  line-height:1;padding:2px 5px;border-radius:3px;
}
.cX button:hover{background:var(--red);color:#fff}

.chFoot{display:flex;gap:8px;padding:10px 16px;background:#fafbfa}
.add{
  border:1px dashed var(--rule);background:transparent;color:var(--slate);
  border-radius:4px;padding:6px 12px;font-size:13px;
}
.add:hover{border-color:var(--ochre);color:var(--ochre)}
.add.primary{border-style:solid;border-color:var(--ochre);color:var(--ochre);background:var(--ochre-soft)}
.addChapter{
  border:1px dashed var(--rule);background:transparent;color:var(--slate);
  border-radius:6px;padding:12px;font-size:14px;
}
.addChapter:hover{border-color:var(--ochre);color:var(--ochre);background:#fff}

/* ---------- drawer ---------- */
.scrim{
  position:fixed;inset:0;background:rgba(19,26,32,.5);
  display:flex;justify-content:center;align-items:flex-start;
  padding:40px 20px;z-index:50;overflow:auto;
}
.drawer{
  background:var(--paper);border-radius:8px;width:min(560px,100%);
  overflow:hidden;box-shadow:0 20px 50px rgba(0,0,0,.3);
}
.drawerHead{
  display:flex;justify-content:space-between;align-items:center;
  padding:14px 18px;border-bottom:2px solid var(--ink);
}
.drawerHead h2{font-family:'Frank Ruhl Libre',serif;font-size:18px;margin:0}
.drawerHead button{border:0;background:transparent;font-size:22px;line-height:1;color:var(--slate)}
.search{width:100%;border:0;border-bottom:1px solid var(--rule);padding:12px 18px;font-size:15px}
.search:focus{outline:none;border-bottom-color:var(--ochre)}
.catList{max-height:52vh;overflow:auto}
.catRow{
  display:grid;grid-template-columns:1fr auto auto;gap:12px;align-items:center;
  width:100%;text-align:right;border:0;background:transparent;
  padding:11px 18px;border-bottom:1px solid #eef0ed;font-size:14px;
}
.catRow:hover{background:var(--ochre-soft)}
.catTag{font-size:11px;color:var(--slate);background:#eef0ed;border-radius:10px;padding:2px 9px}
.catPrice{font-size:13px;color:var(--ochre);font-weight:600;min-width:110px;text-align:left}
.empty{padding:26px 18px;color:var(--mute);text-align:center;font-size:14px}

/* ---------- import ---------- */
.ghost.accent{border-color:var(--ochre);color:#f0c374}
.ghost.accent:hover{background:var(--ochre);color:#fff}

.importWrap{
  background:var(--paper);border-radius:8px;width:min(880px,100%);
  overflow:hidden;box-shadow:0 20px 50px rgba(0,0,0,.3);
}
.impTop{
  display:flex;align-items:flex-end;gap:18px;flex-wrap:wrap;
  padding:14px 18px;border-bottom:1px solid var(--rule);background:#fafbfa;
}
.impFile{font-size:13px;color:var(--slate);font-weight:600}
.impField{display:flex;flex-direction:column;gap:4px;min-width:0}
.impField span{font-size:11px;color:var(--mute)}
.impField select{
  border:1px solid var(--rule);border-radius:4px;padding:5px 8px;
  background:#fff;max-width:280px;font-size:13px;
}
.mapGrid{
  display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));
  gap:12px;padding:14px 18px;border-bottom:1px solid var(--rule);
}
.mapField{display:flex;flex-direction:column;gap:4px}
.mapField span{font-size:11px;color:var(--mute)}
.mapField b{color:var(--ochre)}
.mapField select{border:1px solid var(--rule);border-radius:4px;padding:5px 8px;background:#fff;font-size:13px}
.hierBar{
  display:flex;align-items:center;gap:14px;flex-wrap:wrap;
  padding:12px 18px;border-bottom:1px solid var(--rule);background:#fafbfa;
}
.hierLabel{font-size:11px;color:var(--mute)}
.hierPick{display:flex;gap:6px}
.hierPick button{
  border:1px solid var(--rule);background:#fff;border-radius:4px;
  padding:6px 14px;font-size:13px;color:var(--slate);
}
.hierPick button.on{border-color:var(--ochre);background:var(--ochre-soft);color:var(--ochre);font-weight:500}
.hierNote{font-size:11px;color:var(--mute);flex:1;min-width:180px}
.impStats{
  display:flex;gap:22px;padding:10px 18px;font-size:13px;
  color:var(--slate);background:var(--ochre-soft);
}
.impStats b{color:var(--ink);font-size:15px;margin-left:4px}
.impStats .warn b{color:var(--red)}
.preview{max-height:38vh;overflow:auto;border-bottom:1px solid var(--rule)}
.preview table{width:100%;border-collapse:collapse;font-size:13px}
.preview th{
  position:sticky;top:0;background:#fff;text-align:right;font-size:11px;
  color:var(--mute);font-weight:500;padding:7px 12px;border-bottom:1px solid var(--rule);
}
.preview td{padding:5px 12px;border-bottom:1px solid #f2f4f1}
.preview .pc{width:84px;color:var(--ochre);font-weight:600;font-size:12px}
.preview .pu{width:60px;color:var(--slate)}
.preview .pu.bad{color:var(--red);font-weight:600}
.unitWarn{
  margin:0;padding:11px 18px;font-size:12px;line-height:1.7;color:#7a2e1e;
  background:#fdf3ef;border-bottom:1px solid var(--rule);
}
.pskip td{color:var(--mute);font-size:12px;background:#fbfbfa}
.preview .pn{width:90px;text-align:left}
.preview th.pn{text-align:left}
.pchap td{background:#eef1ed;font-weight:600;font-family:'Frank Ruhl Libre',serif}
.more{padding:9px 12px;font-size:12px;color:var(--mute);margin:0}
.impFoot{
  display:flex;justify-content:space-between;align-items:center;gap:16px;
  padding:14px 18px;flex-wrap:wrap;
}
.impHint{margin:0;font-size:12px;color:var(--mute);max-width:420px;line-height:1.6}
.impBtns{display:flex;gap:8px}
.impBtns .add:disabled{opacity:.4;cursor:default}

/* ---------- quote ---------- */
.quoteWrap{width:min(860px,100%)}
.quoteBar{
  display:flex;gap:8px;flex-wrap:wrap;background:var(--ink);
  padding:10px 14px;border-radius:8px 8px 0 0;
}
.quoteBar input{
  border:1px solid #3a464d;background:#1c252b;color:#eef1ee;
  border-radius:4px;padding:6px 10px;font-size:13px;flex:1;min-width:120px;
}
.quoteBar button{
  border:1px solid #3a464d;background:transparent;color:#eef1ee;
  border-radius:4px;padding:6px 14px;font-size:13px;
}
.quoteBar button:hover{background:#2a353c}
.quote{background:#fff;padding:34px 38px;border-radius:0 0 8px 8px}
.qHead{display:flex;justify-content:space-between;align-items:flex-start;
  border-bottom:2px solid var(--ink);padding-bottom:14px;margin-bottom:18px}
.qHead h2{font-family:'Frank Ruhl Libre',serif;font-size:24px;margin:0}
.qHead p{margin:4px 0 0;font-size:13px;color:var(--slate)}
.qMeta{text-align:left}
.qMeta strong{display:block;font-family:'Frank Ruhl Libre',serif;font-size:17px;color:var(--ochre)}
.qMeta span{font-size:13px;color:var(--slate)}
.qProject{display:flex;gap:26px;flex-wrap:wrap;font-size:13px;margin-bottom:22px;color:var(--slate)}
.qProject p{margin:0}
.qProject b{color:var(--ink)}
.qChapter{margin-bottom:20px}
.qChapter h3{
  font-family:'Frank Ruhl Libre',serif;font-size:15px;margin:0 0 6px;
  padding-bottom:5px;border-bottom:1px solid var(--rule);
}
.qChapter h3 span{color:var(--ochre);margin-left:6px}
.qChapter table{width:100%;border-collapse:collapse;font-size:13px}
.qChapter td{padding:5px 6px;border-bottom:1px solid #f0f2ef}
.qn{color:var(--mute);width:44px;font-size:11px}
.qu{color:var(--slate);width:60px}
.qnum{text-align:left;width:86px}
.qnum.b{font-weight:600}
.qSub td{border-top:1px solid var(--rule);border-bottom:0;font-weight:600;padding-top:7px}
.qTotals{margin-top:26px;border-top:2px solid var(--ink);padding-top:12px}
.qTotals div{display:flex;justify-content:space-between;padding:5px 0;font-size:14px}
.qTotals span{color:var(--slate)}
.qGrand{border-top:1px solid var(--rule);margin-top:8px;padding-top:12px !important}
.qGrand span{font-size:16px !important;color:var(--ink) !important;font-weight:600}
.qGrand b{font-family:'Frank Ruhl Libre',serif;font-size:22px;color:var(--ochre)}
.qNote{margin-top:22px;font-size:11px;color:var(--mute);line-height:1.7}


/* ---------- shell ---------- */
.boot{padding:60px;text-align:center;color:var(--mute)}
.page{max-width:1180px;margin:0 auto;padding:22px}
.page.narrow{max-width:900px}
.pageTitle{font-family:'Frank Ruhl Libre',serif;font-size:21px;margin:0 0 18px}
.secTitle{font-family:'Frank Ruhl Libre',serif;font-size:16px;margin:28px 0 10px}

.tabs{
  display:flex;gap:2px;background:#1c252b;padding:0 22px;overflow:auto;
}
.tabs button{
  border:0;background:transparent;color:#9daaa5;padding:11px 18px;
  font-size:14px;border-bottom:3px solid transparent;white-space:nowrap;
}
.tabs button:hover{color:#eef1ee}
.tabs button.on{color:#f0c374;border-bottom-color:var(--ochre)}

/* ---------- projects list ---------- */
.listHead{display:flex;justify-content:space-between;align-items:center;gap:16px;flex-wrap:wrap;margin-bottom:20px}
.listHead h2{font-family:'Frank Ruhl Libre',serif;font-size:22px;margin:0}
.listTools{display:flex;gap:8px;flex-wrap:wrap}
.searchBox{
  border:1px solid var(--rule);border-radius:5px;padding:8px 12px;
  background:#fff;font-size:14px;min-width:240px;
}
.searchBox:focus{outline:none;border-color:var(--ochre)}
.primaryBtn{
  border:1px solid var(--ochre);background:var(--ochre);color:#fff;
  border-radius:5px;padding:8px 18px;font-size:14px;font-weight:500;
}
.primaryBtn:hover{background:#9a660a}

.cards{display:grid;grid-template-columns:repeat(auto-fill,minmax(268px,1fr));gap:14px}
.pcard{
  background:var(--paper);border:1px solid var(--rule);border-radius:6px;
  overflow:hidden;display:flex;flex-direction:column;
}
.pcard:hover{border-color:var(--ochre)}
.pcardMain{
  border:0;background:transparent;text-align:right;padding:16px;flex:1;
  display:flex;flex-direction:column;gap:4px;
}
.ptag{
  align-self:flex-start;font-size:11px;padding:2px 9px;border-radius:10px;
  background:#eef0ed;color:var(--slate);margin-bottom:4px;
}
.ptag.construction{background:#e4ecea;color:var(--green)}
.ptag.renovation{background:var(--ochre-soft);color:var(--ochre)}
.pcard h3{font-family:'Frank Ruhl Libre',serif;font-size:17px;margin:0}
.pmeta{margin:0;font-size:13px;color:var(--slate)}
.ptotal{margin:6px 0 0;font-size:19px;font-weight:600;color:var(--ochre)}
.pdate{margin:0;font-size:11px;color:var(--mute)}
.pcardTools{display:flex;border-top:1px solid var(--rule)}
.pcardTools button{
  flex:1;border:0;background:transparent;padding:8px;font-size:12px;
  color:var(--slate);border-left:1px solid var(--rule);
}
.pcardTools button:last-child{border-left:0}
.pcardTools button:hover{background:#f2f4f1}
.pcardTools .del:hover{background:var(--red);color:#fff}

.blank{
  background:var(--paper);border:1px dashed var(--rule);border-radius:8px;
  padding:52px 26px;text-align:center;
}
.blank.small{padding:34px 20px;margin-bottom:16px}
.blank h3{font-family:'Frank Ruhl Libre',serif;font-size:19px;margin:0 0 8px}
.blank p{color:var(--slate);font-size:14px;margin:0 0 18px;line-height:1.7}

.kindGrid{display:grid;grid-template-columns:1fr;gap:10px;padding:18px}
.kindCard{
  display:flex;flex-direction:column;gap:2px;align-items:flex-start;
  border:1px solid var(--rule);background:#fff;border-radius:6px;
  padding:14px 18px;text-align:right;
}
.kindCard:hover{border-color:var(--ochre);background:var(--ochre-soft)}
.kindCard b{font-family:'Frank Ruhl Libre',serif;font-size:17px}
.kindCard i{font-style:normal;font-size:13px;color:var(--slate)}
.kindCard em{font-style:normal;font-size:11px;color:var(--mute);margin-top:3px}

/* ---------- tool row ---------- */
.toolRow{
  display:flex;align-items:center;gap:8px;flex-wrap:wrap;
  background:var(--paper);border:1px solid var(--rule);border-radius:6px;padding:10px 14px;
}
.toolRow .spacer{flex:1}
.toolBtn{
  border:1px solid var(--rule);background:#fff;border-radius:4px;
  padding:7px 14px;font-size:13px;color:var(--slate);
}
.toolBtn:hover{border-color:var(--ochre);color:var(--ochre)}
.toolBtn.accent{border-color:var(--ochre);background:var(--ochre-soft);color:var(--ochre);font-weight:500}
.toolBtn.accent:hover{background:var(--ochre);color:#fff}
.kindPicker{display:flex;gap:4px}
.kindBtn{
  border:1px solid var(--rule);background:#fff;border-radius:4px;
  padding:6px 14px;font-size:13px;color:var(--slate);
}
.kindBtn.on{border-color:var(--ochre);background:var(--ochre-soft);color:var(--ochre);font-weight:500}

/* ---------- stats ---------- */
.statRow{display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:12px;margin-bottom:22px}
.stat{background:var(--paper);border:1px solid var(--rule);border-radius:6px;padding:14px 16px}
.stat span{display:block;font-size:11px;color:var(--mute);margin-bottom:5px}
.stat strong{font-family:'Frank Ruhl Libre',serif;font-size:20px;font-weight:700}
.statRow.big .stat strong{font-size:24px}
.stat.good strong{color:var(--green)}
.stat.bad strong{color:var(--red)}
.stat.warn strong{color:var(--ochre)}

.miniGrid{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:10px;margin-bottom:22px}
.mini{background:var(--paper);border:1px solid var(--rule);border-radius:5px;padding:11px 14px}
.mini span{display:block;font-size:11px;color:var(--mute);margin-bottom:3px}
.mini b{font-size:16px}
.mini.bad b{color:var(--red)}

.warnBox{
  background:#fdf3ef;border:1px solid #e7c3b8;border-right:3px solid var(--red);
  border-radius:5px;padding:12px 16px;font-size:13px;color:#7a2e1e;line-height:1.7;
}

/* ---------- budget ---------- */
.budgetGrid{background:var(--paper);border:1px solid var(--rule);border-radius:6px;overflow:hidden;margin-bottom:22px}
.bHead,.bRow{
  display:grid;grid-template-columns:1.1fr 1fr 1fr 1fr 1.6fr;gap:12px;
  align-items:center;padding:9px 16px;
}
.bHead{border-bottom:2px solid var(--ink);font-size:11px;color:var(--mute)}
.bRow{border-bottom:1px solid #eef0ed}
.bRow:last-child{border-bottom:0}
.blabel{font-size:14px;font-weight:500}
.bnum{text-align:left}
.bHead .bnum{text-align:left}
.bnum input{
  width:100%;border:1px solid transparent;background:transparent;border-radius:3px;
  padding:5px 6px;text-align:left;
}
.bnum input:hover{border-color:var(--rule)}
.bnum input:focus{border-color:var(--ochre);outline:none;background:#fff}
.bres{font-weight:600}
.bres.neg{color:var(--red)}
.bbar{position:relative;height:20px;background:#eef0ed;border-radius:10px;overflow:hidden;display:flex;align-items:center}
.bbar i{position:absolute;inset-inline-start:0;top:0;bottom:0;background:var(--green);opacity:.75}
.bbar i.over{background:var(--red)}
.bbar em{position:relative;font-style:normal;font-size:11px;padding:0 9px;color:var(--ink);font-weight:600}
.bRow.bTotal{background:#fafbfa;border-top:2px solid var(--ink);font-weight:600}
.bRow.bTotal .bnum{padding:5px 6px}

/* ---------- generic grid ---------- */
.grid{
  width:100%;border-collapse:collapse;font-size:14px;background:var(--paper);
  border:1px solid var(--rule);border-radius:6px;overflow:hidden;
}
.grid th{
  text-align:right;font-size:11px;font-weight:500;color:var(--mute);
  padding:9px 10px;border-bottom:2px solid var(--ink);background:#fafbfa;
}
.grid td{padding:3px 10px;border-bottom:1px solid #eef0ed}
.grid tr:last-child td{border-bottom:0}
.grid input:not([type=checkbox]),.grid select{
  width:100%;border:1px solid transparent;background:transparent;
  border-radius:3px;padding:6px;
}
.grid input:hover,.grid select:hover{border-color:var(--rule)}
.grid input:focus,.grid select:focus{border-color:var(--ochre);outline:none;background:#fff}
.grid.readonly td{padding:8px 10px}
.gdate{width:140px}
.gstat{width:130px}
.gnum{width:100px;text-align:left}
.grid th.gnum{text-align:left}
.gnum input{text-align:left}
.gdone{width:56px;text-align:center}
.gsup{width:150px}
.gno{width:40px;color:var(--mute)}
.grid tr.cancelled{opacity:.45}
.rowBtns{display:flex;gap:8px;align-items:center;margin-top:12px}

.pill{display:inline-block;font-size:11px;padding:3px 10px;border-radius:10px;background:#eef0ed;color:var(--slate)}
.pill.ok{background:#e0efe8;color:var(--green)}
.pill.warn{background:var(--ochre-soft);color:var(--ochre)}
.pill.bad{background:#fbe6e0;color:var(--red)}
.grid .sumRow td{border-top:2px solid var(--ink);font-weight:600;background:#fafbfa}
.grid .muteRow td{color:var(--mute);font-size:13px}


/* ---------- ביצוע ---------- */
.progressWide{
  position:relative;height:26px;background:#dfe3de;border-radius:13px;
  overflow:hidden;display:flex;align-items:center;margin-bottom:20px;
}
.progressWide i{position:absolute;inset-inline-start:0;top:0;bottom:0;background:var(--green);opacity:.8}
.progressWide em{position:relative;font-style:normal;font-size:12px;padding:0 14px;font-weight:600}

.checkRow{display:flex;align-items:center;gap:7px;font-size:13px;color:var(--slate)}
.checkRow input{width:15px;height:15px}

.chHead.exec{grid-template-columns:auto 1fr 130px auto auto}
.chTitleText{font-family:'Frank Ruhl Libre',serif;font-size:16px;font-weight:600}
.chProgress{
  position:relative;height:18px;background:#e6e9e4;border-radius:9px;
  overflow:hidden;display:flex;align-items:center;
}
.chProgress i{position:absolute;inset-inline-start:0;top:0;bottom:0;background:var(--green);opacity:.75}
.chProgress em{position:relative;font-style:normal;font-size:11px;padding:0 8px;font-weight:600}
.chHead.exec .chTools button{width:auto;padding:0 10px;font-size:11px}

.items.exec td{padding:4px 10px}
.items.exec .descText{font-size:14px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.items.exec .muted{color:var(--slate)}
.items.exec tr.over td{background:#fdf3ef}
.items.exec tr.over .lineTotal{color:var(--red)}
.cBar{width:110px}
.miniBar{display:block;height:8px;background:#e6e9e4;border-radius:4px;overflow:hidden;position:relative}
.miniBar i{position:absolute;inset-inline-start:0;top:0;bottom:0;background:var(--green);opacity:.8}
.miniBar i.over{background:var(--red)}
.allDone{padding:14px 16px;margin:0;font-size:13px;color:var(--green);background:#f4f8f6}

.barField{display:flex;align-items:center;gap:8px;color:#eef1ee;font-size:13px}
.barField input{
  border:1px solid #3a464d;background:#1c252b;color:#eef1ee;
  border-radius:4px;padding:6px 10px;font-size:13px;width:120px;text-align:left;
}


/* ---------- שינויים ---------- */
.chapter.co{border-right:3px solid var(--rule)}
.chapter.co.ok{border-right-color:var(--green)}
.chapter.co.warn{border-right-color:var(--ochre)}
.chapter.co.bad{border-right-color:var(--red)}

.coHead{
  display:grid;grid-template-columns:auto auto 1fr auto auto auto;gap:12px;
  align-items:center;padding:11px 16px;border-bottom:1px solid var(--rule);background:#fafbfa;
}
.coToggle{
  border:1px solid var(--rule);background:#fff;border-radius:3px;
  width:24px;height:24px;line-height:1;color:var(--slate);font-size:15px;
}
.coNo{
  font-family:'Frank Ruhl Libre',serif;font-size:14px;font-weight:700;
  color:#fff;background:var(--slate);border-radius:3px;padding:2px 8px;
}
.coStatus{
  border:1px solid var(--rule);border-radius:4px;padding:5px 9px;
  font-size:12px;background:#fff;color:var(--slate);
}
.coStatus.ok{border-color:var(--green);color:var(--green);background:#f0f6f3}
.coStatus.warn{border-color:var(--ochre);color:var(--ochre);background:var(--ochre-soft)}
.coStatus.bad{border-color:var(--red);color:var(--red);background:#fdf3ef}
.coHead .chTools button{width:auto;padding:0 10px;font-size:11px;height:24px}

.coMeta{
  display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));
  gap:14px;padding:14px 16px;border-bottom:1px solid var(--rule);
}

.coDocHead{margin-bottom:18px;padding-bottom:12px;border-bottom:1px solid var(--rule)}
.coDocHead h3{font-family:'Frank Ruhl Libre',serif;font-size:18px;margin:0 0 6px}
.coDocHead p{margin:2px 0;font-size:13px;color:var(--slate)}
.coDocHead b{color:var(--ink)}

.signRow{display:flex;gap:40px;margin-top:40px}
.signBox{flex:1}
.signBox span{display:block;font-size:12px;color:var(--slate);margin-bottom:34px}
.signBox i{display:block;border-bottom:1px solid var(--ink);margin-bottom:5px}
.signBox em{font-style:normal;font-size:11px;color:var(--mute)}


/* ---------- פיקוח ---------- */
.segmented{display:flex;border:1px solid var(--rule);border-radius:5px;overflow:hidden}
.segmented button{
  border:0;background:#fff;padding:7px 16px;font-size:13px;color:var(--slate);
  border-left:1px solid var(--rule);
}
.segmented button:last-child{border-left:0}
.segmented button.on{background:var(--ochre-soft);color:var(--ochre);font-weight:500}

.gloc{width:150px}
.gsev{width:110px}
.grid tr.overdue td{background:#fdf3ef}
select.sev{border-radius:4px;font-size:12px}
select.sev.bad{color:var(--red);font-weight:500}
select.sev.warn{color:var(--ochre);font-weight:500}
select.sev.ok{color:var(--green);font-weight:500}

.visit{
  background:var(--paper);border:1px solid var(--rule);border-radius:6px;
  margin-bottom:12px;overflow:hidden;
}
.visitHead{
  display:grid;grid-template-columns:150px 1fr 1fr 1.4fr auto;gap:10px;
  padding:10px 14px;border-bottom:1px solid var(--rule);background:#fafbfa;align-items:center;
}
.visitHead input{
  border:1px solid transparent;background:transparent;border-radius:3px;padding:6px 8px;font-size:13px;
}
.visitHead input:hover{border-color:var(--rule)}
.visitHead input:focus{border-color:var(--ochre);outline:none;background:#fff}
.visitDate{font-weight:600;color:var(--ink)}
.visitHead .x{
  border:1px solid var(--rule);background:#fff;border-radius:3px;
  width:24px;height:24px;line-height:1;color:var(--slate);
}
.visitHead .x:hover{background:var(--red);color:#fff;border-color:var(--red)}
.visit textarea{
  width:100%;border:0;padding:14px;font-family:inherit;font-size:14px;
  line-height:1.7;resize:vertical;min-height:88px;background:transparent;color:inherit;
}
.visit textarea:focus{outline:none;background:#fffdf8}


/* ---------- כניסה ---------- */
.authWrap{
  min-height:100vh;display:flex;align-items:center;justify-content:center;
  padding:28px 20px;background:var(--concrete);
}
.authCard{
  width:min(420px,100%);background:var(--paper);
  border:1px solid var(--rule);border-radius:8px;overflow:hidden;
  box-shadow:0 6px 28px rgba(19,26,32,.09);
}
.authBrand{
  display:flex;align-items:center;gap:12px;
  padding:24px;background:var(--ink);color:#eef1ee;
}
.authBrand h1{font-family:'Frank Ruhl Libre',serif;font-size:21px;margin:0;letter-spacing:.3px}
.authBrand p{margin:2px 0 0;font-size:12px;color:#9daaa5}

.authTabs{display:flex;border-bottom:1px solid var(--rule)}
.authTabs button{
  flex:1;border:0;background:transparent;padding:13px;font-size:14px;
  color:var(--slate);border-bottom:2px solid transparent;
}
.authTabs button.on{color:var(--ochre);border-bottom-color:var(--ochre);font-weight:500}

.authForm{padding:22px 24px 26px;display:flex;flex-direction:column;gap:14px}
.authField{display:flex;flex-direction:column;gap:5px}
.authField span{font-size:11px;letter-spacing:.4px;color:var(--mute)}
.authField input{
  border:1px solid var(--rule);border-radius:5px;padding:9px 11px;
  font-size:15px;background:#fff;
}
.authField input:focus{outline:none;border-color:var(--ochre)}
.authField input[dir=ltr]{text-align:left}

.authBtn{
  border:1px solid var(--ochre);background:var(--ochre);color:#fff;
  border-radius:5px;padding:11px;font-size:15px;font-weight:500;margin-top:4px;
}
.authBtn:hover{background:#9a660a}
.authBtn:disabled{opacity:.6;cursor:default}
.authErr{
  margin:0;font-size:13px;color:#7a2e1e;background:#fdf3ef;
  border:1px solid #e7c3b8;border-radius:5px;padding:9px 12px;line-height:1.6;
}
.authNote{margin:2px 0 0;font-size:11px;color:var(--mute);line-height:1.7;text-align:center}

/* ---------- משתמש וצוות ---------- */
.userChip{display:flex;flex-direction:column;line-height:1.25;margin-inline-end:4px}
.uname{font-size:13px;color:#eef1ee}
.urole{font-size:11px;color:#9daaa5}

.drawer.wide{width:min(720px,100%)}
.drawer .grid{border:0;border-radius:0}
.ltr{direction:ltr;text-align:right}
.teamAdd{padding:18px 20px 22px;border-top:1px solid var(--rule);background:#fafbfa}
.teamAdd h3{font-family:'Frank Ruhl Libre',serif;font-size:15px;margin:0 0 12px}
.teamGrid{
  display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));
  gap:12px;margin-bottom:14px;
}
.teamGrid select{
  border:0;border-bottom:1px solid var(--rule);background:transparent;
  padding:5px 0;font-size:15px;
}
.teamGrid select:focus{outline:none;border-bottom-color:var(--ochre)}
.teamNote{margin:12px 0 0;font-size:11px;color:var(--mute);line-height:1.7}


/* ---------- הרשאות ---------- */
.permsBody{padding:18px 20px 22px}
.permsTop{display:flex;gap:18px;align-items:flex-end;flex-wrap:wrap;margin-bottom:18px}
.permsTop .field{min-width:180px}
.permsTop select{
  border:0;border-bottom:1px solid var(--rule);background:transparent;
  padding:5px 0;font-size:15px;
}
.permsTop select:focus{outline:none;border-bottom-color:var(--ochre)}
.permsHint{margin:0;font-size:12px;color:var(--mute);flex:1;min-width:220px;line-height:1.7}
.permsHint b{color:var(--ochre)}

.permsGrid{width:100%;border-collapse:collapse;font-size:14px;margin-bottom:20px}
.permsGrid th{
  text-align:right;font-size:11px;font-weight:500;color:var(--mute);
  padding:8px 10px;border-bottom:2px solid var(--ink);
}
.permsGrid th.plev,.permsGrid td.plev{width:74px;text-align:center}
.permsGrid td{padding:7px 10px;border-bottom:1px solid #eef0ed}
.permsGrid input[type=radio]{width:16px;height:16px;accent-color:var(--ochre)}

.flagList{
  display:grid;grid-template-columns:repeat(auto-fit,minmax(210px,1fr));
  gap:10px;padding-top:16px;border-top:1px solid var(--rule);
}
.flagRow{display:flex;align-items:flex-start;gap:9px;font-size:14px}
.flagRow input{width:16px;height:16px;margin-top:2px;accent-color:var(--ochre)}
.flagRow span{display:flex;flex-direction:column;line-height:1.4}
.flagRow i{font-style:normal;font-size:11px;color:var(--mute);margin-top:1px}
.permsFoot{margin-top:22px}

.youTag{
  font-style:normal;font-size:10px;color:var(--ochre);background:var(--ochre-soft);
  border-radius:8px;padding:1px 7px;margin-right:7px;
}
.rowTools{display:flex;gap:5px;width:auto;white-space:nowrap}
.rowTools button{
  border:1px solid var(--rule);background:#fff;border-radius:3px;
  padding:4px 10px;font-size:12px;color:var(--slate);
}
.rowTools button:hover{border-color:var(--ochre);color:var(--ochre)}
.rowTools .del:hover{background:var(--red);color:#fff;border-color:var(--red)}


.authInfo{
  margin:0;font-size:13px;color:#1a5c46;background:#eef7f2;
  border:1px solid #b8ddc9;border-radius:5px;padding:9px 12px;line-height:1.6;
}

@media print{
  .topbar,.tabs,.tape,.quoteBar,.chFoot,.addChapter,.toolRow,.projectCard{display:none !important}
  .scrim{position:static;padding:0;background:none;overflow:visible}
  .quote{padding:0;border-radius:0}
  .quoteWrap{width:100%}
}
@media (max-width:960px){
  .layout{flex-direction:column}
  .tape{position:static;width:100%;flex:1}
  .projectCard{grid-template-columns:1fr}
}
@media (prefers-reduced-motion:reduce){
  .boq *{transition:none !important;animation:none !important}
}
`;

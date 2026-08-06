/* ============================================================
   supabaseClient.js — שכבת הנתונים
   ------------------------------------------------------------
   כל הקריאות למסד עוברות דרך כאן. שום קומפוננטה לא מדברת
   ישירות עם Supabase, כדי שיהיה מקום אחד לתקן בו.

   הערה על אבטחה: המפתח כאן הוא publishable ומיועד להיות
   גלוי. מה שמגן על הנתונים הוא ה-RLS במסד, לא המפתח.
   ============================================================ */

import { createClient } from "@supabase/supabase-js";

export const SUPABASE_URL = "https://selxcmchrdaiiubvhcpf.supabase.co";
export const SUPABASE_KEY = "sb_publishable_hITPN7iQ8dQQm6ajjYDEoQ_0A_sQuX-";

export const sb = createClient(SUPABASE_URL, SUPABASE_KEY, {
  auth: { persistSession: true, autoRefreshToken: true },
});

/* ---------- עזרים ---------- */

const fail = (error, what) => {
  if (!error) return;
  console.error(what, error);
  throw new Error(hebrewError(error, what));
};

/* הודעות השרת מגיעות באנגלית. מתרגמים את השכיחות. */
function hebrewError(error, what) {
  const m = String(error.message || "");
  if (m.includes("Invalid login credentials")) return "המייל או הסיסמה אינם נכונים.";
  if (m.includes("User already registered")) return "כתובת המייל כבר רשומה.";
  if (m.includes("Password should be")) return "הסיסמה קצרה מדי — 8 תווים לפחות.";
  if (m.includes("Email not confirmed")) return "יש לאשר את המייל לפני הכניסה.";
  if (m.includes("row-level security") || m.includes("violates row-level"))
    return "אין לך הרשאה לפעולה הזאת.";
  if (m.includes("אין הרשאה")) return m;  /* הודעות מהטריגרים שלנו */
  if (m.includes("Failed to fetch")) return "אין חיבור לשרת. בדוק את האינטרנט.";
  return `${what} נכשל: ${m}`;
}

/* ============================================================
   התחברות
   ============================================================ */

export const authApi = {
  async currentSession() {
    const { data } = await sb.auth.getSession();
    return data.session || null;
  },

  onChange(cb) {
    const { data } = sb.auth.onAuthStateChange((_e, session) => cb(session));
    return () => data.subscription.unsubscribe();
  },

  async signIn(email, password) {
    const { data, error } = await sb.auth.signInWithPassword({
      email: email.trim().toLowerCase(),
      password,
    });
    fail(error, "כניסה");
    return data.session;
  },

  /* הרשמה של הבעלים: נוצר משתמש, ומיד אחריו ארגון.
     bootstrap_org רץ בשרת ומגדיר אותו כמנהל מערכת. */
  async signUpOwner({ email, password, fullName, orgName }) {
    const { data, error } = await sb.auth.signUp({
      email: email.trim().toLowerCase(),
      password,
    });
    fail(error, "הרשמה");

    if (!data.session) {
      /* הפרויקט דורש אימות מייל */
      return { needsEmailConfirm: true };
    }

    const { error: bootErr } = await sb.rpc("bootstrap_org", {
      org_name: orgName || fullName,
      full_name: fullName,
    });
    fail(bootErr, "יצירת ארגון");
    return { session: data.session };
  },

  async signOut() {
    await sb.auth.signOut();
  },

  async me() {
    const { data, error } = await sb
      .from("members")
      .select("id, org_id, full_name, role, perms, active")
      .single();
    if (error) return null;
    return data;
  },

  async myOrg() {
    const { data } = await sb.from("orgs").select("*").single();
    return data || null;
  },

  async updateOrg(fields) {
    const { error } = await sb.from("orgs").update(fields).neq("id", "");
    fail(error, "עדכון פרטי החברה");
  },
};

/* ============================================================
   צוות
   ============================================================ */

export const teamApi = {
  async list() {
    const { data, error } = await sb
      .from("members")
      .select("id, full_name, role, perms, active")
      .order("full_name");
    fail(error, "טעינת הצוות");
    return data || [];
  },

  async update(id, fields) {
    const { error } = await sb.from("members").update(fields).eq("id", id);
    fail(error, "עדכון הרשאות");
  },

  async deactivate(id) {
    const { error } = await sb.from("members").update({ active: false }).eq("id", id);
    fail(error, "הסרת משתמש");
  },
};

/* ============================================================
   פרויקטים
   ============================================================ */

const PROJECT_COLS =
  "id, name, description, client_name, address, start_date, end_date, " +
  "site_manager_name, site_manager_phone, kind, overhead, created_at, updated_at";

export const projectsApi = {
  async list() {
    const { data, error } = await sb
      .from("projects")
      .select(PROJECT_COLS)
      .eq("archived", false)
      .order("updated_at", { ascending: false });
    fail(error, "טעינת הפרויקטים");
    return data || [];
  },

  async create(orgId, fields) {
    const { data, error } = await sb
      .from("projects")
      .insert({ org_id: orgId, ...fields })
      .select(PROJECT_COLS)
      .single();
    fail(error, "יצירת פרויקט");
    return data;
  },

  async update(id, fields) {
    const { error } = await sb.from("projects").update(fields).eq("id", id);
    fail(error, "שמירת הפרויקט");
  },

  async remove(id) {
    const { error } = await sb.from("projects").delete().eq("id", id);
    fail(error, "מחיקת הפרויקט");
  },

  /* טעינת פרויקט שלם. כל שאילתה עשויה לחזור ריקה אם אין
     הרשאה — זו התנהגות תקינה, לא שגיאה. */
  async load(id) {
    const [proj, chapters, items, budget, milestones, orders, changes, changeItems, visits, defects] =
      await Promise.all([
        sb.from("projects").select(PROJECT_COLS).eq("id", id).single(),
        sb.from("chapters").select("*").eq("project_id", id).order("position"),
        sb.from("items_safe").select("*").eq("project_id", id).order("position"),
        sb.from("budgets").select("*").eq("project_id", id).maybeSingle(),
        sb.from("milestones").select("*").eq("project_id", id).order("position"),
        sb.from("purchase_orders").select("*").eq("project_id", id).order("created_at"),
        sb.from("change_orders").select("*").eq("project_id", id).order("number"),
        sb.from("change_items").select("*"),
        sb.from("site_visits").select("*").eq("project_id", id).order("visit_date", { ascending: false }),
        sb.from("defects").select("*").eq("project_id", id).order("due_date"),
      ]);

    fail(proj.error, "טעינת הפרויקט");

    const byChapter = {};
    (items.data || []).forEach((it) => {
      (byChapter[it.chapter_id] = byChapter[it.chapter_id] || []).push({
        id: it.id,
        code: it.code || "",
        desc: it.description || "",
        unit: it.unit || "",
        qty: it.qty ?? 0,
        price: it.unit_price ?? 0,
        done: it.done_qty ?? 0,
      });
    });

    const changeIds = new Set((changes.data || []).map((c) => c.id));
    const ciByChange = {};
    (changeItems.data || [])
      .filter((ci) => changeIds.has(ci.change_id))
      .forEach((ci) => {
        (ciByChange[ci.change_id] = ciByChange[ci.change_id] || []).push({
          id: ci.id,
          desc: ci.description || "",
          unit: ci.unit || "",
          qty: ci.qty ?? 0,
          price: ci.unit_price ?? 0,
        });
      });

    return {
      ...proj.data,
      chapters: (chapters.data || []).map((c) => ({
        id: c.id,
        title: c.title || "",
        items: byChapter[c.id] || [],
      })),
      budget: budget.data || null,
      milestones: (milestones.data || []).map((m) => ({
        id: m.id,
        name: m.name || "",
        planned_date: m.planned_date || "",
        actual_date: m.actual_date || "",
        done: !!m.done,
      })),
      orders: (orders.data || []).map((o) => ({
        id: o.id,
        supplier: o.supplier || "",
        description: o.description || "",
        quantity: o.qty ?? 0,
        unit_price: o.unit_price ?? 0,
        status: o.status,
        order_date: o.order_date || "",
        received_date: o.received_date || "",
      })),
      changes: (changes.data || []).map((c) => ({
        id: c.id,
        number: c.number,
        title: c.title || "",
        reason: c.reason || "",
        requested_by: c.requested_by || "",
        date: c.co_date || "",
        status: c.status,
        items: ciByChange[c.id] || [],
      })),
      visits: (visits.data || []).map((v) => ({
        id: v.id,
        date: v.visit_date || "",
        weather: v.weather || "",
        workers: v.workers || "",
        subcontractors: v.subcontractors || "",
        notes: v.notes || "",
      })),
      defects: (defects.data || []).map((d) => ({
        id: d.id,
        title: d.title || "",
        location: d.location || "",
        responsible: d.responsible || "",
        severity: d.severity,
        status: d.status,
        opened: d.opened_at || "",
        due: d.due_date || "",
        closed: d.closed_at || "",
      })),
    };
  },
};

/* ============================================================
   כתב כמויות
   ============================================================ */

export const boqApi = {
  async addChapter(projectId, title, position) {
    const { data, error } = await sb
      .from("chapters")
      .insert({ project_id: projectId, title, position })
      .select()
      .single();
    fail(error, "הוספת פרק");
    return data;
  },

  async updateChapter(id, fields) {
    const { error } = await sb.from("chapters").update(fields).eq("id", id);
    fail(error, "עדכון פרק");
  },

  async removeChapter(id) {
    const { error } = await sb.from("chapters").delete().eq("id", id);
    fail(error, "מחיקת פרק");
  },

  async reorderChapters(pairs) {
    await Promise.all(
      pairs.map(({ id, position }) =>
        sb.from("chapters").update({ position }).eq("id", id)
      )
    );
  },

  async addItem(projectId, chapterId, item, position) {
    const { data, error } = await sb
      .from("items")
      .insert({
        project_id: projectId,
        chapter_id: chapterId,
        code: item.code || null,
        description: item.desc || "",
        unit: item.unit || null,
        qty: +item.qty || 0,
        unit_price: +item.price || 0,
        position,
      })
      .select()
      .single();
    fail(error, "הוספת סעיף");
    return data;
  },

  async updateItem(id, fields) {
    const payload = {};
    if ("desc" in fields) payload.description = fields.desc;
    if ("code" in fields) payload.code = fields.code || null;
    if ("unit" in fields) payload.unit = fields.unit || null;
    if ("qty" in fields) payload.qty = +fields.qty || 0;
    if ("price" in fields) payload.unit_price = +fields.price || 0;
    if ("done" in fields) payload.done_qty = +fields.done || 0;

    const { error } = await sb.from("items").update(payload).eq("id", id);
    fail(error, "שמירת הסעיף");
  },

  async removeItem(id) {
    const { error } = await sb.from("items").delete().eq("id", id);
    fail(error, "מחיקת סעיף");
  },

  /* ייבוא מאקסל: פרקים וסעיפים בבת אחת */
  async importChapters(projectId, chapters, replace) {
    if (replace) {
      const { error } = await sb.from("chapters").delete().eq("project_id", projectId);
      fail(error, "ניקוי כתב הכמויות");
    }

    for (let i = 0; i < chapters.length; i++) {
      const ch = await boqApi.addChapter(projectId, chapters[i].title, i);
      const rows = chapters[i].items.map((it, k) => ({
        project_id: projectId,
        chapter_id: ch.id,
        code: it.code || null,
        description: it.desc || "",
        unit: it.unit || null,
        qty: +it.qty || 0,
        unit_price: +it.price || 0,
        position: k,
      }));
      if (rows.length) {
        const { error } = await sb.from("items").insert(rows);
        fail(error, "ייבוא סעיפים");
      }
    }
  },
};

/* ============================================================
   שאר המודולים
   ============================================================ */

const simpleTable = (table, mapIn, label) => ({
  async add(projectId, row) {
    const { data, error } = await sb
      .from(table)
      .insert({ project_id: projectId, ...mapIn(row) })
      .select()
      .single();
    fail(error, `הוספת ${label}`);
    return data;
  },
  async update(id, row) {
    const { error } = await sb.from(table).update(mapIn(row)).eq("id", id);
    fail(error, `עדכון ${label}`);
  },
  async remove(id) {
    const { error } = await sb.from(table).delete().eq("id", id);
    fail(error, `מחיקת ${label}`);
  },
});

export const budgetApi = {
  async save(projectId, b) {
    const { error } = await sb.from("budgets").upsert(
      {
        project_id: projectId,
        planned_materials: +b.planned_materials || 0,
        planned_labor: +b.planned_labor || 0,
        planned_subcontractors: +b.planned_subcontractors || 0,
        actual_materials: +b.actual_materials || 0,
        actual_labor: +b.actual_labor || 0,
        actual_subcontractors: +b.actual_subcontractors || 0,
        updated_at: new Date().toISOString(),
      },
      { onConflict: "project_id" }
    );
    fail(error, "שמירת התקציב");
  },
};

export const milestonesApi = simpleTable(
  "milestones",
  (m) => ({
    name: m.name || "",
    planned_date: m.planned_date || null,
    actual_date: m.actual_date || null,
    done: !!m.done,
  }),
  "אבן דרך"
);

export const ordersApi = simpleTable(
  "purchase_orders",
  (o) => ({
    supplier: o.supplier || null,
    description: o.description || null,
    qty: +o.quantity || 0,
    unit_price: +o.unit_price || 0,
    status: o.status,
    order_date: o.order_date || null,
    received_date: o.received_date || null,
  }),
  "הזמנה"
);

export const visitsApi = simpleTable(
  "site_visits",
  (v) => ({
    visit_date: v.date || null,
    weather: v.weather || null,
    workers: v.workers || null,
    subcontractors: v.subcontractors || null,
    notes: v.notes || null,
  }),
  "ביקור"
);

export const defectsApi = simpleTable(
  "defects",
  (d) => ({
    title: d.title || "",
    location: d.location || null,
    responsible: d.responsible || null,
    severity: d.severity,
    status: d.status,
    opened_at: d.opened || null,
    due_date: d.due || null,
    closed_at: d.closed || null,
  }),
  "ליקוי"
);

/* שינויים: כותרת וסעיפים בטבלאות נפרדות */
export const changesApi = {
  async add(projectId, number) {
    const { data, error } = await sb
      .from("change_orders")
      .insert({ project_id: projectId, number, title: "" })
      .select()
      .single();
    fail(error, "הוספת שינוי");
    return data;
  },
  async update(id, co) {
    const { error } = await sb
      .from("change_orders")
      .update({
        title: co.title || "",
        reason: co.reason || null,
        requested_by: co.requested_by || null,
        co_date: co.date || null,
        status: co.status,
      })
      .eq("id", id);
    fail(error, "עדכון השינוי");
  },
  async remove(id) {
    const { error } = await sb.from("change_orders").delete().eq("id", id);
    fail(error, "מחיקת השינוי");
  },
  async addItem(changeId, item, position) {
    const { data, error } = await sb
      .from("change_items")
      .insert({
        change_id: changeId,
        description: item.desc || "",
        unit: item.unit || null,
        qty: +item.qty || 0,
        unit_price: +item.price || 0,
        position,
      })
      .select()
      .single();
    fail(error, "הוספת סעיף לשינוי");
    return data;
  },
  async updateItem(id, item) {
    const { error } = await sb
      .from("change_items")
      .update({
        description: item.desc || "",
        unit: item.unit || null,
        qty: +item.qty || 0,
        unit_price: +item.price || 0,
      })
      .eq("id", id);
    fail(error, "עדכון סעיף");
  },
  async removeItem(id) {
    const { error } = await sb.from("change_items").delete().eq("id", id);
    fail(error, "מחיקת סעיף");
  },
};

/* ============================================================
   סנכרון פרויקט שלם
   ------------------------------------------------------------
   האפליקציה מחזיקה את הפרויקט בזיכרון ושומרת אותו כמכלול.
   כאן ממירים את זה לפעולות מסד: מעדכנים מה שהשתנה, מוסיפים
   מה שנוצר, ומוחקים מה שנמחק — בלי למחוק ולבנות מחדש, כדי
   שמזהים לא ישתנו ושהיסטוריה לא תאבד.
   ============================================================ */

export const newId = () =>
  typeof crypto !== "undefined" && crypto.randomUUID
    ? crypto.randomUUID()
    : "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx".replace(/[xy]/g, (c) => {
        const r = (Math.random() * 16) | 0;
        return (c === "x" ? r : (r & 0x3) | 0x8).toString(16);
      });

const diffIds = (before, after) => {
  const now = new Set(after.map((x) => x.id));
  return before.filter((id) => !now.has(id));
};

export async function saveProject(project, snapshot, perms) {
  const id = project.id;
  const may = (mod) => perms?.modules?.[mod] === "edit";

  /* ---------- פרטי הפרויקט ---------- */
  if (may("boq")) {
    await projectsApi.update(id, {
      name: project.name || "",
      description: project.description || null,
      client_name: project.client_name || null,
      address: project.address || null,
      start_date: project.start_date || null,
      end_date: project.end_date || null,
      site_manager_name: project.site_manager_name || null,
      site_manager_phone: project.site_manager_phone || null,
      kind: project.kind,
      overhead: +project.overhead || 0,
    });
  }

  /* ---------- פרקים וסעיפים ---------- */
  const canBoq = may("boq");
  const canExec = may("execution");

  if (canBoq || canExec) {
    const chapterRows = [];
    const itemRows = [];

    project.chapters.forEach((c, ci) => {
      chapterRows.push({ id: c.id, project_id: id, title: c.title || "", position: ci });
      c.items.forEach((it, ii) => {
        itemRows.push({
          id: it.id,
          project_id: id,
          chapter_id: c.id,
          code: it.code || null,
          description: it.desc || "",
          unit: it.unit || null,
          qty: +it.qty || 0,
          unit_price: +it.price || 0,
          done_qty: +it.done || 0,
          position: ii,
        });
      });
    });

    if (canBoq) {
      if (chapterRows.length) {
        const { error } = await sb.from("chapters").upsert(chapterRows);
        fail(error, "שמירת הפרקים");
      }
      const goneCh = diffIds(snapshot.chapterIds || [], project.chapters);
      if (goneCh.length) await sb.from("chapters").delete().in("id", goneCh);

      if (itemRows.length) {
        const { error } = await sb.from("items").upsert(itemRows);
        fail(error, "שמירת הסעיפים");
      }
      const allItems = project.chapters.flatMap((c) => c.items);
      const goneIt = diffIds(snapshot.itemIds || [], allItems);
      if (goneIt.length) await sb.from("items").delete().in("id", goneIt);
    } else {
      /* מי שרשאי רק ביצוע מעדכן כמות מבוצעת בלבד.
         הטריגר בשרת חוסם כל שינוי אחר, גם אם ננסה. */
      for (const it of project.chapters.flatMap((c) => c.items)) {
        const was = (snapshot.done || {})[it.id];
        if (was !== undefined && +was === (+it.done || 0)) continue;
        const { error } = await sb
          .from("items")
          .update({ done_qty: +it.done || 0 })
          .eq("id", it.id);
        fail(error, "עדכון כמות מבוצעת");
      }
    }
  }

  /* ---------- תקציב ---------- */
  if (may("budget") && project.budget) {
    await budgetApi.save(id, project.budget);
  }

  /* ---------- אבני דרך ---------- */
  if (may("schedule")) {
    const rows = (project.milestones || []).map((m, i) => ({
      id: m.id,
      project_id: id,
      name: m.name || "",
      planned_date: m.planned_date || null,
      actual_date: m.actual_date || null,
      done: !!m.done,
      position: i,
    }));
    if (rows.length) {
      const { error } = await sb.from("milestones").upsert(rows);
      fail(error, "שמירת אבני הדרך");
    }
    const gone = diffIds(snapshot.milestoneIds || [], project.milestones || []);
    if (gone.length) await sb.from("milestones").delete().in("id", gone);
  }

  /* ---------- רכש ---------- */
  if (may("procurement")) {
    const rows = (project.orders || []).map((o) => ({
      id: o.id,
      project_id: id,
      supplier: o.supplier || null,
      description: o.description || null,
      qty: +o.quantity || 0,
      unit_price: +o.unit_price || 0,
      status: o.status,
      order_date: o.order_date || null,
      received_date: o.received_date || null,
    }));
    if (rows.length) {
      const { error } = await sb.from("purchase_orders").upsert(rows);
      fail(error, "שמירת הרכש");
    }
    const gone = diffIds(snapshot.orderIds || [], project.orders || []);
    if (gone.length) await sb.from("purchase_orders").delete().in("id", gone);
  }

  /* ---------- פיקוח ---------- */
  if (may("site")) {
    const vRows = (project.visits || []).map((v) => ({
      id: v.id,
      project_id: id,
      visit_date: v.date || null,
      weather: v.weather || null,
      workers: v.workers || null,
      subcontractors: v.subcontractors || null,
      notes: v.notes || null,
    }));
    if (vRows.length) {
      const { error } = await sb.from("site_visits").upsert(vRows);
      fail(error, "שמירת היומן");
    }
    const goneV = diffIds(snapshot.visitIds || [], project.visits || []);
    if (goneV.length) await sb.from("site_visits").delete().in("id", goneV);

    const dRows = (project.defects || []).map((d) => ({
      id: d.id,
      project_id: id,
      title: d.title || "",
      location: d.location || null,
      responsible: d.responsible || null,
      severity: d.severity,
      status: d.status,
      opened_at: d.opened || null,
      due_date: d.due || null,
      closed_at: d.closed || null,
    }));
    if (dRows.length) {
      const { error } = await sb.from("defects").upsert(dRows);
      fail(error, "שמירת הליקויים");
    }
    const goneD = diffIds(snapshot.defectIds || [], project.defects || []);
    if (goneD.length) await sb.from("defects").delete().in("id", goneD);
  }

  /* ---------- שינויים ---------- */
  if (may("changes")) {
    const coRows = (project.changes || []).map((c) => ({
      id: c.id,
      project_id: id,
      number: c.number,
      title: c.title || "",
      reason: c.reason || null,
      requested_by: c.requested_by || null,
      co_date: c.date || null,
      status: c.status,
    }));
    if (coRows.length) {
      const { error } = await sb.from("change_orders").upsert(coRows);
      fail(error, "שמירת השינויים");
    }
    const goneCo = diffIds(snapshot.changeIds || [], project.changes || []);
    if (goneCo.length) await sb.from("change_orders").delete().in("id", goneCo);

    const ciRows = [];
    (project.changes || []).forEach((c) =>
      (c.items || []).forEach((it, i) =>
        ciRows.push({
          id: it.id,
          change_id: c.id,
          description: it.desc || "",
          unit: it.unit || null,
          qty: +it.qty || 0,
          unit_price: +it.price || 0,
          position: i,
        })
      )
    );
    if (ciRows.length) {
      const { error } = await sb.from("change_items").upsert(ciRows);
      fail(error, "שמירת סעיפי שינוי");
    }
    const allCi = (project.changes || []).flatMap((c) => c.items || []);
    const goneCi = diffIds(snapshot.changeItemIds || [], allCi);
    if (goneCi.length) await sb.from("change_items").delete().in("id", goneCi);
  }
}

/* צילום מצב לפני עריכה — משמש כדי לדעת מה נמחק */
export function snapshotOf(project) {
  return {
    chapterIds: project.chapters.map((c) => c.id),
    itemIds: project.chapters.flatMap((c) => c.items.map((i) => i.id)),
    milestoneIds: (project.milestones || []).map((m) => m.id),
    orderIds: (project.orders || []).map((o) => o.id),
    visitIds: (project.visits || []).map((v) => v.id),
    defectIds: (project.defects || []).map((d) => d.id),
    changeIds: (project.changes || []).map((c) => c.id),
    changeItemIds: (project.changes || []).flatMap((c) => (c.items || []).map((i) => i.id)),
    done: Object.fromEntries(
      project.chapters.flatMap((c) => c.items.map((i) => [i.id, +i.done || 0]))
    ),
  };
}

/**
 * منصة الخدمات الرقمية - إدارة الاتصال المؤسسي
 * Google Apps Script (Web App) - يستقبل الطلبات والتقييمات
 *
 * الإعداد:
 * 1) ضع معرف الشيت في SHEET_ID (من الرابط بين /d/ و /edit)
 * 2) انشر: نشر ← نشر جديد ← تطبيق ويب ← الوصول: Anyone
 * 3) انسخ رابط /exec وضعه في config.js داخل formEndpoint
 */

const SHEET_ID = "PASTE_SHEET_ID_HERE";
const ADMIN_EMAIL = "src.prc@sr.edu.sa";
const SENDER_NAME = "إدارة الاتصال المؤسسي | Corporate Communication";
const LOGO_URL = "https://malmosalm.github.io/sru-comms-platform/logo-sru.png";

const REQUEST_HEADERS = [
  "الرقم المرجعي", "تاريخ الإرسال", "نوع الخدمة", "مقدم الطلب", "الإدارة / الكلية",
  "البريد", "الجوال", "تاريخ الحاجة", "الموقع", "الأولوية", "المكلّف بالمتابعة",
  "عنوان الطلب", "التفاصيل", "المرفقات", "التكلفة التقديرية", "حالة الموافقة", "حالة الطلب",
  "وقت البداية", "وقت النهاية"
];
const FEEDBACK_HEADERS = ["الرقم المرجعي", "تاريخ التقييم", "التقييم", "الملاحظات", "البريد"];
const TEAM_SHEET = "الفريق";
const APPROVAL_THRESHOLD = 100; // الحد الذي فوقه يُرسل إشعار للمعتمد (ريال). الاعتماد غير إلزامي: لا يوقف سير الطلب
// ساعات الدوام المسموحة للمواعيد (تُعدَّل هنا فقط)
const WORK_START = "08:00";
const WORK_END = "19:00";
const RATING_LABELS = { satisfied: "راضٍ", neutral: "محايد", dissatisfied: "غير راضٍ" };

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.type === "feedback") {
      saveFeedback(d);
    } else {
      saveRequest(d);
    }
    return json({ ok: true });
  } catch (err) {
    console.error(err);
    return json({ ok: false, error: String(err) });
  }
}

function doGet(e) {
  const action = e && e.parameter ? e.parameter.action : "";
  if (action === "status") return json(getStatus(e.parameter.ref));
  if (action === "check") return json(checkSlot(e.parameter));
  return json({ ok: true, service: "sru-comms-platform" });
}

// حالة الطلب للعامة: لا تُرجع الاسم أو البريد أو التفاصيل
function getStatus(ref) {
  const clean = String(ref || "").trim().toUpperCase();
  if (!/^SRU-\d{8}-\d{6}-\d{3}$/.test(clean)) return { ok: false, error: "bad_format" };
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName("الطلبات");
  if (!sheet) return { ok: false, error: "not_found" };
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).toUpperCase() === clean) {
      return {
        ok: true,
        ref: clean,
        service: rows[i][2],
        date: Utilities.formatDate(new Date(rows[i][1]), "Asia/Riyadh", "yyyy-MM-dd"),
        assignee: rows[i][10] || "",
        approval: rows[i][15] || "",
        status: rows[i][16] || ""
      };
    }
  }
  return { ok: false, error: "not_found" };
}

// تعارض الموعد: نفس اليوم والموقع، وتداخل في الساعات، وطلب غير مغلق ولم يُرفض
function normLoc(v) { return String(v || "").trim().replace(/\s+/g, " ").toLowerCase(); }
function findConflicts(date, location, start, end, excludeRef) {
  if (!date || !location || !start || !end) return [];
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName("الطلبات");
  if (!sheet) return [];
  const rows = sheet.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    const rDate = r[7] instanceof Date ? Utilities.formatDate(r[7], "Asia/Riyadh", "yyyy-MM-dd") : String(r[7]);
    if (rDate !== date || normLoc(r[8]) !== normLoc(location)) continue;
    if (String(r[16]) === "مغلق" || String(r[15]) === "مرفوض") continue;
    if (excludeRef && String(r[0]) === excludeRef) continue;
    const rs = String(r[18] || ""), re = String(r[19] || "");
    if (!rs || !re) continue;
    if (start < re && rs < end) out.push({ ref: String(r[0]), start: rs, end: re });
  }
  return out;
}
// للواجهة: يُرجع الحالة فقط، بدون أي بيانات شخصية
function outsideHours(start, end) {
  return !!(start && end) && (start < WORK_START || end > WORK_END);
}
function checkSlot(p) {
  const list = findConflicts(p.date, p.location, p.start, p.end, "");
  return { ok: true, conflict: list.length > 0, outside: outsideHours(p.start, p.end) };
}

function saveRequest(d) {
  const id = d.ref || ("SRU-" + Utilities.formatDate(new Date(), "Asia/Riyadh", "yyyyMMdd-HHmmss") + "-" + Math.floor(Math.random() * 900 + 100));
  const sheet = getSheet("الطلبات", REQUEST_HEADERS);
  const clash = findConflicts(d.needed_date, d.location, d.start_time, d.end_time, id).length > 0;
  const outside = outsideHours(d.start_time, d.end_time);
  sheet.appendRow([
    id, new Date(), d.service || "", d.requester || "", d.department || "",
    d.email || "", d.phone || "", d.needed_date || "", d.location || "", d.priority || "",
    d.assignee || "", d.title || "", d.details || "", d.attachments || "",
    d.est_cost || "", "معلّق", outside ? "خارج أوقات الدوام" : (clash ? "تعارض موعد" : "جديد"),
    d.start_time || "", d.end_time || ""
  ]);
  if (clash || outside) {
    MailApp.sendEmail({ to: ADMIN_EMAIL, subject: `${outside ? "موعد خارج الدوام" : "تعارض موعد"} للطلب ${id} | ${outside ? "Outside working hours" : "Schedule conflict"} ${id}`,
      htmlBody: buildHtml("يوجد طلب آخر في نفس الموقع والوقت، ويحتاج قراراً.", "Another request overlaps this location and time.",
        [["الرقم المرجعي", id], ["الموقع", d.location], ["التاريخ", d.needed_date], ["الوقت", (d.start_time || "") + " - " + (d.end_time || "")]]),
      name: SENDER_NAME });
  }

  const rows = [
    ["الرقم المرجعي", id], ["نوع الخدمة", d.service], ["عنوان الطلب", d.title],
    ["تاريخ الحاجة", d.needed_date], ["الموقع", d.location || "-"], ["الأولوية", d.priority]
  ];
  if (d.email) {
    MailApp.sendEmail({
      to: d.email,
      subject: `تأكيد استلام الطلب ${id} | Request received ${id}`,
      htmlBody: buildHtml(
        "شكراً لك، تم استلام طلبك وسنتواصل معك قريباً. احتفظ بالرقم المرجعي لمتابعة الطلب.",
        "Thank you, your request has been received. Keep the reference number to track it.",
        rows
      ),
      name: SENDER_NAME
    });
  }
  MailApp.sendEmail({
    to: ADMIN_EMAIL,
    subject: `طلب جديد ${id} | New request ${id}`,
    htmlBody: buildHtml(
      `طلب جديد من ${escapeHtml(d.requester || "-")} - ${escapeHtml(d.department || "-")}`,
      `New request from ${escapeHtml(d.requester || "-")} - ${escapeHtml(d.department || "-")}`,
      rows.concat([["مقدم الطلب", d.requester], ["البريد", d.email], ["الجوال", d.phone || "-"], ["التفاصيل", d.details]])
    ),
    name: SENDER_NAME
  });
}

function saveFeedback(d) {
  const sheet = getSheet("التقييمات", FEEDBACK_HEADERS);
  sheet.appendRow([d.ref || "", new Date(), RATING_LABELS[d.rating] || d.rating || "", d.comment || "", d.email || ""]);
  MailApp.sendEmail({
    to: ADMIN_EMAIL,
    subject: `تقييم خدمة ${d.ref || ""} | Service feedback ${d.ref || ""}`,
    htmlBody: buildHtml(
      `تقييم جديد: ${escapeHtml(RATING_LABELS[d.rating] || "-")}`,
      `New feedback: ${escapeHtml(d.rating || "-")}`,
      [["الرقم المرجعي", d.ref], ["التقييم", RATING_LABELS[d.rating] || d.rating], ["الملاحظات", d.comment || "-"]]
    ),
    name: SENDER_NAME
  });
}

function getSheet(name, headers) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function buildHtml(introAr, introEn, rows) {
  const body = rows.map(r =>
    `<tr><td style="padding:8px 10px;border-bottom:1px solid #e5e0ee;color:#675b72;width:38%">${escapeHtml(String(r[0]))}</td>` +
    `<td style="padding:8px 10px;border-bottom:1px solid #e5e0ee;font-weight:700">${escapeHtml(String(r[1] || "-"))}</td></tr>`
  ).join("");
  return `<div dir="rtl" style="font-family:Arial,sans-serif;max-width:620px;margin:auto;color:#241331">
    <div style="background:#3a1464;padding:18px;text-align:center"><img src="${LOGO_URL}" alt="SRU" style="height:56px"></div>
    <div style="padding:20px">
      <p style="font-size:15px;line-height:1.8">${escapeHtml(introAr)}</p>
      <p dir="ltr" style="font-size:14px;color:#675b72;text-align:left">${escapeHtml(introEn)}</p>
      <table style="width:100%;border-collapse:collapse;font-size:14px;margin-top:12px">${body}</table>
    </div>
    <div style="background:#f7f5fb;padding:12px;text-align:center;font-size:12px;color:#675b72">
      إدارة الاتصال المؤسسي · جامعة سليمان الراجحي · Corporate Communication Administration
    </div>
  </div>`;
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}


// ==================== الإعداد لمرة واحدة ====================
// شغّل setupAll من المحرر مرة واحدة فقط.
function setupAll() {
  const req = getSheet("الطلبات", REQUEST_HEADERS);
  const team = getSheet(TEAM_SHEET, ["الاسم", "البريد", "الدور"]);
  const n = 998;
  const teamNames = team.getRange("A2:A200");
  req.getRange(2, 11, n, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInRange(teamNames, true).build());
  req.getRange(2, 16, n, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(["معلّق", "معتمد", "مرفوض"], true).build());
  req.getRange(2, 17, n, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(["جديد", "مُكلَّف", "بانتظار الاعتماد", "قيد التنفيذ", "مُعاد للتعديل", "مغلق"], true).build());
  installTriggers();
}

// ينشئ المحفّزات: عند التعديل، وتذكير يومي. يحذف القديمة أولاً.
function installTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => {
    const f = t.getHandlerFunction();
    if (f === "onEditInstalled" || f === "dailyReminders") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("onEditInstalled").forSpreadsheet(SpreadsheetApp.openById(SHEET_ID)).onEdit().create();
  ScriptApp.newTrigger("dailyReminders").timeBased().everyDays(1).atHour(9).create();
}

// ==================== الإشعارات عند التعديل ====================
// الأعمدة: K=11 المكلّف، O=15 التكلفة، P=16 الموافقة، Q=17 الحالة
function onEditInstalled(e) {
  const sh = e.range.getSheet();
  const row = e.range.getRow();
  const col = e.range.getColumn();
  if (sh.getName() !== "الطلبات" || row < 2) return;
  const v = sh.getRange(row, 1, 1, 17).getValues()[0];
  const ref = v[0];

  if (col === 11 && v[10]) {
    const m = findTeam(v[10]);
    if (m) sendMail(m.email, `تكليف بالطلب ${ref}`, `تم تكليفك بالطلب ${ref}: ${v[11]}. موعد الحاجة: ${fmtDate(v[7])}.`);
    if (v[16] === "جديد") sh.getRange(row, 17).setValue("مُكلَّف");
  }
  if (col === 15 && Number(v[14]) > APPROVAL_THRESHOLD) {
    sh.getRange(row, 16).setValue("معلّق");
    notifyRole("معتمد", `طلب يحتاج اعتماد ${ref}`, `الطلب ${ref} تكلفته التقديرية ${v[14]} ريال، ويحتاج اعتمادك.`);
  }
  if (col === 16 && (v[15] === "معتمد" || v[15] === "مرفوض")) {
    notifyRequester(v, `قرار الاعتماد للطلب ${ref}`, `تم ${v[15]} الطلب ${ref}.`);
  }
  if (col === 17 && v[16] === "مُعاد للتعديل") {
    const m = findTeam(v[10]);
    if (m) sendMail(m.email, `إعادة الطلب ${ref} للتعديل`, `أُعيد الطلب ${ref} (${v[11]}) للتعديل. راجع ملاحظات المعتمد وأعد الإرسال.`);
    notifyRequester(v, `تحديث على طلبك ${ref}`, `يتم حالياً مراجعة طلبك ${ref}، وسنوافيك عند الانتهاء.`);
  }
  if (col === 17 && v[16] === "مغلق") {
    const link = `https://malmosalm.github.io/sru-comms-platform/feedback.html?ref=${encodeURIComponent(ref)}`;
    notifyRequester(v, `تم إغلاق الطلب ${ref}`, `تم إغلاق طلبك ${ref}. نرجو تقييم الخدمة: <a href="${link}">تقييم الخدمة</a>`);
  }
}

// ==================== التذكير اليومي ====================
// يرسل للإدارة قائمة الطلبات المفتوحة التي تبقى عليها يوم أو أقل، أو تجاوزت موعدها.
function dailyReminders() {
  const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName("الطلبات");
  if (!sh || sh.getLastRow() < 2) return;
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 17).getValues();
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const late = [];
  rows.forEach(v => {
    if (!v[0] || v[16] === "مغلق" || !v[7]) return;
    const need = new Date(v[7]);
    const days = Math.round((need - today) / 86400000);
    if (days <= 1) late.push(`${v[0]} | ${v[2]} | ${fmtDate(v[7])} | ${v[10] || "غير مكلّف"} | ${v[16]}`);
  });
  if (!late.length) return;
  MailApp.sendEmail({
    to: ADMIN_EMAIL,
    subject: `تذكير: ${late.length} طلب قريب من موعده أو متأخر`,
    htmlBody: `<div dir="rtl" style="font-family:Arial"><p>الطلبات التالية تحتاج متابعة اليوم:</p><ul>${late.map(l => `<li>${escapeHtml(l)}</li>`).join("")}</ul></div>`,
    name: SENDER_NAME
  });
}

// ==================== أدوات مساعدة ====================
function findTeam(name) {
  const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TEAM_SHEET);
  if (!sh || sh.getLastRow() < 2) return null;
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues();
  for (const r of rows) if (r[0] === name) return { name: r[0], email: r[1], role: r[2] };
  return null;
}

function notifyRole(role, subject, body) {
  const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TEAM_SHEET);
  if (!sh || sh.getLastRow() < 2) return;
  sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues()
    .filter(r => r[2] === role && r[1])
    .forEach(r => sendMail(r[1], subject, body));
}

function notifyRequester(v, subject, body) {
  if (v[5]) sendMail(v[5], subject, body);
}

function sendMail(to, subject, body) {
  MailApp.sendEmail({ to: to, subject: subject, htmlBody: `<div dir="rtl" style="font-family:Arial">${body}</div>`, name: SENDER_NAME });
}

function fmtDate(d) {
  if (!d) return "-";
  return Utilities.formatDate(new Date(d), "Asia/Riyadh", "yyyy-MM-dd");
}

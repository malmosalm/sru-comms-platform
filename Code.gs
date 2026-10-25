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
  "عنوان الطلب", "التفاصيل", "المرفقات", "التكلفة التقديرية", "حالة الموافقة", "حالة الطلب"
];
const FEEDBACK_HEADERS = ["الرقم المرجعي", "تاريخ التقييم", "التقييم", "الملاحظات", "البريد"];
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

function saveRequest(d) {
  const id = d.ref || ("SRU-" + Utilities.formatDate(new Date(), "Asia/Riyadh", "yyyyMMdd-HHmmss") + "-" + Math.floor(Math.random() * 900 + 100));
  const sheet = getSheet("الطلبات", REQUEST_HEADERS);
  sheet.appendRow([
    id, new Date(), d.service || "", d.requester || "", d.department || "",
    d.email || "", d.phone || "", d.needed_date || "", d.location || "", d.priority || "",
    d.assignee || "", d.title || "", d.details || "", d.attachments || "",
    d.est_cost || "", "معلّق", "جديد"
  ]);

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

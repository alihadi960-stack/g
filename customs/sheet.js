// كتابة بلوك شهادة جمركية في شيت "شيك مطالبة" بنفس معادلات الشيت الأصلي.
// يشتغل مع ExcelJS في المتصفح وفي Node (مفيش import هنا، بياخد الـ worksheet جاهز).

export const SHEET_NAME = "شيك مطالبة";
export const VAT_RATE = 0.14;   // K = I * 14%
export const ATS_RATE = 0.01;   // L = D * F * 1% (أرباح تجارية / أ.ت.ص)

const NUM_FMT = '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)';
const PCT_FMT = "0%";
const FONT = { name: "Aptos Narrow", size: 11 };
const THIN = { style: "thin" };
const BOX = { top: THIN, left: THIN, bottom: THIN, right: THIN };
const HEADER_FILL = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD9E1F2" } };

export const HEADERS = [
  "رقم الشهادة ", "قيمه الفاتوره", "العملة", "سعر التحويل", "القيمة النهائية - بالشهادة",
  "قيمة الفاتورة عملة ", "البند الجمركى ", "نسبة البند الجمركى ", "وعاء الضريبة",
  "ضريبة الوارد ", "القيمة المضافة ", "ارباح تجارية ", "قيمة الشيك المحسوبة", "قيمة المطالبة",
];
const HEADER_ROW = 3;
const WIDTHS = [9.6, 11.4, 6.3, 9.3, 18.6, 15.6, 11, 14.3, 14.3, 14.6, 14.6, 12.7, 17.1, 14.6, 11.3];

export function getOrCreateSheet(wb) {
  let ws = wb.getWorksheet(SHEET_NAME) || wb.worksheets[0];
  if (!ws) {
    ws = wb.addWorksheet(SHEET_NAME, { views: [{ rightToLeft: true }] });
  }
  if (!String(ws.getCell(HEADER_ROW, 1).value || "").trim()) writeHeader(ws);
  return ws;
}

function writeHeader(ws) {
  WIDTHS.forEach((w, i) => { ws.getColumn(i + 1).width = w; });
  const row = ws.getRow(HEADER_ROW);
  HEADERS.forEach((h, i) => {
    const c = row.getCell(i + 1);
    c.value = h;
    c.font = { ...FONT, bold: true };
    c.fill = HEADER_FILL;
    c.border = BOX;
    c.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  });
  row.height = 32;
}

// آخر صف فيه أي قيمة في الأعمدة A..O
function lastUsedRow(ws) {
  let last = HEADER_ROW;
  ws.eachRow({ includeEmpty: false }, (row, n) => {
    for (let col = 1; col <= 15; col++) {
      const v = row.getCell(col).value;
      if (v !== null && v !== undefined && v !== "") { last = Math.max(last, n); break; }
    }
  });
  return last;
}

function asNumberIfDigits(s) {
  const t = String(s ?? "").replace(/\s/g, "");
  return /^\d{1,15}$/.test(t) ? Number(t) : t;
}

// رقم الشهادة زي ما في الشيت: آخر جزء من رقم الإفراج (2026-614-1-453456 → 453456)
export function shortCertNumber(full) {
  const m = String(full ?? "").match(/(\d+)\D*$/);
  return m ? Number(m[1]) : full;
}

/**
 * data = {
 *   certificate_number, currency, invoice_value, final_value,
 *   items: [{ hs_code, duty_rate_percent, invoice_value }],
 *   claim_total   // قيمة المطالبة (اختياري)
 * }
 * بيرجع { startRow, totalRow }
 */
export function appendClaimBlock(wb, data) {
  const items = (data.items || []).filter(it => Number(it.invoice_value) > 0);
  if (!items.length) throw new Error("مفيش بنود في الفاتورة");

  const ws = getOrCreateSheet(wb);
  const start = lastUsedRow(ws) + (lastUsedRow(ws) === HEADER_ROW ? 1 : 2);
  const end = start + items.length - 1;
  const total = end + 1;
  const currency = data.currency || "دولار";

  items.forEach((it, i) => {
    const r = start + i;
    const row = ws.getRow(r);
    const set = (col, value, fmt, center) => {
      const c = row.getCell(col);
      c.value = value;
      c.font = FONT;
      c.border = BOX;
      if (fmt) c.numFmt = fmt;
      if (center) c.alignment = { horizontal: "center", vertical: "middle" };
    };
    if (i === 0) {
      set(1, shortCertNumber(data.certificate_number), null, true);
      set(2, Number(data.invoice_value), NUM_FMT, true);
      set(4, { formula: `+E${start}/B${start}` }, NUM_FMT, true);
      set(5, Number(data.final_value), NUM_FMT, true);
    } else {
      set(1, null, null, true);
      set(2, null, NUM_FMT, true);
      // نفس سعر التحويل الفعلي بتاع أول سطر بدل رقم ثابت
      set(4, { formula: `+$D$${start}` }, NUM_FMT, true);
      set(5, null, NUM_FMT, true);
    }
    set(3, currency, null, true);
    set(6, Number(it.invoice_value), NUM_FMT);
    set(7, asNumberIfDigits(it.hs_code), null, true);
    set(8, Number(it.duty_rate_percent) / 100, PCT_FMT, true);
    set(9, { formula: `+(D${r}*F${r})+J${r}` }, NUM_FMT);
    set(10, { formula: `+F${r}*D${r}*H${r}` }, NUM_FMT);
    set(11, { formula: `+I${r}*${VAT_RATE}` }, NUM_FMT);
    set(12, { formula: `+D${r}*F${r}*${ATS_RATE}` }, NUM_FMT);
    set(13, { formula: `+L${r}+K${r}+J${r}` }, NUM_FMT);
    set(14, null, NUM_FMT);
  });

  if (items.length > 1) ws.mergeCells(start, 1, end, 1);

  const tr = ws.getRow(total);
  const bold = { ...FONT, size: 12, bold: true };
  ["J", "K", "L", "M"].forEach((L, i) => {
    const c = tr.getCell(10 + i);
    c.value = { formula: `SUM(${L}${start}:${L}${end})` };
    c.numFmt = NUM_FMT;
    c.font = bold;
  });
  const n = tr.getCell(14);
  n.value = data.claim_total != null && data.claim_total !== "" ? Number(data.claim_total) : null;
  n.numFmt = NUM_FMT;
  n.font = bold;
  const o = tr.getCell(15);
  o.value = { formula: `+M${total}-N${total}` };
  o.numFmt = NUM_FMT;
  o.font = FONT;

  // Excel يعيد حساب كل المعادلات أول ما الملف يتفتح
  wb.calcProperties = { ...(wb.calcProperties || {}), fullCalcOnLoad: true };
  return { startRow: start, totalRow: total };
}

// نفس معادلات الشيت، للمعاينة قبل الإضافة
export function previewTotals(data) {
  const rate = Number(data.final_value) / Number(data.invoice_value) || 0;
  const rows = (data.items || []).map(it => {
    const base = rate * Number(it.invoice_value || 0);
    const duty = base * Number(it.duty_rate_percent || 0) / 100;
    const vat = (base + duty) * VAT_RATE;
    const ats = base * ATS_RATE;
    return { duty, vat, ats, total: duty + vat + ats };
  });
  const sum = k => rows.reduce((s, r) => s + r[k], 0);
  const total = sum("total");
  const claim = data.claim_total != null && data.claim_total !== "" ? Number(data.claim_total) : null;
  return {
    rate, rows,
    duty: sum("duty"), vat: sum("vat"), ats: sum("ats"), total,
    diff: claim == null ? null : total - claim,
  };
}

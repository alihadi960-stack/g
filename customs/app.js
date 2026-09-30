// شيك المطالبة الجمركية: صورة الشهادة → Claude يستخرج البيانات → بلوك جديد في الشيت بنفس المعادلات
import Anthropic from "https://cdn.jsdelivr.net/npm/@anthropic-ai/sdk@0.130.0/+esm";
import { appendClaimBlock, getOrCreateSheet, previewTotals, SHEET_NAME } from "./sheet.js";

const MODEL = "claude-opus-5-5";
const STORE_KEY = "customs_claude_api_key";
const MAX_IMG_EDGE = 2200; // تصغير صور الموبايل الكبيرة قبل الإرسال

const $ = (id) => document.getElementById(id);
const fmt = (n) => (n == null || isNaN(n)) ? "—" :
  Number(n).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const esc = (s) => String(s ?? "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const num = (v) => {
  const t = String(v ?? "").replace(/[,\s]/g, "").replace(/[٠-٩]/g, d => "٠١٢٣٤٥٦٧٨٩".indexOf(d));
  return t === "" ? null : Number(t);
};

let files = [];
let workbook = null;
let workbookName = "شيك مطالبة.xlsx";
let current = null;      // البيانات اللي بتتراجع دلوقتي
let addedCount = 0;

function toast(msg, bad = false) {
  const el = $("toast");
  el.textContent = msg;
  el.className = "toast" + (bad ? " bad" : "");
  el.style.display = "block";
  clearTimeout(toast._t);
  toast._t = setTimeout(() => el.style.display = "none", 3500);
}

// ---------------- Workbook ----------------
function setWbStatus() {
  if (!workbook) { $("wbStatus").textContent = "لم يتم اختيار ملف."; return; }
  const ws = getOrCreateSheet(workbook);
  $("wbStatus").innerHTML = `📄 <b>${esc(workbookName)}</b> — شيت "${esc(ws.name)}"` +
    (addedCount ? ` — اتضاف ${addedCount} شهادة في الجلسة دي` : "");
  $("btnDownload").disabled = !addedCount;
}

async function loadWorkbook(file) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(await file.arrayBuffer());
  workbook = wb;
  workbookName = file.name;
  addedCount = 0;
  setWbStatus();
  toast("تم فتح الملف ✅");
}

function newWorkbook() {
  workbook = new ExcelJS.Workbook();
  workbook.addWorksheet(SHEET_NAME, { views: [{ rightToLeft: true }] });
  getOrCreateSheet(workbook);
  workbookName = "شيك مطالبة.xlsx";
  addedCount = 0;
  setWbStatus();
}

async function downloadWorkbook() {
  const buf = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = workbookName;
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 5000);
}

// ---------------- Files ----------------
function renderFiles() {
  $("fileList").innerHTML = files.map(f => `<span class="chip">${f.type === "application/pdf" ? "📄" : "🖼"} ${esc(f.name)}</span>`).join("");
  $("btnExtract").disabled = !files.length;
}

function addFiles(list) {
  for (const f of list) {
    if (f.type.startsWith("image/") || f.type === "application/pdf") files.push(f);
  }
  renderFiles();
}

function readAsBase64(blob) {
  return new Promise((resolve, reject) => {
    const r = new FileReader();
    r.onload = () => resolve(String(r.result).split(",")[1]);
    r.onerror = reject;
    r.readAsDataURL(blob);
  });
}

async function imageToJpegBase64(file) {
  const bmp = await createImageBitmap(file);
  const scale = Math.min(1, MAX_IMG_EDGE / Math.max(bmp.width, bmp.height));
  const canvas = document.createElement("canvas");
  canvas.width = Math.round(bmp.width * scale);
  canvas.height = Math.round(bmp.height * scale);
  canvas.getContext("2d").drawImage(bmp, 0, 0, canvas.width, canvas.height);
  const blob = await new Promise(res => canvas.toBlob(res, "image/jpeg", 0.9));
  return readAsBase64(blob);
}

async function fileToBlock(file) {
  if (file.type === "application/pdf") {
    return { type: "document", source: { type: "base64", media_type: "application/pdf", data: await readAsBase64(file) } };
  }
  return { type: "image", source: { type: "base64", media_type: "image/jpeg", data: await imageToJpegBase64(file) } };
}

// ---------------- Claude extraction ----------------
const nullable = (type) => ({ anyOf: [{ type }, { type: "null" }] });

const SCHEMA = {
  type: "object",
  additionalProperties: false,
  required: ["certificate_number", "currency", "exchange_rate", "invoice_value", "final_value",
    "items", "items_taxes_total", "claim_total", "warnings"],
  properties: {
    certificate_number: nullable("string"),
    currency: nullable("string"),
    exchange_rate: nullable("number"),
    invoice_value: nullable("number"),
    final_value: nullable("number"),
    items: {
      type: "array",
      items: {
        type: "object",
        additionalProperties: false,
        required: ["hs_code", "duty_rate_percent", "invoice_value", "description", "duty_amount", "vat_amount", "ats_amount"],
        properties: {
          hs_code: { type: "string" },
          duty_rate_percent: { type: "number" },
          invoice_value: { type: "number" },
          description: nullable("string"),
          duty_amount: nullable("number"),
          vat_amount: nullable("number"),
          ats_amount: nullable("number"),
        },
      },
    },
    items_taxes_total: nullable("number"),
    claim_total: nullable("number"),
    warnings: { type: "array", items: { type: "string" } },
  },
};

const SYSTEM = `You extract data from Egyptian customs documents (منظومة نافذة / Nafeza) for a claim-check spreadsheet.
The user uploads photos or PDFs of one shipment: usually the customs certificate review form ("نموذج مراجعة بيانات شهادة جمركية", possibly several pages) and optionally the payment claim ("تفاصيل المطالبة").

Fill the fields like this:
- certificate_number: field 4 "رقم الافراج", digits and dashes only, e.g. "2026-614-1-453456" (drop letters like "م ن").
- currency: field 24 "العملة" in short Arabic, e.g. "دولار", "يورو", "يوان".
- exchange_rate: field 25 "سعر التحويل".
- invoice_value: field 28 "قيمة الفاتورة" in the invoice currency.
- final_value: field 30 "القيمة النهائية" in EGP (keep all decimals).
- items: one entry per line of section 32 "سطور الفاتورة", in order. hs_code = "البند الجمركي" (10 digits), duty_rate_percent = "قيمة الوارد" percentage as a number (10.000 % -> 10), invoice_value = "قيمة الفاتورة" of that line in the invoice currency. Do not confuse the quantity (قطعة) or weight (كيلو جرام) columns with the invoice value. Descriptions may appear under the line.
  From section 33 "حساب الضرائب و الرسوم" (same line numbers) copy duty_amount = "ضريبه الوارد", vat_amount = "ضريبه قيمه مضافه", ats_amount = "ضريبه أ.ت.ص". Use null when not visible.
- items_taxes_total: field 34 "اجمالي ضرائب البنود".
- claim_total: the grand total "الاجمالي" of the claim document (تفاصيل المطالبة). null if no claim document was uploaded.
- warnings: short Arabic notes about anything unreadable, cut off, missing pages, or values you are unsure about. Empty if everything is clear.

Numbers: output plain numbers (no thousands separators). Convert Arabic-Indic digits. Never guess a value you cannot read — use null and add a warning.`;

async function extract() {
  const apiKey = localStorage.getItem(STORE_KEY);
  if (!apiKey) { openSettings(); return toast("حط الـ API key الأول", true); }

  $("btnExtract").disabled = true;
  $("extractMsg").textContent = "⏳ جاري قراءة الملفات…";
  try {
    const blocks = [];
    for (const f of files) blocks.push(await fileToBlock(f));

    const client = new Anthropic({ apiKey, dangerouslyAllowBrowser: true });
    const response = await client.beta.messages.create({
      model: MODEL,
      max_tokens: 16000,
      betas: ["server-side-fallback-2026-07-01"],
      fallbacks: "default",
      output_config: { effort: "high", format: { type: "json_schema", schema: SCHEMA } },
      system: SYSTEM,
      messages: [{
        role: "user",
        content: [...blocks, { type: "text", text: "استخرج بيانات الشهادة والمطالبة من الملفات دي." }],
      }],
    });

    if (response.stop_reason === "refusal") throw new Error("الطلب اترفض من الموديل، جرب صورة أوضح.");
    if (response.stop_reason === "max_tokens") throw new Error("الرد اتقطع، جرب عدد صفحات أقل.");
    const text = response.content.filter(b => b.type === "text").map(b => b.text).join("");
    const data = JSON.parse(text);

    current = {
      certificate_number: data.certificate_number || "",
      currency: data.currency || "دولار",
      invoice_value: data.invoice_value,
      final_value: data.final_value,
      claim_total: data.claim_total,
      items_taxes_total: data.items_taxes_total,
      items: data.items || [],
      warnings: data.warnings || [],
    };
    $("extractMsg").textContent = "✅ تم الاستخراج — راجع البيانات تحت.";
    renderReview();
  } catch (e) {
    let msg = e.message || String(e);
    if (e instanceof Anthropic.AuthenticationError) msg = "الـ API key غلط";
    else if (e instanceof Anthropic.RateLimitError) msg = "ضغط على الخدمة، استنى شوية وجرب تاني";
    else if (e instanceof Anthropic.APIConnectionError) msg = "مفيش اتصال بالإنترنت أو الخدمة مش متاحة";
    $("extractMsg").textContent = "❌ " + msg;
    toast(msg, true);
  } finally {
    $("btnExtract").disabled = !files.length;
  }
}

// ---------------- Review ----------------
function readForm() {
  current.certificate_number = $("fCert").value.trim();
  current.currency = $("fCur").value.trim();
  current.invoice_value = num($("fInv").value);
  current.final_value = num($("fFinal").value);
  current.claim_total = num($("fClaim").value);
  current.items = [...$("itemsBody").querySelectorAll("tr")].map((tr, i) => ({
    ...current.items[i],
    hs_code: tr.querySelector("[data-k=hs]").value.trim(),
    duty_rate_percent: num(tr.querySelector("[data-k=rate]").value) ?? 0,
    invoice_value: num(tr.querySelector("[data-k=val]").value) ?? 0,
  }));
}

// مقارنة الرقم المحسوب بالرقم المكتوب في الشهادة (فرق أقل من 1 جنيه = مطابق)
function match(calc, cert) {
  if (cert == null) return "";
  return Math.abs(calc - cert) < 1 ? '<span class="ok">✓</span>' : `<span class="bad">✗ ${fmt(cert)}</span>`;
}

function renderReview() {
  $("secReview").style.display = "block";
  $("fCert").value = current.certificate_number ?? "";
  $("fCur").value = current.currency ?? "";
  $("fInv").value = current.invoice_value ?? "";
  $("fFinal").value = current.final_value ?? "";
  $("fClaim").value = current.claim_total ?? "";
  $("itemsBody").innerHTML = current.items.map((it, i) => `
    <tr>
      <td>${i + 1}</td>
      <td><input data-k="hs" value="${esc(it.hs_code)}" inputmode="numeric" style="min-width:120px" /></td>
      <td><input data-k="rate" value="${esc(it.duty_rate_percent)}" inputmode="decimal" style="min-width:60px" /></td>
      <td><input data-k="val" value="${esc(it.invoice_value)}" inputmode="decimal" /></td>
      <td class="num" data-c="duty"></td><td class="num" data-c="vat"></td>
      <td class="num" data-c="ats"></td><td class="num" data-c="total"></td>
      <td data-c="match"></td>
      <td><button class="btn" data-del="${i}" title="حذف">🗑</button></td>
    </tr>`).join("");
  $("warnings").innerHTML = (current.warnings || []).map(w => `<li>⚠ ${esc(w)}</li>`).join("");
  recalc();
  $("secReview").scrollIntoView({ behavior: "smooth" });
}

function recalc() {
  readForm();
  const p = previewTotals(current);
  $("fRate").value = p.rate ? p.rate.toFixed(4) : "";
  const trs = $("itemsBody").querySelectorAll("tr");
  p.rows.forEach((r, i) => {
    const it = current.items[i];
    const tr = trs[i];
    tr.querySelector("[data-c=duty]").textContent = fmt(r.duty);
    tr.querySelector("[data-c=vat]").textContent = fmt(r.vat);
    tr.querySelector("[data-c=ats]").textContent = fmt(r.ats);
    tr.querySelector("[data-c=total]").textContent = fmt(r.total);
    tr.querySelector("[data-c=match]").innerHTML =
      [match(r.duty, it.duty_amount), match(r.vat, it.vat_amount), match(r.ats, it.ats_amount)].filter(Boolean).join(" ");
  });
  $("itemsFoot").innerHTML = `
    <tr>
      <th colspan="4">الإجمالي</th>
      <th>${fmt(p.duty)}</th><th>${fmt(p.vat)}</th><th>${fmt(p.ats)}</th><th>${fmt(p.total)}</th>
      <th colspan="2">${match(p.total, current.items_taxes_total)}</th>
    </tr>`;

  const invSum = current.items.reduce((s, it) => s + (Number(it.invoice_value) || 0), 0);
  const checks = [];
  if (current.invoice_value != null) {
    checks.push(Math.abs(invSum - current.invoice_value) < 0.05
      ? `<span class="ok">✓ مجموع البنود = قيمة الفاتورة (${fmt(invSum)})</span>`
      : `<span class="bad">✗ مجموع البنود ${fmt(invSum)} ≠ قيمة الفاتورة ${fmt(current.invoice_value)}</span>`);
  }
  if (p.diff != null) checks.push(`الفرق بين الشيك المحسوب والمطالبة: <b>${fmt(p.diff)}</b>`);
  $("checks").innerHTML = checks.join("<br>");
}

function addItem() {
  readForm();
  current.items.push({ hs_code: "", duty_rate_percent: 0, invoice_value: 0 });
  renderReview();
}

function appendToWorkbook() {
  if (!workbook) newWorkbook();
  readForm();
  if (!current.invoice_value || !current.final_value) return toast("قيمة الفاتورة والقيمة النهائية مطلوبين", true);
  try {
    const { startRow, totalRow } = appendClaimBlock(workbook, current);
    addedCount++;
    setWbStatus();
    toast(`✅ اتضافت في الصفوف ${startRow}–${totalRow}. دوس "تحميل" لما تخلص.`);
  } catch (e) {
    toast(e.message || String(e), true);
  }
}

// ---------------- Settings ----------------
function openSettings() {
  $("apiKeyInput").value = localStorage.getItem(STORE_KEY) || "";
  $("ov").style.display = "flex";
}
function saveSettings() {
  localStorage.setItem(STORE_KEY, $("apiKeyInput").value.trim());
  $("ov").style.display = "none";
  toast("تم الحفظ ✅");
}

// ---------------- Wiring ----------------
document.addEventListener("DOMContentLoaded", () => {
  $("btnSettings").onclick = openSettings;
  $("btnCloseSettings").onclick = () => $("ov").style.display = "none";
  $("btnSaveSettings").onclick = saveSettings;

  $("xlsxInput").onchange = (e) => e.target.files[0] && loadWorkbook(e.target.files[0]).catch(err => toast("مش قادر أفتح الملف: " + err.message, true));
  $("btnNewWb").onclick = () => { newWorkbook(); toast("ملف جديد ✅"); };
  $("btnDownload").onclick = downloadWorkbook;

  $("imgInput").onchange = (e) => { addFiles(e.target.files); e.target.value = ""; };
  const drop = $("drop");
  drop.ondragover = (e) => { e.preventDefault(); drop.classList.add("over"); };
  drop.ondragleave = () => drop.classList.remove("over");
  drop.ondrop = (e) => { e.preventDefault(); drop.classList.remove("over"); addFiles(e.dataTransfer.files); };
  $("btnClear").onclick = () => { files = []; renderFiles(); $("extractMsg").textContent = ""; };
  $("btnExtract").onclick = extract;

  $("itemsBody").addEventListener("input", recalc);
  $("itemsBody").addEventListener("click", (e) => {
    const i = e.target.closest("[data-del]")?.dataset.del;
    if (i == null) return;
    readForm();
    current.items.splice(Number(i), 1);
    renderReview();
  });
  ["fInv", "fFinal", "fClaim"].forEach(id => $(id).addEventListener("input", recalc));
  $("btnAddItem").onclick = addItem;
  $("btnAppend").onclick = appendToWorkbook;

  if (!localStorage.getItem(STORE_KEY)) openSettings();
});

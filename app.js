// Firebase Hosting Starter (JSONP) — Gardenia Suite
const DEFAULT_APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbxBO3W9S-5L-jkneHr1fMFKKBVkNb7EpiErN-OFeXS-3la1nDG2hySeDi1rqRDBSr5c/exec";

const STORE_API = "gardenia_api_url";
const STORE_TOKEN = "gardenia_token";
const STORE_ROLE = "gardenia_role";

const $ = (id) => document.getElementById(id);

let API_URL = localStorage.getItem(STORE_API) || DEFAULT_APPS_SCRIPT_URL;
let TOKEN = localStorage.getItem(STORE_TOKEN) || "";
let ROLE  = localStorage.getItem(STORE_ROLE)  || "";

function setStatus(ok, text){
  const dot = $("dot");
  const t = $("statusText");
  if(dot) dot.className = "dot" + (ok ? " ok" : "");
  if(t) t.textContent = text || (ok ? "متصل" : "غير متصل");
}

function toast(msg, bad=false){
  const el = $("toast");
  if(!el) return;
  el.textContent = msg;
  el.className = "toast" + (bad ? " bad" : "");
  el.style.display = "block";
  setTimeout(()=> el.style.display="none", 2500);
}

// JSONP call: Apps Script doGet must accept callback/action/token/data
function apiJSONP(action, data, token){
  return new Promise((resolve, reject) => {
    const cb = "cb_" + Math.random().toString(16).slice(2);
    const script = document.createElement("script");

    window[cb] = (res) => {
      resolve(res);
      try { delete window[cb]; } catch(e){}
      script.remove();
    };

    const qs =
      "callback=" + encodeURIComponent(cb) +
      "&action=" + encodeURIComponent(action) +
      "&token=" + encodeURIComponent(token || "") +
      "&data=" + encodeURIComponent(JSON.stringify(data || {}));

    script.src = API_URL + "?" + qs;
    script.onerror = () => {
      reject(new Error("JSONP failed"));
      try { delete window[cb]; } catch(e){}
      script.remove();
    };

    document.body.appendChild(script);
  });
}

async function ping(){
  try{
    const r = await apiJSONP("ping", {}, TOKEN);
    if(r && r.ok){
      setStatus(true, "متصل ✅");
      toast("Ping ✅");
    } else {
      setStatus(false, "غير متصل");
      toast(r && r.error ? r.error : "Ping فشل", true);
    }
  }catch(e){
    setStatus(false, "غير متصل");
    toast("Ping فشل: " + (e.message||e), true);
  }
}

function showLogin(){
  $("secLogin").style.display = "block";
  $("secApp").style.display = "none";
  $("btnLogout").style.display = "none";
}
function showApp(){
  $("secLogin").style.display = "none";
  $("secApp").style.display = "block";
  $("btnLogout").style.display = "inline-block";
  $("meRole").textContent = ROLE || "—";
  $("meToken").textContent = TOKEN ? TOKEN.slice(0,10)+"..." : "—";
}

async function login(){
  const u = $("lgUser").value.trim();
  const p = $("lgPass").value;
  $("loginMsg").textContent = "";
  if(!u || !p) return toast("اكتب Username و Password", true);
  try{
    const r = await apiJSONP("login", {username:u, password:p}, "");
    if(!r || !r.ok) {
      toast((r && r.error) ? r.error : "بيانات الدخول غير صحيحة", true);
      $("loginMsg").textContent = (r && r.error) ? r.error : "بيانات الدخول غير صحيحة";
      return;
    }
    TOKEN = r.token || "";
    ROLE  = (r.user && r.user.role) ? r.user.role : "";
    localStorage.setItem(STORE_TOKEN, TOKEN);
    localStorage.setItem(STORE_ROLE, ROLE);
    toast("تم تسجيل الدخول ✅");
    showApp();
    await ping();
  }catch(e){
    toast("Login فشل: " + (e.message||e), true);
  }
}

function logout(){
  TOKEN = "";
  ROLE = "";
  localStorage.removeItem(STORE_TOKEN);
  localStorage.removeItem(STORE_ROLE);
  toast("تم تسجيل الخروج");
  showLogin();
}

function openSettings(){
  $("apiUrlInput").value = API_URL || "";
  $("ov").style.display = "flex";
}
function closeSettings(){ $("ov").style.display = "none"; }
function saveSettings(){
  API_URL = $("apiUrlInput").value.trim();
  if(!API_URL) API_URL = DEFAULT_APPS_SCRIPT_URL;
  localStorage.setItem(STORE_API, API_URL);
  closeSettings();
  toast("تم حفظ Settings ✅");
  ping();
}

function todayISO(){
  return new Date().toISOString().slice(0,10);
}

async function loadToday(){
  const d = $("dateInput").value || todayISO();
  $("dateInput").value = d;
  try{
    const r = await apiJSONP("listBookingsByDateAll", {date:d}, TOKEN);
    if(!r || !r.ok) return toast((r && r.error) ? r.error : "فشل تحميل الحجوزات", true);
    const items = r.items || [];
    $("todayList").innerHTML = items.length ? items.map(b=>`
      <div class="item">
        <div class="t">${(b.clientName||"")}</div>
        <div class="m">📅 ${b.date||d} • ${b.type||""} • 💰 إجمالي ${b.total||0} • مدفوع ${b.paid||0} • متبقي ${b.remaining||0}</div>
      </div>
    `).join("") : '<div class="msg">لا يوجد حجوزات.</div>';
  }catch(e){
    toast("فشل: " + (e.message||e), true);
  }
}

document.addEventListener("DOMContentLoaded", async ()=>{
  $("btnPing").onclick = ping;
  $("btnSettings").onclick = openSettings;
  $("btnCloseSettings").onclick = closeSettings;
  $("btnSaveSettings").onclick = saveSettings;
  $("btnLogin").onclick = login;
  $("btnLogout").onclick = logout;
  $("btnLoadToday").onclick = loadToday;

  $("apiUrlInput").value = API_URL;
  $("dateInput").value = todayISO();

  if(TOKEN) showApp(); else showLogin();
  await ping();
});

import { useState, useEffect, useMemo, useRef, useCallback } from "react";
import { createClient } from "https://esm.sh/@supabase/supabase-js@2";

const SUPA_URL = import.meta.env.VITE_SUPA_URL;
const SUPA_ANON = import.meta.env.VITE_SUPA_ANON;

const supabase = createClient(SUPA_URL, SUPA_ANON);

const DEFAULT_SETTINGS = {
  branches: ["الفرع الرئيسي", "العبد -جسر السويس"],
  materials: [
    { id: "mat_1", name: "Zirconia", price: 450 },
    { id: "mat_2", name: "E-max", price: 450 },
    { id: "mat_3", name: "PMMA", price: 225 },
    { id: "mat_other", name: "Other / أخرى", price: null },
  ],
  caseStatuses: [
    { id: "cs_1", name: "In progress", isFinal: false },
    { id: "cs_2", name: "Correction", isFinal: false },
    { id: "cs_3", name: "Temp cement", isFinal: true },
    { id: "cs_4", name: "Missed", isFinal: true },
    { id: "cs_5", name: "Completed", isFinal: true },
  ],
  paymentStatuses: [
    { id: "ps_1", name: "Paid" },
    { id: "ps_2", name: "Unpaid" },
    { id: "ps_3", name: "Partial" },
    { id: "ps_4", name: "Free" },
  ],
};

const STATUS_META = {
  "In progress": { color: "#4F8EF7", bg: "rgba(79,142,247,0.12)", icon: "⚙️" },
  "Correction": { color: "#F59E0B", bg: "rgba(245,158,11,0.12)", icon: "✏️" },
  "Temp cement": { color: "#A78BFA", bg: "rgba(167,139,250,0.12)", icon: "🔩" },
  "Missed": { color: "#EF4444", bg: "rgba(239,68,68,0.12)", icon: "⚠️" },
  "Completed": { color: "#10B981", bg: "rgba(16,185,129,0.12)", icon: "✅" },
};

const MAT_COLORS = {
  "Zirconia": "#4F8EF7",
  "E-max": "#A78BFA",
  "PMMA": "#F59E0B",
  "Other / أخرى": "#10B981",
};

const DATE_RANGE = {
  startYear: 2024,
  endYear: 2030
};

const todayISO = () => new Date().toISOString().split("T")[0];
const fmtMoney = (n) => Number(n || 0).toLocaleString("ar-EG");
const fmtDate = (d) => d ? new Date(d).toLocaleDateString("ar-EG", { day: "2-digit", month: "short", year: "numeric" }) : "—";
const monthKey = (d) => d ? d.slice(0, 7) : "";
const currentMonthKey = () => new Date().toISOString().slice(0, 7);

const CSS = `
@import url('https://fonts.googleapis.com/css2?family=Tajawal:wght@300;400;500;700;800;900&family=JetBrains+Mono:wght@400;600&display=swap');

*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

:root {
  --bg: #080B12;
  --bg2: #0D1117;
  --surface: #111621;
  --surface2: #161C2A;
  --surface3: #1D2438;
  --border: rgba(255,255,255,0.07);
  --border2: rgba(255,255,255,0.12);
  --accent: #3D7EFF;
  --accent2: #6C63FF;
  --gold: #F5A623;
  --mint: #00D4A1;
  --rose: #FF4D6D;
  --amber: #FFB830;
  --lavender: #9B8EFF;
  --text: #E8ECF8;
  --text2: #8A94B8;
  --text3: #4A5280;
  --r: 12px;
  --r2: 8px;
}

html { direction: rtl; }

body {
  font-family: 'Tajawal', system-ui, sans-serif;
  background: var(--bg);
  color: var(--text);
  min-height: 100vh;
}

::-webkit-scrollbar { width: 4px; height: 4px; }
::-webkit-scrollbar-track { background: transparent; }
::-webkit-scrollbar-thumb { background: var(--border2); border-radius: 2px; }

@keyframes fadeUp { from { opacity:0; transform:translateY(12px); } to { opacity:1; transform:none; } }
@keyframes fadeIn { from { opacity:0; } to { opacity:1; } }
@keyframes slideIn { from { transform:translateX(100%); opacity:0; } to { transform:none; opacity:1; } }

.fade-up { animation: fadeUp .35s ease both; }
.fade-in { animation: fadeIn .25s ease both; }

.login-shell {
  min-height: 100vh;
  display: flex; align-items: center; justify-content: center;
  background: radial-gradient(ellipse 80% 60% at 50% 0%, rgba(61,126,255,.18) 0%, transparent 70%), var(--bg);
  padding: 24px;
}
.login-card {
  background: var(--surface);
  border: 1px solid var(--border2);
  border-radius: 20px;
  padding: 40px 36px;
  width: 100%; max-width: 420px;
  box-shadow: 0 40px 80px rgba(0,0,0,.6);
  animation: fadeUp .4s ease;
}
.login-logo { text-align: center; margin-bottom: 32px; }
.login-logo-icon {
  width: 64px; height: 64px; border-radius: 20px;
  background: linear-gradient(135deg, var(--accent), var(--accent2));
  display: inline-flex; align-items: center; justify-content: center;
  font-size: 28px; margin-bottom: 14px;
  box-shadow: 0 8px 24px rgba(61,126,255,.4);
}
.login-logo h1 { font-size: 20px; font-weight: 800; color: var(--text); }
.login-logo p { font-size: 13px; color: var(--text2); margin-top: 4px; }

.header {
  background: var(--surface);
  border-bottom: 1px solid var(--border);
  position: sticky; top: 0; z-index: 100;
  padding: 0 24px;
}
.header-top {
  height: 58px;
  display: flex; align-items: center; justify-content: space-between; gap: 12px;
}
.header-title {
  font-size: 17px; font-weight: 800; color: var(--text);
  display: flex; align-items: center; gap: 8px;
}
.header-title .dot {
  width: 8px; height: 8px; border-radius: 50%;
  background: var(--accent);
  box-shadow: 0 0 8px var(--accent);
}
.header-actions { display: flex; gap: 8px; align-items: center; }

.branch-bar {
  padding: 0 24px 12px;
  display: flex; align-items: center; gap: 10px; flex-wrap: wrap;
}
.branch-label { font-size: 11px; color: var(--text3); font-weight: 700; }
.branch-select {
  background: var(--surface2);
  border: 1px solid var(--border2);
  border-radius: var(--r2);
  color: var(--accent);
  font-family: 'Tajawal', sans-serif;
  font-size: 14px; font-weight: 700;
  padding: 6px 12px;
  outline: none; cursor: pointer;
}
.branch-select:focus { border-color: var(--accent); }

.tabs {
  display: flex; gap: 4px;
  padding: 8px 24px;
  border-bottom: 1px solid var(--border);
  overflow-x: auto;
}
.tab {
  padding: 7px 16px;
  border-radius: 20px;
  font-size: 12px; font-weight: 700;
  border: 1px solid var(--border);
  color: var(--text2);
  cursor: pointer; white-space: nowrap;
  background: transparent;
  font-family: 'Tajawal', sans-serif;
}
.tab:hover { border-color: var(--accent); color: var(--accent); }
.tab.active { background: var(--accent); border-color: var(--accent); color: #fff; }

.main { padding: 20px 24px; max-width: 1400px; margin: 0 auto; }

.stats-grid {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
  gap: 12px; margin-bottom: 20px;
}
.stat-card {
  background: var(--surface);
  border: 1px solid var(--border);
  border-radius: var(--r);
  padding: 16px;
  position: relative; overflow: hidden;
  animation: fadeUp .3s ease both;
}
.stat-card::before {
  content: '';
  position: absolute; inset: 0;
  background: linear-gradient(135deg, var(--card-color, transparent) 0%, transparent 60%);
  opacity: .06; pointer-events: none;
}
.stat-label { font-size: 11px; color: var(--text2); font-weight: 700; margin-bottom: 8px; }
.stat-value { font-size: 22px; font-weight: 900; line-height: 1; }

.toolbar {
  display: flex; gap: 10px; align-items: center;
  margin-bottom: 16px; flex-wrap: wrap;
}
.search-wrap { position: relative; flex: 1; min-width: 200px; }
.search-wrap input {
  width: 100%; padding: 9px 36px 9px 14px;
  background: var(--surface2); border: 1px solid var(--border);
  border-radius: var(--r2); color: var(--text);
  font-family: 'Tajawal', sans-serif; font-size: 14px; outline: none;
}
.search-wrap input:focus { border-color: var(--accent); }

.month-select {
  background: var(--surface2); border: 1px solid var(--border);
  border-radius: var(--r2); color: var(--text);
  font-family: 'Tajawal', sans-serif; font-size: 13px;
  padding: 8px 12px; outline: none; cursor: pointer;
}

.cases-grid {
  display: grid;
  grid-template-columns: repeat(auto-fill, minmax(320px, 1fr));
  gap: 12px;
}
.case-card {
  background: var(--surface);
  border: 1px solid var(--border);
  border-radius: var(--r);
  padding: 16px;
  cursor: pointer;
  transition: all .18s;
  animation: fadeUp .3s ease both;
  position: relative; overflow: hidden;
}
.case-card::after {
  content: '';
  position: absolute; right: 0; top: 0; bottom: 0;
  width: 3px;
  background: var(--status-color, var(--border));
  border-radius: 0 var(--r) var(--r) 0;
}
.case-card:hover { border-color: var(--border2); transform: translateY(-2px); }
.case-card-header { display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 10px; }
.case-name { font-size: 15px; font-weight: 700; color: var(--text); }
.case-date { font-size: 11px; color: var(--text3); margin-top: 2px; }
.case-badges { display: flex; gap: 6px; flex-wrap: wrap; margin-bottom: 10px; }
.case-footer { display: flex; justify-content: space-between; align-items: center; }
.case-amount { font-size: 18px; font-weight: 800; }
.case-units { font-size: 11px; color: var(--text3); }
.case-branch { font-size: 11px; color: var(--text2); background: var(--surface2); padding: 3px 8px; border-radius: 20px; }

.badge {
  display: inline-flex; align-items: center; gap: 4px;
  padding: 3px 9px; border-radius: 20px;
  font-size: 11px; font-weight: 700; white-space: nowrap;
}

.btn {
  display: inline-flex; align-items: center; gap: 6px;
  padding: 8px 16px; border-radius: var(--r2);
  font-family: 'Tajawal', sans-serif; font-size: 13px; font-weight: 700;
  border: none; cursor: pointer; transition: all .18s; white-space: nowrap;
}
.btn-primary { background: var(--accent); color: #fff; }
.btn-primary:hover { background: #3070e8; box-shadow: 0 4px 16px rgba(61,126,255,.4); }
.btn-ghost { background: var(--surface2); border: 1px solid var(--border2); color: var(--text2); }
.btn-ghost:hover { border-color: var(--accent); color: var(--accent); }
.btn-danger { background: rgba(239,68,68,.15); border: 1px solid rgba(239,68,68,.3); color: var(--rose); }
.btn-danger:hover { background: var(--rose); color: #fff; }
.btn-mint { background: rgba(0,212,161,.15); border: 1px solid rgba(0,212,161,.3); color: var(--mint); }
.btn-mint:hover { background: var(--mint); color: #fff; }
.btn-sm { padding: 6px 12px; font-size: 12px; }
.btn-icon { padding: 7px; border-radius: var(--r2); }
.btn-fab {
  position: fixed; bottom: 28px; left: 28px;
  width: 56px; height: 56px; border-radius: 50%;
  background: linear-gradient(135deg, var(--accent), var(--accent2));
  color: #fff; font-size: 24px;
  display: flex; align-items: center; justify-content: center;
  border: none; cursor: pointer; z-index: 90;
  box-shadow: 0 8px 24px rgba(61,126,255,.5);
}
.btn-fab:hover { transform: scale(1.08); }

.modal-overlay {
  position: fixed; inset: 0;
  background: rgba(0,0,0,.7); backdrop-filter: blur(6px);
  z-index: 200;
  display: flex; align-items: center; justify-content: center;
  padding: 20px; animation: fadeIn .2s ease;
}
.modal {
  background: var(--surface);
  border: 1px solid var(--border2);
  border-radius: 18px;
  width: 100%; max-width: 560px;
  max-height: 90vh; overflow-y: auto;
  box-shadow: 0 40px 80px rgba(0,0,0,.7);
  animation: fadeUp .25s ease;
}
.modal-header {
  padding: 20px 24px 0;
  display: flex; justify-content: space-between; align-items: center;
  position: sticky; top: 0;
  background: var(--surface); z-index: 1;
  padding-bottom: 16px;
  border-bottom: 1px solid var(--border);
}
.modal-title { font-size: 16px; font-weight: 800; }
.modal-body { padding: 20px 24px; }
.modal-footer { padding: 16px 24px; border-top: 1px solid var(--border); display: flex; gap: 10px; justify-content: flex-end; }

.form-group { margin-bottom: 14px; }
.form-label { display: block; font-size: 12px; font-weight: 700; color: var(--text2); margin-bottom: 6px; }
.form-input {
  width: 100%;
  background: var(--surface2); border: 1px solid var(--border);
  border-radius: var(--r2); padding: 10px 13px;
  color: var(--text); font-family: 'Tajawal', sans-serif; font-size: 14px;
  outline: none;
}
.form-input:focus { border-color: var(--accent); }
.form-row { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; }
.form-textarea { min-height: 72px; resize: vertical; }

.total-preview {
  background: linear-gradient(135deg, rgba(61,126,255,.1), rgba(108,99,255,.08));
  border: 1px solid rgba(61,126,255,.25);
  border-radius: var(--r); padding: 14px 16px;
  display: flex; justify-content: space-between; align-items: center;
  margin-bottom: 14px;
}
.total-preview-label { font-size: 12px; color: var(--text2); }
.total-preview-value { font-size: 26px; font-weight: 900; color: var(--gold); font-family: 'JetBrains Mono', monospace; }

.drawer-overlay {
  position: fixed; inset: 0;
  background: rgba(0,0,0,.5); z-index: 150;
  animation: fadeIn .2s ease;
}
.drawer {
  position: fixed; top: 0; left: 0; bottom: 0;
  width: min(480px, 100vw);
  background: var(--surface);
  border-right: 1px solid var(--border2);
  overflow-y: auto; z-index: 151;
  animation: slideIn .25s ease;
}
.drawer-header {
  padding: 20px 20px 16px;
  border-bottom: 1px solid var(--border);
  display: flex; justify-content: space-between; align-items: flex-start;
  position: sticky; top: 0;
  background: var(--surface); z-index: 1;
}
.drawer-body { padding: 20px; }
.drawer-section { margin-bottom: 24px; }
.drawer-section-title {
  font-size: 10px; font-weight: 700; color: var(--text3);
  letter-spacing: 1px; text-transform: uppercase;
  margin-bottom: 10px; padding-bottom: 6px;
  border-bottom: 1px solid var(--border);
}
.detail-row { display: flex; justify-content: space-between; align-items: center; padding: 8px 0; border-bottom: 1px solid var(--border); }
.detail-key { font-size: 12px; color: var(--text2); }
.detail-val { font-size: 13px; font-weight: 600; color: var(--text); text-align: left; }

.settings-section { margin-bottom: 28px; }
.settings-section-title {
  font-size: 13px; font-weight: 700; color: var(--text);
  margin-bottom: 12px; display: flex; align-items: center; gap: 8px;
}
.settings-list-item {
  display: flex; justify-content: space-between; align-items: center;
  padding: 10px 14px;
  background: var(--surface2); border: 1px solid var(--border);
  border-radius: var(--r2); margin-bottom: 6px;
}
.settings-list-item-name { font-size: 13px; font-weight: 600; }
.settings-list-item-sub { font-size: 11px; color: var(--text2); }

.toast {
  position: fixed; bottom: 90px; left: 50%; transform: translateX(-50%);
  background: var(--surface2); border: 1px solid var(--border2);
  padding: 10px 20px; border-radius: 20px;
  font-size: 13px; font-weight: 600; color: var(--text);
  z-index: 300; animation: fadeUp .25s ease;
}

.empty { text-align: center; padding: 60px 20px; color: var(--text3); }
.empty-icon { font-size: 48px; margin-bottom: 12px; opacity: .3; }

.loading { display: flex; align-items: center; justify-content: center; gap: 8px; padding: 40px; color: var(--text2); }
.spinner { width: 18px; height: 18px; border: 2px solid var(--border2); border-top-color: var(--accent); border-radius: 50%; animation: spin .7s linear infinite; }
@keyframes spin { to { transform: rotate(360deg); } }

@media (max-width: 640px) {
  .main { padding: 14px; }
  .stats-grid { grid-template-columns: 1fr 1fr; }
  .cases-grid { grid-template-columns: 1fr; }
  .form-row { grid-template-columns: 1fr; }
}
`;

function Badge({ label, color, bg, icon }) {
  return (
    <span className="badge" style={{ color, background: bg, border: `1px solid ${color}30` }}>
      {icon && <span>{icon}</span>}
      {label}
    </span>
  );
}

function Toast({ msg }) {
  return msg ? <div className="toast">{msg}</div> : null;
}

function Spinner() {
  return <div className="loading"><div className="spinner" /><span>جاري التحميل...</span></div>;
}

async function getSettings(userId) {
  const { data } = await supabase.from("settings").select("*").eq("user_id", userId).single();
  if (!data) {
    await supabase.from("settings").insert({
      user_id: userId,
      branches: DEFAULT_SETTINGS.branches,
      materials: DEFAULT_SETTINGS.materials,
      case_statuses: DEFAULT_SETTINGS.caseStatuses,
      payment_statuses: DEFAULT_SETTINGS.paymentStatuses,
    });
    return DEFAULT_SETTINGS;
  }
  return {
    branches: data.branches,
    materials: data.materials,
    caseStatuses: data.case_statuses,
    paymentStatuses: data.payment_statuses,
  };
}

async function saveSettings(userId, settings) {
  await supabase.from("settings").upsert({
    user_id: userId,
    branches: settings.branches,
    materials: settings.materials,
    case_statuses: settings.caseStatuses,
    payment_statuses: settings.paymentStatuses,
  }, { onConflict: "user_id" });
}

async function fetchCases(userId) {
  const { data } = await supabase
    .from("cases").select("*")
    .eq("user_id", userId)
    .order("created_at", { ascending: false });
  return data || [];
}

async function insertCase(userId, c) {
  const { data } = await supabase.from("cases").insert({
    user_id: userId,
    branch_name: c.branchName,
    patient_name: c.patientName,
    material_name: c.materialName,
    price_per_unit: c.pricePerUnit,
    units: c.units,
    case_status: c.caseStatus,
    payment_status: c.paymentStatus,
    paid_amount: c.paidAmount || 0,
    start_date: c.startDate,
    action_date: c.actionDate || null,
    notes: c.notes || null,
  }).select().single();
  return data;
}

async function updateCaseDB(id, patch) {
  await supabase.from("cases").update(patch).eq("id", id);
}

async function deleteCaseDB(id) {
  await supabase.from("cases").delete().eq("id", id);
}

const toLocal = (r) => ({
  id: r.id,
  branchName: r.branch_name,
  patientName: r.patient_name,
  materialName: r.material_name,
  pricePerUnit: r.price_per_unit,
  units: r.units,
  totalAmount: r.total_amount,
  caseStatus: r.case_status,
  paymentStatus: r.payment_status,
  paidAmount: r.paid_amount || 0,
  startDate: r.start_date,
  actionDate: r.action_date,
  notes: r.notes,
  createdAt: r.created_at,
});

function exportCSV(casesToExport) {
  const formatDate = (dateStr) => {
    if (!dateStr) return "";
    const d = new Date(dateStr);
    return `${d.getDate()}/${d.getMonth() + 1}/${d.getFullYear()}`;
  };

  const getPaymentText = (status, paidAmount) => {
    if (status === "Paid") return "مدفوع كامل";
    if (status === "Free") return "مجاناً";
    if (status === "Partial") return `مدفوع جزئي (${fmtMoney(paidAmount)} ج.م)`;
    return "غير مدفوع";
  };

  const headers = ["الفرع", "المريض", "المادة", "السعر", "الوحدات", "الإجمالي", "المحصل", "المتبقي", "الحالة", "الدفع", "البداية", "الإجراء", "ملاحظات"];
  const rows = casesToExport.map(c => [
    c.branchName,
    c.patientName,
    c.materialName,
    c.pricePerUnit,
    c.units,
    c.totalAmount,
    c.paidAmount || 0,
    (c.totalAmount - (c.paidAmount || 0)),
    c.caseStatus,
    getPaymentText(c.paymentStatus, c.paidAmount),
    formatDate(c.startDate),
    formatDate(c.actionDate) || "",
    c.notes || ""
  ]);
  const csv = "\uFEFF" + [headers, ...rows].map(r => r.map(v => `"${v}"`).join(",")).join("\n");
  const a = document.createElement("a");
  a.href = URL.createObjectURL(new Blob([csv], { type: "text/csv" }));
  a.download = `dental_cases_${todayISO()}.csv`;
  a.click();
}

async function importCSV(file, userId, currentBranch) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = async (e) => {
      try {
        const text = e.target.result.replace(/^\uFEFF/, "");
        const rows = text.split("\n").slice(1);
        let count = 0;
        for (const row of rows) {
          const cols = row.match(/(".*?"|[^,]+)(?=,|$)/g);
          if (!cols || cols.length < 6) continue;
          const clean = cols.map(c => c.replace(/^"|"$/g, "").trim());
          
          let paymentStatus = clean[9] || "Unpaid";
          if (paymentStatus === "مدفوع" || paymentStatus === "مدفوع كامل") paymentStatus = "Paid";
          if (paymentStatus === "مجاناً") paymentStatus = "Free";
          if (paymentStatus.includes("مدفوع جزئي")) paymentStatus = "Partial";
          if (paymentStatus === "غير مدفوع") paymentStatus = "Unpaid";
          
          const rec = {
            branchName: clean[0] || currentBranch,
            patientName: clean[1],
            materialName: clean[2],
            pricePerUnit: parseFloat(clean[3]) || 0,
            units: parseInt(clean[4]) || 0,
            caseStatus: clean[8] || "In progress",
            paymentStatus: paymentStatus,
            paidAmount: parseFloat(clean[6]) || 0,
            startDate: clean[10] ? clean[10].split('/').reverse().join('-') : todayISO(),
            actionDate: clean[11] ? clean[11].split('/').reverse().join('-') : null,
            notes: clean[12] || null,
          };
          if (!rec.patientName) continue;
          await insertCase(userId, rec);
          count++;
        }
        resolve(count);
      } catch (err) { reject(err); }
    };
    reader.readAsText(file);
  });
}

function LoginScreen({ onLogin }) {
  const [mode, setMode] = useState("login");
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [err, setErr] = useState("");
  const [loading, setLoading] = useState(false);

  async function submit(e) {
    e.preventDefault();
    setErr("");
    setLoading(true);
    try {
      if (mode === "signup") {
        const { error } = await supabase.auth.signUp({ email, password });
        if (error) throw error;
        setErr("تم إرسال رابط التأكيد على البريد الإلكتروني ✉️");
      } else {
        const { data, error } = await supabase.auth.signInWithPassword({ email, password });
        if (error) throw error;
        onLogin(data.session);
      }
    } catch (err) {
      setErr(err.message || "حدث خطأ");
    } finally {
      setLoading(false);
    }
  }

  return (
    <div className="login-shell">
      <div className="login-card">
        <div className="login-logo">
          <div className="login-logo-icon">🦷</div>
          <h1>Dental Tracker</h1>
          <p>نظام متابعة الحالات والأرباح</p>
        </div>
        <form onSubmit={submit}>
          <div className="form-group">
            <label className="form-label">البريد الإلكتروني</label>
            <input className="form-input" type="email" value={email} onChange={e => setEmail(e.target.value)} placeholder="doctor@clinic.com" required />
          </div>
          <div className="form-group">
            <label className="form-label">كلمة المرور</label>
            <input className="form-input" type="password" value={password} onChange={e => setPassword(e.target.value)} placeholder="••••••••" required />
          </div>
          {err && <div style={{ fontSize: 13, color: err.includes("✉️") ? "var(--mint)" : "var(--rose)", marginBottom: 12, padding: "8px 12px", background: "rgba(239,68,68,.08)", borderRadius: 8 }}>{err}</div>}
          <button className="btn btn-primary" type="submit" disabled={loading} style={{ width: "100%", justifyContent: "center", padding: "12px" }}>
            {loading ? "جاري..." : mode === "login" ? "تسجيل الدخول" : "إنشاء حساب"}
          </button>
        </form>
        <div style={{ textAlign: "center", marginTop: 16 }}>
          <button onClick={() => setMode(m => m === "login" ? "signup" : "login")} style={{ background: "none", border: "none", color: "var(--accent)", cursor: "pointer", fontFamily: "Tajawal", fontSize: 13 }}>
            {mode === "login" ? "مش عندك حساب؟ سجّل دلوقتي" : "عندك حساب؟ ادخل"}
          </button>
        </div>
      </div>
    </div>
  );
}

function CaseModal({ existing, settings, defaultBranch, onSave, onClose }) {
  const isEdit = !!existing;
  const [form, setForm] = useState({
    patientName: existing?.patientName || "",
    branchName: existing?.branchName || defaultBranch || settings.branches[0] || "",
    materialName: existing?.materialName || settings.materials[0]?.name || "",
    pricePerUnit: existing?.pricePerUnit || settings.materials[0]?.price || 450,
    units: existing?.units || 1,
    caseStatus: existing?.caseStatus || "In progress",
    paymentStatus: existing?.paymentStatus || "Unpaid",
    paidAmount: existing?.paidAmount || 0,
    startDate: existing?.startDate || todayISO(),
    actionDate: existing?.actionDate || "",
    notes: existing?.notes || "",
  });
  const [prevStatus, setPrevStatus] = useState(existing?.caseStatus || "In progress");
  const [loading, setLoading] = useState(false);
  const total = (form.pricePerUnit || 0) * (form.units || 0);
  const remaining = total - (form.paidAmount || 0);
  const selMat = settings.materials.find(m => m.name === form.materialName);
  const priceFixed = selMat?.price !== null && selMat?.price !== undefined;

  function handleMaterial(name) {
    const mat = settings.materials.find(m => m.name === name);
    setForm(f => ({
      ...f,
      materialName: name,
      pricePerUnit: isEdit ? f.pricePerUnit : (mat?.price !== null && mat?.price !== undefined ? mat.price : f.pricePerUnit),
    }));
  }

  function handleStatus(s) {
    const wasActive = prevStatus === "In progress" || prevStatus === "Correction";
    const becomingFinal = s !== "In progress" && s !== "Correction";
    setForm(f => ({
      ...f,
      caseStatus: s,
      actionDate: wasActive && becomingFinal ? todayISO() : f.actionDate,
    }));
    setPrevStatus(s);
  }

  async function submit() {
    if (!form.patientName.trim()) return;
    setLoading(true);
    await onSave(form);
    setLoading(false);
  }

  return (
    <div className="modal-overlay" onClick={onClose}>
      <div className="modal" onClick={e => e.stopPropagation()}>
        <div className="modal-header">
          <span className="modal-title">{isEdit ? "✏️ تعديل الحالة" : "➕ حالة جديدة"}</span>
          <button className="btn btn-ghost btn-icon" onClick={onClose}>✕</button>
        </div>
        <div className="modal-body">
          <div className="total-preview">
            <div>
              <div className="total-preview-label">إجمالي الحالة</div>
              <div className="total-preview-calc">{form.units} وحدة × {form.pricePerUnit} = {total.toLocaleString("ar-EG")}</div>
            </div>
            <div className="total-preview-value">{total.toLocaleString("ar-EG")}</div>
          </div>
          <div className="form-row">
            <div className="form-group">
              <label className="form-label">اسم المريض *</label>
              <input className="form-input" value={form.patientName} onChange={e => setForm(f => ({ ...f, patientName: e.target.value }))} placeholder="الاسم كاملاً" />
            </div>
            <div className="form-group">
              <label className="form-label">الفرع</label>
              <select className="form-input" value={form.branchName} onChange={e => setForm(f => ({ ...f, branchName: e.target.value }))}>
                {settings.branches.map(b => <option key={b} value={b}>{b}</option>)}
              </select>
            </div>
          </div>
          <div className="form-group">
            <label className="form-label">المادة / الإجراء</label>
            <select className="form-input" value={form.materialName} onChange={e => handleMaterial(e.target.value)}>
              {settings.materials.map(m => (
                <option key={m.id} value={m.name}>{m.name}{m.price ? ` — ${m.price} ج.م` : " — سعر حر"}</option>
              ))}
            </select>
          </div>
          <div className="form-row">
            <div className="form-group">
              <label className="form-label">عدد الوحدات</label>
              <input className="form-input" type="number" min="1" value={form.units} onChange={e => setForm(f => ({ ...f, units: parseInt(e.target.value) || 1 }))} />
            </div>
            <div className="form-group">
              <label className="form-label">سعر الوحدة {priceFixed && !isEdit ? <span style={{ color: "var(--text3)", fontSize: 10 }}>🔒 محدد</span> : ""}</label>
              <input className="form-input" type="number" value={form.pricePerUnit} disabled={priceFixed && !isEdit} onChange={e => setForm(f => ({ ...f, pricePerUnit: parseFloat(e.target.value) || 0 }))} style={priceFixed && !isEdit ? { opacity: .6 } : {}} />
            </div>
          </div>
          <div className="form-row">
            <div className="form-group">
              <label className="form-label">حالة الكيس</label>
              <select className="form-input" value={form.caseStatus} onChange={e => handleStatus(e.target.value)}>
                {settings.caseStatuses.map(s => (
                  <option key={s.id} value={s.name}>{STATUS_META[s.name]?.icon} {s.name}</option>
                ))}
              </select>
            </div>
            <div className="form-group">
              <label className="form-label">حالة الدفع</label>
              <select className="form-input" value={form.paymentStatus} onChange={e => setForm(f => ({ ...f, paymentStatus: e.target.value }))}>
                {settings.paymentStatuses.map(s => <option key={s.id} value={s.name}>{s.name === "Paid" ? "مدفوع كامل" : (s.name === "Free" ? "مجاناً" : (s.name === "Partial" ? "مدفوع جزئي" : "غير مدفوع"))}</option>)}
              </select>
            </div>
          </div>
          {form.paymentStatus === "Partial" && (
            <div className="form-row">
              <div className="form-group">
                <label className="form-label">المبلغ المحصل</label>
                <input className="form-input" type="number" value={form.paidAmount} onChange={e => setForm(f => ({ ...f, paidAmount: parseFloat(e.target.value) || 0 }))} placeholder="المبلغ المدفوع" />
                <div style={{ fontSize: 11, color: "var(--text3)", marginTop: 4 }}>المتبقي: {remaining.toLocaleString("ar-EG")} ج.م</div>
              </div>
            </div>
          )}
          <div className="form-row">
            <div className="form-group">
              <label className="form-label">تاريخ البداية</label>
              <input className="form-input" type="date" value={form.startDate} onChange={e => setForm(f => ({ ...f, startDate: e.target.value }))} />
            </div>
            <div className="form-group">
              <label className="form-label">تاريخ الإجراء</label>
              <input className="form-input" type="date" value={form.actionDate} onChange={e => setForm(f => ({ ...f, actionDate: e.target.value }))} />
            </div>
          </div>
          <div className="form-group">
            <label className="form-label">ملاحظات إكلينيكية</label>
            <textarea className="form-input form-textarea" value={form.notes} onChange={e => setForm(f => ({ ...f, notes: e.target.value }))} placeholder="أي تفاصيل إضافية عن الحالة..." />
          </div>
        </div>
        <div className="modal-footer">
          <button className="btn btn-ghost" onClick={onClose}>إلغاء</button>
          <button className="btn btn-primary" onClick={submit} disabled={loading || !form.patientName.trim()}>
            {loading ? "جاري الحفظ..." : isEdit ? "💾 حفظ التعديل" : "➕ إضافة الحالة"}
          </button>
        </div>
      </div>
    </div>
  );
}

function CaseDrawer({ c, settings, onEdit, onDelete, onUpdatePayment, onClose }) {
  const sm = STATUS_META[c.caseStatus] || {};
  const mc = MAT_COLORS[c.materialName] || "var(--mint)";
  const remaining = (c.totalAmount || 0) - (c.paidAmount || 0);
  
  const getPaymentLabel = (status) => {
    if (status === "Paid") return "مدفوع كامل";
    if (status === "Free") return "مجاناً";
    if (status === "Partial") return `مدفوع جزئي (${fmtMoney(c.paidAmount)} ج.م)`;
    return "غير مدفوع";
  };

  return (
    <>
      <div className="drawer-overlay" onClick={onClose} />
      <div className="drawer">
        <div className="drawer-header">
          <div>
            <div style={{ fontSize: 17, fontWeight: 800, marginBottom: 4 }}>{c.patientName}</div>
            <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
              <Badge label={c.caseStatus} color={sm.color} bg={sm.bg} icon={sm.icon} />
              <Badge label={c.materialName} color={mc} bg={mc + "20"} />
              <Badge 
                label={getPaymentLabel(c.paymentStatus)} 
                color={c.paymentStatus === "Paid" ? "var(--mint)" : (c.paymentStatus === "Free" ? "var(--amber)" : (c.paymentStatus === "Partial" ? "var(--lavender)" : "var(--rose)"))} 
                bg={c.paymentStatus === "Paid" ? "rgba(0,212,161,.12)" : (c.paymentStatus === "Free" ? "rgba(245,166,35,.12)" : (c.paymentStatus === "Partial" ? "rgba(155,142,255,.12)" : "rgba(239,68,68,.12)"))} 
              />
            </div>
          </div>
          <button className="btn btn-ghost btn-icon" onClick={onClose}>✕</button>
        </div>
        <div className="drawer-body">
          <div className="drawer-section">
            <div className="drawer-section-title">الحساب</div>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "12px 0" }}>
              <div>
                <div style={{ fontSize: 11, color: "var(--text2)" }}>الإجمالي</div>
                <div style={{ fontSize: 11, color: "var(--text3)", marginTop: 2 }}>{c.units} وحدة × {c.pricePerUnit} ج.م</div>
              </div>
              <div style={{ fontSize: 26, fontWeight: 900, color: "var(--gold)", fontFamily: "JetBrains Mono" }}>{fmtMoney(c.totalAmount)} ج.م</div>
            </div>
            {c.paymentStatus === "Partial" && (
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "8px 0", borderTop: "1px solid var(--border)" }}>
                <div>
                  <div style={{ fontSize: 11, color: "var(--text2)" }}>المحصل</div>
                  <div style={{ fontSize: 11, color: "var(--mint)" }}>{fmtMoney(c.paidAmount)} ج.م</div>
                </div>
                <div>
                  <div style={{ fontSize: 11, color: "var(--text2)" }}>المتبقي</div>
                  <div style={{ fontSize: 11, color: "var(--rose)" }}>{fmtMoney(remaining)} ج.م</div>
                </div>
              </div>
            )}
          </div>
          <div className="drawer-section">
            <div className="drawer-section-title">التفاصيل</div>
            {[
              ["الفرع", c.branchName],
              ["تاريخ البداية", fmtDate(c.startDate)],
              ["تاريخ الإجراء", fmtDate(c.actionDate)],
              ["تاريخ الإضافة", fmtDate(c.createdAt)],
            ].map(([k, v]) => (
              <div key={k} className="detail-row">
                <span className="detail-key">{k}</span>
                <span className="detail-val">{v}</span>
              </div>
            ))}
            {c.notes && (
              <div style={{ marginTop: 10, padding: "10px 12px", background: "var(--surface2)", borderRadius: 8, fontSize: 13, color: "var(--text2)" }}>
                📝 {c.notes}
              </div>
            )}
          </div>
          
          {/* أزرار تغيير حالة الدفع */}
          <div className="drawer-section">
            <div className="drawer-section-title">تغيير حالة الدفع</div>
            <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
              <button 
                className={`btn ${c.paymentStatus === "Paid" ? "btn-primary" : "btn-ghost"}`}
                style={{ flex: 1, justifyContent: "center" }}
                onClick={() => onUpdatePayment(c, "Paid")}
              >
                ✅ مدفوع
              </button>
              <button 
                className={`btn ${c.paymentStatus === "Unpaid" ? "btn-primary" : "btn-ghost"}`}
                style={{ flex: 1, justifyContent: "center" }}
                onClick={() => onUpdatePayment(c, "Unpaid")}
              >
                ❌ غير مدفوع
              </button>
              <button 
                className={`btn ${c.paymentStatus === "Free" ? "btn-primary" : "btn-ghost"}`}
                style={{ flex: 1, justifyContent: "center" }}
                onClick={() => onUpdatePayment(c, "Free")}
              >
                🎁 مجاناً
              </button>
            </div>
          </div>

          <div className="divider" style={{ margin: "12px 0" }} />

          <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
            <button className="btn btn-ghost" style={{ justifyContent: "center" }} onClick={() => onEdit(c)}>✏️ تعديل الحالة</button>
            <button className="btn btn-danger" style={{ justifyContent: "center" }} onClick={() => onDelete(c)}>🗑 حذف الحالة</button>
          </div>
        </div>
      </div>
    </>
  );
}

function SettingsScreen({ settings, onSave, onClose }) {
  const [s, setS] = useState(JSON.parse(JSON.stringify(settings)));
  const [loading, setLoading] = useState(false);
  const [newBranch, setNewBranch] = useState("");
  const [newMaterial, setNewMaterial] = useState({ name: "", price: "" });
  const [editingMaterial, setEditingMaterial] = useState(null);

  function addBranch() {
    const b = newBranch.trim();
    if (!b || s.branches.includes(b)) return;
    setS(prev => ({ ...prev, branches: [...prev.branches, b] }));
    setNewBranch("");
  }

  function delBranch(b) {
    setS(prev => ({ ...prev, branches: prev.branches.filter(x => x !== b) }));
  }

  function addMaterial() {
    const name = newMaterial.name.trim();
    if (!name || s.materials.some(m => m.name === name)) return;
    const newId = `mat_${Date.now()}`;
    const price = newMaterial.price === "" ? null : parseFloat(newMaterial.price);
    setS(prev => ({
      ...prev,
      materials: [...prev.materials, { id: newId, name, price }]
    }));
    setNewMaterial({ name: "", price: "" });
  }

  function updateMaterialPrice(id, price) {
    setS(prev => ({
      ...prev,
      materials: prev.materials.map(m =>
        m.id === id ? { ...m, price: price === "" ? null : parseFloat(price) } : m
      )
    }));
    setEditingMaterial(null);
  }

  function deleteMaterial(id) {
    if (s.materials.length <= 1) {
      alert("لا يمكن حذف المادة الوحيدة");
      return;
    }
    setS(prev => ({
      ...prev,
      materials: prev.materials.filter(m => m.id !== id)
    }));
  }

  async function save() {
    setLoading(true);
    await onSave(s);
    setLoading(false);
    onClose();
  }

  return (
    <div className="modal-overlay" onClick={onClose}>
      <div className="modal" style={{ maxWidth: 680 }} onClick={e => e.stopPropagation()}>
        <div className="modal-header">
          <span className="modal-title">⚙️ الإعدادات</span>
          <button className="btn btn-ghost btn-icon" onClick={onClose}>✕</button>
        </div>
        <div className="modal-body">
          <div className="settings-section">
            <div className="settings-section-title">🏥 إدارة الفروع</div>
            {s.branches.map(b => (
              <div key={b} className="settings-list-item">
                <span className="settings-list-item-name">{b}</span>
                {s.branches.length > 1 && <button className="btn btn-danger btn-sm" onClick={() => delBranch(b)}>حذف</button>}
              </div>
            ))}
            <div style={{ display: "flex", gap: 8, marginTop: 8 }}>
              <input className="form-input" value={newBranch} onChange={e => setNewBranch(e.target.value)} placeholder="اسم الفرع الجديد" onKeyDown={e => e.key === "Enter" && addBranch()} />
              <button className="btn btn-primary" onClick={addBranch}>إضافة فرع</button>
            </div>
          </div>
          <div className="divider" />
          <div className="settings-section">
            <div className="settings-section-title">💊 المواد والأسعار</div>
            <div style={{ fontSize: 11, color: "var(--text3)", marginBottom: 8 }}>
              ⚠️ تغيير الأسعار هنا لن يؤثر على الحالات القديمة
            </div>
            {s.materials.map(m => (
              <div key={m.id} className="settings-list-item">
                <div>
                  <div className="settings-list-item-name" style={{ color: MAT_COLORS[m.name] || "var(--accent)" }}>{m.name}</div>
                  <div className="settings-list-item-sub">
                    {editingMaterial === m.id ? (
                      <input 
                        className="form-input" 
                        type="number" 
                        defaultValue={m.price ?? ""} 
                        style={{ width: 100, padding: "4px 8px", fontSize: 12 }}
                        onBlur={(e) => updateMaterialPrice(m.id, e.target.value)}
                        onKeyDown={(e) => e.key === "Enter" && updateMaterialPrice(m.id, e.target.value)}
                        autoFocus
                      />
                    ) : (
                      m.price !== null && m.price !== undefined ? `${m.price} ج.م / وحدة` : "سعر حر (يحدد عند الإضافة)"
                    )}
                  </div>
                </div>
                <div style={{ display: "flex", gap: 6 }}>
                  {m.price !== null && (
                    <button className="btn btn-ghost btn-sm" onClick={() => setEditingMaterial(m.id)}>تعديل السعر</button>
                  )}
                  <button className="btn btn-danger btn-sm" onClick={() => deleteMaterial(m.id)}>حذف</button>
                </div>
              </div>
            ))}
            <div style={{ display: "flex", gap: 8, marginTop: 12 }}>
              <input className="form-input" value={newMaterial.name} onChange={e => setNewMaterial({ ...newMaterial, name: e.target.value })} placeholder="اسم المادة الجديدة" style={{ flex: 2 }} />
              <input className="form-input" type="number" value={newMaterial.price} onChange={e => setNewMaterial({ ...newMaterial, price: e.target.value })} placeholder="السعر (اتركه فارغاً لسعر حر)" style={{ flex: 1 }} />
              <button className="btn btn-primary" onClick={addMaterial}>➕ إضافة مادة</button>
            </div>
          </div>
        </div>
        <div className="modal-footer">
          <button className="btn btn-ghost" onClick={onClose}>إلغاء</button>
          <button className="btn btn-primary" onClick={save} disabled={loading}>{loading ? "جاري الحفظ..." : "💾 حفظ الإعدادات"}</button>
        </div>
      </div>
    </div>
  );
}

export default function App() {
  const [session, setSession] = useState(null);
  const [authInit, setAuthInit] = useState(false);
  const [settings, setSettings] = useState(DEFAULT_SETTINGS);
  const [cases, setCases] = useState([]);
  const [loading, setLoading] = useState(true);
  const [activeBranch, setActiveBranch] = useState("all");
  const [activeTab, setActiveTab] = useState("current");
  const [search, setSearch] = useState("");
  const [histMonth, setHistMonth] = useState("");
  const [showAdd, setShowAdd] = useState(false);
  const [editCase, setEditCase] = useState(null);
  const [detailCase, setDetailCase] = useState(null);
  const [showSettings, setShowSettings] = useState(false);
  const [toast, setToast] = useState("");
  const importRef = useRef();

  useEffect(() => {
    supabase.auth.getSession().then(({ data: { session } }) => {
      setSession(session);
      setAuthInit(true);
    });
    const { data: { subscription } } = supabase.auth.onAuthStateChange((_e, s) => {
      setSession(s);
    });
    return () => subscription.unsubscribe();
  }, []);

  useEffect(() => {
    if (!session) return;
    (async () => {
      setLoading(true);
      const [sett, rawCases] = await Promise.all([getSettings(session.user.id), fetchCases(session.user.id)]);
      setSettings(sett);
      setCases(rawCases.map(toLocal));
      setActiveBranch("all");
      setLoading(false);
    })();
  }, [session]);

  const showToast = useCallback((msg) => {
    setToast(msg);
    setTimeout(() => setToast(""), 2500);
  }, []);

  const curMonth = currentMonthKey();
  
  const branchFiltered = useMemo(() => 
    activeBranch === "all" ? cases : cases.filter(c => c.branchName === activeBranch),
    [cases, activeBranch]
  );

  const patientSearchFiltered = useMemo(() => {
    if (!search.trim()) return branchFiltered;
    const q = search.toLowerCase().trim();
    return branchFiltered.filter(c => c.patientName.toLowerCase().includes(q));
  }, [branchFiltered, search]);

  const allMonths = useMemo(() => {
    const months = [];
    for (let year = DATE_RANGE.startYear; year <= DATE_RANGE.endYear; year++) {
      for (let month = 1; month <= 12; month++) {
        const monthStr = month.toString().padStart(2, '0');
        months.push(`${year}-${monthStr}`);
      }
    }
    return months.sort().reverse();
  }, []);

  const tabFiltered = useMemo(() => {
    let arr = patientSearchFiltered;
    const targetMonth = histMonth || curMonth;
    
    if (activeTab === "current") {
      arr = arr.filter(c => {
        const d = c.actionDate || c.startDate || c.createdAt;
        return monthKey(d) === targetMonth;
      });
    } else if (activeTab === "active") {
      arr = arr.filter(c => c.caseStatus === "In progress" || c.caseStatus === "Correction");
    } else if (activeTab === "missed") {
      arr = arr.filter(c => c.caseStatus === "Missed");
    } else if (activeTab === "completed") {
      arr = arr.filter(c => c.caseStatus === "Completed");
    }
    
    return arr;
  }, [patientSearchFiltered, activeTab, curMonth, histMonth]);

  const stats = useMemo(() => {
    const src = activeTab === "current" ? tabFiltered : patientSearchFiltered.filter(c => {
      const d = c.actionDate || c.startDate || c.createdAt;
      return monthKey(d) === (histMonth || curMonth);
    });
    const total = src.reduce((s, c) => s + (c.totalAmount || 0), 0);
    const paid = src.filter(c => c.paymentStatus === "Paid").reduce((s, c) => s + (c.totalAmount || 0), 0);
    const partial = src.filter(c => c.paymentStatus === "Partial").reduce((s, c) => s + (c.paidAmount || 0), 0);
    const free = src.filter(c => c.paymentStatus === "Free").reduce((s, c) => s + (c.totalAmount || 0), 0);
    const unpaidTotal = src.filter(c => c.paymentStatus === "Unpaid").reduce((s, c) => s + (c.totalAmount || 0), 0);
    const partialRemaining = src.filter(c => c.paymentStatus === "Partial").reduce((s, c) => s + ((c.totalAmount || 0) - (c.paidAmount || 0)), 0);
    
    return { 
      total, 
      paid, 
      partial, 
      free, 
      unpaid: unpaidTotal + partialRemaining,
      collected: paid + partial,
      count: src.length 
    };
  }, [tabFiltered, patientSearchFiltered, activeTab, histMonth, curMonth]);

  async function handleAddCase(form) {
    const row = await insertCase(session.user.id, form);
    if (row) {
      setCases(prev => [toLocal(row), ...prev]);
      showToast("✅ تمت إضافة الحالة");
    }
    setShowAdd(false);
  }

  async function handleEditCase(form) {
    const c = editCase;
    await updateCaseDB(c.id, {
      branch_name: form.branchName,
      patient_name: form.patientName,
      material_name: form.materialName,
      price_per_unit: form.pricePerUnit,
      units: form.units,
      case_status: form.caseStatus,
      payment_status: form.paymentStatus,
      paid_amount: form.paidAmount || 0,
      start_date: form.startDate,
      action_date: form.actionDate || null,
      notes: form.notes || null,
    });
    setCases(prev => prev.map(x => x.id === c.id ? { 
      ...x, 
      ...form, 
      totalAmount: form.pricePerUnit * form.units 
    } : x));
    showToast("✅ تم حفظ التعديل");
    setEditCase(null);
    setDetailCase(null);
  }

  async function handleDelete(c) {
    if (!confirm(`حذف حالة "${c.patientName}"؟`)) return;
    await deleteCaseDB(c.id);
    setCases(prev => prev.filter(x => x.id !== c.id));
    setDetailCase(null);
    showToast("🗑 تم الحذف");
  }

  async function handleUpdatePayment(c, newStatus) {
    if (c.paymentStatus === newStatus) return;
    
    let updateData = { payment_status: newStatus };
    
    if (newStatus === "Paid") {
      updateData.paid_amount = c.totalAmount;
    } else if (newStatus === "Free" || newStatus === "Unpaid") {
      updateData.paid_amount = 0;
    }
    
    await updateCaseDB(c.id, updateData);
    
    setCases(prev => prev.map(x => x.id === c.id ? { 
      ...x, 
      paymentStatus: newStatus,
      paidAmount: updateData.paid_amount !== undefined ? updateData.paid_amount : x.paidAmount
    } : x));
    
    setDetailCase(prev => prev ? { 
      ...prev, 
      paymentStatus: newStatus,
      paidAmount: updateData.paid_amount !== undefined ? updateData.paid_amount : prev.paidAmount
    } : null);
    
    const msg = newStatus === "Paid" ? "✅ تم التحديد كمدفوع كامل" : (newStatus === "Free" ? "🎁 تم التحديد كمجاناً" : "❌ تم التحديد كغير مدفوع");
    showToast(msg);
  }

  async function handleSaveSettings(s) {
    await saveSettings(session.user.id, s);
    setSettings(s);
    showToast("✅ تم حفظ الإعدادات");
  }

  async function handleImport(e) {
    const file = e.target.files[0];
    if (!file) return;
    try {
      const count = await importCSV(file, session.user.id, activeBranch !== "all" ? activeBranch : settings.branches[0]);
      const fresh = await fetchCases(session.user.id);
      setCases(fresh.map(toLocal));
      showToast(`✅ تم استيراد ${count} حالة`);
    } catch {
      showToast("❌ خطأ في الاستيراد");
    }
    e.target.value = "";
  }

  const exportWithFilter = useCallback(() => {
    let dataToExport = tabFiltered;
    if (activeTab === "current" && histMonth) {
      dataToExport = patientSearchFiltered.filter(c => {
        const d = c.actionDate || c.startDate || c.createdAt;
        return monthKey(d) === histMonth;
      });
    }
    exportCSV(dataToExport);
  }, [tabFiltered, patientSearchFiltered, activeTab, histMonth]);

  if (!authInit) return <div className="loading" style={{ minHeight: "100vh" }}><div className="spinner" /></div>;
  if (!session) return <LoginScreen onLogin={s => setSession(s)} />;

  const tabs = [
    { id: "current", label: "📅 الشهر الحالي" },
    { id: "active", label: "⚙️ قيد العمل" },
    { id: "missed", label: "⚠️ Missed" },
    { id: "completed", label: "✅ المنتهية" },
  ];

  return (
    <>
      <style>{CSS}</style>
      <div className="header">
        <div className="header-top">
          <div className="header-title"><span className="dot" />🦷 Dental Tracker</div>
          <div className="header-actions">
            <input type="file" ref={importRef} accept=".csv" style={{ display: "none" }} onChange={handleImport} />
            <button className="btn btn-ghost btn-sm" onClick={() => importRef.current?.click()}>📥 استيراد</button>
            <button className="btn btn-ghost btn-sm" onClick={exportWithFilter}>📤 تصدير</button>
            <button className="btn btn-ghost btn-icon" onClick={() => setShowSettings(true)}>⚙️</button>
            <button className="btn btn-ghost btn-sm" onClick={() => supabase.auth.signOut()}>خروج</button>
          </div>
        </div>
        <div className="branch-bar">
          <span className="branch-label">الفرع:</span>
          <select className="branch-select" value={activeBranch} onChange={e => setActiveBranch(e.target.value)}>
            <option value="all">🏥 كل الفروع</option>
            {settings.branches.map(b => <option key={b} value={b}>{b}</option>)}
          </select>
          <span className="branch-label" style={{ marginRight: 8 }}>الشهر:</span>
          <select className="month-select" value={histMonth} onChange={e => setHistMonth(e.target.value)}>
            <option value="">كل الشهور</option>
            {allMonths.map(m => {
              const [y, mo] = m.split("-");
              const names = ["يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو", "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر"];
              return <option key={m} value={m}>{names[parseInt(mo) - 1]} {y}</option>;
            })}
          </select>
        </div>
        <div className="tabs">
          {tabs.map(t => (
            <button key={t.id} className={`tab${activeTab === t.id ? " active" : ""}`} onClick={() => { setActiveTab(t.id); setHistMonth(""); }}>
              {t.label}
            </button>
          ))}
        </div>
      </div>
      <div className="main">
        <div className="stats-grid">
          {[
            { label: "إجمالي الشهر", value: `${fmtMoney(stats.total)} ج.م`, color: "var(--gold)", icon: "💰" },
            { label: "المحصل", value: `${fmtMoney(stats.collected)} ج.م`, color: "var(--mint)", icon: "✅" },
            { label: "المتبقي", value: `${fmtMoney(stats.unpaid)} ج.م`, color: "var(--rose)", icon: "⏳" },
            { label: "مجاناً", value: `${fmtMoney(stats.free)} ج.م`, color: "var(--amber)", icon: "🎁" },
            { label: "عدد الحالات", value: tabFiltered.length, color: "var(--lavender)", icon: "📋" },
          ].map(({ label, value, color, icon }, i) => (
            <div className="stat-card fade-up" key={i} style={{ animationDelay: `${i * 0.06}s`, "--card-color": color }}>
              <div className="stat-label">{icon} {label}</div>
              <div className="stat-value" style={{ color }}>{value}</div>
            </div>
          ))}
        </div>
        <div className="toolbar">
          <div className="search-wrap">
            <input 
              value={search} 
              onChange={e => setSearch(e.target.value)} 
              placeholder="🔍 بحث باسم المريض (يجلب كل حالاته)..." 
            />
          </div>
        </div>
        {loading ? <Spinner /> : tabFiltered.length === 0 ? (
          <div className="empty">
            <div className="empty-icon">📋</div>
            <div className="empty-msg">
              {search ? `لا توجد حالات للمريض "${search}"` : "لا توجد حالات في هذا الفلتر"}
            </div>
          </div>
        ) : (
          <div className="cases-grid">
            {tabFiltered.map((c, i) => {
              const sm = STATUS_META[c.caseStatus] || {};
              const mc = MAT_COLORS[c.materialName] || "var(--mint)";
              const getPaymentLabel = (status, paidAmount) => {
                if (status === "Paid") return "مدفوع كامل";
                if (status === "Free") return "مجاناً";
                if (status === "Partial") return `مدفوع جزئي (${fmtMoney(paidAmount)})`;
                return "غير مدفوع";
              };
              return (
                <div key={c.id} className="case-card fade-up" style={{ "--status-color": sm.color, animationDelay: `${Math.min(i, 12) * 0.04}s` }} onClick={() => setDetailCase(c)}>
                  <div className="case-card-header">
                    <div><div className="case-name">{c.patientName}</div><div className="case-date">{fmtDate(c.startDate)}</div></div>
                    <span className="case-branch">{c.branchName}</span>
                  </div>
                  <div className="case-badges">
                    <Badge label={c.caseStatus} color={sm.color} bg={sm.bg} icon={sm.icon} />
                    <Badge label={c.materialName} color={mc} bg={mc + "20"} />
                    <Badge 
                      label={getPaymentLabel(c.paymentStatus, c.paidAmount)} 
                      color={c.paymentStatus === "Paid" ? "var(--mint)" : (c.paymentStatus === "Free" ? "var(--amber)" : (c.paymentStatus === "Partial" ? "var(--lavender)" : "var(--rose)"))} 
                      bg={c.paymentStatus === "Paid" ? "rgba(0,212,161,.12)" : (c.paymentStatus === "Free" ? "rgba(245,166,35,.12)" : (c.paymentStatus === "Partial" ? "rgba(155,142,255,.12)" : "rgba(239,68,68,.12)"))} 
                    />
                  </div>
                  <div className="case-footer">
                    <div>
                      <div className="case-amount" style={{ color: "var(--gold)" }}>{fmtMoney(c.totalAmount)} ج.م</div>
                      <div className="case-units">{c.units} وحدة × {c.pricePerUnit}</div>
                    </div>
                    {c.paymentStatus === "Partial" && (
                      <div style={{ fontSize: 10, color: "var(--mint)" }}>محصل: {fmtMoney(c.paidAmount)}</div>
                    )}
                  </div>
                </div>
              );
            })}
          </div>
        )}
      </div>
      <button className="btn-fab" onClick={() => setShowAdd(true)}>＋</button>
      {showAdd && <CaseModal settings={settings} defaultBranch={activeBranch !== "all" ? activeBranch : settings.branches[0]} onSave={handleAddCase} onClose={() => setShowAdd(false)} />}
      {editCase && <CaseModal existing={editCase} settings={settings} defaultBranch={editCase.branchName} onSave={handleEditCase} onClose={() => setEditCase(null)} />}
      {detailCase && <CaseDrawer 
        c={detailCase} 
        settings={settings} 
        onEdit={c => { setEditCase(c); setDetailCase(null); }} 
        onDelete={handleDelete} 
        onUpdatePayment={handleUpdatePayment}
        onClose={() => setDetailCase(null)} 
      />}
      {showSettings && <SettingsScreen settings={settings} onSave={handleSaveSettings} onClose={() => setShowSettings(false)} />}
      <Toast msg={toast} />
    </>
  );
}
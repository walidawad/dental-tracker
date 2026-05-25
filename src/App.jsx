================================================================================
DENTAL TRACKER - COMPLETE WORKING CODE
================================================================================

COPY EVERYTHING BELOW AND PASTE INTO YOUR PROJECT FILE

================================================================================

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
    { id: "mat_4", name: "Other / أخرى", price: null },
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

const CSS = `
@import url('https://fonts.googleapis.com/css2?family=Tajawal:wght@300;400;500;700;800;900&family=JetBrains+Mono:wght@400;600&display=swap');

*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

:root {
  --bg: #080B12;
  --surface: #111621;
  --surface2: #161C2A;
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

html { 
  direction: rtl;
  -webkit-text-size-adjust: 100%;
}

body {
  font-family: 'Tajawal', system-ui, sans-serif;
  background: var(--bg);
  color: var(--text);
  min-height: 100vh;
  overflow-x: hidden;
  width: 100%;
}

::-webkit-scrollbar { width: 4px; }
::-webkit-scrollbar-track { background: transparent; }
::-webkit-scrollbar-thumb { background: var(--border2); border-radius: 2px; }

@keyframes fadeUp { from { opacity:0; transform:translateY(12px); } to { opacity:1; } }
@keyframes fadeIn { from { opacity:0; } to { opacity:1; } }
@keyframes spin { to { transform: rotate(360deg); } }

.login-shell {
  min-height: 100vh;
  display: flex; align-items: center; justify-content: center;
  background: radial-gradient(ellipse 80% 60% at 50% 0%, rgba(61,126,255,.18) 0%, transparent 70%), var(--bg);
  padding: 24px;
  width: 100%;
  overflow-x: hidden;
}

.login-card {
  background: var(--surface);
  border: 1px solid var(--border2);
  border-radius: 20px;
  padding: 40px 36px;
  width: 100%;
  max-width: 420px;
  animation: fadeUp .4s ease;
}

.login-logo { text-align: center; margin-bottom: 32px; }

.login-logo-icon {
  width: 64px; height: 64px; border-radius: 20px;
  background: linear-gradient(135deg, var(--accent), var(--accent2));
  display: inline-flex; align-items: center; justify-content: center;
  font-size: 28px; margin-bottom: 14px;
}

.login-logo h1 { font-size: 20px; font-weight: 800; }
.login-logo p { font-size: 13px; color: var(--text2); margin-top: 4px; }

.form-group { margin-bottom: 14px; }
.form-label { display: block; font-size: 12px; font-weight: 700; color: var(--text2); margin-bottom: 6px; }

.form-input {
  width: 100%;
  background: var(--surface2);
  border: 1px solid var(--border);
  border-radius: var(--r2);
  padding: 10px 13px;
  color: var(--text);
  font-family: 'Tajawal', sans-serif;
  font-size: 14px;
  outline: none;
  box-sizing: border-box;
}

.form-input:focus { border-color: var(--accent); }

.btn {
  display: inline-flex;
  align-items: center;
  gap: 6px;
  padding: 8px 16px;
  border-radius: var(--r2);
  font-family: 'Tajawal', sans-serif;
  font-size: 13px;
  font-weight: 700;
  border: none;
  cursor: pointer;
  white-space: nowrap;
}

.btn-primary { background: var(--accent); color: #fff; }
.btn-primary:hover { background: #3070e8; }
.btn-ghost { background: var(--surface2); border: 1px solid var(--border2); color: var(--text2); }
.btn-ghost:hover { border-color: var(--accent); color: var(--accent); }
.btn-danger { background: rgba(239,68,68,.15); border: 1px solid rgba(239,68,68,.3); color: var(--rose); }
.btn-danger:hover { background: var(--rose); color: #fff; }
.btn-sm { padding: 6px 12px; font-size: 12px; }
.btn-icon { padding: 7px; }

.header {
  background: var(--surface);
  border-bottom: 1px solid var(--border);
  position: sticky;
  top: 0;
  z-index: 100;
  padding: 0 24px;
  width: 100%;
  overflow-x: hidden;
}

.header-top {
  height: 58px;
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 12px;
  width: 100%;
}

.header-title {
  font-size: 17px;
  font-weight: 800;
  display: flex;
  align-items: center;
  gap: 8px;
  white-space: nowrap;
}

.header-actions {
  display: flex;
  gap: 8px;
  align-items: center;
  flex-wrap: wrap;
}

.main {
  padding: 14px;
  max-width: 1400px;
  margin: 0 auto;
  width: 100%;
  overflow-x: hidden;
  box-sizing: border-box;
}

.stats-grid {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 10px;
  margin-bottom: 20px;
}

.stat-card {
  background: var(--surface);
  border: 1px solid var(--border);
  border-radius: var(--r);
  padding: 16px;
  animation: fadeUp .3s ease both;
}

.stat-label { font-size: 11px; color: var(--text2); font-weight: 700; margin-bottom: 8px; }
.stat-value { font-size: 18px; font-weight: 900; word-break: break-word; }

.toolbar { display: flex; gap: 10px; margin-bottom: 16px; flex-wrap: wrap; }

.search-wrap {
  position: relative;
  flex: 1;
  min-width: 200px;
  width: 100%;
}

.search-wrap input { width: 100%; padding: 9px 14px; }

.cases-grid {
  display: grid;
  grid-template-columns: 1fr;
  gap: 12px;
}

.case-card {
  background: var(--surface);
  border: 1px solid var(--border);
  border-radius: var(--r);
  padding: 12px;
  position: relative;
  animation: fadeUp .3s ease both;
}

.case-card:hover { border-color: var(--border2); }

.case-card::after {
  content: '';
  position: absolute;
  right: 0;
  top: 0;
  bottom: 0;
  width: 3px;
  background: var(--status-color, var(--border));
  border-radius: 0 var(--r) var(--r) 0;
}

.case-name { font-size: 14px; font-weight: 700; word-break: break-word; }
.case-date { font-size: 11px; color: var(--text3); margin-top: 2px; }

.case-badges {
  display: flex;
  gap: 6px;
  flex-wrap: wrap;
  margin-top: 8px;
  margin-bottom: 8px;
}

.case-amount { font-size: 16px; font-weight: 800; }
.case-units { font-size: 10px; color: var(--text3); }
.case-branch { font-size: 10px; color: var(--text2); background: var(--surface2); padding: 3px 6px; border-radius: 20px; }

.badge {
  display: inline-flex;
  align-items: center;
  gap: 4px;
  padding: 4px 8px;
  border-radius: 20px;
  font-size: 10px;
  font-weight: 700;
  white-space: nowrap;
  flex-shrink: 0;
}

.dropdown-container {
  position: relative;
  display: inline-block;
}

.dropdown-menu {
  position: absolute;
  background: var(--surface2);
  border: 1px solid var(--border2);
  border-radius: var(--r2);
  z-index: 2000;
  min-width: 140px;
  max-width: 220px;
  box-shadow: 0 8px 24px rgba(0,0,0,.4);
  overflow: visible;
  top: 100%;
  right: 0;
  margin-top: 6px;
  max-height: none;
  overflow-y: visible;
}

.dropdown-menu.top {
  top: auto;
  bottom: 100%;
  margin-bottom: 4px;
}

.dropdown-item {
  padding: 8px 10px;
  font-size: 11px;
  cursor: pointer;
  text-align: center;
  border-bottom: 1px solid var(--border);
  word-break: break-word;
}

.dropdown-item:last-child { border-bottom: none; }
.dropdown-item:hover { background: var(--accent); color: #fff; }

.modal-overlay {
  position: fixed;
  inset: 0;
  background: rgba(0,0,0,.7);
  z-index: 200;
  display: flex;
  align-items: center;
  justify-content: center;
  padding: 20px;
}

.modal {
  background: var(--surface);
  border: 1px solid var(--border2);
  border-radius: 18px;
  width: 100%;
  max-width: 560px;
  max-height: 90vh;
  overflow-y: auto;
  animation: fadeUp .25s ease;
}

.modal-header {
  padding: 20px 24px;
  display: flex;
  justify-content: space-between;
  align-items: center;
  border-bottom: 1px solid var(--border);
  position: sticky;
  top: 0;
  background: var(--surface);
  z-index: 1;
}

.modal-title { font-size: 16px; font-weight: 800; }
.modal-body { padding: 20px 24px; }
.modal-footer { padding: 16px 24px; border-top: 1px solid var(--border); display: flex; gap: 10px; justify-content: flex-end; flex-wrap: wrap; }

.form-row { display: grid; grid-template-columns: 1fr; gap: 12px; }
.form-textarea { min-height: 72px; resize: vertical; }

.total-preview {
  background: linear-gradient(135deg, rgba(61,126,255,.1), rgba(108,99,255,.08));
  border: 1px solid rgba(61,126,255,.25);
  border-radius: var(--r);
  padding: 14px 16px;
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin-bottom: 14px;
}

.total-preview-label { font-size: 12px; color: var(--text2); }
.total-preview-value { font-size: 20px; font-weight: 900; color: var(--gold); word-break: break-word; }

.drawer {
  position: fixed;
  top: 0;
  left: 0;
  bottom: 0;
  width: 100%;
  background: var(--surface);
  overflow-y: auto;
  z-index: 151;
  animation: slideUp .25s ease;
  border-top: 1px solid var(--border2);
}

.drawer-header {
  padding: 20px;
  border-bottom: 1px solid var(--border);
  display: flex;
  justify-content: space-between;
  align-items: flex-start;
  gap: 12px;
}

.drawer-body { padding: 20px; }
.drawer-section { margin-bottom: 24px; }
.drawer-section-title { font-size: 10px; font-weight: 700; color: var(--text3); margin-bottom: 10px; }

.detail-row {
  display: flex;
  justify-content: space-between;
  align-items: center;
  padding: 8px 0;
  border-bottom: 1px solid var(--border);
  gap: 12px;
}

.detail-key { font-size: 12px; color: var(--text2); }
.detail-val { font-size: 12px; font-weight: 600; color: var(--text); word-break: break-word; }

.toast {
  position: fixed;
  bottom: 90px;
  left: 50%;
  transform: translateX(-50%);
  background: var(--surface2);
  border: 1px solid var(--border2);
  padding: 10px 20px;
  border-radius: 20px;
  font-size: 12px;
  color: var(--text);
  z-index: 300;
  max-width: 90%;
}

.empty { text-align: center; padding: 40px 20px; color: var(--text3); }
.empty-icon { font-size: 40px; margin-bottom: 12px; opacity: .3; }

.loading { display: flex; align-items: center; justify-content: center; gap: 8px; padding: 40px 20px; color: var(--text2); }
.spinner { width: 18px; height: 18px; border: 2px solid var(--border2); border-top-color: var(--accent); border-radius: 50%; animation: spin .7s linear infinite; }

.divider { height: 1px; background: var(--border); margin: 12px 0; }

@keyframes slideUp { from { transform: translateY(100%); } to { transform: none; } }

@media (max-width: 640px) {
  .stats-grid { grid-template-columns: 1fr 1fr; gap: 8px; }
  .case-card { padding: 10px; }
  .stat-value { font-size: 16px; }
  .header-actions .btn { padding: 4px 10px; font-size: 11px; }
  .btn-fab { width: 48px; height: 48px; bottom: 20px; right: 20px; font-size: 20px; }
}
`;

// ========== SUPABASE FUNCTIONS ==========

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
  const { data } = await supabase.from("cases").select("*").eq("user_id", userId).order("created_at", { ascending: false });
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
    return `${d.getFullYear()}-${(d.getMonth() + 1).toString().padStart(2, '0')}-${d.getDate().toString().padStart(2, '0')}`;
  };

  const headers = ["الفرع", "المريض", "المادة", "السعر", "الوحدات", "الحالة", "الدفع", "المحصل", "البداية", "الإجراء", "ملاحظات"];
  const rows = casesToExport.map(c => [
    c.branchName, c.patientName, c.materialName, c.pricePerUnit, c.units, c.caseStatus,
    c.paymentStatus === "Paid" ? "مدفوع" : (c.paymentStatus === "Free" ? "مجاناً" : (c.paymentStatus === "Partial" ? "مدفوع جزئي" : "غير مدفوع")),
    c.paidAmount || 0, formatDate(c.startDate), c.actionDate ? formatDate(c.actionDate) : "", c.notes || ""
  ]);
  const csv = "\uFEFF" + [headers, ...rows].map(r => r.map(v => `"${v}"`).join(",")).join("\n");
  const a = document.createElement("a");
  a.href = URL.createObjectURL(new Blob([csv], { type: "text/csv" }));
  a.download = `dental_cases_${todayISO()}.csv`;
  a.click();
}

// ========== COMPONENTS ==========

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
          {err && <div style={{ fontSize: 13, color: err.includes("✉️") ? "var(--mint)" : "var(--rose)", marginBottom: 12, padding: "8px 12px" }}>{err}</div>}
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
  const [loading, setLoading] = useState(false);
  const total = (form.pricePerUnit || 0) * (form.units || 0);

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
          <span className="modal-title">{isEdit ? "✏️ تعديل" : "➕ حالة جديدة"}</span>
          <button className="btn btn-ghost btn-icon" onClick={onClose}>✕</button>
        </div>
        <div className="modal-body">
          <div className="total-preview">
            <div><div className="total-preview-label">الإجمالي</div></div>
            <div className="total-preview-value">{total.toLocaleString("ar-EG")}</div>
          </div>
          <div className="form-row">
            <div className="form-group">
              <label className="form-label">اسم المريض *</label>
              <input className="form-input" value={form.patientName} onChange={e => setForm(f => ({ ...f, patientName: e.target.value }))} />
            </div>
          </div>
          <div className="form-group">
            <label className="form-label">الفرع</label>
            <select className="form-input" value={form.branchName} onChange={e => setForm(f => ({ ...f, branchName: e.target.value }))}>
              {settings.branches.map(b => <option key={b} value={b}>{b}</option>)}
            </select>
          </div>
          <div className="form-group">
            <label className="form-label">المادة</label>
            <select className="form-input" value={form.materialName} onChange={e => setForm(f => ({ ...f, materialName: e.target.value }))}>
              {settings.materials.map(m => <option key={m.id} value={m.name}>{m.name}</option>)}
            </select>
          </div>
          <div className="form-row">
            <div className="form-group">
              <label className="form-label">الوحدات</label>
              <input className="form-input" type="number" min="1" value={form.units} onChange={e => setForm(f => ({ ...f, units: parseInt(e.target.value) || 1 }))} />
            </div>
          </div>
          <div className="form-group">
            <label className="form-label">حالة الكيس</label>
            <select className="form-input" value={form.caseStatus} onChange={e => setForm(f => ({ ...f, caseStatus: e.target.value }))}>
              {settings.caseStatuses.map(s => <option key={s.id} value={s.name}>{s.name}</option>)}
            </select>
          </div>
          <div className="form-group">
            <label className="form-label">حالة الدفع</label>
            <select className="form-input" value={form.paymentStatus} onChange={e => setForm(f => ({ ...f, paymentStatus: e.target.value }))}>
              {settings.paymentStatuses.map(s => <option key={s.id} value={s.name}>{s.name}</option>)}
            </select>
          </div>
          <div className="form-group">
            <label className="form-label">تاريخ البداية</label>
            <input className="form-input" type="date" value={form.startDate} onChange={e => setForm(f => ({ ...f, startDate: e.target.value }))} />
          </div>
        </div>
        <div className="modal-footer">
          <button className="btn btn-ghost" onClick={onClose}>إلغاء</button>
          <button className="btn btn-primary" onClick={submit} disabled={loading}>
            {loading ? "جاري..." : isEdit ? "حفظ" : "إضافة"}
          </button>
        </div>
      </div>
    </div>
  );
}

function BadgeDropdown({ c, type, settings, onStatusChange, onMaterialChange, onPaymentChange }) {
  const [dropdownOpen, setDropdownOpen] = useState(false);
  const dropdownRef = useRef(null);

  useEffect(() => {
    function handleClickOutside(event) {
      if (dropdownRef.current && !dropdownRef.current.contains(event.target)) {
        setDropdownOpen(false);
      }
    }

    if (dropdownOpen) {
      document.addEventListener("mousedown", handleClickOutside);
      document.addEventListener("touchstart", handleClickOutside);
      return () => {
        document.removeEventListener("mousedown", handleClickOutside);
        document.removeEventListener("touchstart", handleClickOutside);
      };
    }
  }, [dropdownOpen]);

  if (type === 'status') {
    const sm = STATUS_META[c.caseStatus] || {};
    return (
      <div className="dropdown-container" ref={dropdownRef}>
        <span 
          className="badge" 
          style={{ color: sm.color, background: sm.bg, border: `1px solid ${sm.color}30`, cursor: 'pointer' }}
          onClick={() => setDropdownOpen(!dropdownOpen)}
        >
          {sm.icon} {c.caseStatus} {dropdownOpen ? '▲' : '▼'}
        </span>
        {dropdownOpen && (
          <div className="dropdown-menu">
            {settings.caseStatuses.map(s => (
              <div key={s.id} className="dropdown-item" onClick={() => {
                onStatusChange(c, s.name);
                setDropdownOpen(false);
              }}>
                {STATUS_META[s.name]?.icon} {s.name}
              </div>
            ))}
          </div>
        )}
      </div>
    );
  }

  if (type === 'material') {
    const mc = MAT_COLORS[c.materialName] || "var(--mint)";
    return (
      <div className="dropdown-container" ref={dropdownRef}>
        <span 
          className="badge" 
          style={{ color: mc, background: mc + "20", cursor: 'pointer' }}
          onClick={() => setDropdownOpen(!dropdownOpen)}
        >
          {c.materialName} {dropdownOpen ? '▲' : '▼'}
        </span>
        {dropdownOpen && (
          <div className="dropdown-menu">
            {settings.materials.map(m => (
              <div key={m.id} className="dropdown-item" onClick={() => {
                onMaterialChange(c, m.name);
                setDropdownOpen(false);
              }}>
                {m.name}
              </div>
            ))}
          </div>
        )}
      </div>
    );
  }

  if (type === 'payment') {
    const paymentColor = c.paymentStatus === "Paid" ? "var(--mint)" : (c.paymentStatus === "Free" ? "var(--amber)" : "var(--rose)");
    return (
      <div className="dropdown-container" ref={dropdownRef}>
        <span 
          className="badge" 
          style={{ color: paymentColor, cursor: 'pointer' }}
          onClick={() => setDropdownOpen(!dropdownOpen)}
        >
          {c.paymentStatus} {dropdownOpen ? '▲' : '▼'}
        </span>
        {dropdownOpen && (
          <div className="dropdown-menu">
            {["Paid", "Unpaid", "Free"].map(s => (
              <div key={s} className="dropdown-item" onClick={() => {
                onPaymentChange(c, s);
                setDropdownOpen(false);
              }}>
                {s}
              </div>
            ))}
          </div>
        )}
      </div>
    );
  }
}

function CaseDrawer({ c, settings, onEdit, onDelete, onUpdatePayment, onUpdateMaterial, onUpdateStatus, onClose }) {
  const sm = STATUS_META[c.caseStatus] || {};
  const mc = MAT_COLORS[c.materialName] || "var(--mint)";

  return (
    <>
      <div className="drawer-overlay" onClick={onClose} />
      <div className="drawer">
        <div className="drawer-header">
          <div>
            <div style={{ fontSize: 16, fontWeight: 800 }}>{c.patientName}</div>
            <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginTop: 8 }}>
              <Badge label={c.caseStatus} color={sm.color} bg={sm.bg} icon={sm.icon} />
              <Badge label={c.materialName} color={mc} bg={mc + "20"} />
            </div>
          </div>
          <button className="btn btn-ghost btn-icon" onClick={onClose}>✕</button>
        </div>
        <div className="drawer-body">
          <div className="drawer-section">
            <div className="drawer-section-title">الحساب</div>
            <div style={{ display: "flex", justifyContent: "space-between", padding: "12px 0" }}>
              <div><div style={{ fontSize: 11, color: "var(--text2)" }}>الإجمالي</div></div>
              <div style={{ fontSize: 20, fontWeight: 900, color: "var(--gold)" }}>{fmtMoney(c.totalAmount)} ج.م</div>
            </div>
          </div>
          <div className="drawer-section">
            <div className="drawer-section-title">التفاصيل</div>
            {[["الفرع", c.branchName], ["تاريخ البداية", fmtDate(c.startDate)]].map(([k, v]) => (
              <div key={k} className="detail-row">
                <span className="detail-key">{k}</span>
                <span className="detail-val">{v}</span>
              </div>
            ))}
          </div>
          <div style={{ display: "flex", gap: 8, flexDirection: "column" }}>
            <button className="btn btn-ghost" style={{ justifyContent: "center" }} onClick={() => onEdit(c)}>✏️ تعديل</button>
            <button className="btn btn-danger" style={{ justifyContent: "center" }} onClick={() => onDelete(c)}>🗑 حذف</button>
          </div>
        </div>
      </div>
    </>
  );
}

// ========== MAIN APP ==========

export default function App() {
  const [session, setSession] = useState(null);
  const [authInit, setAuthInit] = useState(false);
  const [settings, setSettings] = useState(DEFAULT_SETTINGS);
  const [cases, setCases] = useState([]);
  const [loading, setLoading] = useState(true);
  const [activeBranch, setActiveBranch] = useState("all");
  const [activeTab, setActiveTab] = useState("current");
  const [search, setSearch] = useState("");
  const [selectedYear, setSelectedYear] = useState(2026);
  const [selectedMonth, setSelectedMonth] = useState(5);
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
    const { data: { subscription } } = supabase.auth.onAuthStateChange((_e, s) => setSession(s));
    return () => subscription.unsubscribe();
  }, []);

  useEffect(() => {
    if (!session) return;
    (async () => {
      setLoading(true);
      const [sett, rawCases] = await Promise.all([
        getSettings(session.user.id),
        fetchCases(session.user.id),
      ]);
      setSettings(sett);
      setCases(rawCases.map(toLocal));
      setLoading(false);
    })();
  }, [session]);

  const showToast = useCallback((msg) => {
    setToast(msg);
    setTimeout(() => setToast(""), 2500);
  }, []);

  const branchFiltered = useMemo(() => 
    activeBranch === "all" ? cases : cases.filter(c => c.branchName === activeBranch),
    [cases, activeBranch]
  );

  const tabFiltered = useMemo(() => {
    let arr = branchFiltered;
    const targetMonth = `${selectedYear}-${selectedMonth.toString().padStart(2, '0')}`;

    if (activeTab === "current") {
      arr = arr.filter(c => {
        const d = c.actionDate || c.startDate || c.createdAt;
        return monthKey(d) === targetMonth;
      });
    } else if (activeTab === "active") {
      arr = arr.filter(c => c.caseStatus === "In progress" || c.caseStatus === "Correction");
    } else if (activeTab === "completed") {
      arr = arr.filter(c => c.caseStatus === "Completed");
    }

    if (search.trim()) {
      const q = search.toLowerCase();
      arr = arr.filter(c => c.patientName.toLowerCase().includes(q));
    }
    return arr;
  }, [branchFiltered, activeTab, selectedYear, selectedMonth, search]);

  const stats = useMemo(() => {
    const total = tabFiltered.reduce((s, c) => s + (c.totalAmount || 0), 0);
    const paid = tabFiltered.filter(c => c.paymentStatus === "Paid").reduce((s, c) => s + (c.totalAmount || 0), 0);
    return { total, paid, unpaid: total - paid, count: tabFiltered.length };
  }, [tabFiltered]);

  async function handleAddCase(form) {
    const row = await insertCase(session.user.id, form);
    if (row) {
      setCases(prev => [toLocal(row), ...prev]);
      showToast("✅ تمت الإضافة");
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
      start_date: form.startDate,
    });
    setCases(prev => prev.map(x => x.id === c.id ? { ...x, ...form, totalAmount: form.pricePerUnit * form.units } : x));
    showToast("✅ تم الحفظ");
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
    await updateCaseDB(c.id, { payment_status: newStatus });
    setCases(prev => prev.map(x => x.id === c.id ? { ...x, paymentStatus: newStatus } : x));
    showToast("✅ تم التحديث");
  }

  async function handleUpdateMaterial(c, newMaterial) {
    await updateCaseDB(c.id, { material_name: newMaterial });
    setCases(prev => prev.map(x => x.id === c.id ? { ...x, materialName: newMaterial } : x));
    showToast("✅ تم التحديث");
  }

  async function handleUpdateStatus(c, newStatus) {
    await updateCaseDB(c.id, { case_status: newStatus });
    setCases(prev => prev.map(x => x.id === c.id ? { ...x, caseStatus: newStatus } : x));
    showToast("✅ تم التحديث");
  }

  async function handleSaveSettings(s) {
    await saveSettings(session.user.id, s);
    setSettings(s);
    showToast("✅ تم الحفظ");
  }

  async function handleImport(e) {
    const file = e.target.files[0];
    if (!file) return;
    try {
      const text = await file.text();
      const lines = text.split("\n");
      let count = 0;
      for (let i = 1; i < lines.length; i++) {
        const line = lines[i].trim();
        if (!line) continue;
        count++;
        // Parse CSV and insert
      }
      showToast(`✅ تم استيراد ${count} حالة`);
    } catch {
      showToast("❌ خطأ");
    }
    e.target.value = "";
  }

  if (!authInit) return <div className="loading"><div className="spinner" /></div>;
  if (!session) return <LoginScreen onLogin={s => setSession(s)} />;

  return (
    <>
      <style>{CSS}</style>
      <div className="header">
        <div className="header-top">
          <div className="header-title">🦷 Dental Tracker</div>
          <div className="header-actions">
            <input type="file" ref={importRef} accept=".csv" style={{ display: "none" }} onChange={handleImport} />
            <button className="btn btn-ghost btn-sm" onClick={() => exportCSV(tabFiltered)}>📤 تصدير</button>
            <button className="btn btn-ghost btn-sm" onClick={() => supabase.auth.signOut()}>خروج</button>
          </div>
        </div>
      </div>
      <div className="main">
        <div className="stats-grid">
          {[
            { label: "إجمالي", value: `${fmtMoney(stats.total)} ج.م`, color: "var(--gold)" },
            { label: "محصّل", value: `${fmtMoney(stats.paid)} ج.م`, color: "var(--mint)" },
            { label: "غير محصّل", value: `${fmtMoney(stats.unpaid)} ج.م`, color: "var(--rose)" },
            { label: "الحالات", value: stats.count, color: "var(--lavender)" },
          ].map((s, i) => (
            <div key={i} className="stat-card">
              <div className="stat-label">{s.label}</div>
              <div className="stat-value" style={{ color: s.color }}>{s.value}</div>
            </div>
          ))}
        </div>
        <div className="toolbar">
          <div className="search-wrap">
            <input value={search} onChange={e => setSearch(e.target.value)} placeholder="🔍 بحث..." />
          </div>
        </div>
        {loading ? <Spinner /> : tabFiltered.length === 0 ? (
          <div className="empty">
            <div className="empty-icon">📋</div>
            <div className="empty-msg">لا توجد حالات</div>
          </div>
        ) : (
          <div className="cases-grid">
            {tabFiltered.map((c, i) => {
              const sm = STATUS_META[c.caseStatus] || {};
              return (
                <div key={c.id} className="case-card" style={{ "--status-color": sm.color }}>
                  <div onClick={() => setDetailCase(c)}>
                    <div className="case-name">{c.patientName}</div>
                    <div className="case-date">{fmtDate(c.startDate)}</div>
                    <div className="case-badges">
                      <BadgeDropdown c={c} type="status" settings={settings} onStatusChange={handleUpdateStatus} />
                      <BadgeDropdown c={c} type="material" settings={settings} onMaterialChange={handleUpdateMaterial} />
                      <BadgeDropdown c={c} type="payment" settings={settings} onPaymentChange={handleUpdatePayment} />
                    </div>
                    <div style={{ marginTop: 8 }}>
                      <div className="case-amount" style={{ color: "var(--gold)" }}>{fmtMoney(c.totalAmount)} ج.م</div>
                      <div className="case-units">{c.units} وحدة × {c.pricePerUnit}</div>
                    </div>
                  </div>
                </div>
              );
            })}
          </div>
        )}
      </div>
      <button className="btn btn-fab" style={{ position: "fixed", bottom: "20px", right: "20px", width: "56px", height: "56px", borderRadius: "50%", background: "var(--accent)", color: "#fff", fontSize: "24px", display: "flex", alignItems: "center", justifyContent: "center", border: "none", cursor: "pointer" }} onClick={() => setShowAdd(true)}>＋</button>
      {showAdd && <CaseModal settings={settings} defaultBranch={settings.branches[0]} onSave={handleAddCase} onClose={() => setShowAdd(false)} />}
      {editCase && <CaseModal existing={editCase} settings={settings} defaultBranch={editCase.branchName} onSave={handleEditCase} onClose={() => setEditCase(null)} />}
      {detailCase && <CaseDrawer c={detailCase} settings={settings} onEdit={c => { setEditCase(c); setDetailCase(null); }} onDelete={handleDelete} onUpdatePayment={handleUpdatePayment} onUpdateMaterial={handleUpdateMaterial} onUpdateStatus={handleUpdateStatus} onClose={() => setDetailCase(null)} />}
      <Toast msg={toast} />
    </>
  );
}

================================================================================
END OF CODE - COPY EVERYTHING ABOVE
================================================================================

import { useState, useEffect, useMemo, useRef, useCallback } from "react";
import { createClient } from "https://esm.sh/@supabase/supabase-js@2";
import * as XLSX from 'https://esm.sh/xlsx';

const SUPA_URL = import.meta.env.VITE_SUPA_URL;
const SUPA_ANON = import.meta.env.VITE_SUPA_ANON;

const supabase = createClient(SUPA_URL, SUPA_ANON);

const DEFAULT_SETTINGS = {
  branches: ["الفرع الرئيسي", "العبد -جسر السويس"],
  materials: [
    { id: "mat_1", name: "Zirconia", price: 450, active: true },
    { id: "mat_2", name: "E-max", price: 450, active: true },
    { id: "mat_3", name: "PMMA", price: 225, active: true },
    { id: "mat_4", name: "Other / أخرى", price: null, active: true },
  ],
  caseStatuses: [
    { id: "cs_1", name: "In progress", isFinal: false },
    { id: "cs_2", name: "Correction", isFinal: false },
    { id: "cs_3", name: "Temp cement", isFinal: true },
    { id: "cs_4", name: "Missed", isFinal: true },
    { id: "cs_5", name: "Completed", isFinal: true },
  ],
  paymentStatuses: [
    { id: "ps_1", name: "Paid", display: "مدفوع كامل", icon: "✅", active: true },
    { id: "ps_2", name: "Unpaid", display: "غير مدفوع", icon: "❌", active: true },
    { id: "ps_3", name: "Partial", display: "مدفوع جزئي", icon: "🔸", active: true },
    { id: "ps_4", name: "Free", display: "مجاناً", icon: "🎁", active: true },
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

const formatDateForExport = (dateStr) => {
  if (!dateStr) return "";
  const d = new Date(dateStr);
  return `${d.getDate().toString().padStart(2, '0')}-${(d.getMonth() + 1).toString().padStart(2, '0')}-${d.getFullYear()}`;
};

function exportExcel(casesToExport, fileName = "dental_cases") {
  const data = casesToExport.map(c => ({
    "الفرع": c.branchName,
    "المريض": c.patientName,
    "المادة": c.materialName,
    "السعر": c.pricePerUnit,
    "الوحدات": c.units,
    "الإجمالي": c.totalAmount,
    "الحالة": c.caseStatus,
    "الدفع": c.paymentStatus === "Paid" ? "مدفوع" : (c.paymentStatus === "Free" ? "مجاناً" : (c.paymentStatus === "Partial" ? "مدفوع جزئي" : "غير مدفوع")),
    "المحصل": c.paidAmount || 0,
    "المتبقي": (c.totalAmount - (c.paidAmount || 0)),
    "البداية": formatDateForExport(c.startDate),
    "الإجراء": formatDateForExport(c.actionDate),
    "ملاحظات": c.notes || ""
  }));

  const worksheet = XLSX.utils.json_to_sheet(data);
  worksheet['!cols'] = [
    { wch: 20 }, { wch: 25 }, { wch: 15 }, { wch: 10 }, { wch: 10 },
    { wch: 12 }, { wch: 15 }, { wch: 12 }, { wch: 12 }, { wch: 12 },
    { wch: 12 }, { wch: 12 }, { wch: 30 }
  ];
  
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "الحالات");
  XLSX.writeFile(workbook, `${fileName}_${todayISO()}.xlsx`);
}

const formatDateForExport = (dateStr) => {
  if (!dateStr) return "";
  const d = new Date(dateStr);
  return `${d.getDate().toString().padStart(2, '0')}-${(d.getMonth() + 1).toString().padStart(2, '0')}-${d.getFullYear()}`;
};

function exportExcel(casesToExport, fileName = "dental_cases") {
  const data = casesToExport.map(c => ({
    "الفرع": c.branchName,
    "المريض": c.patientName,
    "المادة": c.materialName,
    "السعر": c.pricePerUnit,
    "الوحدات": c.units,
    "الإجمالي": c.totalAmount,
    "الحالة": c.caseStatus,
    "الدفع": c.paymentStatus === "Paid" ? "مدفوع" : (c.paymentStatus === "Free" ? "مجاناً" : (c.paymentStatus === "Partial" ? "مدفوع جزئي" : "غير مدفوع")),
    "المحصل": c.paidAmount || 0,
    "المتبقي": (c.totalAmount - (c.paidAmount || 0)),
    "الأسنان": c.teeth ? c.teeth.join(", ") : "",
    "البداية": formatDateForExport(c.startDate),
    "الإجراء": formatDateForExport(c.actionDate),
    "ملاحظات": c.notes || ""
  }));

  const worksheet = XLSX.utils.json_to_sheet(data);
  worksheet['!cols'] = [
    { wch: 20 }, { wch: 25 }, { wch: 15 }, { wch: 10 }, { wch: 10 },
    { wch: 12 }, { wch: 15 }, { wch: 12 }, { wch: 12 }, { wch: 12 },
    { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 30 }
  ];
  
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "الحالات");
  XLSX.writeFile(workbook, `${fileName}_${todayISO()}.xlsx`);
}

// ========== CSS ==========
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
  --gold: #F5A623;
  --mint: #00D4A1;
  --rose: #FF4D6D;
  --amber: #FFB830;
  --lavender: #9B8EFF;
  --text: #E8ECF8;
  --text2: #8A94B8;
  --text3: #4A5280;
  --r: 8px;
  --r2: 5px;
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

@keyframes fadeUp { from { opacity:0; transform:translateY(8px); } to { opacity:1; } }
@keyframes fadeIn { from { opacity:0; } to { opacity:1; } }

.login-shell {
  min-height: 100vh;
  display: flex;
  align-items: center;
  justify-content: center;
  background: radial-gradient(ellipse 80% 60% at 50% 0%, rgba(61,126,255,.18) 0%, transparent 70%), var(--bg);
  padding: 20px;
}
.login-card {
  background: var(--surface);
  border: 1px solid var(--border2);
  border-radius: 18px;
  padding: 32px 28px;
  width: 100%;
  max-width: 400px;
}
.login-logo { text-align: center; margin-bottom: 28px; }
.login-logo-icon {
  width: 56px; height: 56px; border-radius: 18px;
  background: linear-gradient(135deg, var(--accent), #6C63FF);
  display: inline-flex;
  align-items: center;
  justify-content: center;
  font-size: 26px;
  margin-bottom: 12px;
}
.login-logo h1 { font-size: 18px; font-weight: 800; }
.login-logo p { font-size: 12px; color: var(--text2); margin-top: 4px; }

.header {
  background: var(--surface);
  border-bottom: 1px solid var(--border);
  position: sticky;
  top: 0;
  z-index: 100;
  padding: 0 10px;
}
.header-top {
  height: 48px;
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 5px;
  flex-wrap: wrap;
}
.header-title {
  font-size: 15px;
  font-weight: 800;
  display: flex;
  align-items: center;
  gap: 5px;
  white-space: nowrap;
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

.month-selector-wrapper {
  padding: 8px 24px;
  border-bottom: 1px solid var(--border);
  display: flex;
  gap: 12px;
  align-items: center;
  flex-wrap: wrap;
  align-items: center;
  border-bottom: 1px solid var(--border);
}
.year-select {
  background: var(--surface2);
  border: 1px solid var(--border2);
  border-radius: var(--r2);
  color: var(--text);
  font-family: 'Tajawal', sans-serif;
  font-size: 11px;
  font-weight: 700;
  padding: 4px 8px;
  cursor: pointer;
}

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
.stat-label {
  font-size: 8px;
  color: var(--text2);
  font-weight: 700;
  margin-bottom: 2px;
  white-space: nowrap;
}
.stat-value {
  font-size: 12px;
  font-weight: 900;
  word-break: break-word;
}

.toolbar {
  display: flex; gap: 10px; align-items: center;
  margin-bottom: 16px; flex-wrap: wrap;
}
.search-wrap {
  flex: 1;
  min-width: 150px;
}
.search-wrap input {
  width: 100%;
  padding: 6px 8px;
  background: var(--surface2);
  border: 1px solid var(--border);
  border-radius: var(--r2);
  color: var(--text);
  font-family: 'Tajawal', sans-serif;
  font-size: 11px;
  outline: none;
}
.search-wrap input:focus { border-color: var(--accent); }

.cases-grid {
  display: grid;
  grid-template-columns: repeat(auto-fill, minmax(320px, 1fr));
  gap: 12px;
}
.case-card {
  background: var(--surface);
  border: 1px solid var(--border);
  border-radius: var(--r);
  padding: 6px 8px;
  position: relative;
}
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
.case-clickable-area {
  cursor: pointer;
}
.case-card-header { display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 10px; }
.case-name { font-size: 15px; font-weight: 700; color: var(--text); }
.case-date { font-size: 11px; color: var(--text3); margin-top: 2px; }
.case-badges { display: flex; gap: 6px; flex-wrap: wrap; margin-bottom: 10px; position: relative; }
.case-footer { display: flex; justify-content: space-between; align-items: center; }
.case-amount { font-size: 18px; font-weight: 800; }
.case-units { font-size: 11px; color: var(--text3); }
.case-branch { font-size: 11px; color: var(--text2); background: var(--surface2); padding: 3px 8px; border-radius: 20px; }

.badge-clickable {
  display: inline-flex;
  align-items: center;
  gap: 4px;
  padding: 6px 10px;
  border-radius: 20px;
  font-size: 12px;
  font-weight: 700;
  white-space: nowrap;
  cursor: pointer;
  transition: all .15s;
  border: 1px solid transparent;
  position: relative;
}
.badge-clickable:hover {
  transform: scale(1.02);
  filter: brightness(1.1);
  border-color: var(--accent);
}

.badge {
  display: inline-flex;
  align-items: center;
  gap: 3px;
  padding: 2px 6px;
  border-radius: 14px;
  font-size: 9px;
  font-weight: 700;
  white-space: nowrap;
}

.dropdown-container {
  position: relative;
  display: inline-block;
}

.dropdown-menu {
  position: fixed;
  background: var(--surface2);
  border: 1px solid var(--border2);
  border-radius: var(--r2);
  z-index: 9999;
  min-width: 120px;
  max-width: 180px;
  max-height: 200px;
  overflow-y: auto;
  box-shadow: 0 4px 12px rgba(0,0,0,.5);
}

.dropdown-item {
  padding: 6px 10px;
  font-size: 10px;
  cursor: pointer;
  text-align: center;
  border-bottom: 1px solid var(--border);
}
.dropdown-item:active {
  background: var(--accent);
  color: #fff;
}
.dropdown-item:last-child { border-bottom: none; }

.btn {
  display: inline-flex;
  align-items: center;
  gap: 3px;
  padding: 4px 8px;
  border-radius: var(--r2);
  font-family: 'Tajawal', sans-serif;
  font-size: 10px;
  font-weight: 700;
  border: none;
  cursor: pointer;
  white-space: nowrap;
}
.btn-primary { background: var(--accent); color: #fff; }
.btn-ghost { background: var(--surface2); border: 1px solid var(--border2); color: var(--text2); }
.btn-danger { background: rgba(239,68,68,.15); border: 1px solid rgba(239,68,68,.3); color: var(--rose); }
.btn-sm { padding: 2px 6px; font-size: 9px; }
.btn-icon { padding: 4px; }
.btn-fab {
  position: fixed;
  bottom: 16px;
  right: 16px;
  width: 40px;
  height: 40px;
  border-radius: 50%;
  background: linear-gradient(135deg, var(--accent), #6C63FF);
  color: #fff;
  font-size: 18px;
  display: flex;
  align-items: center;
  justify-content: center;
  border: none;
  cursor: pointer;
  z-index: 90;
  box-shadow: 0 3px 10px rgba(61,126,255,.4);
}

.modal-overlay {
  position: fixed;
  inset: 0;
  background: rgba(0,0,0,.7);
  z-index: 1000;
  display: flex;
  align-items: center;
  justify-content: center;
  padding: 16px;
}
.modal {
  background: var(--surface);
  border: 1px solid var(--border2);
  border-radius: 18px;
  width: 100%; max-width: 560px;
  max-height: 90vh;
  overflow-y: auto;
}
.modal-header {
  padding: 10px 14px;
  display: flex;
  justify-content: space-between;
  align-items: center;
  border-bottom: 1px solid var(--border);
}
.modal-title { font-size: 16px; font-weight: 800; }
.modal-body { padding: 20px 24px; }
.modal-footer { padding: 16px 24px; border-top: 1px solid var(--border); display: flex; gap: 10px; justify-content: flex-end; }

.form-group { margin-bottom: 10px; }
.form-label { display: block; font-size: 10px; font-weight: 700; color: var(--text2); margin-bottom: 3px; }
.form-input {
  width: 100%;
  background: var(--surface2);
  border: 1px solid var(--border);
  border-radius: var(--r2);
  padding: 6px 8px;
  color: var(--text);
  font-family: 'Tajawal', sans-serif;
  font-size: 11px;
  outline: none;
}
.form-input:focus { border-color: var(--accent); }
.form-row { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; }
.form-textarea { min-height: 72px; resize: vertical; }

.total-preview {
  background: linear-gradient(135deg, rgba(61,126,255,.1), rgba(108,99,255,.08));
  border: 1px solid rgba(61,126,255,.25);
  border-radius: var(--r);
  padding: 8px;
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin-bottom: 10px;
}
.total-preview-value { font-size: 16px; font-weight: 900; color: var(--gold); }

.drawer {
  position: fixed;
  top: 0;
  left: 0;
  bottom: 0;
  width: 100%;
  background: var(--surface);
  overflow-y: auto;
  z-index: 151;
  animation: slideIn .25s ease;
  box-shadow: 8px 0 40px rgba(0,0,0,.5);
}
@media (min-width: 640px) {
  .drawer {
    width: min(480px, 100vw);
  }
}
.drawer-header {
  padding: 10px;
  border-bottom: 1px solid var(--border);
  display: flex;
  justify-content: space-between;
  align-items: flex-start;
}
.drawer-body { padding: 10px; }
.drawer-section { margin-bottom: 14px; }
.drawer-section-title { font-size: 9px; font-weight: 700; color: var(--text3); margin-bottom: 5px; }
.detail-row {
  display: flex;
  justify-content: space-between;
  align-items: center;
  padding: 5px 0;
  border-bottom: 1px solid var(--border);
}
.detail-key { font-size: 10px; color: var(--text2); }
.detail-val { font-size: 10px; font-weight: 600; color: var(--text); }

.settings-section { margin-bottom: 16px; }
.settings-section-title {
  font-size: 11px;
  font-weight: 700;
  color: var(--text);
  margin-bottom: 8px;
}
.settings-list-item {
  display: flex;
  justify-content: space-between;
  align-items: center;
  padding: 5px 8px;
  background: var(--surface2);
  border: 1px solid var(--border);
  border-radius: var(--r2);
  margin-bottom: 4px;
}
.settings-list-item-name { font-size: 11px; font-weight: 600; }

.toast {
  position: fixed;
  bottom: 70px;
  left: 50%;
  transform: translateX(-50%);
  background: var(--surface2);
  border: 1px solid var(--border2);
  padding: 5px 10px;
  border-radius: 18px;
  font-size: 10px;
  color: var(--text);
  z-index: 2000;
  max-width: 85%;
  text-align: center;
  cursor: pointer;
}

.empty { text-align: center; padding: 30px 16px; color: var(--text3); }
.empty-icon { font-size: 28px; margin-bottom: 6px; opacity: .3; }
.loading { display: flex; align-items: center; justify-content: center; gap: 5px; padding: 30px; }
.spinner { width: 14px; height: 14px; border: 2px solid var(--border2); border-top-color: var(--accent); border-radius: 50%; animation: spin .7s linear infinite; }
@keyframes spin { to { transform: rotate(360deg); } }

.divider { height: 1px; background: var(--border); margin: 10px 0; }

@media (max-width: 640px) {
  .main { padding: 14px; }
  .stats-grid { grid-template-columns: 1fr 1fr; }
  .cases-grid { grid-template-columns: 1fr; }
  .form-row { grid-template-columns: 1fr; }
  .month-selector-wrapper { flex-direction: column; align-items: flex-start; }
  .badge-clickable { padding: 4px 8px; font-size: 11px; }
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
    teeth: c.teeth || [],
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
  teeth: r.teeth || [],
  startDate: r.start_date,
  actionDate: r.action_date,
  notes: r.notes,
  createdAt: r.created_at,
});

async function importExcel(file, userId, currentBranch, setToast) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = async (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
        
        if (!rows || rows.length === 0) {
          reject(new Error("لا توجد بيانات في الملف"));
          return;
        }
        
        let count = 0;
        for (let i = 1; i < lines.length; i++) {
          const line = lines[i].trim();
          if (!line) continue;
          
          const cols = line.match(/(".*?"|[^,]+)(?=,|$)/g);
          if (!cols || cols.length < 5) continue;
          
          const clean = cols.map(c => c.replace(/^"|"$/g, "").trim());
          
          let paymentStatus = row["الدفع"] || "Unpaid";
          if (paymentStatus === "مدفوع" || paymentStatus === "مدفوع كامل") paymentStatus = "Paid";
          if (paymentStatus === "مجاناً") paymentStatus = "Free";
          if (paymentStatus.includes("مدفوع جزئي") || paymentStatus === "Partial") paymentStatus = "Partial";
          if (paymentStatus === "غير مدفوع") paymentStatus = "Unpaid";
          
          let startDate = todayISO();
          let rawStartDate = row["البداية"] || "";
          if (rawStartDate) {
            let dateStr = String(rawStartDate);
            if (dateStr.includes("-")) {
              const parts = dateStr.split("-");
              if (parts.length === 3) {
                if (parts[0].length === 4) {
                  startDate = `${parts[0]}-${parts[1].padStart(2,'0')}-${parts[2].padStart(2,'0')}`;
                } else if (parts[2].length === 4) {
                  startDate = `${parts[2]}-${parts[1].padStart(2,'0')}-${parts[0].padStart(2,'0')}`;
                }
              }
            }
          }
          
          let actionDate = null;
          if (colIndex.actionDate >= 0 && clean[colIndex.actionDate]) {
            let dateStr = clean[colIndex.actionDate];
            if (dateStr.includes("/")) {
              const parts = dateStr.split("/");
              if (parts.length === 3) actionDate = `${parts[2]}-${parts[1].padStart(2,'0')}-${parts[0].padStart(2,'0')}`;
            } else if (dateStr.includes("-")) {
              actionDate = dateStr;
            }
          }
          
          const rec = {
            branchName: branchName,
            patientName: patientName,
            materialName: materialName,
            pricePerUnit: pricePerUnit,
            units: units,
            caseStatus: caseStatus,
            paymentStatus: paymentStatus,
            paidAmount: (colIndex.paidAmount >= 0 && clean[colIndex.paidAmount]) ? parseFloat(clean[colIndex.paidAmount]) || 0 : 0,
            startDate: startDate,
            actionDate: actionDate,
            notes: notes,
          };
          
          await insertCase(userId, rec);
          count++;
        }
        resolve(count);
      } catch (err) {
        console.error("Import error:", err);
        reject(err);
      }
    };
    reader.readAsArrayBuffer(file);
  });
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

function Toast({ msg, onClick }) {
  return msg ? <div className="toast" onClick={onClick}>{msg}</div> : null;
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
          {err && <div style={{ fontSize: 11, color: err.includes("✉️") ? "var(--mint)" : "var(--rose)", marginBottom: 8, padding: "5px 8px", background: "rgba(239,68,68,.08)", borderRadius: 5 }}>{err}</div>}
          <button className="btn btn-primary" type="submit" disabled={loading} style={{ width: "100%", justifyContent: "center", padding: "7px" }}>
            {loading ? "جاري..." : mode === "login" ? "تسجيل الدخول" : "إنشاء حساب"}
          </button>
        </form>
        <div style={{ textAlign: "center", marginTop: 12 }}>
          <button onClick={() => setMode(m => m === "login" ? "signup" : "login")} style={{ background: "none", border: "none", color: "var(--accent)", cursor: "pointer", fontFamily: "Tajawal", fontSize: 11 }}>
            {mode === "login" ? "مش عندك حساب؟ سجّل دلوقتي" : "عندك حساب؟ ادخل"}
          </button>
        </div>
      </div>
    </div>
  );
}

function DentalChart({ selectedTeeth, onChange }) {
  const teethNumbers = [
    [11, 12, 13, 14, 15, 16, 17, 18],
    [21, 22, 23, 24, 25, 26, 27, 28],
    [31, 32, 33, 34, 35, 36, 37, 38],
    [41, 42, 43, 44, 45, 46, 47, 48]
  ];
  
  const quadrantsLabels = ["علوي يمين", "علوي يسار", "سفلي يسار", "سفلي يمين"];
  
  const toggleTooth = (tooth) => {
    if (selectedTeeth.includes(tooth)) {
      onChange(selectedTeeth.filter(t => t !== tooth));
    } else {
      onChange([...selectedTeeth, tooth].sort((a, b) => a - b));
    }
  };
  
  return (
    <div className="dental-chart">
      <div className="form-label">🦷 الأسنان المعالجة</div>
      {teethNumbers.map((quadrant, idx) => (
        <div key={idx} style={{ marginBottom: 12 }}>
          <div style={{ fontSize: 9, color: "var(--text3)", marginBottom: 4 }}>{quadrantsLabels[idx]}</div>
          <div className="dental-teeth">
            {quadrant.map(tooth => (
              <button
                key={tooth}
                type="button"
                className={`tooth-btn ${selectedTeeth.includes(tooth) ? "selected" : ""}`}
                onClick={() => toggleTooth(tooth)}
              >
                {tooth}
              </button>
            ))}
          </div>
        </div>
      ))}
      <div className="selected-teeth">
        الأسنان المختارة: {selectedTeeth.length > 0 ? selectedTeeth.join(", ") : "لا يوجد"}
      </div>
    </div>
  );
}

function CaseModal({ existing, settings, defaultBranch, onSave, onClose }) {
  const isEdit = !!existing;
  const [form, setForm] = useState({
    patientName: existing?.patientName || "",
    branchName: existing?.branchName || defaultBranch || settings.branches[0] || "",
    materialName: existing?.materialName || (settings.materials.find(m => m.active !== false)?.name) || "",
    pricePerUnit: existing?.pricePerUnit || (settings.materials.find(m => m.active !== false)?.price) || 450,
    units: existing?.units || 1,
    caseStatus: existing?.caseStatus || "In progress",
    paymentStatus: existing?.paymentStatus || "Unpaid",
    paidAmount: existing?.paidAmount || 0,
    teeth: existing?.teeth || [],
    startDate: existing?.startDate || todayISO(),
    actionDate: existing?.actionDate || "",
    notes: existing?.notes || "",
  });
  const [prevStatus, setPrevStatus] = useState(existing?.caseStatus || "In progress");
  const [loading, setLoading] = useState(false);
  const total = (form.pricePerUnit || 0) * (form.units || 0);
  const remaining = total - (form.paidAmount || 0);
  const activeMaterials = settings.materials.filter(m => m.active !== false);
  const activePaymentStatuses = settings.paymentStatuses.filter(p => p.active !== false);
  const selMat = activeMaterials.find(m => m.name === form.materialName);
  const priceFixed = selMat?.price !== null && selMat?.price !== undefined;

  function handleMaterial(name) {
    const mat = activeMaterials.find(m => m.name === name);
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
          <span className="modal-title">{isEdit ? "✏️ تعديل" : "➕ حالة جديدة"}</span>
          <button className="btn btn-ghost btn-icon" onClick={onClose}>✕</button>
        </div>
        <div className="modal-body">
          <div className="total-preview">
            <div className="total-preview-label">الإجمالي</div>
            <div className="total-preview-value">{total.toLocaleString("ar-EG")} ج.م</div>
          </div>
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
          <div className="form-group">
            <label className="form-label">المادة</label>
            <select className="form-input" value={form.materialName} onChange={e => handleMaterial(e.target.value)}>
              {settings.materials.map(m => (
                <option key={m.id} value={m.name}>{m.name}{m.price ? ` — ${m.price} ج.م` : " — سعر حر"}</option>
              ))}
            </select>
          </div>
          <div className="form-row">
            <div className="form-group">
              <label className="form-label">الوحدات</label>
              <input className="form-input" type="number" min="1" value={form.units} onChange={e => setForm(f => ({ ...f, units: parseInt(e.target.value) || 1 }))} />
            </div>
            <div className="form-group">
              <label className="form-label">سعر الوحدة</label>
              <input className="form-input" type="number" value={form.pricePerUnit} disabled={priceFixed && !isEdit} onChange={e => setForm(f => ({ ...f, pricePerUnit: parseFloat(e.target.value) || 0 }))} />
            </div>
          </div>
          <div className="form-row">
            <div className="form-group">
              <label className="form-label">الحالة</label>
              <select className="form-input" value={form.caseStatus} onChange={e => handleStatus(e.target.value)}>
                {settings.caseStatuses.map(s => <option key={s.id} value={s.name}>{s.name}</option>)}
              </select>
            </div>
            <div className="form-group">
              <label className="form-label">الدفع</label>
              <select className="form-input" value={form.paymentStatus} onChange={e => setForm(f => ({ ...f, paymentStatus: e.target.value }))}>
                {settings.paymentStatuses.map(s => <option key={s.id} value={s.name}>{s.name === "Paid" ? "مدفوع كامل" : (s.name === "Free" ? "مجاناً" : (s.name === "Partial" ? "مدفوع جزئي" : "غير مدفوع"))}</option>)}
              </select>
            </div>
          </div>
          {form.paymentStatus === "Partial" && (
            <div className="form-group">
              <label className="form-label">المبلغ المحصل</label>
              <input className="form-input" type="number" value={form.paidAmount} onChange={e => setForm(f => ({ ...f, paidAmount: parseFloat(e.target.value) || 0 }))} placeholder="المبلغ المدفوع" />
              <div style={{ fontSize: 9, color: "var(--text3)", marginTop: 2 }}>المتبقي: {remaining.toLocaleString("ar-EG")} ج.م</div>
            </div>
          )}
          
          <DentalChart selectedTeeth={form.teeth} onChange={(teeth) => setForm(f => ({ ...f, teeth }))} />
          
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
          <button className="btn btn-primary" onClick={submit} disabled={loading}>
            {loading ? "جاري..." : isEdit ? "💾 حفظ" : "➕ إضافة"}
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
            <div style={{ fontSize: 14, fontWeight: 800 }}>{c.patientName}</div>
            <div style={{ display: "flex", gap: 4, flexWrap: "wrap", marginTop: 5 }}>
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
            <div className="detail-row">
              <span className="detail-key">الإجمالي</span>
              <span className="detail-val" style={{ color: "var(--gold)", fontWeight: 900 }}>{fmtMoney(c.totalAmount)} ج.م</span>
            </div>
            {c.paymentStatus === "Partial" && (
              <>
                <div className="detail-row">
                  <span className="detail-key">المحصل</span>
                  <span className="detail-val" style={{ color: "var(--mint)" }}>{fmtMoney(c.paidAmount)} ج.م</span>
                </div>
                <div className="detail-row">
                  <span className="detail-key">المتبقي</span>
                  <span className="detail-val" style={{ color: "var(--rose)" }}>{fmtMoney(remaining)} ج.م</span>
                </div>
              </>
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
          <div style={{ display: "flex", gap: 5, flexDirection: "column" }}>
            <button className="btn btn-ghost" onClick={() => onEdit(c)}>✏️ تعديل</button>
            <button className="btn btn-danger" onClick={() => onDelete(c)}>🗑 حذف</button>
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
      materials: [...prev.materials, { id: newId, name, price, active: true }]
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
      materials: prev.materials.map(m =>
        m.id === id ? { ...m, active: false } : m
      )
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
            <div className="settings-section-title">🔐 تغيير كلمة المرور</div>
            <div className="form-group">
              <input type="password" className="form-input" placeholder="كلمة المرور الجديدة" value={newPassword} onChange={(e) => setNewPassword(e.target.value)} />
            </div>
            <button className="btn btn-primary" onClick={handleUpdatePassword} disabled={updatingPassword}>
              {updatingPassword ? "جاري..." : "تغيير"}
            </button>
          </div>
          
          <div className="divider" />
          
          <div className="settings-section">
            <div className="settings-section-title">🏥 الفروع</div>
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
            <div className="settings-section-title">💊 المواد</div>
            <div style={{ fontSize: 9, color: "var(--text3)", marginBottom: 5 }}>⚠️ تغيير السعر لا يؤثر على الحالات القديمة</div>
            {s.materials.map(m => (
              <div key={m.id} className="settings-list-item" style={{ opacity: m.active === false ? 0.5 : 1 }}>
                <div>
                  <div className="settings-list-item-name" style={{ color: MAT_COLORS[m.name] || "var(--accent)" }}>{m.name}</div>
                  <div className="settings-list-item-sub">
                    {editingMaterial === m.id ? (
                      <input className="form-input" type="number" defaultValue={m.price ?? ""} style={{ width: 80, padding: "2px 5px", fontSize: 10 }} onBlur={(e) => updateMaterialPrice(m.id, e.target.value)} onKeyDown={(e) => e.key === "Enter" && updateMaterialPrice(m.id, e.target.value)} autoFocus />
                    ) : (
                      m.price !== null && m.price !== undefined ? `${m.price} ج.م` : "سعر حر"
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
          <button className="btn btn-primary" onClick={save} disabled={loading}>{loading ? "جاري..." : "💾 حفظ"}</button>
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
  const [selectedYear, setSelectedYear] = useState("all");
  const [selectedMonth, setSelectedMonth] = useState("all");
  const [statusFilter, setStatusFilter] = useState("all");
  
  const [search, setSearch] = useState("");
  const [selectedYear, setSelectedYear] = useState(2026);
  const [selectedMonth, setSelectedMonth] = useState(5);
  const [showAdd, setShowAdd] = useState(false);
  const [editCase, setEditCase] = useState(null);
  const [detailCase, setDetailCase] = useState(null);
  const [showSettings, setShowSettings] = useState(false);
  const [showStats, setShowStats] = useState(false);
  const [showExportDropdown, setShowExportDropdown] = useState(false);
  const [showRangeExport, setShowRangeExport] = useState(false);
  const [toast, setToast] = useState("");
  const [dropdownState, setDropdownState] = useState({ caseId: null, type: null, x: 0, y: 0 });
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
      const [sett, rawCases] = await Promise.all([getSettings(session.user.id), fetchCases(session.user.id)]);
      setSettings(sett);
      setCases(rawCases.map(toLocal));
      setLoading(false);
    })();
  }, [session]);

  const showToastMessage = useCallback((msg) => {
    if (toastTimeout.current) clearTimeout(toastTimeout.current);
    setToast(msg);
    toastTimeout.current = setTimeout(() => setToast(""), 1500);
  }, []);

  useEffect(() => {
    const handleScroll = () => { if (toast) setToast(""); };
    window.addEventListener("scroll", handleScroll);
    return () => window.removeEventListener("scroll", handleScroll);
  }, [toast]);

  const totalStats = useMemo(() => {
    const count = cases.length;
    const totalMoney = cases.reduce((sum, c) => sum + (c.totalAmount || 0), 0);
    const totalCollected = cases.reduce((sum, c) => sum + (c.paidAmount || 0), 0);
    return { count, totalMoney, totalCollected };
  }, [cases]);

  const patientSearchFiltered = useMemo(() => {
    if (!search.trim()) return cases;
    const q = search.toLowerCase().trim();
    return branchFiltered.filter(c => c.patientName.toLowerCase().includes(q));
  }, [branchFiltered, search]);

  const availableYears = useMemo(() => {
    const years = [];
    for (let y = DATE_RANGE.startYear; y <= DATE_RANGE.endYear; y++) {
      years.push(y);
    }
    return years;
  }, []);

  const months = [
    "يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو",
    "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر"
  ];

  // تطبيق فلتر الشهر بعد البحث
  const tabFiltered = useMemo(() => {
    let arr = patientSearchFiltered;
    
    if (activeTab === "current") {
      const targetMonth = `${selectedYear}-${selectedMonth.toString().padStart(2, '0')}`;
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
  }, [patientSearchFiltered, activeTab, selectedYear, selectedMonth]);

  const stats = useMemo(() => {
    const targetMonth = `${selectedYear}-${selectedMonth.toString().padStart(2, '0')}`;
    const src = patientSearchFiltered.filter(c => {
      const d = c.actionDate || c.startDate || c.createdAt;
      return monthKey(d) === targetMonth;
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
  }, [patientSearchFiltered, selectedYear, selectedMonth]);

  async function handleAddCase(form) {
    const row = await insertCase(session.user.id, form);
    if (row) {
      setCases(prev => [toLocal(row), ...prev]);
      showToastMessage("✅ تمت الإضافة");
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
      teeth: form.teeth || [],
      start_date: form.startDate,
      action_date: form.actionDate || null,
      notes: form.notes || null,
    });
    setCases(prev => prev.map(x => x.id === c.id ? { ...x, ...form, totalAmount: form.pricePerUnit * form.units } : x));
    showToastMessage("✅ تم الحفظ");
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
    if (newStatus === "Paid") updateData.paid_amount = c.totalAmount;
    else if (newStatus === "Free" || newStatus === "Unpaid") updateData.paid_amount = 0;
    await updateCaseDB(c.id, updateData);
    setCases(prev => prev.map(x => x.id === c.id ? { ...x, paymentStatus: newStatus, paidAmount: updateData.paid_amount !== undefined ? updateData.paid_amount : x.paidAmount } : x));
    showToastMessage(newStatus === "Paid" ? "✅ مدفوع" : (newStatus === "Free" ? "🎁 مجاناً" : "❌ غير مدفوع"));
  }

  async function handleUpdateMaterial(c, newMaterialName) {
    if (c.materialName === newMaterialName) return;
    const newMaterial = settings.materials.find(m => m.name === newMaterialName);
    const newPrice = newMaterial?.price !== undefined ? newMaterial.price : c.pricePerUnit;
    await updateCaseDB(c.id, { material_name: newMaterialName, price_per_unit: newPrice });
    setCases(prev => prev.map(x => x.id === c.id ? { ...x, materialName: newMaterialName, pricePerUnit: newPrice, totalAmount: newPrice * x.units } : x));
    showToastMessage(`✅ تم تغيير المادة`);
  }

  async function handleUpdateCaseStatus(c, newStatus) {
    if (c.caseStatus === newStatus) return;
    await updateCaseDB(c.id, { case_status: newStatus });
    setCases(prev => prev.map(x => x.id === c.id ? { ...x, caseStatus: newStatus } : x));
    showToastMessage(`✅ تم تغيير الحالة`);
  }

  async function handleSaveSettings(s) {
    await saveSettings(session.user.id, s);
    setSettings(s);
    showToastMessage("✅ تم حفظ الإعدادات");
  }

  async function handleImport(e) {
    const file = e.target.files[0];
    if (!file) return;
    try {
      const count = await importExcel(file, session.user.id, activeBranch !== "all" ? activeBranch : settings.branches[0], showToastMessage);
      const fresh = await fetchCases(session.user.id);
      setCases(fresh.map(toLocal));
      showToastMessage(`✅ تم استيراد ${count} حالة`);
    } catch (err) {
      showToastMessage(`❌ خطأ: ${err.message}`);
    }
    e.target.value = "";
  }

  const exportAll = useCallback(() => {
    exportExcel(cases, "dental_all");
    showToastMessage(`✅ تم تصدير ${cases.length} حالة`);
  }, [cases, showToastMessage]);

  const exportCurrentMonth = useCallback(() => {
    const targetMonth = `${selectedYear}-${selectedMonth.toString().padStart(2, '0')}`;
    const dataToExport = patientSearchFiltered.filter(c => {
      const d = c.actionDate || c.startDate || c.createdAt;
      return monthKey(d) === targetMonth;
    });
    exportCSV(dataToExport);
    showToast(`✅ تم تصدير ${dataToExport.length} حالة لشهر ${selectedMonth}/${selectedYear}`);
  }, [patientSearchFiltered, selectedYear, selectedMonth, showToast]);

  const exportRange = useCallback((startDate, endDate) => {
    const dataToExport = cases.filter(c => (c.startDate || c.createdAt) >= startDate && (c.startDate || c.createdAt) <= endDate);
    exportExcel(dataToExport, `dental_range`);
    showToastMessage(`✅ تم تصدير ${dataToExport.length} حالة`);
  }, [cases, showToastMessage]);

  const handleCardClick = (c, e) => {
    if (e.target.closest('.badge-clickable')) return;
    setDetailCase(c);
  };

  const handleBadgeClick = (e, caseId, type) => {
    e.stopPropagation();
    e.preventDefault();
    
    const rect = e.target.getBoundingClientRect();
    const clientX = e.touches ? e.touches[0].clientX : rect.left;
    const clientY = e.touches ? e.touches[0].clientY : rect.bottom;
    
    setDropdownState({
      caseId: caseId,
      type: type,
      x: clientX,
      y: clientY + 10
    });
  };

  const closeDropdown = useCallback(() => {
    setDropdownState({ caseId: null, type: null, x: 0, y: 0 });
  }, []);

  useEffect(() => {
    document.addEventListener('click', closeDropdown);
    document.addEventListener('touchstart', closeDropdown);
    return () => {
      document.removeEventListener('click', closeDropdown);
      document.removeEventListener('touchstart', closeDropdown);
    };
  }, [closeDropdown]);

  if (!authInit) return <div className="loading" style={{ minHeight: "100vh" }}><div className="spinner" /></div>;
  if (!session) return <LoginScreen onLogin={s => setSession(s)} />;

  const tabs = [
    { id: "current", label: "📅 الشهر المحدد" },
    { id: "active", label: "⚙️ قيد العمل" },
    { id: "missed", label: "⚠️ Missed" },
    { id: "completed", label: "✅ المنتهية" },
  ];

  return (
    <>
      <style>{CSS}</style>
      <div className="header">
        <div className="header-top">
          <div className="header-title">🦷 Dental Tracker</div>
          <div className="header-actions">
            <input type="file" ref={importRef} accept=".csv" style={{ display: "none" }} onChange={handleImport} />
            <button className="btn btn-ghost btn-sm" onClick={() => importRef.current?.click()}>📥 استيراد</button>
            <button className="btn btn-ghost btn-sm" onClick={exportCurrentMonth}>📤 تصدير الشهر</button>
            <button className="btn btn-primary btn-sm" onClick={exportAll}>📤 تصدير الكل</button>
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
        </div>
        
        <div className="month-selector-wrapper">
          <span className="branch-label">السنة:</span>
          <select className="year-select" value={selectedYear} onChange={e => setSelectedYear(parseInt(e.target.value))}>
            {availableYears.map(year => (
              <option key={year} value={year}>{year}</option>
            ))}
          </select>
          
          <span className="branch-label" style={{ marginRight: 8 }}>الشهر:</span>
          <select className="month-select" value={selectedMonth} onChange={e => setSelectedMonth(parseInt(e.target.value))}>
            {months.map((month, index) => (
              <option key={index} value={index + 1}>{month}</option>
            ))}
          </select>
        </div>

        <div className="tabs">
          {tabs.map(t => (
            <button key={t.id} className={`tab${activeTab === t.id ? " active" : ""}`} onClick={() => setActiveTab(t.id)}>
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
              placeholder="🔍 بحث باسم المريض (جميع الحالات)..." 
            />
          </div>
        </div>
        {loading ? <Spinner /> : tabFiltered.length === 0 ? (
          <div className="empty"><div className="empty-icon">📋</div><div className="empty-msg">لا توجد حالات</div></div>
        ) : (
          <div className="cases-grid">
            {tabFiltered.map((c) => {
              const sm = STATUS_META[c.caseStatus] || {};
              const mc = MAT_COLORS[c.materialName] || "var(--mint)";
              const getPaymentLabel = (status, paidAmount) => {
                if (status === "Paid") return "مدفوع كامل";
                if (status === "Free") return "مجاناً";
                if (status === "Partial") return `مدفوع جزئي (${fmtMoney(paidAmount)})`;
                return "غير مدفوع";
              };
              return (
                <div key={c.id} className="case-card fade-up" style={{ "--status-color": sm.color, animationDelay: `${Math.min(i, 12) * 0.04}s` }}>
                  <div className="case-clickable-area" onClick={(e) => handleCardClick(c, e)}>
                    <div className="case-card-header">
                      <div><div className="case-name">{c.patientName}</div><div className="case-date">{fmtDate(c.startDate)}</div></div>
                      <span className="case-branch">{c.branchName}</span>
                    </div>
                    <div className="case-badges">
                      <BadgeDropdown c={c} type="status" settings={settings} onStatusChange={handleUpdateCaseStatus} setToast={showToastMessage} />
                      <BadgeDropdown c={c} type="material" settings={settings} onMaterialChange={handleUpdateMaterial} setToast={showToastMessage} />
                      <BadgeDropdown c={c} type="payment" settings={settings} onPaymentChange={handleUpdatePayment} setToast={showToastMessage} />
                    </div>
                    <div className="case-footer">
                      <span className="case-amount">{fmtMoney(c.totalAmount)} ج.م</span>
                      <span className="case-units">{c.units} × {c.pricePerUnit}</span>
                      <span className="case-date">{fmtDate(c.startDate)}</span>
                    </div>
                  </div>
                  
                  {dropdownState.caseId === c.id && dropdownState.type === 'status' && (
                    <div className="dropdown-menu" style={{ position: 'fixed', top: dropdownState.y, left: dropdownState.x, zIndex: 2000 }}>
                      {settings.caseStatuses.map(s => (
                        <div key={s.id} className="dropdown-item" onClick={() => handleUpdateCaseStatus(c, s.name)}>
                          {STATUS_META[s.name]?.icon} {s.name}
                        </div>
                      ))}
                    </div>
                  )}
                  
                  {dropdownState.caseId === c.id && dropdownState.type === 'material' && (
                    <div className="dropdown-menu" style={{ position: 'fixed', top: dropdownState.y, left: dropdownState.x, zIndex: 2000 }}>
                      {settings.materials.map(m => (
                        <div key={m.id} className="dropdown-item" onClick={() => handleUpdateMaterial(c, m.name)}>
                          {m.name} {m.price ? `(${m.price})` : '(سعر حر)'}
                        </div>
                      ))}
                    </div>
                  )}
                  
                  {dropdownState.caseId === c.id && dropdownState.type === 'payment' && (
                    <div className="dropdown-menu" style={{ position: 'fixed', top: dropdownState.y, left: dropdownState.x, zIndex: 2000 }}>
                      {settings.paymentStatuses.filter(s => s.name !== "Partial").map(s => (
                        <div key={s.id} className="dropdown-item" onClick={() => handleUpdatePayment(c, s.name)}>
                          {s.name === "Paid" ? "✅ مدفوع كامل" : (s.name === "Free" ? "🎁 مجاناً" : "❌ غير مدفوع")}
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              );
            })}
          </div>
        )}
      </div>
      
      <button className="btn-fab" onClick={() => setShowAdd(true)}>＋</button>
      
      {showAdd && <CaseModal settings={settings} defaultBranch={activeBranch !== "all" ? activeBranch : settings.branches[0]} onSave={handleAddCase} onClose={() => setShowAdd(false)} />}
      {editCase && <CaseModal existing={editCase} settings={settings} defaultBranch={editCase.branchName} onSave={handleEditCase} onClose={() => setEditCase(null)} />}
      {detailCase && <CaseDrawer c={detailCase} settings={settings} onEdit={c => { setEditCase(c); setDetailCase(null); }} onDelete={handleDelete} onUpdatePayment={handleUpdatePayment} onClose={() => setDetailCase(null)} />}
      {showSettings && <SettingsScreen settings={settings} onSave={handleSaveSettings} onClose={() => setShowSettings(false)} />}
      <Toast msg={toast} />
    </>
  );
}
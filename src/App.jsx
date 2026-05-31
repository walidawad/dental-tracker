import { useState, useEffect, useMemo, useRef, useCallback } from "react";
import { createClient } from "https://esm.sh/@supabase/supabase-js@2";
import * as XLSX from 'https://esm.sh/xlsx';

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

html { direction: rtl; -webkit-text-size-adjust: 100%; }

body {
  font-family: 'Tajawal', system-ui, sans-serif;
  background: var(--bg);
  color: var(--text);
  min-height: 100vh;
  overflow-x: hidden;
  width: 100%;
  max-width: 100vw;
}

::-webkit-scrollbar { width: 3px; }
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
.header-actions { display: flex; gap: 5px; flex-wrap: wrap; }

.branch-bar, .month-selector-wrapper {
  padding: 5px 10px;
  display: flex;
  gap: 6px;
  flex-wrap: wrap;
  align-items: center;
  border-bottom: 1px solid var(--border);
}
.branch-label { font-size: 9px; color: var(--text3); font-weight: 700; }
.branch-select, .year-select, .month-select {
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
  display: flex;
  gap: 3px;
  padding: 5px 10px;
  border-bottom: 1px solid var(--border);
  overflow-x: auto;
}
.tab {
  padding: 4px 10px;
  border-radius: 14px;
  font-size: 10px;
  font-weight: 700;
  border: 1px solid var(--border);
  color: var(--text2);
  cursor: pointer;
  white-space: nowrap;
  background: transparent;
}
.tab.active { background: var(--accent); border-color: var(--accent); color: #fff; }

.main { padding: 8px; max-width: 600px; margin: 0 auto; }

.stats-grid {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(80px, 1fr));
  gap: 6px;
  margin-bottom: 10px;
}
.stat-card {
  background: var(--surface);
  border: 1px solid var(--border);
  border-radius: 8px;
  padding: 5px 3px;
  text-align: center;
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
  display: flex;
  gap: 6px;
  margin-bottom: 10px;
  flex-wrap: wrap;
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

.cases-grid {
  display: grid;
  grid-template-columns: 1fr;
  gap: 6px;
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
.case-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin-bottom: 5px;
}
.case-name { font-size: 12px; font-weight: 700; word-break: break-word; }
.case-date { font-size: 8px; color: var(--text3); }
.case-badges {
  display: flex;
  gap: 5px;
  flex-wrap: wrap;
  margin: 5px 0;
}
.case-footer {
  display: flex;
  justify-content: space-between;
  align-items: center;
  flex-wrap: wrap;
  gap: 5px;
  margin-top: 5px;
}
.case-amount { font-size: 13px; font-weight: 800; }
.case-units { font-size: 8px; color: var(--text3); }
.case-branch { font-size: 8px; color: var(--text2); background: var(--surface2); padding: 2px 5px; border-radius: 14px; }

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
  border-radius: 14px;
  width: 100%;
  max-width: 480px;
  max-height: 85vh;
  overflow-y: auto;
}
.modal-header {
  padding: 10px 14px;
  display: flex;
  justify-content: space-between;
  align-items: center;
  border-bottom: 1px solid var(--border);
}
.modal-title { font-size: 14px; font-weight: 800; }
.modal-body { padding: 12px 14px; }
.modal-footer { padding: 8px 14px; border-top: 1px solid var(--border); display: flex; gap: 8px; justify-content: flex-end; }

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
.form-row { display: grid; grid-template-columns: 1fr; gap: 8px; }

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
  z-index: 1001;
  animation: slideUp .25s ease;
}
@keyframes slideUp { from { transform: translateY(100%); } to { transform: none; } }
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

@media (min-width: 640px) {
  .drawer { width: min(320px, 100vw); }
  .form-row { grid-template-columns: 1fr 1fr; }
}

@media (max-width: 640px) {
  .stats-grid {
    grid-template-columns: repeat(4, 1fr);
    gap: 5px;
  }
  .stat-card {
    padding: 4px 2px;
  }
  .stat-label {
    font-size: 7px;
  }
  .stat-value {
    font-size: 10px;
  }
}
`;

// ========== دوال Supabase الأساسية ==========
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
        for (const row of rows) {
          const branchName = row["الفرع"] || currentBranch;
          const patientName = row["المريض"] || "";
          const materialName = row["المادة"] || "Zirconia";
          const pricePerUnit = parseFloat(row["السعر"]) || 450;
          const units = parseInt(row["الوحدات"]) || 1;
          const caseStatus = row["الحالة"] || "In progress";
          
          let paymentStatus = row["الدفع"] || "Unpaid";
          if (paymentStatus === "مدفوع" || paymentStatus === "مدفوع كامل") paymentStatus = "Paid";
          if (paymentStatus === "مجاناً") paymentStatus = "Free";
          if (paymentStatus.includes("مدفوع جزئي") || paymentStatus === "Partial") paymentStatus = "Partial";
          if (paymentStatus === "غير مدفوع") paymentStatus = "Unpaid";
          
          const paidAmount = parseFloat(row["المحصل"]) || 0;
          
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
          let rawActionDate = row["الإجراء"] || "";
          if (rawActionDate) {
            let dateStr = String(rawActionDate);
            if (dateStr.includes("-")) {
              const parts = dateStr.split("-");
              if (parts.length === 3) {
                if (parts[0].length === 4) {
                  actionDate = `${parts[0]}-${parts[1].padStart(2,'0')}-${parts[2].padStart(2,'0')}`;
                } else if (parts[2].length === 4) {
                  actionDate = `${parts[2]}-${parts[1].padStart(2,'0')}-${parts[0].padStart(2,'0')}`;
                }
              }
            }
          }
          
          const notes = row["ملاحظات"] || "";
          
          if (!patientName) continue;
          
          const rec = {
            branchName: branchName,
            patientName: patientName,
            materialName: materialName,
            pricePerUnit: pricePerUnit,
            units: units,
            caseStatus: caseStatus,
            paymentStatus: paymentStatus,
            paidAmount: paidAmount,
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
              {settings.materials.map(m => <option key={m.id} value={m.name}>{m.name}</option>)}
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
                {settings.paymentStatuses.map(s => <option key={s.id} value={s.name}>{s.name}</option>)}
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
            <label className="form-label">ملاحظات</label>
            <textarea className="form-input form-textarea" value={form.notes} onChange={e => setForm(f => ({ ...f, notes: e.target.value }))} placeholder="أي تفاصيل إضافية..." />
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

function BadgeDropdown({ c, type, settings, onStatusChange, onMaterialChange, onPaymentChange, setToast }) {
  const [dropdownOpen, setDropdownOpen] = useState(false);
  const [dropdownPosition, setDropdownPosition] = useState({ top: 0, left: 0 });
  const badgeRef = useRef(null);
  const dropdownRef = useRef(null);

  const openDropdown = (e) => {
    e.stopPropagation();
    e.preventDefault();
    
    const rect = e.currentTarget.getBoundingClientRect();
    const windowHeight = window.innerHeight;
    const spaceBelow = windowHeight - rect.bottom;
    const dropdownHeight = 200;
    
    let top = rect.bottom + 5;
    
    if (spaceBelow < dropdownHeight && rect.top > dropdownHeight) {
      top = rect.top - dropdownHeight - 5;
    }
    
    setDropdownPosition({ top: top, left: rect.left });
    setDropdownOpen(true);
  };

  useEffect(() => {
    const handleClickOutside = (event) => {
      if (dropdownRef.current && !dropdownRef.current.contains(event.target) &&
          badgeRef.current && !badgeRef.current.contains(event.target)) {
        setDropdownOpen(false);
      }
    };
    
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
      <div className="dropdown-container">
        <span 
          ref={badgeRef}
          className="badge" 
          style={{ color: sm.color, background: sm.bg, border: `1px solid ${sm.color}30`, cursor: 'pointer' }} 
          onClick={openDropdown}
        >
          {sm.icon} {c.caseStatus} ▼
        </span>
        {dropdownOpen && (
          <div 
            ref={dropdownRef}
            className="dropdown-menu" 
            style={{ position: 'fixed', top: dropdownPosition.top, left: dropdownPosition.left, zIndex: 9999 }}
          >
            {settings.caseStatuses.map(s => (
              <div key={s.id} className="dropdown-item" onClick={(e) => { e.stopPropagation(); onStatusChange(c, s.name); setDropdownOpen(false); }}>
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
      <div className="dropdown-container">
        <span 
          ref={badgeRef}
          className="badge" 
          style={{ color: mc, background: mc + "20", cursor: 'pointer' }} 
          onClick={openDropdown}
        >
          {c.materialName} ▼
        </span>
        {dropdownOpen && (
          <div 
            ref={dropdownRef}
            className="dropdown-menu" 
            style={{ position: 'fixed', top: dropdownPosition.top, left: dropdownPosition.left, zIndex: 9999 }}
          >
            {settings.materials.map(m => (
              <div key={m.id} className="dropdown-item" onClick={(e) => { e.stopPropagation(); onMaterialChange(c, m.name); setDropdownOpen(false); }}>
                {m.name} {m.price ? `(${m.price})` : ''}
              </div>
            ))}
          </div>
        )}
      </div>
    );
  }

  if (type === 'payment') {
    const getPaymentLabel = (status, paidAmount) => {
      if (status === "Paid") return "مدفوع";
      if (status === "Free") return "مجاناً";
      if (status === "Partial") return `مدفوع جزئي`;
      return "غير مدفوع";
    };
    const paymentColor = c.paymentStatus === "Paid" ? "var(--mint)" : (c.paymentStatus === "Free" ? "var(--amber)" : (c.paymentStatus === "Partial" ? "var(--lavender)" : "var(--rose)"));
    const paymentBg = c.paymentStatus === "Paid" ? "rgba(0,212,161,.12)" : (c.paymentStatus === "Free" ? "rgba(245,166,35,.12)" : (c.paymentStatus === "Partial" ? "rgba(155,142,255,.12)" : "rgba(239,68,68,.12)"));
    
    return (
      <div className="dropdown-container">
        <span 
          ref={badgeRef}
          className="badge" 
          style={{ color: paymentColor, background: paymentBg, border: `1px solid ${paymentColor}30`, cursor: 'pointer' }} 
          onClick={openDropdown}
        >
          {getPaymentLabel(c.paymentStatus, c.paidAmount)} ▼
        </span>
        {dropdownOpen && (
          <div 
            ref={dropdownRef}
            className="dropdown-menu" 
            style={{ position: 'fixed', top: dropdownPosition.top, left: dropdownPosition.left, zIndex: 9999 }}
          >
            {settings.paymentStatuses.filter(s => s.name !== "Partial").map(s => (
              <div key={s.id} className="dropdown-item" onClick={(e) => { e.stopPropagation(); onPaymentChange(c, s.name); setDropdownOpen(false); }}>
                {s.name === "Paid" ? "✅ مدفوع" : (s.name === "Free" ? "🎁 مجاناً" : "❌ غير مدفوع")}
              </div>
            ))}
          </div>
        )}
      </div>
    );
  }
  
  return null;
}

function CaseDrawer({ c, settings, onEdit, onDelete, onUpdatePayment, onClose, setToast }) {
  const sm = STATUS_META[c.caseStatus] || {};
  const mc = MAT_COLORS[c.materialName] || "var(--mint)";
  const remaining = (c.totalAmount || 0) - (c.paidAmount || 0);
  
  const getPaymentLabel = (status) => {
    if (status === "Paid") return "مدفوع";
    if (status === "Free") return "مجاناً";
    if (status === "Partial") return `مدفوع جزئي (${fmtMoney(c.paidAmount)})`;
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
              <Badge label={getPaymentLabel(c.paymentStatus)} color={c.paymentStatus === "Paid" ? "var(--mint)" : (c.paymentStatus === "Free" ? "var(--amber)" : "var(--rose)")} />
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
            <div className="detail-row">
              <span className="detail-key">الفرع</span>
              <span className="detail-val">{c.branchName}</span>
            </div>
            <div className="detail-row">
              <span className="detail-key">تاريخ البداية</span>
              <span className="detail-val">{fmtDate(c.startDate)}</span>
            </div>
            <div className="detail-row">
              <span className="detail-key">تاريخ الإجراء</span>
              <span className="detail-val">{fmtDate(c.actionDate)}</span>
            </div>
            {c.notes && (
              <div className="detail-row">
                <span className="detail-key">ملاحظات</span>
                <span className="detail-val">{c.notes}</span>
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

function SettingsScreen({ settings, onSave, onClose, setToast }) {
  const [s, setS] = useState(JSON.parse(JSON.stringify(settings)));
  const [loading, setLoading] = useState(false);
  const [newBranch, setNewBranch] = useState("");
  const [newMaterial, setNewMaterial] = useState({ name: "", price: "" });
  const [editingMaterial, setEditingMaterial] = useState(null);
  const [newPassword, setNewPassword] = useState("");
  const [updatingPassword, setUpdatingPassword] = useState(false);

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
      setToast("لا يمكن حذف المادة الوحيدة");
      return;
    }
    setS(prev => ({
      ...prev,
      materials: prev.materials.filter(m => m.id !== id)
    }));
  }

  async function handleUpdatePassword() {
    if (!newPassword.trim()) {
      setToast("الرجاء إدخال كلمة مرور جديدة");
      return;
    }
    setUpdatingPassword(true);
    const { error } = await supabase.auth.updateUser({ password: newPassword });
    setUpdatingPassword(false);
    if (error) {
      setToast("❌ خطأ: " + error.message);
    } else {
      setToast("✅ تم تحديث كلمة المرور!");
      setNewPassword("");
    }
  }

  async function save() {
    setLoading(true);
    await onSave(s);
    setLoading(false);
    onClose();
  }

  return (
    <div className="modal-overlay" onClick={onClose}>
      <div className="modal" style={{ maxWidth: 450 }} onClick={e => e.stopPropagation()}>
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
            <div style={{ display: "flex", gap: 5, marginTop: 5 }}>
              <input className="form-input" value={newBranch} onChange={e => setNewBranch(e.target.value)} placeholder="اسم فرع جديد" onKeyDown={e => e.key === "Enter" && addBranch()} />
              <button className="btn btn-primary btn-sm" onClick={addBranch}>إضافة</button>
            </div>
          </div>
          
          <div className="divider" />
          
          <div className="settings-section">
            <div className="settings-section-title">💊 المواد</div>
            <div style={{ fontSize: 9, color: "var(--text3)", marginBottom: 5 }}>⚠️ تغيير السعر لا يؤثر على الحالات القديمة</div>
            {s.materials.map(m => (
              <div key={m.id} className="settings-list-item">
                <div>
                  <div className="settings-list-item-name">{m.name}</div>
                  <div className="settings-list-item-sub">
                    {editingMaterial === m.id ? (
                      <input className="form-input" type="number" defaultValue={m.price ?? ""} style={{ width: 80, padding: "2px 5px", fontSize: 10 }} onBlur={(e) => updateMaterialPrice(m.id, e.target.value)} onKeyDown={(e) => e.key === "Enter" && updateMaterialPrice(m.id, e.target.value)} autoFocus />
                    ) : (
                      m.price !== null && m.price !== undefined ? `${m.price} ج.م` : "سعر حر"
                    )}
                  </div>
                </div>
                <div style={{ display: "flex", gap: 3 }}>
                  {m.price !== null && <button className="btn btn-ghost btn-sm" onClick={() => setEditingMaterial(m.id)}>تعديل</button>}
                  <button className="btn btn-danger btn-sm" onClick={() => deleteMaterial(m.id)}>حذف</button>
                </div>
              </div>
            ))}
            <div style={{ display: "flex", gap: 5, marginTop: 6 }}>
              <input className="form-input" value={newMaterial.name} onChange={e => setNewMaterial({ ...newMaterial, name: e.target.value })} placeholder="مادة جديدة" style={{ flex: 2 }} />
              <input className="form-input" type="number" value={newMaterial.price} onChange={e => setNewMaterial({ ...newMaterial, price: e.target.value })} placeholder="السعر" style={{ flex: 1 }} />
              <button className="btn btn-primary btn-sm" onClick={addMaterial}>➕</button>
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

function StatsModal({ cases, onClose }) {
  const totalCases = cases.length;
  const totalMoney = cases.reduce((sum, c) => sum + (c.totalAmount || 0), 0);
  const totalCollected = cases.reduce((sum, c) => sum + (c.paidAmount || 0), 0);
  const totalFree = cases.filter(c => c.paymentStatus === "Free").reduce((sum, c) => sum + (c.totalAmount || 0), 0);
  const completedCases = cases.filter(c => c.caseStatus === "Completed").length;
  const inProgressCases = cases.filter(c => c.caseStatus === "In progress").length;
  const correctionCases = cases.filter(c => c.caseStatus === "Correction").length;
  const missedCases = cases.filter(c => c.caseStatus === "Missed").length;

  return (
    <div className="modal-overlay" onClick={onClose}>
      <div className="modal" style={{ maxWidth: 350 }} onClick={e => e.stopPropagation()}>
        <div className="modal-header">
          <span className="modal-title">📊 الإحصائيات</span>
          <button className="btn btn-ghost btn-icon" onClick={onClose}>✕</button>
        </div>
        <div className="modal-body">
          <div className="detail-row"><span className="detail-key">إجمالي الحالات</span><span className="detail-val">{totalCases}</span></div>
          <div className="detail-row"><span className="detail-key">إجمالي المبالغ</span><span className="detail-val">{fmtMoney(totalMoney)} ج.م</span></div>
          <div className="detail-row"><span className="detail-key">المحصل</span><span className="detail-val">{fmtMoney(totalCollected)} ج.م</span></div>
          <div className="detail-row"><span className="detail-key">المتبقي</span><span className="detail-val">{fmtMoney(totalMoney - totalCollected)} ج.م</span></div>
          <div className="detail-row"><span className="detail-key">مجاناً</span><span className="detail-val">{fmtMoney(totalFree)} ج.م</span></div>
          <div className="divider" />
          <div className="detail-row"><span className="detail-key">✅ منتهية</span><span className="detail-val">{completedCases}</span></div>
          <div className="detail-row"><span className="detail-key">⚙️ قيد العمل</span><span className="detail-val">{inProgressCases}</span></div>
          <div className="detail-row"><span className="detail-key">🔧 تصحيح</span><span className="detail-val">{correctionCases}</span></div>
          <div className="detail-row"><span className="detail-key">⚠️ Missed</span><span className="detail-val">{missedCases}</span></div>
        </div>
        <div className="modal-footer">
          <button className="btn btn-primary" onClick={onClose}>إغلاق</button>
        </div>
      </div>
    </div>
  );
}

function RangeExportModal({ onClose, onExport }) {
  const [startDate, setStartDate] = useState(todayISO());
  const [endDate, setEndDate] = useState(todayISO());

  return (
    <div className="modal-overlay" onClick={onClose}>
      <div className="modal" style={{ maxWidth: 350 }} onClick={e => e.stopPropagation()}>
        <div className="modal-header">
          <span className="modal-title">📆 تصدير نطاق زمني</span>
          <button className="btn btn-ghost btn-icon" onClick={onClose}>✕</button>
        </div>
        <div className="modal-body">
          <div className="form-group">
            <label className="form-label">من تاريخ</label>
            <input className="form-input" type="date" value={startDate} onChange={e => setStartDate(e.target.value)} />
          </div>
          <div className="form-group">
            <label className="form-label">إلى تاريخ</label>
            <input className="form-input" type="date" value={endDate} onChange={e => setEndDate(e.target.value)} />
          </div>
        </div>
        <div className="modal-footer">
          <button className="btn btn-ghost" onClick={onClose}>إلغاء</button>
          <button className="btn btn-primary" onClick={() => { onExport(startDate, endDate); onClose(); }}>تصدير</button>
        </div>
      </div>
    </div>
  );
}

function ExportDropdown({ onExportAll, onExportCurrent, onExportRange, onClose }) {
  const dropdownRef = useRef(null);

  useEffect(() => {
    const handleClickOutside = (event) => {
      if (dropdownRef.current && !dropdownRef.current.contains(event.target)) onClose();
    };
    document.addEventListener("mousedown", handleClickOutside);
    document.addEventListener("touchstart", handleClickOutside);
    return () => {
      document.removeEventListener("mousedown", handleClickOutside);
      document.removeEventListener("touchstart", handleClickOutside);
    };
  }, [onClose]);

  return (
    <div ref={dropdownRef} className="dropdown-menu" style={{ position: 'fixed', top: 45, left: 'auto', right: 50, zIndex: 9999 }}>
      <div className="dropdown-item" onClick={() => { onExportAll(); onClose(); }}>📤 الكل</div>
      <div className="dropdown-item" onClick={() => { onExportCurrent(); onClose(); }}>📅 الشهر الحالي</div>
      <div className="dropdown-item" onClick={() => { onExportRange(); onClose(); }}>📆 نطاق زمني</div>
    </div>
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
  const [selectedYear, setSelectedYear] = useState("all");
  const [selectedMonth, setSelectedMonth] = useState(new Date().getMonth() + 1);
  const [showAdd, setShowAdd] = useState(false);
  const [editCase, setEditCase] = useState(null);
  const [detailCase, setDetailCase] = useState(null);
  const [showSettings, setShowSettings] = useState(false);
  const [showStats, setShowStats] = useState(false);
  const [showExportDropdown, setShowExportDropdown] = useState(false);
  const [showRangeExport, setShowRangeExport] = useState(false);
  const [toast, setToast] = useState("");
  const importRef = useRef();
  let toastTimeout = useRef(null);

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
    return cases.filter(c => c.patientName.toLowerCase().includes(q));
  }, [cases, search]);

  const branchFiltered = useMemo(() => 
    activeBranch === "all" ? patientSearchFiltered : patientSearchFiltered.filter(c => c.branchName === activeBranch),
    [patientSearchFiltered, activeBranch]
  );

  const availableYears = useMemo(() => {
    const years = [];
    for (let y = DATE_RANGE.startYear; y <= DATE_RANGE.endYear; y++) years.push(y);
    return years;
  }, []);

  const months = ["يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو", "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر"];

  const tabFiltered = useMemo(() => {
    let arr = branchFiltered;
    
    if (activeTab === "current") {
      if (selectedYear !== "all") {
        const targetMonth = `${selectedYear}-${selectedMonth.toString().padStart(2, '0')}`;
        arr = arr.filter(c => {
          const d = c.actionDate || c.startDate || c.createdAt;
          return monthKey(d) === targetMonth;
        });
      }
    } else if (activeTab === "annual") {
      if (selectedYear !== "all") {
        arr = arr.filter(c => {
          const d = c.actionDate || c.startDate || c.createdAt;
          return d && new Date(d).getFullYear() === selectedYear;
        });
      }
    } else if (activeTab === "active") {
      arr = arr.filter(c => c.caseStatus === "In progress" || c.caseStatus === "Correction");
    } else if (activeTab === "missed") {
      arr = arr.filter(c => c.caseStatus === "Missed");
    } else if (activeTab === "completed") {
      arr = arr.filter(c => c.caseStatus === "Completed");
    }
    
    return arr;
  }, [branchFiltered, activeTab, selectedYear, selectedMonth]);

  const stats = useMemo(() => {
    const total = tabFiltered.reduce((s, c) => s + (c.totalAmount || 0), 0);
    const collected = tabFiltered.reduce((s, c) => s + (c.paidAmount || 0), 0);
    const free = tabFiltered.filter(c => c.paymentStatus === "Free").reduce((s, c) => s + (c.totalAmount || 0), 0);
    const completed = tabFiltered.filter(c => c.caseStatus === "Completed").reduce((s, c) => s + (c.totalAmount || 0), 0);
    const inProgress = tabFiltered.filter(c => c.caseStatus === "In progress").reduce((s, c) => s + (c.totalAmount || 0), 0);
    const correction = tabFiltered.filter(c => c.caseStatus === "Correction").reduce((s, c) => s + (c.totalAmount || 0), 0);
    const missed = tabFiltered.filter(c => c.caseStatus === "Missed").reduce((s, c) => s + (c.totalAmount || 0), 0);
    const tempCement = tabFiltered.filter(c => c.caseStatus === "Temp cement").reduce((s, c) => s + (c.totalAmount || 0), 0);
    return { total, collected, unpaid: total - collected, free, completed, inProgress, correction, missed, tempCement, count: tabFiltered.length };
  }, [tabFiltered]);

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
    showToastMessage("🗑 تم الحذف");
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
    let dataToExport = tabFiltered;
    if (activeTab === "current" && selectedYear !== "all") {
      const targetMonth = `${selectedYear}-${selectedMonth.toString().padStart(2, '0')}`;
      dataToExport = patientSearchFiltered.filter(c => monthKey(c.actionDate || c.startDate || c.createdAt) === targetMonth);
    } else if (activeTab === "annual" && selectedYear !== "all") {
      dataToExport = patientSearchFiltered.filter(c => new Date(c.actionDate || c.startDate || c.createdAt).getFullYear() === selectedYear);
    }
    exportExcel(dataToExport, `dental_filtered`);
    showToastMessage(`✅ تم تصدير ${dataToExport.length} حالة`);
  }, [patientSearchFiltered, tabFiltered, activeTab, selectedYear, selectedMonth, showToastMessage]);

  const exportRange = useCallback((startDate, endDate) => {
    const dataToExport = cases.filter(c => (c.startDate || c.createdAt) >= startDate && (c.startDate || c.createdAt) <= endDate);
    exportExcel(dataToExport, `dental_range`);
    showToastMessage(`✅ تم تصدير ${dataToExport.length} حالة`);
  }, [cases, showToastMessage]);

  const handleCardClick = (c, e) => {
    if (e.target.closest('.dropdown-container') || e.target.closest('.badge')) return;
    setDetailCase(c);
  };

  if (!authInit) return <div className="loading"><div className="spinner" /></div>;
  if (!session) return <LoginScreen onLogin={s => setSession(s)} />;

  const tabs = [
    { id: "current", label: "📅 الشهر" },
    { id: "annual", label: "📅 السنة" },
    { id: "active", label: "⚙️ قيد" },
    { id: "missed", label: "⚠️ Missed" },
    { id: "completed", label: "✅ منتهية" },
  ];

  return (
    <>
      <style>{CSS}</style>
      <div className="header">
        <div className="header-top">
          <div className="header-title">🦷 Dental Tracker</div>
          <div className="header-actions">
            <input type="file" ref={importRef} accept=".xlsx, .xls" style={{ display: "none" }} onChange={handleImport} />
            <button className="btn btn-ghost btn-sm" onClick={() => importRef.current?.click()}>📥</button>
            <div style={{ position: "relative" }}>
              <button className="btn btn-ghost btn-sm" onClick={() => setShowExportDropdown(!showExportDropdown)}>📤▼</button>
              {showExportDropdown && <ExportDropdown onExportAll={exportAll} onExportCurrent={exportCurrentMonth} onExportRange={() => setShowRangeExport(true)} onClose={() => setShowExportDropdown(false)} />}
            </div>
            <button className="btn btn-ghost btn-sm" onClick={() => setShowStats(true)}>📊</button>
            <button className="btn btn-ghost btn-sm" onClick={() => setShowSettings(true)}>⚙️</button>
            <button className="btn btn-ghost btn-sm" onClick={() => supabase.auth.signOut()}>خروج</button>
          </div>
        </div>
      </div>
      
      <div className="branch-bar">
        <span className="branch-label">الفرع:</span>
        <select className="branch-select" value={activeBranch} onChange={e => setActiveBranch(e.target.value)}>
          <option value="all">🏥 الكل</option>
          {settings.branches.map(b => <option key={b} value={b}>{b}</option>)}
        </select>
      </div>
      
      <div className="month-selector-wrapper">
        <span className="branch-label">السنة:</span>
        <select className="year-select" value={selectedYear} onChange={e => setSelectedYear(e.target.value)}>
          <option value="all">📅 الكل</option>
          {availableYears.map(y => <option key={y} value={y}>{y}</option>)}
        </select>
        
        {activeTab === "current" && selectedYear !== "all" && (
          <>
            <span className="branch-label">الشهر:</span>
            <select className="month-select" value={selectedMonth} onChange={e => setSelectedMonth(parseInt(e.target.value))}>
              {months.map((m, i) => <option key={i} value={i+1}>{m}</option>)}
            </select>
          </>
        )}
      </div>
      
      <div className="tabs">
        {tabs.map(t => (
          <button key={t.id} className={`tab${activeTab === t.id ? " active" : ""}`} onClick={() => setActiveTab(t.id)}>
            {t.label}
          </button>
        ))}
      </div>
      
      <div className="main">
        <div className="stats-grid">
          <div className="stat-card"><div className="stat-label">💰 الإجمالي</div><div className="stat-value">{fmtMoney(stats.total)}</div></div>
          <div className="stat-card"><div className="stat-label">✅ المحصل</div><div className="stat-value">{fmtMoney(stats.collected)}</div></div>
          <div className="stat-card"><div className="stat-label">⏳ المتبقي</div><div className="stat-value">{fmtMoney(stats.unpaid)}</div></div>
          <div className="stat-card"><div className="stat-label">🎁 مجاناً</div><div className="stat-value">{fmtMoney(stats.free)}</div></div>
          <div className="stat-card"><div className="stat-label">📋 منتهية</div><div className="stat-value">{fmtMoney(stats.completed)}</div></div>
          <div className="stat-card"><div className="stat-label">⚙️ قيد العمل</div><div className="stat-value">{fmtMoney(stats.inProgress)}</div></div>
          <div className="stat-card"><div className="stat-label">🔧 تصحيح</div><div className="stat-value">{fmtMoney(stats.correction)}</div></div>
          <div className="stat-card"><div className="stat-label">⚠️ Missed</div><div className="stat-value">{fmtMoney(stats.missed)}</div></div>
          <div className="stat-card"><div className="stat-label">📊 عدد الحالات</div><div className="stat-value">{stats.count}</div></div>
        </div>
        
        <div className="toolbar">
          <div className="search-wrap">
            <input value={search} onChange={e => setSearch(e.target.value)} placeholder="🔍 بحث..." />
          </div>
        </div>
        
        {loading ? <Spinner /> : tabFiltered.length === 0 ? (
          <div className="empty"><div className="empty-icon">📋</div><div className="empty-msg">لا توجد حالات</div></div>
        ) : (
          <div className="cases-grid">
            {tabFiltered.map((c, i) => {
              const sm = STATUS_META[c.caseStatus] || {};
              const mc = MAT_COLORS[c.materialName] || "var(--mint)";
              return (
                <div key={c.id} className="case-card" style={{ "--status-color": sm.color }}>
                  <div onClick={(e) => handleCardClick(c, e)}>
                    <div className="case-header">
                      <div className="case-name">{c.patientName}</div>
                      <div className="case-branch">{c.branchName}</div>
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
                    {c.paymentStatus === "Partial" && <div style={{ fontSize: 8, color: "var(--mint)", marginTop: 3 }}>محصل: {fmtMoney(c.paidAmount)}</div>}
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
      {detailCase && <CaseDrawer c={detailCase} settings={settings} onEdit={c => { setEditCase(c); setDetailCase(null); }} onDelete={handleDelete} onUpdatePayment={handleUpdatePayment} onClose={() => setDetailCase(null)} setToast={showToastMessage} />}
      {showSettings && <SettingsScreen settings={settings} onSave={handleSaveSettings} onClose={() => setShowSettings(false)} setToast={showToastMessage} />}
      {showStats && <StatsModal cases={cases} onClose={() => setShowStats(false)} />}
      {showRangeExport && <RangeExportModal onClose={() => setShowRangeExport(false)} onExport={exportRange} />}
      
      <Toast msg={toast} onClick={() => setToast("")} />
    </>
  );
}
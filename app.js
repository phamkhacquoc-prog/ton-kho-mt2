/* ===================== Tồn Kho MT2 - app.js ===================== */

/* ---------- IndexedDB helper ---------- */
const DB_NAME = "tonkho_db";
const DB_VERSION = 1;
let dbInstance = null;

function openDB() {
  return new Promise((resolve, reject) => {
    if (dbInstance) return resolve(dbInstance);
    const req = indexedDB.open(DB_NAME, DB_VERSION);
    req.onupgradeneeded = (e) => {
      const db = e.target.result;
      if (!db.objectStoreNames.contains("records")) {
        const store = db.createObjectStore("records", { keyPath: "id" });
        store.createIndex("date", "date", { unique: false });
        store.createIndex("ma_hang", "ma_hang", { unique: false });
        store.createIndex("kho", "kho", { unique: false });
        store.createIndex("synced", "synced", { unique: false });
      }
      if (!db.objectStoreNames.contains("settings")) {
        db.createObjectStore("settings", { keyPath: "key" });
      }
    };
    req.onsuccess = (e) => { dbInstance = e.target.result; resolve(dbInstance); };
    req.onerror = (e) => reject(e.target.error);
  });
}

function tx(storeName, mode) {
  return openDB().then((db) => db.transaction(storeName, mode).objectStore(storeName));
}

async function putRecord(rec) {
  const store = await tx("records", "readwrite");
  return new Promise((resolve, reject) => {
    const r = store.put(rec);
    r.onsuccess = () => resolve();
    r.onerror = (e) => reject(e.target.error);
  });
}

async function putRecords(recs) {
  const store = await tx("records", "readwrite");
  return new Promise((resolve, reject) => {
    let remaining = recs.length;
    if (remaining === 0) return resolve();
    recs.forEach((r) => {
      const req = store.put(r);
      req.onsuccess = () => { remaining--; if (remaining === 0) resolve(); };
      req.onerror = (e) => reject(e.target.error);
    });
  });
}

async function getAllRecords() {
  const store = await tx("records", "readonly");
  return new Promise((resolve, reject) => {
    const r = store.getAll();
    r.onsuccess = () => resolve(r.result);
    r.onerror = (e) => reject(e.target.error);
  });
}

async function deleteRecordById(id) {
  const store = await tx("records", "readwrite");
  return new Promise((resolve, reject) => {
    const r = store.delete(id);
    r.onsuccess = () => resolve();
    r.onerror = (e) => reject(e.target.error);
  });
}

async function clearAllRecords() {
  const store = await tx("records", "readwrite");
  return new Promise((resolve, reject) => {
    const r = store.clear();
    r.onsuccess = () => resolve();
    r.onerror = (e) => reject(e.target.error);
  });
}

async function getSetting(key, def = null) {
  const store = await tx("settings", "readonly");
  return new Promise((resolve) => {
    const r = store.get(key);
    r.onsuccess = () => resolve(r.result ? r.result.value : def);
    r.onerror = () => resolve(def);
  });
}

async function setSetting(key, value) {
  const store = await tx("settings", "readwrite");
  return new Promise((resolve, reject) => {
    const r = store.put({ key, value });
    r.onsuccess = () => resolve();
    r.onerror = (e) => reject(e.target.error);
  });
}

/* ---------- Utilities ---------- */
function todayISO() {
  const d = new Date();
  return d.toISOString().slice(0, 10);
}

function fmtDateVN(iso) {
  if (!iso) return "";
  const [y, m, d] = iso.split("-");
  return `${d}/${m}/${y}`;
}

function num(v) {
  const n = Number(v);
  return isNaN(n) ? 0 : n;
}

function fmtNum(v) {
  return num(v).toLocaleString("vi-VN");
}

function slug(str) {
  return String(str || "")
    .trim()
    .replace(/\s+/g, "_")
    .replace(/[^\w\-]/g, "");
}

function makeId(date, kho, ma) {
  return `${date}__${slug(kho)}__${slug(ma)}`;
}

let toastTimer = null;
function showToast(msg) {
  const el = document.getElementById("toast");
  el.textContent = msg;
  el.classList.add("show");
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => el.classList.remove("show"), 2200);
}

/* ---------- Navigation ---------- */
document.querySelectorAll(".nav-btn").forEach((btn) => {
  btn.addEventListener("click", () => {
    document.querySelectorAll(".nav-btn").forEach((b) => b.classList.remove("active"));
    document.querySelectorAll(".view").forEach((v) => v.classList.remove("active"));
    btn.classList.add("active");
    document.getElementById("view-" + btn.dataset.view).classList.add("active");
    if (btn.dataset.view === "dashboard") renderDashboard();
    if (btn.dataset.view === "inventory") renderInventory();
    if (btn.dataset.view === "history") renderHistorySelect();
    if (btn.dataset.view === "settings") renderSettings();
    // Mở tab nào cũng tranh thủ đồng bộ ngầm để lấy số liệu mới nhất từ người khác
    if (["dashboard", "inventory", "history"].includes(btn.dataset.view)) maybeAutoSync();
  });
});

document.getElementById("tabExcelBtn").addEventListener("click", () => {
  document.getElementById("tabExcelBtn").classList.add("active");
  document.getElementById("tabManualBtn").classList.remove("active");
  document.getElementById("panelExcel").style.display = "";
  document.getElementById("panelManual").style.display = "none";
});
document.getElementById("tabManualBtn").addEventListener("click", () => {
  document.getElementById("tabManualBtn").classList.add("active");
  document.getElementById("tabExcelBtn").classList.remove("active");
  document.getElementById("panelManual").style.display = "";
  document.getElementById("panelExcel").style.display = "none";
});

/* ---------- Excel import parsing ---------- */
const COLUMN_ALIASES = {
  kho: ["tên kho", "ten kho"],
  ma_hang: ["mã hàng", "ma hang", "mã sp", "mã sản phẩm"],
  ten_hang: ["tên hàng", "ten hang", "tên sp", "tên sản phẩm"],
  dvt: ["đvt", "dvt", "đơn vị tính", "don vi tinh"],
  dau_ky: ["đầu kỳ", "dau ky"],
  nhap_kho: ["nhập kho", "nhap kho", "nhập", "nhap"],
  xuat_kho: ["xuất kho", "xuat kho", "xuất", "xuat"],
  cuoi_ky: ["cuối kỳ", "cuoi ky"],
  nhom: ["nhóm vthh", "nhom vthh", "nhóm", "nhom"],
};

function normalizeHeader(s) {
  return String(s || "").trim().toLowerCase();
}

function findHeaderRow(rows) {
  for (let i = 0; i < Math.min(rows.length, 15); i++) {
    const row = rows[i].map(normalizeHeader);
    if (row.some((c) => c.includes("mã hàng") || c.includes("ma hang"))) {
      return i;
    }
  }
  return -1;
}

function buildColumnMap(headerRow) {
  const norm = headerRow.map(normalizeHeader);
  const map = {};
  Object.keys(COLUMN_ALIASES).forEach((key) => {
    const aliases = COLUMN_ALIASES[key];
    let idx = -1;
    for (let i = 0; i < norm.length; i++) {
      if (aliases.some((a) => norm[i] === a || norm[i].includes(a))) { idx = i; break; }
    }
    map[key] = idx;
  });
  return map;
}

function parseWorkbook(arrayBuffer) {
  const wb = XLSX.read(arrayBuffer, { type: "array" });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: null });
  const headerIdx = findHeaderRow(rows);
  if (headerIdx === -1) {
    throw new Error("Không tìm thấy dòng tiêu đề (Mã hàng) trong file. Vui lòng dùng đúng mẫu.");
  }
  const colMap = buildColumnMap(rows[headerIdx]);
  if (colMap.ma_hang === -1) {
    throw new Error("Không xác định được cột 'Mã hàng'. Vui lòng kiểm tra lại file.");
  }

  // detect default kho from a title row above header, e.g. "Kho: Kho Thành phẩm nhà máy Quảng Nam (1551), Tháng 7 năm 2026"
  let defaultKho = "";
  for (let i = 0; i < headerIdx; i++) {
    const cell = rows[i] && rows[i][0];
    if (cell && String(cell).toLowerCase().includes("kho:")) {
      const match = String(cell).split(",")[0].replace(/kho:\s*/i, "").trim();
      defaultKho = match;
    }
  }

  const items = [];
  for (let i = headerIdx + 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row) continue;
    const maRaw = colMap.ma_hang >= 0 ? row[colMap.ma_hang] : null;
    const khoRaw = colMap.kho >= 0 ? row[colMap.kho] : null;

    // skip sub-header rows like "Số lượng"
    if (maRaw == null && khoRaw && String(khoRaw).toLowerCase().includes("số lượng")) continue;
    // stop at "Tổng cộng" row or fully empty rows
    if (khoRaw && String(khoRaw).trim().toLowerCase().startsWith("tổng cộng")) break;
    if (maRaw == null || String(maRaw).trim() === "") continue;

    items.push({
      kho: (khoRaw && String(khoRaw).trim()) || defaultKho || "Kho chưa đặt tên",
      ma_hang: String(maRaw).trim(),
      ten_hang: colMap.ten_hang >= 0 ? String(row[colMap.ten_hang] || "").trim() : "",
      dvt: colMap.dvt >= 0 ? String(row[colMap.dvt] || "").trim() : "",
      dau_ky: colMap.dau_ky >= 0 ? num(row[colMap.dau_ky]) : 0,
      nhap_kho: colMap.nhap_kho >= 0 ? num(row[colMap.nhap_kho]) : 0,
      xuat_kho: colMap.xuat_kho >= 0 ? num(row[colMap.xuat_kho]) : 0,
      cuoi_ky: colMap.cuoi_ky >= 0 ? num(row[colMap.cuoi_ky]) : 0,
      nhom: colMap.nhom >= 0 ? String(row[colMap.nhom] || "").trim() : "",
    });
  }
  return items;
}

let pendingImportItems = [];

document.getElementById("importDate").value = todayISO();

document.getElementById("importFile").addEventListener("change", async (e) => {
  const file = e.target.files[0];
  const previewEl = document.getElementById("importPreview");
  const confirmBtn = document.getElementById("btnConfirmImport");
  if (!file) { previewEl.innerHTML = ""; confirmBtn.disabled = true; return; }
  try {
    const buf = await file.arrayBuffer();
    const items = parseWorkbook(buf);
    pendingImportItems = items;
    if (items.length === 0) {
      previewEl.innerHTML = `<p class="hint" style="color:var(--danger)">Không đọc được dòng dữ liệu nào.</p>`;
      confirmBtn.disabled = true;
      return;
    }
    const sample = items.slice(0, 5);
    previewEl.innerHTML = `
      <p class="hint">Đọc được <b>${items.length}</b> mã hàng. Xem trước 5 dòng đầu:</p>
      <div class="table-wrap" style="max-height:200px;">
        <table>
          <thead><tr><th>Mã hàng</th><th>Tên hàng</th><th class="num-cell">Cuối kỳ</th></tr></thead>
          <tbody>
            ${sample.map(it => `<tr><td>${it.ma_hang}</td><td>${it.ten_hang}</td><td class="num-cell">${fmtNum(it.cuoi_ky)}</td></tr>`).join("")}
          </tbody>
        </table>
      </div>`;
    confirmBtn.disabled = false;
  } catch (err) {
    previewEl.innerHTML = `<p class="hint" style="color:var(--danger)">Lỗi đọc file: ${err.message}</p>`;
    confirmBtn.disabled = true;
  }
});

document.getElementById("btnConfirmImport").addEventListener("click", async () => {
  const date = document.getElementById("importDate").value || todayISO();
  const userName = (await getSetting("userName", "")) || "";
  const recs = pendingImportItems.map((it) => ({
    id: makeId(date, it.kho, it.ma_hang),
    date, ...it,
    updatedBy: userName,
    updatedAt: new Date().toISOString(),
    source: "excel",
    synced: false,
  }));
  await putRecords(recs);
  showToast(`Đã lưu ${recs.length} dòng cho ngày ${fmtDateVN(date)}`);
  document.getElementById("importFile").value = "";
  document.getElementById("importPreview").innerHTML = "";
  document.getElementById("btnConfirmImport").disabled = true;
  pendingImportItems = [];
  await refreshAllDateSelectors();
  await maybeAutoSync();
});

/* ---------- Manual entry ---------- */
document.getElementById("mNgay").value = todayISO();

document.getElementById("btnSaveManual").addEventListener("click", async () => {
  const date = document.getElementById("mNgay").value || todayISO();
  const kho = document.getElementById("mKho").value.trim();
  const ma_hang = document.getElementById("mMa").value.trim();
  if (!kho || !ma_hang) { showToast("Vui lòng nhập Kho và Mã hàng"); return; }
  const userName = (await getSetting("userName", "")) || "";
  const rec = {
    id: makeId(date, kho, ma_hang),
    date, kho, ma_hang,
    ten_hang: document.getElementById("mTen").value.trim(),
    dvt: document.getElementById("mDvt").value.trim(),
    nhom: document.getElementById("mNhom").value.trim(),
    dau_ky: 0,
    nhap_kho: 0,
    xuat_kho: 0,
    cuoi_ky: num(document.getElementById("mCuoi").value),
    updatedBy: userName,
    updatedAt: new Date().toISOString(),
    source: "manual",
    synced: false,
  };
  await putRecord(rec);
  showToast("Đã lưu dòng tồn kho");
  ["mKho","mMa","mTen","mDvt","mNhom"].forEach(id => document.getElementById(id).value = "");
  document.getElementById("mCuoi").value = 0;
  await refreshAllDateSelectors();
  await maybeAutoSync();
});

/* ---------- Export Excel ---------- */
function buildExportRows(records, dateLabel) {
  const rows = [];
  rows.push(["TỔNG HỢP TỒN KHO"]);
  rows.push([`Xuất từ app Tồn Kho MT2, Ngày: ${dateLabel}`]);
  rows.push([]);
  rows.push(["Tên kho", "Mã hàng", "Tên hàng", "ĐVT", "Đầu kỳ", "Nhập kho", "Xuất kho", "Cuối kỳ", "Nhóm VTHH"]);
  rows.push([null, null, null, null, "Số lượng", "Số lượng", "Số lượng", "Số lượng", null]);
  let tDau = 0, tNhap = 0, tXuat = 0, tCuoi = 0;
  records.forEach((r) => {
    rows.push([r.kho, r.ma_hang, r.ten_hang, r.dvt, r.dau_ky, r.nhap_kho, r.xuat_kho, r.cuoi_ky, r.nhom]);
    tDau += num(r.dau_ky); tNhap += num(r.nhap_kho); tXuat += num(r.xuat_kho); tCuoi += num(r.cuoi_ky);
  });
  rows.push(["Tổng cộng", null, null, null, tDau, tNhap, tXuat, tCuoi, null]);
  return rows;
}

function downloadExcel(rows, filename) {
  const ws = XLSX.utils.aoa_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Tồn kho");
  XLSX.writeFile(wb, filename);
}

document.getElementById("btnExport").addEventListener("click", async () => {
  const date = document.getElementById("exportDate").value;
  if (!date) { showToast("Không có dữ liệu để xuất"); return; }
  const all = await getAllRecords();
  const recs = all.filter((r) => r.date === date);
  downloadExcel(buildExportRows(recs, fmtDateVN(date)), `TonKho_${date}.xlsx`);
});

document.getElementById("btnExportAll").addEventListener("click", async () => {
  const all = await getAllRecords();
  if (all.length === 0) { showToast("Chưa có dữ liệu"); return; }
  all.sort((a, b) => a.date.localeCompare(b.date));
  const rows = [["TỔNG HỢP TỒN KHO - TOÀN BỘ LỊCH SỬ"], [], ["Ngày", "Tên kho", "Mã hàng", "Tên hàng", "ĐVT", "Đầu kỳ", "Nhập kho", "Xuất kho", "Cuối kỳ", "Nhóm VTHH"]];
  all.forEach(r => rows.push([fmtDateVN(r.date), r.kho, r.ma_hang, r.ten_hang, r.dvt, r.dau_ky, r.nhap_kho, r.xuat_kho, r.cuoi_ky, r.nhom]));
  downloadExcel(rows, `TonKho_ToanBo_${todayISO()}.xlsx`);
});

/* ---------- Dashboard ---------- */
async function getDistinctDates() {
  const all = await getAllRecords();
  const dates = [...new Set(all.map(r => r.date))].sort().reverse();
  return dates;
}
async function getDistinctKho() {
  const all = await getAllRecords();
  return [...new Set(all.map(r => r.kho))].filter(Boolean).sort();
}
async function getDistinctGroups() {
  const all = await getAllRecords();
  return [...new Set(all.map(r => r.nhom))].filter(Boolean).sort();
}

async function refreshAllDateSelectors() {
  const dates = await getDistinctDates();
  const khoList = await getDistinctKho();
  const groups = await getDistinctGroups();

  fillSelect("dashDate", dates.map(d => ({ v: d, l: fmtDateVN(d) })), dates[0]);
  fillSelect("invDate", dates.map(d => ({ v: d, l: fmtDateVN(d) })), dates[0]);
  fillSelect("exportDate", dates.map(d => ({ v: d, l: fmtDateVN(d) })), dates[0]);
  fillSelect("dashKho", [{ v: "", l: "Tất cả kho" }, ...khoList.map(k => ({ v: k, l: k }))], "");
  fillSelect("invKho", [{ v: "", l: "Tất cả kho" }, ...khoList.map(k => ({ v: k, l: k }))], "");
  fillSelect("invGroup", [{ v: "", l: "Tất cả nhóm" }, ...groups.map(g => ({ v: g, l: g }))], "");

  renderDashboard();
  renderInventory();
}

function fillSelect(id, options, selected) {
  const el = document.getElementById(id);
  const prevValue = el.value;
  el.innerHTML = options.map(o => `<option value="${o.v}">${o.l}</option>`).join("");
  if (selected !== undefined && options.some(o => o.v === selected)) {
    el.value = selected;
  } else if (options.some(o => o.v === prevValue)) {
    el.value = prevValue;
  }
}

async function renderDashboard() {
  const date = document.getElementById("dashDate").value;
  const khoFilter = document.getElementById("dashKho").value;
  const all = await getAllRecords();
  let recs = all.filter(r => r.date === date);
  if (khoFilter) recs = recs.filter(r => r.kho === khoFilter);

  const totalSKU = recs.length;
  const totalCuoi = recs.reduce((s, r) => s + num(r.cuoi_ky), 0);
  document.getElementById("statSKU").textContent = fmtNum(totalSKU);
  document.getElementById("statCuoiKy").textContent = fmtNum(totalCuoi);

  const threshold = num(await getSetting("threshold", 20));
  const low = recs.filter(r => num(r.cuoi_ky) <= threshold).sort((a, b) => num(a.cuoi_ky) - num(b.cuoi_ky));
  document.getElementById("lowStockCount").textContent = low.length;
  const lowList = document.getElementById("lowStockList");
  if (low.length === 0) {
    lowList.innerHTML = `<div class="empty-state">Không có mã hàng nào dưới ngưỡng ${fmtNum(threshold)}</div>`;
  } else {
    lowList.innerHTML = low.slice(0, 30).map(r => `
      <div class="alert-item">
        <span>${r.ma_hang} - ${r.ten_hang}</span>
        <span class="low-stock">${fmtNum(r.cuoi_ky)} ${r.dvt || ""}</span>
      </div>`).join("");
  }

  // Compare with previous available date
  const dates = await getDistinctDates();
  const idx = dates.indexOf(date);
  const compareBox = document.getElementById("compareBox");
  if (idx === -1 || idx === dates.length - 1) {
    compareBox.innerHTML = `<div class="empty-state">Chưa đủ dữ liệu để so sánh</div>`;
  } else {
    const prevDate = dates[idx + 1];
    let prevRecs = all.filter(r => r.date === prevDate);
    if (khoFilter) prevRecs = prevRecs.filter(r => r.kho === khoFilter);
    const prevTotal = prevRecs.reduce((s, r) => s + num(r.cuoi_ky), 0);
    const diff = totalCuoi - prevTotal;
    const pct = prevTotal !== 0 ? ((diff / prevTotal) * 100).toFixed(1) : "—";
    const cls = diff > 0 ? "ok" : diff < 0 ? "danger" : "warn";
    const arrow = diff > 0 ? "▲" : diff < 0 ? "▼" : "—";
    compareBox.innerHTML = `
      <div class="row between">
        <span>So với ngày ${fmtDateVN(prevDate)}</span>
        <span class="badge ${cls}">${arrow} ${fmtNum(Math.abs(diff))} (${pct}%)</span>
      </div>`;
  }
}

document.getElementById("dashDate").addEventListener("change", renderDashboard);
document.getElementById("dashKho").addEventListener("change", renderDashboard);

/* ---------- Inventory list ---------- */
async function renderInventory() {
  const search = document.getElementById("invSearch").value.trim().toLowerCase();
  const group = document.getElementById("invGroup").value;
  const kho = document.getElementById("invKho").value;
  const date = document.getElementById("invDate").value;
  const all = await getAllRecords();
  let recs = all.filter(r => r.date === date);
  if (group) recs = recs.filter(r => r.nhom === group);
  if (kho) recs = recs.filter(r => r.kho === kho);
  if (search) recs = recs.filter(r => r.ma_hang.toLowerCase().includes(search) || (r.ten_hang || "").toLowerCase().includes(search));
  recs.sort((a, b) => a.ma_hang.localeCompare(b.ma_hang));

  const threshold = num(await getSetting("threshold", 20));
  const tbody = document.getElementById("invTableBody");
  const emptyEl = document.getElementById("invEmpty");
  if (recs.length === 0) {
    tbody.innerHTML = "";
    emptyEl.style.display = "block";
  } else {
    emptyEl.style.display = "none";
    tbody.innerHTML = recs.map(r => `
      <tr data-id="${r.id}">
        <td>${r.ma_hang}</td>
        <td>${r.ten_hang || ""}</td>
        <td>${r.dvt || ""}</td>
        <td class="num-cell ${num(r.cuoi_ky) <= threshold ? 'low-stock' : ''}">${fmtNum(r.cuoi_ky)}</td>
      </tr>`).join("");
  }
}
["invSearch", "invGroup", "invKho", "invDate"].forEach(id => {
  document.getElementById(id).addEventListener("input", renderInventory);
  document.getElementById(id).addEventListener("change", renderInventory);
});

/* ---------- History / trend ---------- */
let histChartInstance = null;

async function renderHistorySelect() {
  const all = await getAllRecords();
  const products = {};
  all.forEach(r => { products[r.ma_hang] = r.ten_hang || products[r.ma_hang] || ""; });
  const keys = Object.keys(products).sort();
  const sel = document.getElementById("histSelect");
  const prev = sel.value;
  sel.innerHTML = `<option value="">-- chọn mã hàng --</option>` + keys.map(k => `<option value="${k}">${k} - ${products[k]}</option>`).join("");
  if (keys.includes(prev)) { sel.value = prev; renderHistoryFor(prev); }
}

document.getElementById("histSelect").addEventListener("change", (e) => renderHistoryFor(e.target.value));

async function renderHistoryFor(maHang) {
  const tbody = document.getElementById("histTableBody");
  const emptyEl = document.getElementById("histEmpty");
  const fallback = document.getElementById("histChartFallback");
  if (!maHang) {
    tbody.innerHTML = "";
    emptyEl.style.display = "block";
    if (histChartInstance) { histChartInstance.destroy(); histChartInstance = null; }
    return;
  }
  const all = await getAllRecords();
  const recs = all.filter(r => r.ma_hang === maHang).sort((a, b) => a.date.localeCompare(b.date));
  if (recs.length === 0) {
    emptyEl.style.display = "block";
    tbody.innerHTML = "";
    return;
  }
  emptyEl.style.display = "none";
  tbody.innerHTML = recs.slice().reverse().map(r => `
    <tr>
      <td>${fmtDateVN(r.date)}</td>
      <td class="num-cell">${fmtNum(r.cuoi_ky)}</td>
    </tr>`).join("");

  const canvas = document.getElementById("histChart");
  if (typeof Chart === "undefined") {
    fallback.textContent = "Không thể tải biểu đồ (cần kết nối mạng lần đầu). Xem bảng số liệu bên dưới.";
    return;
  }
  fallback.textContent = "";
  if (histChartInstance) histChartInstance.destroy();
  histChartInstance = new Chart(canvas, {
    type: "line",
    data: {
      labels: recs.map(r => fmtDateVN(r.date)),
      datasets: [{
        label: "Tồn cuối kỳ",
        data: recs.map(r => r.cuoi_ky),
        borderColor: "#0f766e",
        backgroundColor: "rgba(15,118,110,0.1)",
        tension: 0.25,
        fill: true,
      }],
    },
    options: {
      responsive: true,
      plugins: { legend: { display: false } },
      scales: { y: { beginAtZero: true } },
    },
  });
}

/* ---------- Settings ---------- */
async function renderSettings() {
  document.getElementById("setUserName").value = (await getSetting("userName", "")) || "";
  document.getElementById("setThreshold").value = await getSetting("threshold", 20);
  document.getElementById("setSheetUrl").value = (await getSetting("sheetUrl", "")) || "";
  const lastSync = await getSetting("lastSync", null);
  document.getElementById("lastSyncLabel").textContent = lastSync
    ? `Lần đồng bộ gần nhất: ${new Date(lastSync).toLocaleString("vi-VN")}`
    : "Chưa đồng bộ lần nào";
}

document.getElementById("btnSaveUser").addEventListener("click", async () => {
  await setSetting("userName", document.getElementById("setUserName").value.trim());
  showToast("Đã lưu tên người dùng");
});
document.getElementById("btnSaveThreshold").addEventListener("click", async () => {
  await setSetting("threshold", num(document.getElementById("setThreshold").value));
  showToast("Đã lưu ngưỡng cảnh báo");
  renderDashboard();
});
document.getElementById("btnSaveSheetUrl").addEventListener("click", async () => {
  await setSetting("sheetUrl", document.getElementById("setSheetUrl").value.trim());
  showToast("Đã lưu URL đồng bộ");
});
document.getElementById("btnResetAll").addEventListener("click", async () => {
  if (!confirm("Xoá toàn bộ dữ liệu tồn kho đã lưu trên máy này? Hành động không thể hoàn tác.")) return;
  await clearAllRecords();
  showToast("Đã xoá toàn bộ dữ liệu");
  await refreshAllDateSelectors();
});

/* ---------- Google Sheet sync ---------- */
const GS_CODE = `function doGet(e) {
  var sheet = getSheet_();
  var data = sheet.getDataRange().getValues();
  var headers = data.shift();
  var records = data.map(function(row) {
    var obj = {};
    headers.forEach(function(h, i) { obj[h] = row[i]; });
    return obj;
  });
  return ContentService.createTextOutput(JSON.stringify(records))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var payload = JSON.parse(e.postData.contents);
  var records = payload.records || [];
  var sheet = getSheet_();
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var idCol = headers.indexOf("id");
  var idToRow = {};
  for (var i = 1; i < data.length; i++) idToRow[data[i][idCol]] = i + 1;

  records.forEach(function(rec) {
    var row = headers.map(function(h) { return rec[h] !== undefined ? rec[h] : ""; });
    if (idToRow[rec.id]) {
      sheet.getRange(idToRow[rec.id], 1, 1, headers.length).setValues([row]);
    } else {
      sheet.appendRow(row);
    }
  });
  return ContentService.createTextOutput(JSON.stringify({ ok: true, count: records.length }))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Data");
  if (!sheet) {
    sheet = ss.insertSheet("Data");
    sheet.appendRow(["id","date","kho","ma_hang","ten_hang","dvt","dau_ky","nhap_kho","xuat_kho","cuoi_ky","nhom","updatedBy","updatedAt","source"]);
  }
  return sheet;
}`;

document.getElementById("gsCodeBox").value = GS_CODE;
document.getElementById("btnCopyGsCode").addEventListener("click", async () => {
  try {
    await navigator.clipboard.writeText(GS_CODE);
    showToast("Đã copy code Apps Script");
  } catch {
    document.getElementById("gsCodeBox").select();
    document.execCommand("copy");
    showToast("Đã copy code Apps Script");
  }
});

async function syncNow(silent) {
  const url = await getSetting("sheetUrl", "");
  if (!url) { if (!silent) showToast("Chưa cấu hình URL Google Sheet"); return; }
  if (!navigator.onLine) { if (!silent) showToast("Không có kết nối mạng"); return; }
  try {
    // push unsynced
    const all = await getAllRecords();
    const unsynced = all.filter(r => !r.synced);
    if (unsynced.length > 0) {
      await fetch(url, {
        method: "POST",
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        body: JSON.stringify({ records: unsynced }),
      });
      const marked = unsynced.map(r => ({ ...r, synced: true }));
      await putRecords(marked);
    }
    // pull all
    const res = await fetch(url + (url.includes("?") ? "&" : "?") + "action=list");
    if (res.ok) {
      const remote = await res.json();
      const remoteRecs = remote
        .filter(r => r.id && r.date && r.ma_hang)
        .map(r => ({
          id: r.id, date: r.date, kho: r.kho, ma_hang: r.ma_hang, ten_hang: r.ten_hang,
          dvt: r.dvt, dau_ky: num(r.dau_ky), nhap_kho: num(r.nhap_kho), xuat_kho: num(r.xuat_kho),
          cuoi_ky: num(r.cuoi_ky), nhom: r.nhom, updatedBy: r.updatedBy, updatedAt: r.updatedAt,
          source: r.source || "sheet", synced: true,
        }));
      if (remoteRecs.length > 0) await putRecords(remoteRecs);
    }
    await setSetting("lastSync", new Date().toISOString());
    if (!silent) showToast("Đồng bộ thành công");
    await refreshAllDateSelectors();
    renderSettings();
  } catch (err) {
    if (!silent) showToast("Lỗi đồng bộ: " + err.message);
  }
}

document.getElementById("btnSyncNow").addEventListener("click", () => syncNow(false));

async function maybeAutoSync() {
  const url = await getSetting("sheetUrl", "");
  if (url && navigator.onLine) syncNow(true);
}

/* ---------- Online/offline indicator ---------- */
function updateOnlineStatus() {
  const dot = document.getElementById("syncDot");
  const label = document.getElementById("syncLabel");
  if (navigator.onLine) {
    dot.classList.add("online"); dot.classList.remove("offline");
    label.textContent = "Online";
  } else {
    dot.classList.add("offline"); dot.classList.remove("online");
    label.textContent = "Offline";
  }
}
window.addEventListener("online", () => { updateOnlineStatus(); maybeAutoSync(); });
window.addEventListener("offline", updateOnlineStatus);

// Đồng bộ ngay khi người dùng quay lại app (mở lại tab, mở lại từ nền)
document.addEventListener("visibilitychange", () => {
  if (document.visibilityState === "visible") maybeAutoSync();
});

/* ---------- PWA install prompt ---------- */
let deferredPrompt = null;
window.addEventListener("beforeinstallprompt", (e) => {
  e.preventDefault();
  deferredPrompt = e;
  document.getElementById("installBanner").style.display = "flex";
});
document.getElementById("btnInstall").addEventListener("click", async () => {
  if (!deferredPrompt) return;
  deferredPrompt.prompt();
  await deferredPrompt.userChoice;
  deferredPrompt = null;
  document.getElementById("installBanner").style.display = "none";
});
window.addEventListener("appinstalled", () => {
  document.getElementById("installBanner").style.display = "none";
});

/* ---------- Service worker ---------- */
if ("serviceWorker" in navigator) {
  window.addEventListener("load", () => {
    navigator.serviceWorker.register("sw.js").catch(() => {});
  });
}

/* ---------- Init ---------- */
(async function init() {
  updateOnlineStatus();
  await refreshAllDateSelectors();
  renderSettings();
  renderHistorySelect();
  // Tự động đồng bộ định kỳ mỗi 60 giây khi app đang mở và có mạng
  setInterval(() => { if (navigator.onLine) maybeAutoSync(); }, 60 * 1000);
})();

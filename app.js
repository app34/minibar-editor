
const $ = (id) => document.getElementById(id);

const state = {
  workbook: null,
  fileName: "Minibar Consumption August 2026.xlsx",
  dirty: false,
  zone: "Beach",
  daySheet: null,
  rooms: [],
  items: [],
  selectedRoom: null,
  sheetMode: false,
};

function toast(msg) {
  const el = $("toast");
  el.textContent = msg;
  el.classList.add("show");
  clearTimeout(toast._t);
  toast._t = setTimeout(() => el.classList.remove("show"), 2200);
}

function zoneOf(room) {
  const n = Number(String(room).replace(/\D/g, ""));
  const block = Math.floor(n / 100);
  return block <= 3 ? "Beach" : "Water";
}

function sheetDayNumber(name) {
  const m = String(name).match(/^(\d+)/);
  return m ? Number(m[1]) : null;
}

function dailySheets() {
  return state.workbook.worksheets
    .filter((ws) => sheetDayNumber(ws.name) != null)
    .sort((a, b) => sheetDayNumber(a.name) - sheetDayNumber(b.name));
}

function findSheet(name) {
  return state.workbook.worksheets.find((ws) => ws.name === name);
}

function cellVal(ws, r, c) {
  const v = ws.getRow(r).getCell(c).value;
  if (v == null || v === "") return null;
  if (typeof v === "object") {
    if (v.result != null) return v.result;
    if (v.richText) return v.richText.map((t) => t.text).join("");
    if (v.text) return v.text;
    if (v.formula) return v.result ?? null;
  }
  return v;
}

function setValueKeepStyle(ws, r, c, value) {
  ws.getRow(r).getCell(c).value = value === "" || value == null ? null : value;
}


function colLetter(n) {
  let s = "";
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function findHeaderCol(ws, name) {
  const needle = String(name).toLowerCase();
  let found = null;
  ws.getRow(2).eachCell({ includeEmpty: false }, (cell, col) => {
    if (String(cell.value || "").toLowerCase().trim() === needle) found = col;
  });
  return found;
}

function roomColumns(ws) {
  return state.rooms.map((r) => r.col).sort((a, b) => a - b);
}

function sumItemRow(ws, row) {
  let t = 0;
  state.rooms.forEach((r) => {
    t += Number(cellVal(ws, row, r.col)) || 0;
  });
  return t;
}

function refreshItemTotal(ws, row) {
  const totalCol = findHeaderCol(ws, "Total") || 140;
  const cols = roomColumns(ws);
  if (!cols.length) return 0;
  const start = colLetter(cols[0]);
  const end = colLetter(cols[cols.length - 1]);
  const sum = sumItemRow(ws, row);
  const cell = ws.getRow(row).getCell(totalCol);
  const existing = cell.value;
  const formula = (existing && typeof existing === "object" && existing.formula)
    ? existing.formula
    : `SUM(${start}${row}:${end}${row})`;
  cell.value = { formula, result: sum };
  return sum;
}

function refreshDayTotals() {
  if (!state.daySheet) return 0;
  let grand = 0;
  state.items.forEach((item) => {
    grand += refreshItemTotal(state.daySheet, item.row);
  });
  updateSummaryForDay();
  const el = $("dayTotal");
  if (el) el.textContent = `Day total: ${grand} units`;
  return grand;
}

function updateSummaryForDay() {
  const summary = findSheet("Summary");
  if (!summary || !state.daySheet) return;
  const day = sheetDayNumber(state.daySheet.name);
  if (!day) return;
  let dayCol = null;
  summary.getRow(2).eachCell({ includeEmpty: false }, (cell, col) => {
    if (Number(cell.value) === day) dayCol = col;
  });
  if (!dayCol) return;
  state.items.forEach((item) => {
    const total = sumItemRow(state.daySheet, item.row);
    let targetRow = null;
    summary.eachRow({ includeEmpty: false }, (row, r) => {
      if (r < 3) return;
      const name = cellVal(summary, r, 3);
      if (name && String(name).toLowerCase() === item.name.toLowerCase()) targetRow = r;
    });
    if (!targetRow) return;
    const cell = summary.getRow(targetRow).getCell(dayCol);
    const existing = cell.value;
    if (existing && typeof existing === "object" && existing.formula) {
      cell.value = { formula: existing.formula, result: total };
    } else {
      cell.value = total || null;
    }
  });
}

function classifyItem(name) {
  const n = String(name || "").toLowerCase();
  if (/peanut|crisp|cashew|pistachio|beet/.test(n)) return "snack";
  if (/coca|sprite|softdrink|diet|zero|tonic|soda|ginger/.test(n)) return "soft";
  if (/grani/.test(n)) return "snack";
  if (/beer|wine|mancura|barone|bottega|pinot|merlot|sicil|tagus|lion|saigon|xibeca|pilsner|royal|333|alc/.test(n)) return "alc";
  return "food";
}

function loadDailyMeta(ws) {
  const rooms = [];
  ws.getRow(2).eachCell({ includeEmpty: false }, (cell, col) => {
    const v = cell.value;
    if (typeof v === "number" || /^\d+$/.test(String(v))) rooms.push({ col, room: String(v) });
  });
  const items = [];
  ws.eachRow({ includeEmpty: false }, (row, r) => {
    if (r < 3) return;
    const article = cellVal(ws, r, 2);
    if (!article || String(article).toLowerCase() === "total") return;
    items.push({ row: r, name: String(article), cat: classifyItem(article) });
  });
  state.rooms = rooms;
  state.items = items;
  state.daySheet = ws;
}

function roomUsedCount(room) {
  let n = 0;
  state.items.forEach((item) => {
    const q = Number(cellVal(state.daySheet, item.row, room.col)) || 0;
    if (q) n += q;
  });
  return n;
}

function renderDaySelect() {
  const sel = $("daySelect");
  sel.innerHTML = "";
  dailySheets().forEach((ws) => {
    const opt = document.createElement("option");
    opt.value = ws.name;
    opt.textContent = ws.name.replace(/\s+/g, " ");
    sel.appendChild(opt);
  });
  sel.disabled = false;
  if (sel.options.length) {
    sel.value = sel.options[0].value;
    onDayChange();
  }
}

function onDayChange() {
  const ws = findSheet($("daySelect").value);
  if (!ws) return;
  loadDailyMeta(ws);
  refreshDayTotals();
  renderGrid();
}

function renderGrid() {
  const q = String($("roomFilter").value || "").trim();
  const list = state.rooms.filter((r) => zoneOf(r.room) === state.zone && (!q || r.room.includes(q)));
  $("zoneTitle").textContent = `${state.zone.toUpperCase()} ZONE (${list.length})`;
  const grid = $("villaGrid");
  grid.innerHTML = "";
  if (!state.workbook) {
    grid.innerHTML = "";
    return;
  }
  list.forEach((r) => {
    const used = roomUsedCount(r);
    const b = document.createElement("button");
    b.className = "villa-tile" + (used ? " done" : "");
    b.textContent = r.room;
    if (used) {
      const badge = document.createElement("span");
      badge.className = "count-badge";
      badge.textContent = used;
      b.appendChild(badge);
    }
    b.onclick = () => openRoom(r);
    grid.appendChild(b);
  });
}

function openRoom(room) {
  state.selectedRoom = room;
  $("modalTitle").textContent = `${room.room} (${state.zone})`;
  renderItems();
  refreshDayTotals();
  $("roomModal").classList.add("open");
}

function closeRoom() {
  $("roomModal").classList.remove("open");
  renderGrid();
}

function renderItems() {
  const box = $("itemGrid");
  box.innerHTML = "";
  if (!state.selectedRoom) return;
  state.items.forEach((item) => {
    const qty = Number(cellVal(state.daySheet, item.row, state.selectedRoom.col)) || 0;
    const card = document.createElement("article");
    card.className = "item-card cat-" + item.cat + (qty ? " has-qty" : "");
    card.innerHTML = `
      <div class="item-name">${escapeHtml(item.name)}</div>
      <div class="stepper">
        <button class="btn-step" data-act="-">-</button>
        <span class="step-val">${qty}</span>
        <button class="btn-step" data-act="+">+</button>
      </div>`;
    const val = card.querySelector(".step-val");
    const write = (n) => {
      const v = Math.max(0, n);
      val.textContent = v;
      card.classList.toggle("has-qty", v > 0);
      setValueKeepStyle(state.daySheet, item.row, state.selectedRoom.col, v || null);
      state.dirty = true;
      refreshDayTotals();
    };
    card.querySelector('[data-act="-"]').onclick = () => write((Number(val.textContent) || 0) - 1);
    card.querySelector('[data-act="+"]').onclick = () => write((Number(val.textContent) || 0) + 1);
    box.appendChild(card);
  });
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, (ch) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[ch]));
}

async function openFile(file) {
  $("loader").classList.remove("hidden");
  try {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(await file.arrayBuffer());
    state.workbook = wb;
    state.fileName = file.name || state.fileName;
    $("fileStatus").textContent = state.fileName;
    $("saveBtn").disabled = false;
    $("finishBtn").disabled = false;
    if ($("waBtn")) $("waBtn").disabled = false;
    renderDaySelect();
    toast("Workbook loaded");
  } catch (e) {
    console.error(e);
    toast("Could not open file");
  } finally {
    $("loader").classList.add("hidden");
  }
}

async function saveSameFile() {
  if (!state.workbook) return;
  $("loader").classList.remove("hidden");
  try {
    refreshDayTotals();
    const buf = await state.workbook.xlsx.writeBuffer();
    const a = document.createElement("a");
    a.href = URL.createObjectURL(new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
    a.download = state.fileName;
    a.click();
    URL.revokeObjectURL(a.href);
    state.dirty = false;
    toast("Saved " + state.fileName);
  } finally {
    $("loader").classList.add("hidden");
  }
}

function cleanZone() {
  if (!state.daySheet) return;
  if (!confirm("Clear quantities for visible " + state.zone + " villas on this date?")) return;
  const q = String($("roomFilter").value || "").trim();
  state.rooms.filter((r) => zoneOf(r.room) === state.zone && (!q || r.room.includes(q))).forEach((room) => {
    state.items.forEach((item) => setValueKeepStyle(state.daySheet, item.row, room.col, null));
  });
  state.dirty = true;
  refreshDayTotals();
  renderGrid();
  toast("Zone cleared");
}


async function workbookFile() {
  refreshDayTotals();
  const buf = await state.workbook.xlsx.writeBuffer();
  return new File([buf], state.fileName, {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
}

async function shareWhatsApp() {
  if (!state.workbook) return;
  $("loader").classList.remove("hidden");
  try {
    const file = await workbookFile();
    if (navigator.canShare && navigator.canShare({ files: [file] })) {
      await navigator.share({
        files: [file],
        title: state.fileName,
        text: "Minibar " + (state.daySheet ? state.daySheet.name.replace(/\s+/g, " ") : ""),
      });
      state.dirty = false;
      toast("Shared to WhatsApp / apps");
      return;
    }
    const a = document.createElement("a");
    a.href = URL.createObjectURL(file);
    a.download = state.fileName;
    a.click();
    URL.revokeObjectURL(a.href);
    toast("XL downloaded — attach it in WhatsApp");
  } catch (e) {
    if (String(e.name) === "AbortError") return;
    console.error(e);
    toast("Share cancelled");
  } finally {
    $("loader").classList.add("hidden");
  }
}

function setZone(zone) {
  state.zone = zone;
  $("btnBeach").classList.toggle("active", zone === "Beach");
  $("btnWater").classList.toggle("active", zone === "Water");
  renderGrid();
}

$("btnBeach").onclick = () => setZone("Beach");
$("btnWater").onclick = () => setZone("Water");
$("daySelect").onchange = onDayChange;
$("roomFilter").oninput = renderGrid;
$("fileInput").onchange = (e) => { const f = e.target.files && e.target.files[0]; if (f) openFile(f); };
$("saveBtn").onclick = saveSameFile;
$("finishBtn").onclick = saveSameFile;
$("cleanBtn").onclick = cleanZone;
$("closeModalBtn").onclick = closeRoom;
$("saveEntryBtn").onclick = () => { closeRoom(); toast("Local entry kept in file"); };
$("sheetViewBtn").onclick = () => {
  state.sheetMode = true;
  $("modalCard").classList.add("sheet");
  $("sheetViewBtn").classList.add("active");
  $("fullViewBtn").classList.remove("active");
};
$("fullViewBtn").onclick = () => {
  state.sheetMode = false;
  $("modalCard").classList.remove("sheet");
  $("fullViewBtn").classList.add("active");
  $("sheetViewBtn").classList.remove("active");
};
$("roomModal").addEventListener("click", (e) => { if (e.target.id === "roomModal") closeRoom(); });
window.addEventListener("beforeunload", (e) => { if (state.dirty) { e.preventDefault(); e.returnValue = ""; } });

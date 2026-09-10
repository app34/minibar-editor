
const $ = (id) => document.getElementById(id);
const APP_VERSION = "3.1.0";

const MONTH_NAMES = ["January","February","March","April","May","June","July","August","September","October","November","December"];
const HISTORY_LIMIT = 12;

function openDb() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open("minibar-xl", 1);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains("kv")) db.createObjectStore("kv");
      if (!db.objectStoreNames.contains("history")) db.createObjectStore("history", { keyPath: "id" });
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

function kvSet(key, value) {
  return openDb().then((db) => new Promise((resolve, reject) => {
    const tx = db.transaction("kv", "readwrite");
    tx.objectStore("kv").put(value, key);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  }));
}

function kvGet(key) {
  return openDb().then((db) => new Promise((resolve, reject) => {
    const tx = db.transaction("kv", "readonly");
    const req = tx.objectStore("kv").get(key);
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  }));
}

function historyPut(row) {
  return openDb().then((db) => new Promise((resolve, reject) => {
    const tx = db.transaction("history", "readwrite");
    tx.objectStore("history").put(row);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  }));
}

function historyAll() {
  return openDb().then((db) => new Promise((resolve, reject) => {
    const tx = db.transaction("history", "readonly");
    const req = tx.objectStore("history").getAll();
    req.onsuccess = () => resolve(req.result || []);
    req.onerror = () => reject(req.error);
  }));
}

function historyDel(id) {
  return openDb().then((db) => new Promise((resolve, reject) => {
    const tx = db.transaction("history", "readwrite");
    tx.objectStore("history").delete(id);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  }));
}

function downloadFileName() {
  const custom = (state.downloadName || "").trim();
  if (custom) return custom.endsWith(".xlsx") ? custom : custom + ".xlsx";
  return state.fileName || "Minibar.xlsx";
}

function markReady() {
  $("fileStatus").textContent = state.fileName + " · kept on this device";
  $("saveBtn").disabled = false;
  $("finishBtn").disabled = false;
  if ($("waBtn")) $("waBtn").disabled = false;
  if ($("importTextBtn")) $("importTextBtn").disabled = false;
  if ($("tableBtn")) $("tableBtn").disabled = false;
}

async function persistCurrent(alsoHistory, note) {
  if (!state.workbook) return null;
  refreshDayTotals();
  const buf = await state.workbook.xlsx.writeBuffer();
  state.cachedBuf = buf;
  const name = downloadFileName();
  state.cachedFile = new File([buf], name, { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  state.cachedOctet = new File([buf], name, { type: "application/octet-stream" });
  await kvSet("current", { fileName: state.fileName, downloadName: state.downloadName || "", buf, savedAt: Date.now() });
  await kvSet("settings", { downloadName: state.downloadName || "" });
  if (alsoHistory) {
    const row = { id: Date.now(), fileName: name, note: note || "snapshot", buf, savedAt: Date.now() };
    await historyPut(row);
    const all = (await historyAll()).sort((a, b) => b.id - a.id);
    for (const extra of all.slice(HISTORY_LIMIT)) await historyDel(extra.id);
  }
  return buf;
}

function schedulePersist() {
  clearTimeout(schedulePersist._t);
  schedulePersist._t = setTimeout(() => persistCurrent(false).catch(console.error), 700);
}

async function loadPersistedWorkbook() {
  const settings = await kvGet("settings");
  if (settings && settings.downloadName) state.downloadName = settings.downloadName;
  const cur = await kvGet("current");
  if (!cur || !cur.buf) return false;
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(cur.buf);
  state.workbook = wb;
  state.fileName = cur.fileName || state.fileName;
  state.cachedBuf = cur.buf;
  markReady();
  renderDaySelect();
  toast("Loaded saved workbook");
  return true;
}



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
  usedOnly: false,
  downloadName: "",
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

const MONTH_INDEX = {
  jan: 0, january: 0, feb: 1, february: 1, mar: 2, march: 2,
  apr: 3, april: 3, may: 4, jun: 5, june: 5, jul: 6, july: 6,
  aug: 7, august: 7, sep: 8, sept: 8, september: 8,
  oct: 9, october: 9, nov: 10, november: 10, dec: 11, december: 11,
};

function parseSheetDate(name) {
  const m = String(name).match(/^(\d+)\s*([A-Za-z]+)?/);
  if (!m) return null;
  const monthName = (m[2] || "").toLowerCase();
  return {
    day: Number(m[1]),
    month: MONTH_INDEX[monthName],
    name,
  };
}

function fileMonthYear() {
  const fromName = String(state.fileName || "").match(/(January|February|March|April|May|June|July|August|September|October|November|December)\s+(\d{4})/i);
  if (fromName) return { month: MONTH_INDEX[fromName[1].toLowerCase()], year: Number(fromName[2]) };
  return null;
}

function pickDefaultSheet(names) {
  const now = new Date();
  const todayD = now.getDate();
  const todayM = now.getMonth();
  const todayY = now.getFullYear();
  const rows = names.map(parseSheetDate).filter((x) => x && x.day);
  if (!rows.length) return names[names.length - 1] || names[0];

  const fileMeta = fileMonthYear();
  const months = rows.map((r) => r.month).filter((m) => m != null);
  const fileMonth = fileMeta ? fileMeta.month : months.sort((a, b) =>
    months.filter((x) => x === a).length - months.filter((x) => x === b).length
  ).pop();
  const fileYear = fileMeta ? fileMeta.year : todayY;

  const sameMonth = fileMonth === todayM && fileYear === todayY;
  if (sameMonth) {
    const exact = rows.find((r) => r.day === todayD);
    if (exact) return exact.name;
    const maxDay = Math.max(...rows.map((r) => r.day));
    if (todayD > maxDay) return rows.sort((a, b) => b.day - a.day)[0].name;
    rows.sort((a, b) => Math.abs(a.day - todayD) - Math.abs(b.day - todayD) || b.day - a.day);
    return rows[0].name;
  }

  const filePast = fileYear < todayY || (fileYear === todayY && fileMonth < todayM);
  const ordered = rows.slice().sort((a, b) => a.day - b.day);
  return filePast ? ordered[ordered.length - 1].name : ordered[0].name;
}

function fillDayOptions(sel, selected) {
  if (!sel) return;
  sel.innerHTML = "";
  dailySheets().forEach((ws) => {
    const opt = document.createElement("option");
    opt.value = ws.name;
    opt.textContent = ws.name.replace(/\s+/g, " ");
    sel.appendChild(opt);
  });
  sel.disabled = !sel.options.length;
  if (selected && [...sel.options].some((o) => o.value === selected)) sel.value = selected;
}

function renderDaySelect() {
  const sel = $("daySelect");
  fillDayOptions(sel);
  fillDayOptions($("tableDaySelect"), state.daySheet && state.daySheet.name);
  sel.disabled = false;
  if (sel.options.length) {
    const names = [...sel.options].map((o) => o.value);
    sel.value = pickDefaultSheet(names);
    onDayChange();
    fillDayOptions($("tableDaySelect"), sel.value);
  }
}

function onDayChange() {
  const ws = findSheet($("daySelect").value);
  if (!ws) return;
  loadDailyMeta(ws);
  refreshDayTotals();
  renderGrid();
  if ($("tableDaySelect")) fillDayOptions($("tableDaySelect"), ws.name);
}

async function onTableDayChange() {
  const name = $("tableDaySelect") && $("tableDaySelect").value;
  if (!name || (state.daySheet && state.daySheet.name === name)) return;
  await persistCurrent(true, "before " + name);
  if ($("daySelect")) $("daySelect").value = name;
  onDayChange();
  state.tableFocus = null;
  state.tableUndo = [];
  state.tableRedo = [];
  renderTableView();
  toast("Saved previous day · " + name.replace(/\s+/g, " "));
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
    if (state.usedOnly && !qty) return;
    const card = document.createElement("article");
    card.className = "item-card cat-" + item.cat + (qty ? " has-qty" : "");
    card.innerHTML = `
      <div class="item-name">${escapeHtml(shortName(item.name))}</div>
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
      schedulePersist();
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
    if ($("importTextBtn")) $("importTextBtn").disabled = false;
  if ($("tableBtn")) $("tableBtn").disabled = false;
    renderDaySelect();
    toast("Workbook saved on this device");
    persistCurrent(true, "imported").then(markReady);
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
    await persistCurrent(true, "downloaded");
    const buf = state.cachedBuf || await state.workbook.xlsx.writeBuffer();
    const a = document.createElement("a");
    a.href = URL.createObjectURL(new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
    a.download = downloadFileName();
    a.click();
    URL.revokeObjectURL(a.href);
    state.dirty = false;
    toast("Saved " + downloadFileName());
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
  schedulePersist();
  toast("Zone cleared");
}



function normName(s) {
  return String(s || "")
    .toLowerCase()
    .replace(/softdrinks?/g, " ")
    .replace(/\bmb\b/g, " ")
    .replace(/\bww\b/g, " ")
    .replace(/\bcan\b/g, " ")
    .replace(/\bbeer\b/g, " ")
    .replace(/[\d.]+(\s)?(ml|grm|g|gram|grams)\b/g, " ")
    .replace(/[^a-z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

const NAME_ALIAS = {
  "potato crisps": "potato crisps",
  "salted peanuts": "salted peanuts",
  "baked beetroots crips": "baked beetroots",
  "baked beetroot crisps": "baked beetroots",
  "cashew salted": "cashew salted",
  "roasted salted pistachios": "roasted salted pistachios",
  "coca cola": "coca cola",
  "sprite": "sprite",
  "diet coke": "diet coke",
  "zero coke": "zero coke",
  "tonic water": "tonic water",
  "soda water": "soda water",
  "ginger ale": "ginger ale",
  "pilsner extra smooth": "xibeca",
  "beer pilsner extra smooth": "xibeca",
  "xibeca damm": "xibeca",
  "saigon lager": "saigon",
  "tagus cerveza lager": "tagus",
  "lion": "lion",
  "beer lion": "lion",
  "333 triple": "333",
  "333 triple can beer": "333",
  "grani mix fit": "grani mix",
  "grani orgng": "grani orgng",
  "grani orange": "grani orgng",
  "grani apple": "grani apple",
  "royal dutch": "xibeca",
};

function matchItem(reportName) {
  const raw = normName(reportName);
  const alias = NAME_ALIAS[raw] || raw;
  let best = null;
  let bestScore = 0;
  state.items.forEach((item) => {
    const n = normName(item.name);
    let score = 0;
    if (n === raw || n === alias) score = 100;
    else if (n.includes(alias) || alias.includes(n) || n.includes(raw) || raw.includes(n)) score = 82;
    else {
      const ta = alias.split(" ").filter((w) => w.length > 2);
      const tb = new Set(n.split(" "));
      const hit = ta.filter((w) => tb.has(w)).length;
      score = ta.length ? (hit / ta.length) * 70 : 0;
    }
    if (score > bestScore) {
      bestScore = score;
      best = item;
    }
  });
  return bestScore >= 45 ? best : null;
}

function parseMinibarReport(text) {
  const rows = [];
  let zone = null;
  String(text || "").split(/\r?\n/).forEach((raw) => {
    const line = raw.trim();
    if (!line) return;
    const head = line.match(/MINIBAR REPORT\s*\(([^)]+)\)/i);
    if (head) {
      zone = /water/i.test(head[1]) ? "Water" : /beach/i.test(head[1]) ? "Beach" : head[1];
      return;
    }
    if (/^---/.test(line)) return;
    const m = line.match(/^(\d+)\s*:\s*(.*)$/);
    if (!m) return;
    const items = [];
    m[2].split(",").forEach((part) => {
      const bit = part.trim();
      if (!bit) return;
      const eq = bit.match(/^(.+?)\s*=\s*(\d+)\s*$/);
      if (eq) items.push({ name: eq[1].trim(), qty: Number(eq[2]) });
    });
    rows.push({ room: m[1], items, zone });
  });
  return { zone, rows };
}

function applyTextReport(text, replaceRooms) {
  if (!state.daySheet) throw new Error("Select date first");
  const parsed = parseMinibarReport(text);
  if (!parsed.rows.length) throw new Error("No villa lines found");
  if (parsed.zone) setZone(parsed.zone);
  const log = [];
  let filled = 0;
  let missed = 0;
  parsed.rows.forEach((row) => {
    const room = state.rooms.find((r) => r.room === String(row.room));
    if (!room) {
      log.push("Villa " + row.room + " not in sheet");
      missed++;
      return;
    }
    if (replaceRooms) {
      state.items.forEach((item) => setValueKeepStyle(state.daySheet, item.row, room.col, null));
    }
    row.items.forEach((it) => {
      const item = matchItem(it.name);
      if (!item) {
        log.push(row.room + ": no match for " + it.name);
        missed++;
        return;
      }
      setValueKeepStyle(state.daySheet, item.row, room.col, it.qty || null);
      filled++;
    });
  });
  state.dirty = true;
  refreshDayTotals();
  renderGrid();
  cacheWorkbook();
  return { filled, missed, villas: parsed.rows.length, day: state.daySheet.name.replace(/\s+/g, " "), log };
}

function buildDayReportText() {
  if (!state.daySheet) return "";
  const zones = ["Beach", "Water"];
  const parts = [];
  zones.forEach((zone) => {
    const lines = [];
    state.rooms.filter((r) => zoneOf(r.room) === zone).forEach((room) => {
      const bits = [];
      state.items.forEach((item) => {
        const q = Number(cellVal(state.daySheet, item.row, room.col)) || 0;
        if (q) bits.push(shortName(item.name) + "=" + q);
      });
      if (bits.length) lines.push(room.room + ": " + bits.join(", "));
    });
    if (lines.length) {
      parts.push("--- MINIBAR REPORT (" + zone.toUpperCase() + ") ---");
      parts.push(...lines);
      parts.push("");
    }
  });
  const day = state.daySheet.name.replace(/\s+/g, " ");
  return (day + "\n" + parts.join("\n")).trim();
}

function shortName(name) {
  return String(name)
    .replace(/^MB\s+/i, "")
    .replace(/^Softdrink\s+/i, "")
    .replace(/\s*-\s*\d+\s*(Grm|Ml).*$/i, "")
    .replace(/\s+\d+\s*(Ml|ml).*$/i, "")
    .trim();
}

async function cacheWorkbook() {
  try {
    await persistCurrent(false);
    return state.cachedFile;
  } catch (e) {
    console.error(e);
    return null;
  }
}

function downloadBlob(file) {
  const a = document.createElement("a");
  a.href = URL.createObjectURL(file);
  a.download = file.name;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 1500);
}


function crc32(buf) {
  const table = crc32.t || (crc32.t = (function () {
    const t = new Uint32Array(256);
    for (let i = 0; i < 256; i++) {
      let c = i;
      for (let k = 0; k < 8; k++) c = (c & 1) ? (0xedb88320 ^ (c >>> 1)) : (c >>> 1);
      t[i] = c >>> 0;
    }
    return t;
  })());
  let crc = 0 ^ -1;
  const u8 = buf instanceof Uint8Array ? buf : new Uint8Array(buf);
  for (let i = 0; i < u8.length; i++) crc = (crc >>> 8) ^ table[(crc ^ u8[i]) & 0xff];
  return (crc ^ -1) >>> 0;
}

function zipOneFile(name, data) {
  const fileName = new TextEncoder().encode(name);
  const payload = data instanceof Uint8Array ? data : new Uint8Array(data);
  const crc = crc32(payload);
  const local = new Uint8Array(30 + fileName.length + payload.length);
  const dv = new DataView(local.buffer);
  dv.setUint32(0, 0x04034b50, true);
  dv.setUint16(4, 20, true);
  dv.setUint16(26, fileName.length, true);
  dv.setUint32(14, crc, true);
  dv.setUint32(18, payload.length, true);
  dv.setUint32(22, payload.length, true);
  local.set(fileName, 30);
  local.set(payload, 30 + fileName.length);
  const central = new Uint8Array(46 + fileName.length);
  const cv = new DataView(central.buffer);
  cv.setUint32(0, 0x02014b50, true);
  cv.setUint16(4, 20, true);
  cv.setUint16(6, 20, true);
  cv.setUint16(28, fileName.length, true);
  cv.setUint32(16, crc, true);
  cv.setUint32(20, payload.length, true);
  cv.setUint32(24, payload.length, true);
  central.set(fileName, 46);
  const end = new Uint8Array(22);
  const ev = new DataView(end.buffer);
  ev.setUint32(0, 0x06054b50, true);
  ev.setUint16(8, 1, true);
  ev.setUint16(10, 1, true);
  ev.setUint32(12, central.length, true);
  ev.setUint32(16, local.length, true);
  const out = new Uint8Array(local.length + central.length + end.length);
  out.set(local, 0);
  out.set(central, local.length);
  out.set(end, local.length + central.length);
  return out;
}


function isIOS() {
  return /iPad|iPhone|iPod/.test(navigator.userAgent) || (navigator.platform === "MacIntel" && navigator.maxTouchPoints > 1);
}
function isStandalone() {
  return window.matchMedia("(display-mode: standalone)").matches || window.navigator.standalone === true;
}

state.installEvent = null;
window.addEventListener("beforeinstallprompt", (e) => {
  e.preventDefault();
  state.installEvent = e;
});

function makeShareFiles() {
  const buf = state.cachedBuf;
  if (!buf) return [];
  const base = (state.fileName || "Minibar.xlsx").replace(/\.xlsx$/i, "");
  const xlsx = new File([buf], base + ".xlsx", {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const excel = new File([buf], base + ".xlsx", { type: "application/vnd.ms-excel" });
  const raw = new File([buf], base + ".xlsx", { type: "application/octet-stream" });
  const zip = new File([zipOneFile(base + ".xlsx", buf)], base + ".zip", { type: "application/zip" });
  return [xlsx, excel, raw, zip];
}

async function tryShareFile(file) {
  if (!navigator.share) throw new Error("no share");
  try {
    await navigator.share({ files: [file], title: file.name });
    return true;
  } catch (e) {
    if (e && e.name === "AbortError") return "abort";
    try {
      await navigator.share({ files: [file] });
      return true;
    } catch (e2) {
      if (e2 && e2.name === "AbortError") return "abort";
      throw e2 || e;
    }
  }
}

async function sharePreparedFile() {
  const files = makeShareFiles();
  if (!files.length) throw new Error("File not ready");
  let lastErr = null;
  for (const file of files) {
    try {
      const res = await tryShareFile(file);
      if (res === "abort") return "abort";
      if (res) return true;
    } catch (e) {
      lastErr = e;
    }
  }
  throw lastErr || new Error("share failed");
}

function openShareSheet() {
  if (!state.workbook) {
    toast("Import XL first");
    return;
  }
  if (!state.cachedBuf) cacheWorkbook();
  const hint = $("shareHint");
  const installHint = $("installHint");
  if (!navigator.share) {
    hint.textContent = "This browser has no share sheet. On PC use Chrome or Edge. The file can still be saved, then sent from WhatsApp Desktop.";
  } else if (isIOS()) {
    hint.textContent = "On iPhone / iPad tap Allow. Choose WhatsApp in the iOS share sheet. If WhatsApp is missing, install Minibar XL to the Home Screen first (Share → Add to Home Screen).";
  } else {
    hint.textContent = "Tap Allow to open the phone share sheet, then choose WhatsApp. The Excel file goes attached — you should not need to download it first.";
  }
  installHint.textContent = isStandalone()
    ? "App is already installed on this device."
    : (isIOS()
      ? "iPhone: Safari → Share → Add to Home Screen."
      : "Android / PC: tap Install app, or Chrome menu → Add to Home screen / Install.");
  $("shareSheet").classList.add("open");
}

function closeShareSheet() {
  $("shareSheet").classList.remove("open");
}


async function requestAppPermissions() {
  const out = [];
  try {
    if (navigator.storage && navigator.storage.persist) {
      const ok = await navigator.storage.persist();
      out.push(ok ? "storage: allowed" : "storage: default");
    }
  } catch (e) {}
  try {
    if (navigator.permissions && navigator.permissions.query) {
      const q = await navigator.permissions.query({ name: "persistent-storage" });
      out.push("persistent-storage=" + q.state);
    }
  } catch (e) {}
  return out;
}

async function allowAndShare() {
  if (!state.cachedBuf) {
    $("loader").classList.remove("hidden");
    await cacheWorkbook();
    $("loader").classList.add("hidden");
  }
  try {
    const ok = await sharePreparedFile();
    if (ok === "abort") return;
    closeShareSheet();
    toast("Pick WhatsApp — file attached");
  } catch (e) {
    console.error(e);
    if (!navigator.share) {
      toast("Use Chrome or Edge to share the file");
      return;
    }
    fallbackDownloadAndWhatsApp();
  }
}

async function installApp() {
  if (state.installEvent) {
    state.installEvent.prompt();
    const choice = await state.installEvent.userChoice;
    if (choice && choice.outcome === "accepted") toast("App installed");
    state.installEvent = null;
    return;
  }
  if (isIOS()) {
    toast("Safari → Share → Add to Home Screen");
    return;
  }
  toast("Chrome menu → Add to Home screen / Install app");
}

function platform() {
  const ua = navigator.userAgent || "";
  if (/iPhone|iPad|iPod/i.test(ua)) return "ios";
  if (/Android/i.test(ua)) return "android";
  return "pc";
}

function canShareFiles(file) {
  try {
    return !!(navigator.share && navigator.canShare && navigator.canShare({ files: [file] }));
  } catch (e) {
    return false;
  }
}

function readyShareFiles() {
  const buf = state.cachedBuf;
  if (!buf) return [];
  const base = (state.fileName || "Minibar.xlsx").replace(/\.xlsx$/i, "");
  const xlsx = new File([buf], base + ".xlsx", {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const xls = new File([buf], base + ".xlsx", { type: "application/vnd.ms-excel" });
  const zip = new File([zipOneFile(base + ".xlsx", buf)], base + ".zip", { type: "application/zip" });
  return [xlsx, xls, zip];
}

function openShareModal() {
  const p = platform();
  const hint = {
    android: "Tap Allow. On the next sheet choose WhatsApp — the workbook stays attached.",
    ios: "Tap Allow. On iPhone pick WhatsApp or Save to Files. Use Safari or the Home Screen icon.",
    pc: "Tap Allow. Chrome / Edge can send the file to an app. You can also use Save XL.",
  }[p];
  $("shareHint").textContent = hint;
  $("shareStatus").textContent = state.cachedBuf ? "Workbook is ready to attach." : "Preparing workbook…";
  $("shareModal").classList.add("open");
  if (!state.cachedBuf) cacheWorkbook().then(() => {
    if ($("shareStatus")) $("shareStatus").textContent = "Workbook is ready to attach.";
  });
}

function closeShareModal() {
  $("shareModal").classList.remove("open");
}


function fallbackDownloadAndWhatsApp() {
  const files = makeShareFiles();
  const file = (files && files[0]) || state.cachedFile;
  if (!file) {
    toast("Import XL first");
    return;
  }
  downloadBlob(file);
  const day = state.daySheet ? state.daySheet.name.replace(/\s+/g, " ") : "";
  const msg = "Minibar XL file downloaded: " + file.name + (day ? " (" + day + ")" : "") + ". Attach that Excel from Downloads and send.";
  setTimeout(() => {
    window.location.href = "https://wa.me/?text=" + encodeURIComponent(msg);
  }, 400);
  closeShareSheet();
  toast("XL saved to Downloads — WhatsApp opening. Attach the file.");
}

function shareWhatsApp() {
  if (!state.workbook) {
    toast("Import XL first");
    return;
  }
  if (platform() === "pc") {
    openShareSheet();
    return;
  }
  if (!state.cachedBuf) {
    cacheWorkbook().then(() => fallbackDownloadAndWhatsApp());
    toast("Saving XL…");
    return;
  }
  fallbackDownloadAndWhatsApp();
}


async function requestAppPermissions() {
  const out = [];
  try {
    if (navigator.storage && navigator.storage.persist) {
      const ok = await navigator.storage.persist();
      out.push(ok ? "storage: allowed" : "storage: default");
    }
  } catch (e) {}
  try {
    if (navigator.permissions && navigator.permissions.query) {
      const q = await navigator.permissions.query({ name: "persistent-storage" });
      out.push("persistent-storage=" + q.state);
    }
  } catch (e) {}
  return out;
}

async function allowAndShare() {
  await requestAppPermissions();
  if (state.installEvent && !isStandalone()) {
    try { state.installEvent.prompt(); } catch (e) {}
  }
  if (!navigator.share) {
    toast("Use Chrome or Edge, or tap Save XL");
    return;
  }
  if (!state.cachedBuf) {
    toast("Preparing file…");
    await cacheWorkbook();
  }
  const files = makeShareFiles();
  if (!files.length) {
    toast("File not ready — tap Allow again");
    return;
  }
  let lastErr = null;
  for (const file of files) {
    try {
      await navigator.share({ files: [file], title: file.name });
      closeShareSheet();
      toast("Pick WhatsApp — file attached");
      return;
    } catch (e) {
      if (e && e.name === "AbortError") return;
      lastErr = e;
    }
  }
  console.error(lastErr);
  fallbackDownloadAndWhatsApp();
}

function openTextModal() {
  if (!state.workbook || !state.daySheet) {
    toast("Import XL and select a date first");
    return;
  }
  $("textHint").textContent = "Fill into " + state.daySheet.name.replace(/\s+/g, " ") + " — same file, colours kept.";
  $("textModal").classList.add("open");
}

function closeTextModal() {
  $("textModal").classList.remove("open");
}

function applyPastedText() {
  try {
    const res = applyTextReport($("textArea").value, $("replaceRooms").checked);
    $("importLog").textContent = res.filled + " cells filled on " + res.day +
      (res.missed ? " · " + res.missed + " unmatched" : "") +
      (res.log.length ? "\n" + res.log.join("\n") : "");
    toast(res.filled + " items written to " + res.day);
    schedulePersist();
  } catch (e) {
    $("importLog").textContent = String(e.message || e);
    toast(String(e.message || e));
  }
}



function daysInMonth(year, monthIdx) {
  return new Date(year, monthIdx + 1, 0).getDate();
}

function defaultNameFor(monthIdx, year) {
  return "Minibar Consumption " + MONTH_NAMES[monthIdx] + " " + year + ".xlsx";
}

function rewriteMonthFormulas(ws, newMonth) {
  if (!ws) return;
  ws.eachRow({ includeEmpty: true }, (row) => {
    row.eachCell({ includeEmpty: true }, (cell) => {
      const v = cell.value;
      if (v && typeof v === "object" && v.formula) {
        cell.value = {
          formula: String(v.formula).replace(/'\d+\s+[A-Za-z]+'/g, (tok) => {
            const d = String(tok).match(/(\d+)/);
            return d ? "'" + d[1] + " " + newMonth + "'" : tok;
          }),
          result: v.result,
        };
      }
    });
  });
}

function clearRoomQtys(ws) {
  const rooms = [];
  ws.getRow(2).eachCell({ includeEmpty: false }, (cell, col) => {
    if (typeof cell.value === "number" || /^\d+$/.test(String(cell.value))) rooms.push(col);
  });
  ws.eachRow({ includeEmpty: false }, (row, r) => {
    if (r < 3) return;
    rooms.forEach((col) => setValueKeepStyle(ws, r, col, null));
  });
}

function copySheet(wb, src, newName) {
  const dest = wb.addWorksheet(newName);
  src.eachRow({ includeEmpty: true }, (row, rn) => {
    const r = dest.getRow(rn);
    row.eachCell({ includeEmpty: true }, (cell, cn) => {
      const c = r.getCell(cn);
      c.value = cell.value;
      try { if (cell.style) c.style = JSON.parse(JSON.stringify(cell.style)); } catch (e) {}
    });
    if (row.height) r.height = row.height;
  });
  (src.columns || []).forEach((col, i) => {
    if (col && col.width) dest.getColumn(i + 1).width = col.width;
  });
  return dest;
}

async function createMonthWorkbook(monthIdx, year) {
  if (!state.workbook) throw new Error("Import a workbook first");
  await persistCurrent(true, "before " + MONTH_NAMES[monthIdx] + " " + year);
  const days = daysInMonth(year, monthIdx);
  const monthName = MONTH_NAMES[monthIdx];
  const dailies = dailySheets();
  if (!dailies.length) throw new Error("No daily sheet to copy");
  const template = dailies[0];
  const created = [];
  for (let d = 1; d <= days; d++) {
    const name = d + " " + monthName;
    const ws = copySheet(state.workbook, template, name);
    clearRoomQtys(ws);
    created.push(name);
  }
  dailies.forEach((ws) => {
    if (created.indexOf(ws.name) === -1) {
      try { state.workbook.removeWorksheet(ws.id); } catch (e) {}
    }
  });
  rewriteMonthFormulas(findSheet("Summary"), monthName);
  const summary = findSheet("Summary");
  if (summary) {
    summary.getRow(2).eachCell({ includeEmpty: false }, (cell, col) => {
      if (typeof cell.value === "number" && cell.value > days) return;
    });
  }
  state.fileName = defaultNameFor(monthIdx, year);
  state.downloadName = state.fileName;
  state.dirty = true;
  await persistCurrent(true, "created " + monthName);
  markReady();
  renderDaySelect();
}

function fillMonthYearInputs() {
  const sel = $("newMonthSelect");
  if (!sel) return;
  sel.innerHTML = MONTH_NAMES.map((n, i) => "<option value=\"" + i + "\">" + n + "</option>").join("");
  const now = new Date();
  const meta = fileMonthYear();
  let m = now.getMonth();
  let y = now.getFullYear();
  if (meta) {
    m = meta.month + 1;
    y = meta.year;
    if (m > 11) { m = 0; y += 1; }
  }
  sel.value = String(m);
  $("newYearInput").value = y;
  $("downloadNameInput").value = state.downloadName || state.fileName || defaultNameFor(m, y);
}

async function renderHistory() {
  const box = $("historyList");
  if (!box) return;
  const rows = (await historyAll()).sort((a, b) => b.id - a.id);
  if (!rows.length) {
    box.innerHTML = "<p class=\"file-hint\">No snapshots yet. Save XL or create a month to keep history.</p>";
    return;
  }
  box.innerHTML = "";
  rows.forEach((row) => {
    const when = new Date(row.savedAt || row.id).toLocaleString();
    const el = document.createElement("div");
    el.className = "history-row";
    el.innerHTML = "<span><b>" + escapeHtml(row.fileName || "workbook") + "</b>" + escapeHtml(row.note || "") + " · " + when + "</span>";
    const rest = document.createElement("button");
    rest.className = "btn-cool";
    rest.textContent = "Restore";
    rest.onclick = () => restoreHistory(row.id);
    const dl = document.createElement("button");
    dl.className = "btn-cool btn-success";
    dl.textContent = "Get XL";
    dl.onclick = () => {
      const f = new File([row.buf], row.fileName || "Minibar.xlsx", { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      downloadBlob(f);
    };
    el.appendChild(rest);
    el.appendChild(dl);
    box.appendChild(el);
  });
}

async function restoreHistory(id) {
  const rows = await historyAll();
  const row = rows.find((r) => r.id === id);
  if (!row || !row.buf) return;
  if (state.workbook) await persistCurrent(true, "before restore");
  const file = new File([row.buf], row.fileName || "Minibar.xlsx", { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  await openFile(file);
  toast("Restored " + file.name);
}

function openSettings() {
  fillMonthYearInputs();
  renderHistory();
  $("settingsModal").classList.add("open");
}

function closeSettings() {
  $("settingsModal").classList.remove("open");
}

async function saveSettings() {
  state.downloadName = ($("downloadNameInput").value || "").trim();
  await kvSet("settings", { downloadName: state.downloadName });
  if (state.workbook) await persistCurrent(false);
  toast("Download name saved");
}

async function onCreateMonth() {
  if (!state.workbook) {
    toast("Import XL once first");
    return;
  }
  const monthIdx = Number($("newMonthSelect").value);
  const year = Number($("newYearInput").value);
  if (!confirm("Create " + MONTH_NAMES[monthIdx] + " " + year + "? Current month is stored in history.")) return;
  $("loader").classList.remove("hidden");
  try {
    await createMonthWorkbook(monthIdx, year);
    $("downloadNameInput").value = state.downloadName;
    await renderHistory();
    toast("Created " + state.fileName);
  } catch (e) {
    console.error(e);
    toast(String(e.message || e));
  } finally {
    $("loader").classList.add("hidden");
  }
}


function tableDisplayName(name) {
  let s = shortName(name).replace(/\s+/g, " ").trim();
  s = s.replace(/\b(mb|softdrink)\b/ig, "").replace(/\s+/g, " ").trim();
  return s.replace(/\b([a-z])/g, (m) => m.toUpperCase());
}

function setTableViewport(allowZoom) {
  const meta = document.querySelector('meta[name="viewport"]');
  if (!meta) return;
  meta.setAttribute("content", allowZoom
    ? "width=device-width, initial-scale=1, maximum-scale=4, user-scalable=yes, viewport-fit=cover"
    : "width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no, viewport-fit=cover");
}

function tableKey(row, col) { return row + ":" + col; }

function refreshTableChrome() {
  const edit = !!state.tableEdit;
  if ($("tableEditBtn")) {
    $("tableEditBtn").classList.toggle("active", edit);
    $("tableEditBtn").textContent = edit ? "Done" : "Edit";
  }
  const canUndo = edit && !!(state.tableUndo && state.tableUndo.length);
  const canRedo = edit && !!(state.tableRedo && state.tableRedo.length);
  ["tableUndoBtn", "tableBarUndoBtn"].forEach((id) => { if ($(id)) $(id).disabled = !canUndo; });
  ["tableRedoBtn", "tableBarRedoBtn"].forEach((id) => { if ($(id)) $(id).disabled = !canRedo; });
  if ($("tableWrap")) $("tableWrap").classList.toggle("editing", edit);
  if ($("suggestBar")) $("suggestBar").classList.toggle("hidden", !edit);
}

function applyTableZoom() {
  const z = state.tableZoom || 1;
  const table = document.querySelector("#tableWrap .sheet-table");
  if (table) table.style.zoom = String(z);
}

function paintTableCell(row, col, val, edited) {
  const cell = document.querySelector('#tableWrap td[data-row="' + row + '"][data-col="' + col + '"]');
  if (!cell) return;
  cell.textContent = val ? String(val) : "";
  cell.classList.toggle("has", val > 0);
  cell.classList.toggle("edited", !!edited);
}

function writeTableCell(row, col, next, record) {
  const prev = Number(cellVal(state.daySheet, row, col)) || 0;
  const val = Math.max(0, parseInt(next, 10) || 0);
  if (val === prev) {
    paintTableCell(row, col, val, state.tableEdited && state.tableEdited[tableKey(row, col)]);
    return;
  }
  setValueKeepStyle(state.daySheet, row, col, val || null);
  state.dirty = true;
  refreshItemTotal(state.daySheet, row);
  updateSummaryForDay();
  if (record) {
    state.tableUndo.push({ row, col, prev, val });
    if (state.tableUndo.length > 80) state.tableUndo.shift();
    state.tableRedo = [];
  }
  state.tableEdited[tableKey(row, col)] = true;
  paintTableCell(row, col, val, true);
  updateTableTotals();
  refreshTableChrome();
}

function undoTableEdit() {
  const last = state.tableUndo && state.tableUndo.pop();
  if (!last) return;
  writeTableCell(last.row, last.col, last.prev, false);
  state.tableRedo.push(last);
  refreshTableChrome();
  if (state.tableEdit) focusTableCell(last.row, last.col, true);
}

function redoTableEdit() {
  const last = state.tableRedo && state.tableRedo.pop();
  if (!last) return;
  writeTableCell(last.row, last.col, last.val, false);
  state.tableUndo.push(last);
  refreshTableChrome();
  if (state.tableEdit) focusTableCell(last.row, last.col, true);
}

function renderTableView() {
  const day = state.daySheet ? state.daySheet.name.replace(/\s+/g, " ") : "";
  $("tableTitle").textContent = "Consumption · " + day;
  fillDayOptions($("tableDaySelect"), state.daySheet && state.daySheet.name);
  const rooms = state.rooms.slice().sort((a, b) => Number(a.room) - Number(b.room));
  const wrap = $("tableWrap");
  let html = "<table class=\"sheet-table\"><thead><tr><th class=\"item-col\">Item</th>";
  rooms.forEach((r) => {
    const zone = zoneOf(r.room) === "Water" ? "th-water" : "th-beach";
    html += "<th class=\"" + zone + "\">" + r.room + "</th>";
  });
  html += "<th class=\"th-total\">Total</th></tr></thead><tbody>";
  state.items.forEach((item) => {
    html += "<tr><td class=\"item-col cat-" + item.cat + "\"><span class=\"item-label\"><i class=\"item-swatch\"></i>" +
      escapeHtml(tableDisplayName(item.name)) + "</span></td>";
    rooms.forEach((r) => {
      const q = Number(cellVal(state.daySheet, item.row, r.col)) || 0;
      const edited = state.tableEdited && state.tableEdited[tableKey(item.row, r.col)];
      html += "<td data-row=\"" + item.row + "\" data-col=\"" + r.col + "\" class=\"" +
        (q ? "has" : "") + (edited ? " edited" : "") + "\">" + (q ? q : "") + "</td>";
    });
    html += "<td class=\"row-total\" data-total-row=\"" + item.row + "\"></td></tr>";
  });
  html += "<tr class=\"col-total-row\"><td class=\"item-col\">Total</td>";
  rooms.forEach((r) => {
    html += "<td class=\"col-total\" data-total-col=\"" + r.col + "\"></td>";
  });
  html += "<td class=\"grand-total\" id=\"tableGrandTotal\"></td></tr></tbody></table>";
  wrap.innerHTML = html;
  applyTableZoom();
  updateTableTotals();
  wrap.classList.toggle("pin-names", !!state.tablePinNames);
  if ($("tablePinBtn")) {
    $("tablePinBtn").classList.toggle("active", !!state.tablePinNames);
    $("tablePinBtn").textContent = state.tablePinNames ? "Scroll names" : "Pin names";
  }
  refreshTableChrome();
}

function updateTableTotals() {
  if (!state.daySheet || !state.items.length) return;
  let grand = 0;
  state.items.forEach((item) => {
    let rowSum = 0;
    state.rooms.forEach((r) => {
      rowSum += Number(cellVal(state.daySheet, item.row, r.col)) || 0;
    });
    const el = document.querySelector('#tableWrap td[data-total-row="' + item.row + '"]');
    if (el) el.textContent = rowSum || "";
    grand += rowSum;
  });
  state.rooms.forEach((r) => {
    let colSum = 0;
    state.items.forEach((item) => {
      colSum += Number(cellVal(state.daySheet, item.row, r.col)) || 0;
    });
    const el = document.querySelector('#tableWrap td[data-total-col="' + r.col + '"]');
    if (el) el.textContent = colSum || "";
  });
  if ($("tableGrandTotal")) $("tableGrandTotal").textContent = grand || "";
  const title = $("tableTitle");
  if (title && state.daySheet) {
    title.textContent = "Consumption · " + state.daySheet.name.replace(/\s+/g, " ") + " · Total " + grand;
  }
}

function clearActiveCell() {
  document.querySelectorAll("#tableWrap td.active-cell").forEach((el) => el.classList.remove("active-cell"));
}

function focusTableCell(row, col, skipInput) {
  const td = document.querySelector('#tableWrap td[data-row="' + row + '"][data-col="' + col + '"]');
  if (!td) return;
  clearActiveCell();
  state.tableFocus = { row, col };
  td.classList.add("active-cell");
  td.scrollIntoView({ block: "nearest", inline: "nearest" });
  if (skipInput) return;
  if (td.querySelector("input")) {
    td.querySelector("input").focus();
    return;
  }
  const current = Number(cellVal(state.daySheet, row, col)) || 0;
  const input = document.createElement("input");
  input.className = "cell-input";
  input.type = "number";
  input.min = "0";
  input.inputMode = "numeric";
  input.value = current ? String(current) : "";
  td.textContent = "";
  td.appendChild(input);
  input.focus();
  input.select();
  const commit = () => {
    const typed = input.value;
    writeTableCell(row, col, typed, true);
    if (String(typed).trim() !== "" && Number(typed) > 0) moveTableNext(true);
  };
  input.addEventListener("blur", commit, { once: true });
  input.addEventListener("keydown", (ev) => {
    if (ev.key === "Enter") input.blur();
    if (ev.key === "Escape") {
      input.value = String(current);
      input.blur();
    }
  });
}

function moveTableNext(fromFill) {
  if (!state.tableFocus || !state.items.length) return;
  const rows = state.items.map((it) => it.row);
  const i = rows.indexOf(state.tableFocus.row);
  if (i < 0 || i >= rows.length - 1) {
    toast(fromFill ? "Last item in this villa" : "Last item");
    return;
  }
  focusTableCell(rows[i + 1], state.tableFocus.col);
}

function applySuggestQty(qty) {
  if (!state.tableEdit) return;
  if (!state.tableFocus) {
    toast("Tap a cell first");
    return;
  }
  writeTableCell(state.tableFocus.row, state.tableFocus.col, qty, true);
  moveTableNext(true);
}

function onTableCellClick(e) {
  if (!state.tableEdit) return;
  if (e.target.tagName === "INPUT") return;
  const td = e.target.closest("td[data-row]");
  if (!td) return;
  focusTableCell(Number(td.getAttribute("data-row")), Number(td.getAttribute("data-col")));
}

function toggleTableEdit() {
  if (state.tableEdit) {
    state.tableEdit = false;
    persistCurrent(true, "table edit");
    toast("Edits saved");
    refreshTableChrome();
    return;
  }
  state.tableEdit = true;
  refreshTableChrome();
  toast("Edit mode — tap a number");
}

function openTableView() {
  if (!state.workbook || !state.daySheet) {
    toast("Load a workbook first");
    return;
  }
  state.tableEdit = false;
  state.tableUndo = [];
  state.tableRedo = [];
  state.tableEdited = {};
  state.tableZoom = state.tableZoom || 1;
  setTableViewport(true);
  $("tableModal").classList.add("open");
  renderTableView();
}

function closeTableView() {
  persistCurrent(true, "table close");
  state.tableEdit = false;
  state.tableEdited = {};
  setTableViewport(false);
  $("tableModal").classList.remove("open");
}

function setZone(zone) {
  state.zone = zone;
  $("btnBeach").classList.toggle("active", zone === "Beach");
  $("btnWater").classList.toggle("active", zone === "Water");
  renderGrid();
}


document.querySelectorAll(".suggest-num").forEach((btn) => {
  btn.onclick = () => applySuggestQty(btn.getAttribute("data-qty"));
});
$("tableNextItemBtn") && ($("tableNextItemBtn").onclick = () => moveTableNext(false));
$("tableBtn") && ($("tableBtn").onclick = openTableView);
$("tableEditBtn") && ($("tableEditBtn").onclick = toggleTableEdit);
$("tableUndoBtn") && ($("tableUndoBtn").onclick = undoTableEdit);
$("tableRedoBtn") && ($("tableRedoBtn").onclick = redoTableEdit);
$("tableBarUndoBtn") && ($("tableBarUndoBtn").onclick = undoTableEdit);
$("tableBarRedoBtn") && ($("tableBarRedoBtn").onclick = redoTableEdit);
$("tableZoomInBtn") && ($("tableZoomInBtn").onclick = () => {
  state.tableZoom = Math.min(1.8, (state.tableZoom || 1) + 0.1);
  applyTableZoom();
});
$("tableZoomOutBtn") && ($("tableZoomOutBtn").onclick = () => {
  state.tableZoom = Math.max(0.7, (state.tableZoom || 1) - 0.1);
  applyTableZoom();
});
$("tablePinBtn") && ($("tablePinBtn").onclick = () => {
  state.tablePinNames = !state.tablePinNames;
  if ($("tableWrap")) $("tableWrap").classList.toggle("pin-names", !!state.tablePinNames);
  $("tablePinBtn").classList.toggle("active", !!state.tablePinNames);
  $("tablePinBtn").textContent = state.tablePinNames ? "Scroll names" : "Pin names";
});
$("tableWrap") && $("tableWrap").addEventListener("click", onTableCellClick);
$("closeTableBtn") && ($("closeTableBtn").onclick = closeTableView);
$("tableModal") && $("tableModal").addEventListener("click", (e) => { if (e.target.id === "tableModal") closeTableView(); });
async function forceLatestVersion() {
  try {
    if ("serviceWorker" in navigator) {
      const regs = await navigator.serviceWorker.getRegistrations();
      await Promise.all(regs.map((r) => r.unregister()));
    }
    if (window.caches) {
      const keys = await caches.keys();
      await Promise.all(keys.map((k) => caches.delete(k)));
    }
  } catch (e) {}
  const url = new URL(location.href);
  url.searchParams.set("v", Date.now().toString());
  location.replace(url.toString());
}

function parseVer(s) {
  return String(s || "0").split(".").map((n) => parseInt(n, 10) || 0);
}

function verOlder(a, b) {
  const x = parseVer(a), y = parseVer(b);
  for (let i = 0; i < 3; i++) {
    if ((x[i] || 0) < (y[i] || 0)) return true;
    if ((x[i] || 0) > (y[i] || 0)) return false;
  }
  return false;
}

async function checkAppVersion() {
  if ($("appVersionLabel")) $("appVersionLabel").textContent = "This device: v" + APP_VERSION;
  try {
    const res = await fetch("version.json?t=" + Date.now(), { cache: "no-store" });
    if (!res.ok) return;
    const info = await res.json();
    if ($("appVersionLabel")) {
      $("appVersionLabel").textContent = "This device: v" + APP_VERSION + " · Live: v" + info.version;
    }
    const need = info.minVersion || info.version;
    if (info.version !== APP_VERSION || verOlder(APP_VERSION, need)) {
      toast("Updating to v" + info.version);
      setTimeout(forceLatestVersion, 400);
    }
  } catch (e) {}
}

$("forceLatestBtn") && ($("forceLatestBtn").onclick = forceLatestVersion);
$("settingsBtn") && ($("settingsBtn").onclick = openSettings);
$("closeSettingsBtn") && ($("closeSettingsBtn").onclick = closeSettings);
$("settingsModal") && $("settingsModal").addEventListener("click", (e) => { if (e.target.id === "settingsModal") closeSettings(); });
$("saveSettingsBtn") && ($("saveSettingsBtn").onclick = saveSettings);
$("createMonthBtn") && ($("createMonthBtn").onclick = onCreateMonth);
$("btnBeach").onclick = () => setZone("Beach");
$("btnWater").onclick = () => setZone("Water");
$("daySelect").onchange = onDayChange;
$("tableDaySelect") && ($("tableDaySelect").onchange = onTableDayChange);
$("roomFilter").oninput = renderGrid;
$("fileInput").onchange = (e) => { const f = e.target.files && e.target.files[0]; if (f) openFile(f); };
$("saveBtn").onclick = saveSameFile;
$("importTextBtn") && ($("importTextBtn").onclick = openTextModal);
$("closeTextBtn") && ($("closeTextBtn").onclick = closeTextModal);
$("applyTextBtn") && ($("applyTextBtn").onclick = applyPastedText);
$("textModal") && $("textModal").addEventListener("click", (e) => { if (e.target.id === "textModal") closeTextModal(); });
$("waBtn") && ($("waBtn").onclick = shareWhatsApp);
$("allowShareBtn") && ($("allowShareBtn").onclick = allowAndShare);
$("installAppBtn") && ($("installAppBtn").onclick = installApp);
$("closeShareBtn") && ($("closeShareBtn").onclick = closeShareSheet);
$("shareSheet") && $("shareSheet").addEventListener("click", (e) => { if (e.target.id === "shareSheet") closeShareSheet(); });
$("waEntryBtn") && ($("waEntryBtn").onclick = shareWhatsApp);
$("finishBtn").onclick = saveSameFile;
$("cleanBtn").onclick = cleanZone;
$("closeModalBtn").onclick = closeRoom;
$("saveEntryBtn").onclick = () => { closeRoom(); persistCurrent(false); toast("Saved on this device"); };
$("usedOnlyBtn") && ($("usedOnlyBtn").onclick = () => {
  state.usedOnly = !state.usedOnly;
  $("usedOnlyBtn").classList.toggle("active", state.usedOnly);
  renderItems();
});
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


function fitAppHeight() {
  const h = (window.visualViewport && window.visualViewport.height) || window.innerHeight;
  document.documentElement.style.setProperty("--app-h", Math.round(h) + "px");
}
fitAppHeight();
window.addEventListener("resize", fitAppHeight);
if (window.visualViewport) {
  window.visualViewport.addEventListener("resize", fitAppHeight);
}

if (window.parent && window.parent !== window) {
  document.body.classList.add("embedded");
}

loadPersistedWorkbook().catch((e) => console.error(e));
checkAppVersion();

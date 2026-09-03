/* Minibar workbook editor
   Loads the EXISTING .xlsx and writes cell VALUES only.
   Styles, fills, fonts, merged cells, column widths and formulas stay in the same file.
*/
const $ = (id) => document.getElementById(id);

const state = {
  workbook: null,
  fileName: "Minibar Consumption August 2026.xlsx",
  dirty: false,
  mode: "daily",
  daySheet: null,
  rooms: [],
  items: [],
  selectedRoom: null,
  pickupHeaders: [],
};

function setDirty(flag) {
  state.dirty = flag;
  $("dirtyLabel").textContent = flag ? "Unsaved changes" : "No changes";
  document.querySelector(".status-bar").classList.toggle("dirty", flag);
  $("saveBtn").disabled = !state.workbook;
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
  const cell = ws.getRow(r).getCell(c);
  const next = value === "" || value === null || value === undefined ? null : value;
  cell.value = next;
}

function classifyItem(name) {
  const n = String(name || "").toLowerCase();
  if (/diet|zero|tonic|soda|ginger/.test(n)) return "flag";
  if (/coca|sprite|softdrink/.test(n)) return "soft";
  if (/peanut|crisp|cashew|pistachio|beet/.test(n)) return "food";
  if (/wine|mancura|barone|bottega|pinot|merlot|sicil/.test(n)) return "wine";
  if (/beer|grani|tagus|lion|saigon|xibeca|pilsner|royal|333/.test(n)) return "beer";
  return "ok";
}

function loadDailyMeta(ws) {
  const rooms = [];
  const header = ws.getRow(2);
  header.eachCell({ includeEmpty: false }, (cell, col) => {
    const v = cell.value;
    if (typeof v === "number" || /^\d+$/.test(String(v))) {
      rooms.push({ col, room: String(v) });
    }
  });
  const items = [];
  ws.eachRow({ includeEmpty: false }, (row, r) => {
    if (r < 3) return;
    const article = cellVal(ws, r, 2);
    if (!article || String(article).toLowerCase() === "total") return;
    items.push({
      row: r,
      sno: cellVal(ws, r, 1),
      name: String(article),
      cat: classifyItem(article),
    });
  });
  state.rooms = rooms;
  state.items = items;
  state.daySheet = ws;
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
  $("roomSearch").disabled = false;
  renderRoomChips($("roomSearch").value);
  if (state.selectedRoom && state.rooms.some((r) => r.room === state.selectedRoom.room)) {
    const keep = state.rooms.find((r) => r.room === state.selectedRoom.room);
    selectRoom(keep);
  } else {
    state.selectedRoom = null;
    $("roomBanner").textContent = `${state.rooms.length} villas on ${ws.name.replace(/\s+/g, " ")} — pick a room`;
    $("itemList").innerHTML = "";
  }
  $("metaLabel").textContent = `${state.items.length} items · ${state.rooms.length} rooms · ${state.workbook.worksheets.length} sheets`;
}

function renderRoomChips(filter) {
  const q = String(filter || "").trim();
  const list = q ? state.rooms.filter((r) => r.room.includes(q)) : state.rooms;
  const box = $("roomChips");
  box.innerHTML = "";
  list.slice(0, 40).forEach((r) => {
    const b = document.createElement("button");
    b.className = "chip" + (state.selectedRoom && state.selectedRoom.room === r.room ? " active" : "");
    b.textContent = r.room;
    b.onclick = () => selectRoom(r);
    box.appendChild(b);
  });
  if (list.length > 40) {
    const more = document.createElement("span");
    more.className = "chip";
    more.textContent = `+${list.length - 40}`;
    box.appendChild(more);
  }
}

function selectRoom(room) {
  state.selectedRoom = room;
  $("roomBanner").textContent = `Villa ${room.room}  ·  ${state.daySheet.name.replace(/\s+/g, " ")}`;
  renderRoomChips($("roomSearch").value);
  renderItems();
}

function renderItems() {
  const ws = state.daySheet;
  const room = state.selectedRoom;
  const list = $("itemList");
  list.innerHTML = "";
  if (!ws || !room) return;
  state.items.forEach((item) => {
    const raw = cellVal(ws, item.row, room.col);
    const qty = raw == null || raw === "" ? 0 : Number(raw) || 0;
    const el = document.createElement("article");
    el.className = "item";
    el.innerHTML = `
      <div class="swatch cat-${item.cat}"></div>
      <div>
        <h3>${escapeHtml(item.name)}</h3>
        <p>#${item.sno ?? "—"}</p>
      </div>
      <div class="stepper">
        <button type="button" data-act="-">−</button>
        <input type="number" min="0" step="1" value="${qty || ""}" />
        <button type="button" data-act="+">+</button>
      </div>`;
    const input = el.querySelector("input");
    const write = (n) => {
      const v = Math.max(0, Number(n) || 0);
      input.value = v ? String(v) : "";
      setValueKeepStyle(ws, item.row, room.col, v ? v : null);
      setDirty(true);
    };
    el.querySelector('[data-act="-"]').onclick = () => write((Number(input.value) || 0) - 1);
    el.querySelector('[data-act="+"]').onclick = () => write((Number(input.value) || 0) + 1);
    input.onchange = () => write(input.value);
    list.appendChild(el);
  });
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, (ch) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"
  }[ch]));
}

function renderPickup() {
  const ws = findSheet("Pick up");
  const table = $("pickupTable");
  table.innerHTML = "";
  if (!ws) { table.innerHTML = "<tr><td>No Pick up sheet</td></tr>"; return; }
  const maxCol = ws.actualColumnCount || 24;
  const maxRow = ws.actualRowCount || 32;
  const head = document.createElement("thead");
  const body = document.createElement("tbody");
  for (let r = 1; r <= maxRow; r++) {
    const tr = document.createElement("tr");
    for (let c = 1; c <= maxCol; c++) {
      const tag = r <= 3 ? "th" : "td";
      const td = document.createElement(tag);
      const val = cellVal(ws, r, c);
      const cell = ws.getRow(r).getCell(c);
      const isFormula = cell.value && typeof cell.value === "object" && cell.value.formula;
      if (r > 3 && c >= 4 && !isFormula && typeof val !== "string") {
        const input = document.createElement("input");
        input.type = "number";
        input.value = val == null ? "" : val;
        input.onchange = () => {
          const n = input.value === "" ? null : Number(input.value);
          setValueKeepStyle(ws, r, c, Number.isFinite(n) ? n : null);
          setDirty(true);
        };
        td.appendChild(input);
      } else {
        td.textContent = val == null ? "" : val;
        if (isFormula) td.title = "Formula kept as-is";
      }
      tr.appendChild(td);
    }
    (r <= 3 ? head : body).appendChild(tr);
  }
  table.appendChild(head);
  table.appendChild(body);
}

function renderRooms() {
  const ws = findSheet("Rooms");
  const table = $("roomsTable");
  table.innerHTML = "";
  if (!ws) return;
  const qRoom = $("roomsFilter").value.trim();
  const qItem = $("roomsItemFilter").value.trim().toLowerCase();
  const header = ws.getRow(2);
  const roomCols = [];
  header.eachCell({ includeEmpty: false }, (cell, col) => {
    if (col === 1) return;
    const v = cell.value;
    if (v == null) return;
    if (qRoom && !String(v).includes(qRoom)) return;
    roomCols.push({ col, room: String(v) });
  });
  const shownCols = roomCols.slice(0, 25);
  const thead = document.createElement("thead");
  const hr = document.createElement("tr");
  hr.innerHTML = "<th>Article</th>" + shownCols.map((c) => `<th>${escapeHtml(c.room)}</th>`).join("");
  thead.appendChild(hr);
  const body = document.createElement("tbody");
  ws.eachRow({ includeEmpty: false }, (row, r) => {
    if (r < 3) return;
    const name = cellVal(ws, r, 1);
    if (!name) return;
    if (qItem && !String(name).toLowerCase().includes(qItem)) return;
    const tr = document.createElement("tr");
    const first = document.createElement("td");
    first.textContent = name;
    tr.appendChild(first);
    shownCols.forEach((rc) => {
      const td = document.createElement("td");
      const cell = ws.getRow(r).getCell(rc.col);
      const isFormula = cell.value && typeof cell.value === "object" && cell.value.formula;
      const val = cellVal(ws, r, rc.col);
      if (isFormula) {
        td.textContent = val == null ? "" : val;
        td.title = "Formula";
      } else {
        const input = document.createElement("input");
        input.type = "number";
        input.value = val == null ? "" : val;
        input.onchange = () => {
          const n = input.value === "" ? null : Number(input.value);
          setValueKeepStyle(ws, r, rc.col, Number.isFinite(n) ? n : null);
          setDirty(true);
        };
        td.appendChild(input);
      }
      tr.appendChild(td);
    });
    body.appendChild(tr);
  });
  table.appendChild(thead);
  table.appendChild(body);
  if (roomCols.length > 25) {
    const note = document.createElement("caption");
    note.textContent = `Showing 25 of ${roomCols.length} matching rooms — type a room number to narrow`;
    table.prepend(note);
  }
}

function renderSummary() {
  const ws = findSheet("Summary");
  const table = $("summaryTable");
  table.innerHTML = "";
  if (!ws) return;
  const maxCol = Math.min(ws.actualColumnCount || 48, 44);
  const maxRow = ws.actualRowCount || 31;
  for (let r = 1; r <= maxRow; r++) {
    const tr = document.createElement("tr");
    for (let c = 2; c <= maxCol; c++) {
      const td = document.createElement(r <= 2 ? "th" : "td");
      const val = cellVal(ws, r, c);
      td.textContent = val == null ? "" : val;
      tr.appendChild(td);
    }
    table.appendChild(tr);
  }
}

function showMode(mode) {
  state.mode = mode;
  document.querySelectorAll(".tab").forEach((t) => t.classList.toggle("active", t.dataset.mode === mode));
  $("dailyView").classList.toggle("hidden", mode !== "daily");
  $("pickupView").classList.toggle("hidden", mode !== "pickup");
  $("roomsView").classList.toggle("hidden", mode !== "rooms");
  $("summaryView").classList.toggle("hidden", mode !== "summary");
  $("dailyToolbar").classList.toggle("hidden", mode !== "daily");
  if (!state.workbook) return;
  if (mode === "pickup") renderPickup();
  if (mode === "rooms") renderRooms();
  if (mode === "summary") renderSummary();
}

async function openFile(file) {
  const buf = await file.arrayBuffer();
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buf);
  state.workbook = wb;
  state.fileName = file.name || state.fileName;
  $("fileStatus").textContent = `Editing ${state.fileName} in place (styles kept)`;
  $("emptyState").classList.add("hidden");
  $("dailyView").classList.remove("hidden");
  renderDaySelect();
  setDirty(false);
  showMode("daily");
}

async function saveSameFile() {
  if (!state.workbook) return;
  const buf = await state.workbook.xlsx.writeBuffer();
  const blob = new Blob([buf], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = state.fileName;
  a.click();
  URL.revokeObjectURL(a.href);
  setDirty(false);
}

$("fileInput").addEventListener("change", (e) => {
  const file = e.target.files && e.target.files[0];
  if (file) openFile(file);
});
$("saveBtn").addEventListener("click", saveSameFile);
$("daySelect").addEventListener("change", onDayChange);
$("roomSearch").addEventListener("input", (e) => renderRoomChips(e.target.value));
$("roomsFilter").addEventListener("input", renderRooms);
$("roomsItemFilter").addEventListener("input", renderRooms);
document.querySelectorAll(".tab").forEach((t) => t.addEventListener("click", () => showMode(t.dataset.mode)));

window.addEventListener("beforeunload", (e) => {
  if (state.dirty) {
    e.preventDefault();
    e.returnValue = "";
  }
});

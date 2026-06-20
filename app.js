// ====================================================
// PRINTING PRODUCTION SUITE - v2.0
// Multi-device responsive, multi-format import
// ====================================================

// ===== Shared Utilities =====
const fmtInt = (n) => Number(n).toLocaleString("en-US");
const fmtPrice = (n) => "$" + Number(n).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const getVal = (id, fallback) => {
  const el = document.getElementById(id);
  if (!el) return fallback;
  const v = parseFloat(el.value);
  return isNaN(v) || v <= 0 ? fallback : v;
};
const sanitize = (str) => String(str || "").replace(/[<>"'&]/g, "").trim();
const isPositiveInt = (v) => Number.isInteger(v) && v > 0;

// ====================================================
// FILE IMPORT ENGINE (CSV, TXT, XLSX/XLS)
// ====================================================
const FileImporter = {
  /**
   * Parse a file into rows of data
   * Supports: .csv, .txt (tab/comma separated), .xlsx, .xls
   * Returns: { rows: [[col1, col2, ...], ...], errors: [], warnings: [] }
   */
  async parseFile(file) {
    const ext = file.name.split(".").pop().toLowerCase();
    const result = { rows: [], errors: [], warnings: [] };

    if (!["csv", "txt", "xlsx", "xls"].includes(ext)) {
      result.errors.push(`Unsupported file type: .${ext}. Use .csv, .txt, .xlsx, or .xls`);
      return result;
    }

    // File size check (max 5MB)
    if (file.size > 5 * 1024 * 1024) {
      result.errors.push("File too large (max 5MB)");
      return result;
    }

    try {
      if (ext === "xlsx" || ext === "xls") {
        return await this.parseExcel(file, result);
      } else {
        return await this.parseText(file, ext, result);
      }
    } catch (e) {
      result.errors.push("Failed to read file: " + e.message);
      return result;
    }
  },

  async parseExcel(file, result) {
    return new Promise((resolve) => {
      const reader = new FileReader();
      reader.onload = function (e) {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: "array" });
          const sheet = workbook.Sheets[workbook.SheetNames[0]];
          const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });

          // Filter out completely empty rows
          result.rows = rows.filter((r) => r.some((cell) => String(cell).trim() !== ""));
          resolve(result);
        } catch (err) {
          result.errors.push("Excel parse error: " + err.message);
          resolve(result);
        }
      };
      reader.onerror = () => {
        result.errors.push("Could not read file");
        resolve(result);
      };
      reader.readAsArrayBuffer(file);
    });
  },

  async parseText(file, ext, result) {
    return new Promise((resolve) => {
      const reader = new FileReader();
      reader.onload = function (e) {
        const text = e.target.result;
        const lines = text.trim().split(/\r?\n/);

        lines.forEach((line) => {
          if (!line.trim()) return;
          // Auto-detect delimiter: tab, semicolon, or comma
          let delimiter = ",";
          if (line.includes("\t")) delimiter = "\t";
          else if (line.includes(";")) delimiter = ";";

          const cols = line.split(delimiter).map((c) => c.trim());
          if (cols.some((c) => c !== "")) {
            result.rows.push(cols);
          }
        });

        resolve(result);
      };
      reader.onerror = () => {
        result.errors.push("Could not read file");
        resolve(result);
      };
      reader.readAsText(file);
    });
  },

  /**
   * Detect and skip header row
   * Heuristic: if first row has no numeric columns (except possibly an index), it's a header
   */
  detectHeader(rows) {
    if (rows.length === 0) return { header: null, dataRows: rows };
    const first = rows[0];
    const numericCount = first.filter((c) => !isNaN(Number(c)) && String(c).trim() !== "").length;
    // If less than half the columns are numeric, treat as header
    if (numericCount < first.length / 2) {
      return { header: first, dataRows: rows.slice(1) };
    }
    return { header: null, dataRows: rows };
  },

  /**
   * Validate rows for Cost Calculator import
   * Expected columns: Title, Pages, Qty, Paper(A4/A3), Material(WF/MT)
   */
  validateCost(rows) {
    const { header, dataRows } = this.detectHeader(rows);
    const valid = [], errors = [], warnings = [];
    const validPapers = ["A4", "A3"];
    const validMaterials = ["WF", "MT"];

    if (dataRows.length === 0) {
      errors.push("No data rows found in file");
      return { valid, errors, warnings, header };
    }

    dataRows.forEach((row, i) => {
      const lineNum = i + (header ? 2 : 1);
      if (row.length < 3) {
        errors.push(`Row ${lineNum}: Need at least 3 columns (Title, Pages, Qty). Found ${row.length}`);
        return;
      }

      const title = sanitize(row[0]) || `Book ${i + 1}`;
      const pages = parseInt(row[1]);
      const qty = parseInt(row[2]);
      const paper = row[3] ? String(row[3]).toUpperCase().trim() : "A4";
      const material = row[4] ? String(row[4]).toUpperCase().trim() : "WF";

      if (!isPositiveInt(pages)) { errors.push(`Row ${lineNum}: Pages "${row[1]}" must be a positive number`); return; }
      if (!isPositiveInt(qty)) { errors.push(`Row ${lineNum}: Qty "${row[2]}" must be a positive number`); return; }
      if (!validPapers.includes(paper)) { errors.push(`Row ${lineNum}: Paper "${row[3]}" must be A4 or A3`); return; }
      if (!validMaterials.includes(material)) { errors.push(`Row ${lineNum}: Material "${row[4]}" must be WF or MT`); return; }
      if (pages > 5000) { warnings.push(`Row ${lineNum}: High page count (${pages})`); }
      if (qty > 1000000) { warnings.push(`Row ${lineNum}: High quantity (${qty})`); }

      valid.push({ title, pages, qty, paper, material });
    });

    return { valid, errors, warnings, header };
  },

  /**
   * Validate rows for Production Calculator import
   * Expected columns: Title, Pages, Qty, Material(WF/MT/GL), Binding(Perfect/Staple)
   */
  validateProduction(rows) {
    const { header, dataRows } = this.detectHeader(rows);
    const valid = [], errors = [], warnings = [];
    const validMaterials = ["WF", "MT", "GL"];
    const validBindings = ["PERFECT", "STAPLE"];

    if (dataRows.length === 0) {
      errors.push("No data rows found in file");
      return { valid, errors, warnings, header };
    }

    dataRows.forEach((row, i) => {
      const lineNum = i + (header ? 2 : 1);
      if (row.length < 3) {
        errors.push(`Row ${lineNum}: Need at least 3 columns (Title, Pages, Qty). Found ${row.length}`);
        return;
      }

      const title = sanitize(row[0]) || `Book ${i + 1}`;
      const pages = parseInt(row[1]);
      const qty = parseInt(row[2]);
      const material = row[3] ? String(row[3]).toUpperCase().trim() : "WF";
      const binding = row[4] ? String(row[4]).toUpperCase().trim() : "PERFECT";

      if (!isPositiveInt(pages)) { errors.push(`Row ${lineNum}: Pages "${row[1]}" must be a positive number`); return; }
      if (pages < 4) { errors.push(`Row ${lineNum}: Pages must be at least 4`); return; }
      if (!isPositiveInt(qty)) { errors.push(`Row ${lineNum}: Qty "${row[2]}" must be a positive number`); return; }
      if (!validMaterials.includes(material)) { errors.push(`Row ${lineNum}: Material "${row[3]}" must be WF, MT, or GL`); return; }
      if (!validBindings.includes(binding)) { errors.push(`Row ${lineNum}: Binding "${row[4]}" must be Perfect or Staple`); return; }

      // Binding page validation - auto-round (add blank pages)
      if (binding === "PERFECT" && pages % 8 !== 0) {
        const fixed = Math.ceil(pages / 8) * 8;
        warnings.push(`Row ${lineNum}: "${title}" - ${pages}pg → ${fixed}pg (+${fixed - pages} blank for Perfect Binding)`);
        valid.push({ title, pages: fixed, qty, material, binding: binding.charAt(0) + binding.slice(1).toLowerCase() });
        return;
      }
      if (binding === "STAPLE" && pages % 4 !== 0) {
        const fixed = Math.ceil(pages / 4) * 4;
        warnings.push(`Row ${lineNum}: "${title}" - ${pages}pg → ${fixed}pg (+${fixed - pages} blank for Staple Binding)`);
        valid.push({ title, pages: fixed, qty, material, binding: binding.charAt(0) + binding.slice(1).toLowerCase() });
        return;
      }

      valid.push({ title, pages, qty, material, binding: binding.charAt(0) + binding.slice(1).toLowerCase() });
    });

    return { valid, errors, warnings, header };
  }
};

// ====================================================
// IMPORT UI & PREVIEW
// ====================================================
let pendingImport = { type: null, data: [] };

function showImportHelp(type) {
  let html = "";
  if (type === "cost") {
    html = `
      <h6>Cost Calculator Import Format</h6>
      <p class="small">Accepts <b>.csv</b>, <b>.txt</b>, <b>.xlsx</b>, <b>.xls</b></p>
      <table class="table table-sm table-bordered small">
        <thead class="table-light"><tr><th>Column</th><th>Description</th><th>Required</th></tr></thead>
        <tbody>
          <tr><td>A</td><td>Title (book name)</td><td>Yes</td></tr>
          <tr><td>B</td><td>Pages per book</td><td>Yes</td></tr>
          <tr><td>C</td><td>Quantity</td><td>Yes</td></tr>
          <tr><td>D</td><td>Paper (A4 or A3)</td><td>No, default: A4</td></tr>
          <tr><td>E</td><td>Material (WF or MT)</td><td>No, default: WF</td></tr>
        </tbody>
      </table>
      <p class="small text-muted">Example: <code>Math Grade 7, 120, 500, A4, WF</code></p>`;
  } else {
    html = `
      <h6>Production Calculator Import Format</h6>
      <p class="small">Accepts <b>.csv</b>, <b>.txt</b>, <b>.xlsx</b>, <b>.xls</b></p>
      <table class="table table-sm table-bordered small">
        <thead class="table-light"><tr><th>Column</th><th>Description</th><th>Required</th></tr></thead>
        <tbody>
          <tr><td>A</td><td>Book Title</td><td>Yes</td></tr>
          <tr><td>B</td><td>Pages</td><td>Yes</td></tr>
          <tr><td>C</td><td>Quantity</td><td>Yes</td></tr>
          <tr><td>D</td><td>Material (WF/MT/GL)</td><td>No, default: WF</td></tr>
          <tr><td>E</td><td>Binding (Perfect/Staple)</td><td>No, default: Perfect</td></tr>
        </tbody>
      </table>
      <p class="small text-muted">Perfect Binding: pages must be divisible by 8<br>
      Staple: pages auto-adjusted to multiple of 4</p>`;
  }

  Swal.fire({ html, icon: "info", confirmButtonText: "Got it", customClass: { popup: "text-start" } });
}

function showImportPreview(validData, errors, warnings, type, headers) {
  const errEl = document.getElementById("importErrors");
  const warnEl = document.getElementById("importWarnings");
  const headEl = document.getElementById("importPreviewHead");
  const bodyEl = document.getElementById("importPreviewBody");
  const summaryEl = document.getElementById("importSummary");

  // Show errors
  if (errors.length > 0) {
    errEl.classList.remove("d-none");
    errEl.innerHTML = `<b>${errors.length} Error(s):</b><br>` + errors.slice(0, 10).map((e) => `• ${e}`).join("<br>") +
      (errors.length > 10 ? `<br><i>... and ${errors.length - 10} more</i>` : "");
  } else {
    errEl.classList.add("d-none");
  }

  // Show warnings
  if (warnings.length > 0) {
    warnEl.classList.remove("d-none");
    warnEl.innerHTML = `<b>${warnings.length} Warning(s):</b><br>` + warnings.slice(0, 5).map((w) => `• ${w}`).join("<br>");
  } else {
    warnEl.classList.add("d-none");
  }

  // Build preview table
  let headHtml = "<tr>";
  let bodyHtml = "";

  if (type === "cost") {
    headHtml += "<th>#</th><th>Title</th><th>Pages</th><th>Qty</th><th>Paper</th><th>Material</th></tr>";
    validData.forEach((d, i) => {
      bodyHtml += `<tr><td>${i + 1}</td><td>${d.title}</td><td>${d.pages}</td><td>${fmtInt(d.qty)}</td><td>${d.paper}</td><td>${d.material}</td></tr>`;
    });
  } else {
    headHtml += "<th>#</th><th>Title</th><th>Pages</th><th>Qty</th><th>Material</th><th>Binding</th></tr>";
    validData.forEach((d, i) => {
      bodyHtml += `<tr><td>${i + 1}</td><td>${d.title}</td><td>${d.pages}</td><td>${fmtInt(d.qty)}</td><td>${d.material}</td><td>${d.binding}</td></tr>`;
    });
  }

  headEl.innerHTML = headHtml;
  bodyEl.innerHTML = bodyHtml || '<tr><td colspan="6" class="text-muted">No valid rows</td></tr>';
  summaryEl.textContent = `${validData.length} valid row(s) ready to import`;

  // Store pending import
  pendingImport = { type, data: validData };

  // Enable/disable confirm button
  const btn = document.getElementById("confirmImportBtn");
  btn.disabled = validData.length === 0;

  new bootstrap.Modal("#importPreviewModal").show();
}

// Confirm import button handler
document.getElementById("confirmImportBtn").addEventListener("click", function () {
  const { type, data } = pendingImport;
  if (data.length === 0) return;

  if (type === "cost") {
    executeCostImport(data);
  } else if (type === "production") {
    executeProductionImport(data);
  }

  bootstrap.Modal.getInstance(document.getElementById("importPreviewModal")).hide();
  pendingImport = { type: null, data: [] };
});

// ===== Import Handlers =====
async function handleCostImport(input) {
  const file = input.files[0];
  if (!file) return;
  input.value = "";

  const parsed = await FileImporter.parseFile(file);
  if (parsed.errors.length > 0 && parsed.rows.length === 0) {
    Swal.fire("Import Error", parsed.errors.join("<br>"), "error");
    return;
  }

  const { valid, errors, warnings } = FileImporter.validateCost(parsed.rows);
  showImportPreview(valid, errors, warnings, "cost");
}

async function handleProdImport(input) {
  const file = input.files[0];
  if (!file) return;
  input.value = "";

  const parsed = await FileImporter.parseFile(file);
  if (parsed.errors.length > 0 && parsed.rows.length === 0) {
    Swal.fire("Import Error", parsed.errors.join("<br>"), "error");
    return;
  }

  const { valid, errors, warnings } = FileImporter.validateProduction(parsed.rows);
  showImportPreview(valid, errors, warnings, "production");
}

// ===== Execute Import Actions =====
function executeCostImport(data) {
  $("#numBooks").val(data.length);
  createBookTable(data.length);

  data.forEach((item, i) => {
    $(`#costTitle${i + 1}`).val(item.title || "");
    $(`#costPages${i + 1}`).val(item.pages);
    $(`#costQty${i + 1}`).val(item.qty);
    $(`#costPaper${i + 1}`).val(item.paper);
    $(`#costMaterial${i + 1}`).val(item.material);
  });

  calculateCostBooks(false);
  Swal.fire("Success", `${data.length} book(s) loaded into Cost Calculator`, "success");
}

function executeProductionImport(data) {
  const container = document.getElementById("bookContainer");
  container.innerHTML = "";

  data.forEach((item) => {
    const row = createProdRow();
    row.querySelector(".prod-title").value = item.title;
    row.querySelector(".prod-pages").value = item.pages;
    row.querySelector(".prod-qty").value = item.qty;
    row.querySelector(".prod-material").value = item.material;
    row.querySelector(".prod-binding").value = item.binding;
    container.appendChild(row);
  });

  Swal.fire("Success", `${data.length} book(s) loaded into Production Calculator`, "success");
}

function createProdRow() {
  const div = document.createElement("div");
  div.className = "book-row row align-items-center g-1 g-md-2";
  div.innerHTML = `
    <div class="col-12 col-sm-6 col-md-3 mb-1">
      <input class="form-control form-control-sm prod-title" placeholder="Book Title">
    </div>
    <div class="col-4 col-sm-3 col-md-2 mb-1">
      <input class="form-control form-control-sm prod-pages" type="number" placeholder="Pages" min="4" inputmode="numeric">
    </div>
    <div class="col-4 col-sm-3 col-md-2 mb-1">
      <input class="form-control form-control-sm prod-qty" type="number" placeholder="Qty" min="1" inputmode="numeric">
    </div>
    <div class="col-4 col-sm-3 col-md-2 mb-1">
      <select class="form-select form-select-sm prod-material">
        <option value="WF">WF</option><option value="MT">MT</option><option value="GL">GL</option>
      </select>
    </div>
    <div class="col-8 col-sm-4 col-md-2 mb-1">
      <select class="form-select form-select-sm prod-binding">
        <option value="Perfect">Perfect</option><option value="Staple">Staple</option>
      </select>
    </div>
    <div class="col-4 col-sm-2 col-md-1 mb-1 text-end">
      <button class="btn btn-outline-danger btn-sm removeRow" title="Remove"><i class="fa fa-trash"></i></button>
    </div>
  `;
  return div;
}

// ====================================================
// TAB 1: PAPER PACKAGE & COST CALCULATOR
// ====================================================
const COST = {
  TONER_PER_SHEET: 962 / 90000,
  INSIDE_A4: 31 / 4000,
  INSIDE_A3: 31 / 2000,
  COVER: 18 / 400,
  LAMINATE: 8 / 400,
  BINDING: 150 / 10000,
  STAPLE: (17 / 20 / 1000) * 2,
  PACK_WF_A4: 500 * 8,
  PACK_MT_A4: 250 * 8,
  PACK_A3: 500 * 4,
  COVER_PACK: 400
};

let costHistory = JSON.parse(localStorage.getItem("paperHistory")) || [];
let lastJobHash = "";

function getCostCapacity() {
  return parseInt($("#costDailyCapacity").val()) || 50000;
}

function createBookTable(n) {
  if (!n || n < 1) {
    $("#bookTableContainer").html("");
    $("#resultArea").hide();
    return;
  }
  let html = "";
  for (let i = 1; i <= n; i++) {
    html += `
      <div class="book-card row g-1 align-items-end p-2 mb-2 rounded" id="costCard${i}">
        <div class="col-2 col-md-1"><span class="badge bg-secondary">${i}</span></div>
        <div class="col-10 col-md-2"><input type="text" id="costTitle${i}" class="form-control form-control-sm calc-input" placeholder="Title"></div>
        <div class="col-4 col-md-1"><input type="number" id="costPages${i}" class="form-control form-control-sm calc-input" min="1" placeholder="Pages" inputmode="numeric"></div>
        <div class="col-4 col-md-2"><input type="number" id="costQty${i}" class="form-control form-control-sm calc-input" min="1" placeholder="Qty" inputmode="numeric"></div>
        <div class="col-4 col-md-2"><select id="costPaper${i}" class="form-select form-select-sm calc-input"><option value="A4">A4</option><option value="A3">A3</option></select></div>
        <div class="col-6 col-md-2"><select id="costMaterial${i}" class="form-select form-select-sm calc-input"><option value="WF">WF</option><option value="MT">MT</option></select></div>
        <div class="col-6 col-md-2"><input type="text" id="costBinding${i}" class="form-control form-control-sm" readonly placeholder="Bind"></div>
      </div>`;
  }
  $("#bookTableContainer").html(html);
}

function calculateCostBooks(saveHistory = false) {
  const n = parseInt($("#numBooks").val());
  if (!n || n < 1) { $("#resultArea").hide(); $("#btnExportExcel").hide(); return; }

  $("#resultBody,#packBody,#summaryCards").html("");

  let totalBooksQty = 0, totalCost = 0, totalCoverSheets = 0, totalA4Pages = 0;
  let sheetsWF = 0, sheetsMT = 0, sheetsA3 = 0;
  let bookRows = [];

  for (let i = 1; i <= n; i++) {
    const title = $(`#costTitle${i}`).val() || `Book ${i}`;
    const pages = parseInt($(`#costPages${i}`).val()) || 0;
    const qty = parseInt($(`#costQty${i}`).val()) || 0;
    const paper = $(`#costPaper${i}`).val();
    const material = $(`#costMaterial${i}`).val();
    const binding = paper === "A4" ? "Perfect" : "Staple";
    $(`#costBinding${i}`).val(binding);

    if (pages <= 0 || qty <= 0) continue;

    const sheetsPerBook = Math.ceil(pages / 2);
    const totalSheetsBook = sheetsPerBook * qty;
    const tonerSheets = paper === "A4" ? sheetsPerBook + 4 : sheetsPerBook * 2 + 4;
    const tonerCost = tonerSheets * COST.TONER_PER_SHEET;
    const insideCost = paper === "A4" ? sheetsPerBook * COST.INSIDE_A4 : sheetsPerBook * COST.INSIDE_A3;
    const pricePerBook = tonerCost + insideCost + COST.COVER + COST.LAMINATE +
      (binding === "Perfect" ? COST.BINDING : COST.STAPLE);

    totalBooksQty += qty;
    totalCost += pricePerBook * qty;
    totalCoverSheets += qty;
    totalA4Pages += paper === "A4" ? totalSheetsBook : totalSheetsBook * 2;

    if (paper === "A4" && material === "WF") sheetsWF += totalSheetsBook;
    if (paper === "A4" && material === "MT") sheetsMT += totalSheetsBook;
    if (paper === "A3") sheetsA3 += totalSheetsBook;

    bookRows.push({ title, pages, qty, paper, material, binding, pricePerBook, idx: i });
  }

  bookRows.forEach((b, i) => {
    $(`#costCard${b.idx}`).toggleClass("expensive", b.pricePerBook > 50);
    $("#resultBody").append(`<tr>
      <td>${i + 1}</td><td class="text-start">${sanitize(b.title)}</td><td>${fmtInt(b.pages)}</td><td>${fmtInt(b.qty)}</td>
      <td>${b.paper}</td><td>${b.material}</td><td>${b.binding}</td>
      <td>${fmtPrice(b.pricePerBook)}</td>
    </tr>`);
  });

  const packsWF = sheetsWF > 0 ? Math.ceil(sheetsWF / COST.PACK_WF_A4) : 0;
  const packsMT = sheetsMT > 0 ? Math.ceil(sheetsMT / COST.PACK_MT_A4) : 0;
  const packsA3 = sheetsA3 > 0 ? Math.ceil(sheetsA3 / COST.PACK_A3) : 0;
  const coverPacks = totalCoverSheets > 0 ? Math.ceil(totalCoverSheets / COST.COVER_PACK) : 0;

  const capacity = getCostCapacity();
  const totalDays = totalA4Pages > 0 && capacity > 0 ? Math.ceil(totalA4Pages / capacity) : 0;
  $("#costTotalDays").text(totalDays);

  if (sheetsWF > 0) $("#packBody").append(`<tr><td>A4-WF</td><td>${fmtInt(sheetsWF)}</td><td>${fmtInt(packsWF)}</td></tr>`);
  if (sheetsMT > 0) $("#packBody").append(`<tr><td>A4-MT</td><td>${fmtInt(sheetsMT)}</td><td>${fmtInt(packsMT)}</td></tr>`);
  if (sheetsA3 > 0) $("#packBody").append(`<tr><td>A3-WF</td><td>${fmtInt(sheetsA3)}</td><td>${fmtInt(packsA3)}</td></tr>`);
  if (totalCoverSheets > 0) $("#packBody").append(`<tr><td>Cover</td><td>${fmtInt(totalCoverSheets)}</td><td>${fmtInt(coverPacks)}</td></tr>`);

  // Summary cards
  const cards = [
    { title: "Books", icon: "fa-solid fa-book", value: fmtInt(totalBooksQty), color: "primary" },
    { title: "Days", icon: "fa-solid fa-clock", value: fmtInt(totalDays), color: "info" },
    { title: "Cost", icon: "fa-solid fa-dollar-sign", value: fmtPrice(totalCost), color: "danger" },
    { title: "Cover", icon: "fa-solid fa-box", value: fmtInt(coverPacks), color: "warning" }
  ];
  if (packsWF > 0) cards.push({ title: "WF Packs", icon: "fa-solid fa-cubes", value: fmtInt(packsWF), color: "success" });
  if (packsMT > 0) cards.push({ title: "MT Packs", icon: "fa-solid fa-cubes-stacked", value: fmtInt(packsMT), color: "secondary" });
  if (packsA3 > 0) cards.push({ title: "A3 Packs", icon: "fa-solid fa-boxes-stacked", value: fmtInt(packsA3), color: "dark" });

  $("#summaryCards").html(cards.map((c) => `
    <div class="col-4 col-sm-3 col-md-2 mb-2">
      <div class="card text-white bg-${c.color} h-100 shadow-sm">
        <div class="card-body text-center p-1 p-md-2">
          <i class="${c.icon} mb-1"></i>
          <div class="tiny-label">${c.title}</div>
          <div class="fw-bold small">${c.value}</div>
        </div>
      </div>
    </div>
  `).join(""));

  $("#resultArea").show();
  $("#btnExportExcel").show();

  // Store bookRows globally for export
  window._costBookRows = bookRows;
  window._costTotalCost = totalCost;

  if (saveHistory) saveCostHistory(n, totalBooksQty, totalCost, totalDays, bookRows);
}

function saveCostHistory(n, totalBooksQty, totalCost, totalDays, bookRows) {
  const hash = btoa(bookRows.map((b) => `${b.title}-${b.pages}-${b.qty}-${b.paper}-${b.material}`).join("|"));
  if (hash === lastJobHash) return;
  lastJobHash = hash;

  const entry = {
    timestamp: new Date().toLocaleString("en-US"),
    books: n,
    totalBooksQty,
    totalCost,
    totalDays,
    bookRows: bookRows.map((b) => ({ title: b.title, pages: b.pages, qty: b.qty, paper: b.paper, material: b.material }))
  };

  costHistory.unshift(entry);
  if (costHistory.length > 50) costHistory.pop();
  localStorage.setItem("paperHistory", JSON.stringify(costHistory));
  updateCostHistoryUI();
}

function updateCostHistoryUI() {
  const html = costHistory.map((h, i) => `
    <div class="history-card p-2 mb-1 rounded border small" data-index="${i}" role="button">
      <b>${h.timestamp}</b> | ${h.books} books | ${fmtPrice(h.totalCost)} | ${h.totalDays}d
    </div>
  `).join("");
  $("#historyCards").html(html);
}

// --- Cost Calculator Events ---
$("#numBooks").on("input", function () {
  createBookTable(parseInt($(this).val()));
  calculateCostBooks(false);
});

$("#btnCalculate").on("click", () => calculateCostBooks(true));
$("#toggleHistory").on("click", () => $("#historyArea").toggle());
$("#hideHistoryBtn").on("click", () => $("#historyArea").hide());

$("#historyCards").on("click", ".history-card", function () {
  const h = costHistory[$(this).data("index")];
  $("#numBooks").val(h.books);
  createBookTable(h.books);
  h.bookRows.forEach((b, i) => {
    $(`#costTitle${i + 1}`).val(b.title || "");
    $(`#costPages${i + 1}`).val(b.pages);
    $(`#costQty${i + 1}`).val(b.qty);
    $(`#costPaper${i + 1}`).val(b.paper);
    $(`#costMaterial${i + 1}`).val(b.material);
  });
  calculateCostBooks(false);
  $("#historyArea").hide();
});

$(document).on("input change", ".calc-input, #costDailyCapacity, #costDailySlider", function () {
  if (this.id === "costDailyCapacity") $("#costDailySlider").val($(this).val());
  if (this.id === "costDailySlider") $("#costDailyCapacity").val($(this).val());
  calculateCostBooks(false);
});

updateCostHistoryUI();

// --- Export Excel ---
$("#btnExportExcel").on("click", function () {
  const bookRows = window._costBookRows;
  const totalCost = window._costTotalCost;
  if (!bookRows || !bookRows.length) {
    Swal.fire("Info", "No data to export. Please calculate first.", "info");
    return;
  }

  const data = [];
  data.push(["Print Job Calculation"]);
  data.push(["Generated:", new Date().toLocaleString("en-US")]);
  data.push([]);
  data.push(["#", "Title", "Pages", "Qty", "Paper", "Material", "Binding", "Price/Book"]);
  bookRows.forEach((b, i) => {
    data.push([i + 1, b.title, b.pages, b.qty, b.paper, b.material, b.binding, Number(b.pricePerBook.toFixed(2))]);
  });
  data.push([]);
  data.push(["", "", "", "", "", "", "Total Cost:", Number(totalCost.toFixed(2))]);

  const ws = XLSX.utils.aoa_to_sheet(data);
  ws["!cols"] = [{ wch: 4 }, { wch: 20 }, { wch: 8 }, { wch: 10 }, { wch: 6 }, { wch: 8 }, { wch: 10 }, { wch: 10 }];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Calculation");
  XLSX.writeFile(wb, "print_job_" + new Date().toISOString().split("T")[0] + ".xlsx");
});

// ====================================================
// TAB 3: OFFSET PRODUCTION CALCULATOR
// ====================================================

// --- Add/Remove Book Rows ---
$("#addRow").on("click", function () {
  const container = document.getElementById("bookContainer");
  container.appendChild(createProdRow());
});

$("#bookContainer").on("click", ".removeRow", function () {
  const rows = document.querySelectorAll("#bookContainer .book-row");
  if (rows.length > 1) {
    $(this).closest(".book-row").remove();
  } else {
    Swal.fire("Info", "At least one book row is required", "info");
  }
});

// --- Print ---
// (moved to bottom with custom print logic)

// --- Schedule: Holidays ---
let prodHolidays = [];

// Set default start date to today
document.getElementById("prodStartDate").value = new Date().toISOString().split("T")[0];

// Add holidays when user types and presses Enter or comma
$("#prodHolidays").on("keydown", function (e) {
  if (e.key === "Enter" || e.key === ",") {
    e.preventDefault();
    addHolidays($(this).val());
    $(this).val("");
  }
}).on("blur", function () {
  if ($(this).val().trim()) {
    addHolidays($(this).val());
    $(this).val("");
  }
});

function addHolidays(text) {
  const dates = text.split(",").map((d) => d.trim()).filter((d) => d);
  dates.forEach((d) => {
    // Validate date format
    const parsed = new Date(d);
    if (isNaN(parsed.getTime())) return;
    const dateStr = parsed.toISOString().split("T")[0];
    if (!prodHolidays.includes(dateStr)) {
      prodHolidays.push(dateStr);
    }
  });
  renderHolidayTags();
}

function renderHolidayTags() {
  const html = prodHolidays.map((d, i) => `
    <span class="badge bg-danger d-inline-flex align-items-center gap-1" style="font-size:0.7rem;">
      ${d} <i class="fa fa-xmark" style="cursor:pointer;" onclick="removeHoliday(${i})"></i>
    </span>
  `).join("");
  $("#holidayTags").html(html);
}

function removeHoliday(index) {
  prodHolidays.splice(index, 1);
  renderHolidayTags();
}

/**
 * Calculate finish date: skip Sundays and holidays
 * @param {number} workDays - number of working days needed
 * @returns {{ finishDate: string, calendarDays: number }}
 */
function calcFinishDate(workDays) {
  const startStr = document.getElementById("prodStartDate").value;
  const start = startStr ? new Date(startStr) : new Date();
  let current = new Date(start);
  let counted = 0;

  while (counted < workDays) {
    current.setDate(current.getDate() + 1);
    const day = current.getDay(); // 0=Sunday
    const dateStr = current.toISOString().split("T")[0];
    // Skip Sunday and holidays
    if (day === 0) continue;
    if (prodHolidays.includes(dateStr)) continue;
    counted++;
  }

  return {
    finishDate: current.toISOString().split("T")[0],
    calendarDays: Math.round((current - start) / 86400000)
  };
}

// --- Production Calculate ---
$("#prodCalculate").on("click", function () {
  const waste = getVal("prodWaste", 0) / 100;
  const formsPerDay = getVal("prodFormsPerDay", 10); // forms (children) that can be printed per day (both sides)
  const dailyCap = getVal("prodDailyCap", 50000);    // sheets/day for material calculation
  const inkCap = getVal("prodInkCap", 33000);
  const aluCap = getVal("prodAluCap", 15000);
  const powderCap = getVal("prodPowderCap", 22000);
  const blanketLife = getVal("prodBlanketLife", 500000);
  const glueCap = getVal("prodGlueCap", 10000);
  const laminateCap = getVal("prodLamCap", 4000);
  const stapleCap = getVal("prodStapleCap", 10000);
  const a1Pack = getVal("prodA1Pack", 500);
  const coverPack = getVal("prodCoverPack", 100);

  let totalInsideA2 = 0, totalCoverA3 = 0;
  let totalPerfectQty = 0, totalStapleQty = 0;
  let totalCTP = 0;
  let totalForms = 0, totalImpressions = 0, totalDaysAll = 0;
  let totalQty = 0;
  let hasError = false;
  const errorMessages = [];
  const notices = [];

  const detailTable = document.getElementById("prodDetailTable");
  detailTable.innerHTML = "";

  const rows = document.querySelectorAll("#bookContainer .book-row");
  let processedCount = 0;

  rows.forEach((row, index) => {
    row.classList.remove("error-row");

    const titleInput = row.querySelector(".prod-title").value;
    const title = sanitize(titleInput) || "Book " + (index + 1);
    const pagesInput = row.querySelector(".prod-pages");
    let pages = parseInt(pagesInput.value) || 0;
    const qty = parseInt(row.querySelector(".prod-qty").value) || 0;
    const binding = row.querySelector(".prod-binding").value;

    if (pages === 0 && qty === 0) return; // Skip empty rows

    // Validation
    if (pages < 4 && pages > 0) {
      errorMessages.push(`[${title}]: Pages must be at least 4`);
      row.classList.add("error-row");
      hasError = true;
      return;
    }
    if (pages === 0 || qty === 0) {
      errorMessages.push(`[${title}]: Both pages and quantity are required`);
      row.classList.add("error-row");
      hasError = true;
      return;
    }

    // Auto-round pages to fit binding requirement (add blank pages)
    if (binding === "Perfect" && pages % 8 !== 0) {
      const adjusted = Math.ceil(pages / 8) * 8;
      notices.push(`[${title}]: ${pages}pg → ${adjusted}pg (+${adjusted - pages} blank pages for Perfect Binding)`);
      pages = adjusted;
      pagesInput.value = pages;
    }
    if (binding === "Staple" && pages % 4 !== 0) {
      const adjusted = Math.ceil(pages / 4) * 4;
      notices.push(`[${title}]: ${pages}pg → ${adjusted}pg (+${adjusted - pages} blank pages for Staple Binding)`);
      pages = adjusted;
      pagesInput.value = pages;
    }

    // ===== OFFSET PRINTING CALCULATION =====
    // 1 form (child) = 8 pages (printed on ONE side of A2)
    // 1 A2 sheet = 2 forms (front + back) = 16 pages total
    // forms = pages ÷ 8 (number of different plate sets)
    // Physical A2 sheets per book = forms ÷ 2 (2 forms per sheet, front & back)
    // Total A2 sheets = (forms ÷ 2) × qty
    // Paper: we buy A1, cut to A2. 1 A1 = 2 A2.

    let forms = 0;
    let insideA2 = 0;
    let plateCount = 0;

    if (binding === "Perfect") {
      forms = Math.ceil(pages / 8);
      const sheetsPerBook = Math.ceil(forms / 2);
      insideA2 = sheetsPerBook * qty;
      insideA2 = Math.ceil(insideA2 * (1 + waste));
      totalInsideA2 += insideA2;
      plateCount = forms * 8;
      totalCTP += plateCount;
      totalPerfectQty += qty;
    } else {
      const totalPages = pages + 4;
      forms = Math.ceil(totalPages / 8);
      const sheetsPerBook = Math.ceil(forms / 2);
      insideA2 = sheetsPerBook * qty;
      insideA2 = Math.ceil(insideA2 * (1 + waste));
      totalInsideA2 += insideA2;
      plateCount = forms * 8;
      totalCTP += plateCount;
      totalStapleQty += qty;
    }

    // Every book needs 1 cover (A3) regardless of binding
    totalCoverA3 += qty;

    // Days calculation for this book (if printed alone, for reference only)
    const impressions = forms * qty;
    const daysForBook = Math.ceil(forms / formsPerDay);

    totalForms += forms;
    totalImpressions += impressions;
    totalQty += qty;

    processedCount++;
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${processedCount}</td>
      <td class="text-start text-truncate" style="max-width:200px"><strong>${title}</strong></td>
      <td>${pages.toLocaleString("en-US")}</td>
      <td style="color:#2980b9; font-weight:600;">${qty.toLocaleString("en-US")}</td>
      <td style="color:#8e44ad;">${forms.toLocaleString("en-US")}</td>
      <td style="color:#d35400;">${impressions.toLocaleString("en-US")}</td>
      <td style="color:#27ae60; font-weight:600;">${insideA2.toLocaleString("en-US")}</td>
      <td style="color:#c0392b;">${plateCount.toLocaleString("en-US")}</td>
      <td style="background:#eaf6ff; color:#2980b9; font-weight:700;">${daysForBook.toLocaleString("en-US")}</td>
    `;
    detailTable.appendChild(tr);
  });

  if (hasError) {
    Swal.fire("Validation Errors", errorMessages.join("<br>"), "error");
    return;
  }

  // Show page adjustment notices (non-blocking, calculation continues)
  if (notices.length > 0) {
    Swal.fire({
      title: "Pages Adjusted (blank pages added)",
      html: notices.join("<br>"),
      icon: "info",
      confirmButtonText: "OK"
    });
  }

  if (processedCount === 0) {
    detailTable.innerHTML = '<tr><td colspan="9" class="text-muted">Please enter book data above</td></tr>';
    $("#prodPrint").hide();
    $("#prodDetailFoot").hide();
    return;
  }

  // Show footer totals
  // Total days = total forms ÷ formsPerDay (continuous production, no wasted capacity between books)
  totalDaysAll = Math.ceil(totalForms / formsPerDay);
  
  // Add extra days (buffer/maintenance/other)
  const extraDays = parseInt(document.getElementById("prodExtraDays").value) || 0;
  totalDaysAll += extraDays;
  
  // Calculate finish date (skip Sundays + holidays)
  const schedule = calcFinishDate(totalDaysAll);
  
  $("#prodDetailFoot").show();
  $("#prodTotalQty2").text(totalQty.toLocaleString("en-US"));
  $("#prodTotalForms").text(totalForms.toLocaleString("en-US"));
  $("#prodTotalImpress").text(totalImpressions.toLocaleString("en-US"));
  $("#prodTotalA2").text(totalInsideA2.toLocaleString("en-US"));
  $("#prodTotalPlates").text(totalCTP.toLocaleString("en-US"));
  $("#prodTotalDaysCol").text(totalDaysAll.toLocaleString("en-US"));

  // Final material calculations
  // Inside paper: A2 → A1 (1 A1 = 2 A2), pack = a1Pack sheets
  const insideA1 = Math.ceil(totalInsideA2 / 2);
  const insidePack = Math.ceil(insideA1 / a1Pack);
  
  // Cover paper: each book = 1 cover (A3 size)
  // 1 A1 = 4 A3 (A1 → 2 A2 → 4 A3)
  // So: total A3 covers ÷ 4 = A1 sheets needed
  const coverA1 = Math.ceil(totalCoverA3 / 4);
  const coverPackNeed = Math.ceil(coverA1 / coverPack);

  const ink = Math.ceil(totalInsideA2 / inkCap);
  const alu = Math.ceil(totalInsideA2 / aluCap);
  const powder = Math.ceil(totalInsideA2 / powderCap);
  const blanket = Math.ceil(totalInsideA2 / blanketLife);
  const totalGlue = Math.ceil(totalPerfectQty / glueCap);
  const totalLaminate = Math.ceil(totalQty / laminateCap);
  const totalStaple = Math.ceil(totalStapleQty / stapleCap);

  // Update UI
  $("#prodTotalBooks").text(processedCount.toLocaleString("en-US"));
  $("#prodTotalQty").text(totalQty.toLocaleString("en-US"));
  $("#prodInsideSheets").text(totalInsideA2.toLocaleString("en-US"));
  $("#prodInsidePack").text(insidePack.toLocaleString("en-US"));
  $("#prodCoverPackNeed").text(coverPackNeed.toLocaleString("en-US"));
  $("#prodCtpResult").text(totalCTP.toLocaleString("en-US"));
  $("#prodDays").text(totalDaysAll.toLocaleString("en-US") + " Work Days");
  const startDate = document.getElementById("prodStartDate").value || new Date().toISOString().split("T")[0];
  $("#prodFinishDate").text(startDate + " → " + schedule.finishDate);

  $("#prodInk").text(ink.toLocaleString("en-US"));
  $("#prodAlu").text(alu.toLocaleString("en-US"));
  $("#prodPowder").text(powder.toLocaleString("en-US"));
  $("#prodBlanket").text(blanket.toLocaleString("en-US"));
  $("#prodGlue").text(totalGlue.toLocaleString("en-US"));
  $("#prodLam").text(totalLaminate.toLocaleString("en-US"));
  $("#prodStaple").text(totalStaple.toLocaleString("en-US"));

  // Store for export/print
  window._prodData = {
    books: [],
    totalBooks: processedCount, totalQty,
    totalForms, totalImpressions, totalInsideA2, totalCTP, totalDaysAll,
    insidePack, coverPackNeed, ink, alu, powder, blanket,
    totalGlue, totalLaminate, totalStaple,
    finishDate: schedule.finishDate, calendarDays: schedule.calendarDays,
    holidays: prodHolidays.length
  };
  // Collect book details from table
  detailTable.querySelectorAll("tr").forEach((tr) => {
    const cells = tr.querySelectorAll("td");
    if (cells.length >= 9) {
      window._prodData.books.push({
        num: cells[0].textContent,
        title: cells[1].textContent,
        pages: cells[2].textContent,
        qty: cells[3].textContent,
        forms: cells[4].textContent,
        impressions: cells[5].textContent,
        a2: cells[6].textContent,
        plates: cells[7].textContent,
        days: cells[8].textContent
      });
    }
  });

  $("#prodPrint").show();
  $("#prodExportExcel").show();
});

// ====================================================
// PRODUCTION: EXPORT EXCEL
// ====================================================
$("#prodExportExcel").on("click", function () {
  const d = window._prodData;
  if (!d || !d.books.length) {
    Swal.fire("Info", "No data. Calculate first.", "info");
    return;
  }

  const data = [];
  data.push(["Production Job Calculation"]);
  data.push(["Generated:", new Date().toLocaleString("en-US")]);
  data.push([]);

  // Summary
  data.push(["SUMMARY"]);
  data.push(["Total Books", d.totalBooks]);
  data.push(["Total Qty", d.totalQty]);
  data.push(["A2 Sheets", d.totalInsideA2]);
  data.push(["Inside Packs", d.insidePack]);
  data.push(["Cover Packs", d.coverPackNeed]);
  data.push(["CTP Plates", d.totalCTP]);
  data.push(["Total Days (Work)", d.totalDaysAll]);
  data.push(["Finish Date", d.finishDate]);
  data.push(["Holidays Excluded", d.holidays]);
  data.push([]);

  // Job Breakdown
  data.push(["JOB BREAKDOWN"]);
  data.push(["#", "Title", "Pages", "Qty", "Forms", "Impressions", "A2 Sheets", "Plates", "Days"]);
  d.books.forEach((b) => {
    data.push([b.num, b.title, b.pages, b.qty, b.forms, b.impressions, b.a2, b.plates, b.days]);
  });
  data.push(["", "TOTAL", "", d.totalQty, d.totalForms, d.totalImpressions, d.totalInsideA2, d.totalCTP, d.totalDaysAll]);
  data.push([]);

  // Consumables
  data.push(["CONSUMABLES"]);
  data.push(["Ink", "Alu", "Powder", "Blanket", "Glue", "Laminate", "Staples"]);
  data.push([d.ink, d.alu, d.powder, d.blanket, d.totalGlue, d.totalLaminate, d.totalStaple]);

  const ws = XLSX.utils.aoa_to_sheet(data);
  ws["!cols"] = [
    { wch: 5 }, { wch: 22 }, { wch: 10 }, { wch: 10 },
    { wch: 8 }, { wch: 14 }, { wch: 12 }, { wch: 8 }, { wch: 6 }
  ];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Production");
  XLSX.writeFile(wb, "production_" + new Date().toISOString().split("T")[0] + ".xlsx");
});

// ====================================================
// PRODUCTION: CUSTOM PRINT
// ====================================================
$("#prodPrint").on("click", function () {
  const d = window._prodData;
  if (!d || !d.books.length) return;

  buildPrintPreview();
  new bootstrap.Modal("#printPreviewModal").show();
});

// Rebuild preview when checkboxes change
$(document).on("change", ".print-opt", function () {
  buildPrintPreview();
});

function buildPrintPreview() {
  const d = window._prodData;
  if (!d) return;

  const showHeader = $(".print-opt[value='header']").is(":checked");
  const showSummary = $(".print-opt[value='summary']").is(":checked");
  const showBreakdown = $(".print-opt[value='breakdown']").is(":checked");
  const showConsumables = $(".print-opt[value='consumables']").is(":checked");

  let html = "";

  if (showHeader) {
    html += `<div class="text-center mb-3">
      <h5 class="fw-bold">Production Job Ticket</h5>
      <small class="text-muted">${new Date().toLocaleString("en-US")} | ${d.totalDaysAll} Work Days (excl. Sundays${d.holidays > 0 ? " + " + d.holidays + " holidays" : ""}) | Finish: ${d.finishDate}</small>
    </div>`;
  }

  if (showSummary) {
    html += `<table class="table table-bordered table-sm text-center mb-3">
      <thead><tr><th>Books</th><th>Total Qty</th><th>A2 Sheets</th><th>Inside Packs</th><th>Cover Packs</th><th>CTP Plates</th><th>Days</th></tr></thead>
      <tbody><tr>
        <td><b>${d.totalBooks.toLocaleString("en-US")}</b></td>
        <td><b>${d.totalQty.toLocaleString("en-US")}</b></td>
        <td><b>${d.totalInsideA2.toLocaleString("en-US")}</b></td>
        <td><b>${d.insidePack.toLocaleString("en-US")}</b></td>
        <td><b>${d.coverPackNeed.toLocaleString("en-US")}</b></td>
        <td><b>${d.totalCTP.toLocaleString("en-US")}</b></td>
        <td><b>${d.totalDaysAll.toLocaleString("en-US")}</b></td>
      </tr></tbody>
    </table>`;
  }

  if (showBreakdown) {
    html += `<table class="table table-bordered table-sm text-center small mb-3">
      <thead><tr><th>#</th><th>Title</th><th>Pages</th><th>Qty</th><th>Forms</th><th>Impress.</th><th>A2</th><th>Plates</th><th>Days</th></tr></thead>
      <tbody>`;
    d.books.forEach((b) => {
      html += `<tr><td>${b.num}</td><td class="text-start">${b.title}</td><td>${b.pages}</td><td>${b.qty}</td><td>${b.forms}</td><td>${b.impressions}</td><td>${b.a2}</td><td>${b.plates}</td><td><b>${b.days}</b></td></tr>`;
    });
    html += `<tr class="fw-bold"><td colspan="3" class="text-end">TOTAL</td><td>${d.totalQty.toLocaleString("en-US")}</td><td>${d.totalForms.toLocaleString("en-US")}</td><td>${d.totalImpressions.toLocaleString("en-US")}</td><td>${d.totalInsideA2.toLocaleString("en-US")}</td><td>${d.totalCTP.toLocaleString("en-US")}</td><td>${d.totalDaysAll.toLocaleString("en-US")}</td></tr>`;
    html += `</tbody></table>`;
  }

  if (showConsumables) {
    html += `<table class="table table-bordered table-sm text-center mb-3">
      <thead><tr><th>Ink</th><th>Alu</th><th>Powder</th><th>Blanket</th><th>Glue</th><th>Laminate</th><th>Staples</th></tr></thead>
      <tbody><tr>
        <td>${d.ink.toLocaleString("en-US")}</td>
        <td>${d.alu.toLocaleString("en-US")}</td>
        <td>${d.powder.toLocaleString("en-US")}</td>
        <td>${d.blanket.toLocaleString("en-US")}</td>
        <td>${d.totalGlue.toLocaleString("en-US")}</td>
        <td>${d.totalLaminate.toLocaleString("en-US")}</td>
        <td>${d.totalStaple.toLocaleString("en-US")}</td>
      </tr></tbody>
    </table>`;
  }

  $("#printArea").html(html);
}

// Actual print
$("#doPrint").on("click", function () {
  const content = document.getElementById("printArea").innerHTML;
  const win = window.open("", "_blank");
  win.document.write(`<!DOCTYPE html><html><head>
    <title>Production Job Ticket</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet">
    <style>
      body { font-family: 'Poppins', sans-serif; padding: 15px; font-size: 12px; }
      table { width: 100%; }
      th { font-size: 11px; font-weight: 700; }
      td { font-size: 11px; }
      @media print { body { padding: 5mm; } }
    </style>
  </head><body>${content}</body></html>`);
  win.document.close();
  setTimeout(() => { win.print(); }, 500);
});

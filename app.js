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
  runProductionCalc();
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
function runProductionCalc() {
  const waste = getVal("prodWaste", 0) / 100;
  const formsPerDay = getVal("prodFormsPerDay", 10); // forms that can be printed per day (both sides)
  const dailyCap = getVal("prodDailyCap", 50000);    // sheets/day for press capacity
  const inkCap = getVal("prodInkCap", 33000);
  const aluCap = getVal("prodAluCap", 15000);
  const powderCap = getVal("prodPowderCap", 22000);
  const blanketLife = getVal("prodBlanketLife", 500000);
  const glueCap = getVal("prodGlueCap", 10000);
  const laminateCap = getVal("prodLamCap", 4000);
  const stapleCap = getVal("prodStapleCap", 10000);
  const wfPackCap = getVal("prodWfPack", 500);
  const mtPackCap = getVal("prodMtPack", 250);
  const glPackCap = getVal("prodGlPack", 100);
  const coverPack = getVal("prodCoverPack", 100);

  let totalInsideA2 = 0, totalCoverA3 = 0;
  let totalPerfectQty = 0, totalStapleQty = 0;
  let totalCTP = 0;
  let totalForms = 0, totalImpressions = 0, totalDaysAll = 0;
  let totalQty = 0;
  let materialSheets = { WF: 0, MT: 0, GL: 0 };
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
    const material = (row.querySelector(".prod-material") ? row.querySelector(".prod-material").value : "WF") || "WF";
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
    // 1 Form = 8 pages (BOTH sides)
    // 1 A2 sheet = 2 forms = 16 pages
    // Forms = Math.ceil(pages / 8)
    // Formula to Calculate Paper Inside: (forms(A2)/2)*QTY OR ((PAGE/16)*QTY)
    // With waste: insideA2 = Math.ceil(((forms / 2) * qty) * (1 + waste))

    const forms = Math.ceil(pages / 8);
    const netInsideA2 = (forms / 2) * qty;
    const insideA2 = Math.ceil(netInsideA2 * (1 + waste));

    // CTP Plates: 4 plates per form for inside (CMYK) + 4 plates for cover (CMYK)
    const insidePlates = forms * 4;
    const coverPlates = 4;
    const plateCount = insidePlates + coverPlates;

    totalInsideA2 += insideA2;
    materialSheets[material] = (materialSheets[material] || 0) + insideA2;
    totalCTP += plateCount;

    if (binding === "Perfect") {
      totalPerfectQty += qty;
    } else {
      totalStapleQty += qty;
    }

    // Every book needs 1 cover (A3)
    totalCoverA3 += qty;

    // Impressions & Days calculation
    const impressions = forms * qty;
    const daysForBook = Math.max(Math.ceil(forms / formsPerDay), Math.ceil(insideA2 / dailyCap));

    totalForms += forms;
    totalImpressions += impressions;
    totalQty += qty;

    processedCount++;
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${processedCount}</td>
      <td class="text-start text-truncate" style="max-width:180px"><strong>${title}</strong></td>
      <td>${pages.toLocaleString("en-US")}</td>
      <td style="color:#2980b9; font-weight:600;">${qty.toLocaleString("en-US")}</td>
      <td><span class="badge bg-secondary">${material}</span></td>
      <td><span class="badge ${binding === "Perfect" ? "bg-primary" : "bg-info"}">${binding}</span></td>
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
    detailTable.innerHTML = '<tr><td colspan="11" class="text-muted">Please enter book data above</td></tr>';
    $("#prodPrint").hide();
    $("#prodDetailFoot").hide();
    return;
  }

  // Show footer totals
  totalDaysAll = Math.max(Math.ceil(totalForms / formsPerDay), Math.ceil(totalInsideA2 / dailyCap));
  
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
  // Inside paper packs per material:
  // WF: wfSheets / prodWfPack (default 500)
  // MT: mtSheets / prodMtPack (default 250)
  // GL: glSheets / prodGlPack (default 100)
  const wfPacks = Math.ceil((materialSheets["WF"] || 0) / wfPackCap);
  const mtPacks = Math.ceil((materialSheets["MT"] || 0) / mtPackCap);
  const glPacks = Math.ceil((materialSheets["GL"] || 0) / glPackCap);
  const insidePack = wfPacks + mtPacks + glPacks;
  
  // Cover paper: each book = 1 cover (A3 size) with waste
  // 1 A1 = 4 A3 (A1 → 2 A2 → 4 A3)
  const coverA3WithWaste = Math.ceil(totalCoverA3 * (1 + waste));
  const coverA1 = Math.ceil(coverA3WithWaste / 4);
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
  $("#prodTotalFormsSummary").text(totalForms.toLocaleString("en-US"));
  $("#prodInsideSheets").text(totalInsideA2.toLocaleString("en-US"));
  $("#prodInsidePack").text(insidePack.toLocaleString("en-US"));
  $("#prodCoverPackNeed").text(coverPackNeed.toLocaleString("en-US"));
  $("#prodCtpResult").text(totalCTP.toLocaleString("en-US"));
  $("#prodSpeedBadge").html(`<i class="fa-solid fa-gauge-high"></i> ${formsPerDay.toLocaleString("en-US")} Child/Day`);
  $("#prodDays").text(totalDaysAll.toLocaleString("en-US") + " Work Days");
  const startDate = document.getElementById("prodStartDate").value || new Date().toISOString().split("T")[0];
  $("#prodFinishDate").text(startDate + " → " + schedule.finishDate);

  // Material breakdown badges
  $("#prodWfPackBadge").html(`<i class="fa-solid fa-layer-group"></i> WF: ${(materialSheets["WF"] || 0).toLocaleString("en-US")} sh (${wfPacks.toLocaleString("en-US")} pk @${wfPackCap})`);
  $("#prodMtPackBadge").html(`<i class="fa-solid fa-layer-group"></i> MT: ${(materialSheets["MT"] || 0).toLocaleString("en-US")} sh (${mtPacks.toLocaleString("en-US")} pk @${mtPackCap})`);
  $("#prodGlPackBadge").html(`<i class="fa-solid fa-layer-group"></i> GL: ${(materialSheets["GL"] || 0).toLocaleString("en-US")} sh (${glPacks.toLocaleString("en-US")} pk @${glPackCap})`);
  $("#prodCoverBadge").html(`<i class="fa-solid fa-box"></i> Cover: ${coverA3WithWaste.toLocaleString("en-US")} A3 (${coverPackNeed.toLocaleString("en-US")} pk @${coverPack})`);

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
    materialSheets, wfPacks, mtPacks, glPacks,
    wfPackCap, mtPackCap, glPackCap, coverPack,
    formsPerDay,
    finishDate: schedule.finishDate, calendarDays: schedule.calendarDays,
    holidays: prodHolidays.length
  };
  // Collect book details from table
  detailTable.querySelectorAll("tr").forEach((tr) => {
    const cells = tr.querySelectorAll("td");
    if (cells.length >= 11) {
      window._prodData.books.push({
        num: cells[0].textContent,
        title: cells[1].textContent,
        pages: cells[2].textContent,
        qty: cells[3].textContent,
        material: cells[4].textContent,
        binding: cells[5].textContent,
        forms: cells[6].textContent,
        impressions: cells[7].textContent,
        a2: cells[8].textContent,
        plates: cells[9].textContent,
        days: cells[10].textContent
      });
    }
  });

  $("#prodPrint").show();
  $("#prodExportExcel").show();
}

// Function to update speed badge dynamically in real time
function updateSpeedBadge() {
  const speed = getVal("prodFormsPerDay", 10);
  $("#prodSpeedBadge").html(`<i class="fa-solid fa-gauge-high"></i> ${speed.toLocaleString("en-US")} Child/Day`);
}

// Initial speed badge setup
updateSpeedBadge();

// --- Button click + auto-calculate on input change ---
$("#prodCalculate").on("click", runProductionCalc);

// Auto-calculate & update speed badge dynamically whenever any input changes
$(document).on("input change keyup", "#bookContainer input, #bookContainer select, #prodSettings input, #prodFormsPerDay, #prodDailyCap, #prodWaste, #prodBlanketLife, #prodWfPack, #prodMtPack, #prodGlPack, #prodCoverPack, #prodInkCap, #prodAluCap, #prodPowderCap, #prodGlueCap, #prodLamCap, #prodStapleCap, #prodExtraDays, #prodStartDate", function () {
  updateSpeedBadge();
  clearTimeout(window._prodCalcTimer);
  window._prodCalcTimer = setTimeout(function () {
    const hasData = $("#bookContainer .prod-pages").toArray().some((el) => el.value > 0);
    if (hasData) runProductionCalc();
  }, 300);
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
  data.push(["Total Children", d.totalForms]);
  data.push(["Producing Speed", `${d.formsPerDay || 10} Child/Day`]);
  data.push(["A2 Sheets (Total)", d.totalInsideA2]);
  data.push(["Inside Packs (Total)", d.insidePack]);
  data.push(["WF Sheets / Packs", `${d.materialSheets.WF || 0} sheets / ${d.wfPacks || 0} packs (@${d.wfPackCap} sh/pk)`]);
  data.push(["MT Sheets / Packs", `${d.materialSheets.MT || 0} sheets / ${d.mtPacks || 0} packs (@${d.mtPackCap} sh/pk)`]);
  data.push(["GL Sheets / Packs", `${d.materialSheets.GL || 0} sheets / ${d.glPacks || 0} packs (@${d.glPackCap} sh/pk)`]);
  data.push(["Cover Packs", d.coverPackNeed]);
  data.push(["CTP Plates", d.totalCTP]);
  data.push(["Total Days (Work)", d.totalDaysAll]);
  data.push(["Finish Date", d.finishDate]);
  data.push(["Holidays Excluded", d.holidays]);
  data.push([]);

  // Job Breakdown
  data.push(["JOB BREAKDOWN"]);
  data.push(["#", "Title", "Pages", "Qty", "Material", "Binding", "Child", "Impressions", "A2 Sheets", "Plates", "Days"]);
  d.books.forEach((b) => {
    data.push([b.num, b.title, b.pages, b.qty, b.material, b.binding, b.forms, b.impressions, b.a2, b.plates, b.days]);
  });
  data.push(["", "TOTAL", "", d.totalQty, "", "", d.totalForms, d.totalImpressions, d.totalInsideA2, d.totalCTP, d.totalDaysAll]);
  data.push([]);

  // Consumables
  data.push(["CONSUMABLES"]);
  data.push(["Ink", "Alu", "Powder", "Blanket", "Glue", "Laminate", "Staples"]);
  data.push([d.ink, d.alu, d.powder, d.blanket, d.totalGlue, d.totalLaminate, d.totalStaple]);

  const ws = XLSX.utils.aoa_to_sheet(data);
  ws["!cols"] = [
    { wch: 5 }, { wch: 22 }, { wch: 10 }, { wch: 10 }, { wch: 8 }, { wch: 10 },
    { wch: 8 }, { wch: 14 }, { wch: 12 }, { wch: 8 }, { wch: 6 }
  ];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Production");
  XLSX.writeFile(wb, "production_" + new Date().toISOString().split("T")[0] + ".xlsx");
});

// ====================================================
// PRODUCTION: CUSTOM PRINT & SELECTION
// ====================================================
$("#prodPrint").on("click", function () {
  const d = window._prodData;
  if (!d || !d.books || !d.books.length) {
    if (typeof Swal !== "undefined") {
      Swal.fire({ icon: "warning", title: "No Data", text: "Please calculate production before printing." });
    } else {
      alert("Please calculate production before printing.");
    }
    return;
  }

  // Populate book checkboxes
  let bookHtml = "";
  d.books.forEach((b, i) => {
    bookHtml += `
      <div class="form-check form-check-inline mb-1">
        <input class="form-check-input print-book-opt" type="checkbox" id="printBook_${i}" value="${i}" checked>
        <label class="form-check-label small" for="printBook_${i}">${b.num}. ${b.title || "Book " + b.num} <span class="badge bg-light text-dark border ms-1">${b.forms} Child</span></label>
      </div>`;
  });
  $("#printBookCheckboxes").html(bookHtml);

  buildPrintPreview();
  const modalEl = document.getElementById("printPreviewModal");
  if (modalEl) {
    const modal = bootstrap.Modal.getOrCreateInstance(modalEl);
    modal.show();
  }
});

// Rebuild preview when section checkboxes change
$(document).on("change", ".print-opt, .print-book-opt", function () {
  buildPrintPreview();
});

// Select / Deselect all sections
$(document).on("click", "#selectAllPrintOpts", function () {
  $(".print-opt").prop("checked", true);
  buildPrintPreview();
});

$(document).on("click", "#clearAllPrintOpts", function () {
  $(".print-opt").prop("checked", false);
  buildPrintPreview();
});

// Select / Deselect all books
$(document).on("click", "#selectAllPrintBooks", function (e) {
  e.preventDefault();
  $(".print-book-opt").prop("checked", true);
  buildPrintPreview();
});

$(document).on("click", "#clearAllPrintBooks", function (e) {
  e.preventDefault();
  $(".print-book-opt").prop("checked", false);
  buildPrintPreview();
});

function buildPrintPreview() {
  const d = window._prodData;
  if (!d) return;

  const showHeader = $("#optHeader").is(":checked");
  const showSummary = $("#optSummary").is(":checked");
  const showMaterials = $("#optMaterials").is(":checked");
  const showBreakdown = $("#optBreakdown").is(":checked");
  const showConsumables = $("#optConsumables").is(":checked");
  const showSignatures = $("#optSignatures").is(":checked");

  // Selected books
  const selectedBookIndices = [];
  $(".print-book-opt:checked").each(function () {
    selectedBookIndices.push(parseInt($(this).val(), 10));
  });

  const allBooksCount = d.books.length;
  const isSubset = selectedBookIndices.length > 0 && selectedBookIndices.length < allBooksCount;
  const booksToDisplay = d.books.filter((_, idx) => selectedBookIndices.includes(idx));

  // Compute stats for displayed books if subset, otherwise full project
  let dispBooks = booksToDisplay.length;
  let dispQty = 0;
  let dispForms = 0;
  let dispA2 = 0;
  let dispImpressions = 0;
  let dispPlates = 0;
  let dispDays = 0;

  if (isSubset) {
    booksToDisplay.forEach(b => {
      dispQty += b.rawQty || 0;
      dispForms += b.forms || 0;
      dispA2 += b.a2Raw || 0;
      dispImpressions += b.impressionsRaw || 0;
      dispPlates += b.platesRaw || 0;
      dispDays += b.days || 0;
    });
  } else {
    dispBooks = d.totalBooks;
    dispQty = d.totalQty;
    dispForms = d.totalForms;
    dispA2 = d.totalInsideA2;
    dispImpressions = d.totalImpressions;
    dispPlates = d.totalCTP;
    dispDays = d.totalDaysAll;
  }

  let html = "";

  // 1. Header
  if (showHeader) {
    html += `
      <div class="print-header text-center mb-3 pb-2 border-bottom">
        <div class="d-flex justify-content-between align-items-center mb-1">
          <div class="text-start">
            <h5 class="fw-bold mb-0 text-dark"><i class="fa-solid fa-industry text-primary"></i> Production Job Ticket</h5>
            <div class="d-flex gap-1 mt-1">
              <span class="badge bg-dark">${dispBooks} Job${dispBooks > 1 ? "s" : ""}</span>
              <span class="badge bg-primary">${dispQty.toLocaleString("en-US")} Total Qty</span>
              <span class="badge bg-purple" style="background:#8e44ad; color:#fff;">${dispForms.toLocaleString("en-US")} Children</span>
            </div>
          </div>
          <div class="text-end small text-muted">
            <div><b>Date:</b> ${new Date().toLocaleDateString("en-US", { year: "numeric", month: "short", day: "numeric" })}</div>
            <div><b>Target Finish:</b> <span class="badge bg-success">${d.finishDate}</span></div>
          </div>
        </div>
        <div class="d-flex justify-content-center gap-3 small text-muted flex-wrap mt-1">
          <span><i class="fa-solid fa-gauge-high text-primary"></i> <b>Speed:</b> ${d.formsPerDay || 10} Child/Day</span>
          <span><i class="fa-solid fa-calendar-day text-secondary"></i> <b>Work Days:</b> ${dispDays} Days</span>
          <span><i class="fa-solid fa-calendar-check text-success"></i> <b>Timeline:</b> ${d.startDate || "N/A"} → ${d.finishDate}</span>
        </div>
      </div>`;
  }

  // 2. 7 Summary Cards
  if (showSummary) {
    html += `
      <div class="prod-summary-grid mb-3">
        <div class="summary-box" style="background:#f0f0f0; border:2px solid #7f8c8d;">
          <div class="tiny-label" style="font-weight:700; color:#555;">BOOKS</div>
          <h5 class="mb-0 fw-bold" style="color:#2c3e50;">${dispBooks.toLocaleString("en-US")}</h5>
        </div>
        <div class="summary-box" style="background:#eaf6ff; border:2px solid #2980b9;">
          <div class="tiny-label" style="font-weight:700; color:#2980b9;">TOTAL QTY</div>
          <h5 class="mb-0 fw-bold" style="color:#2980b9;">${dispQty.toLocaleString("en-US")}</h5>
        </div>
        <div class="summary-box" style="background:#f5eef8; border:2px solid #8e44ad;">
          <div class="tiny-label" style="font-weight:700; color:#8e44ad;">TOTAL CHILDREN</div>
          <h5 class="mb-0 fw-bold" style="color:#8e44ad;">${dispForms.toLocaleString("en-US")}</h5>
        </div>
        <div class="summary-box" style="background:#e8fdf5; border:2px solid #27ae60;">
          <div class="tiny-label" style="font-weight:700; color:#27ae60;">A2 SHEETS</div>
          <h5 class="mb-0 fw-bold" style="color:#27ae60;">${dispA2.toLocaleString("en-US")}</h5>
        </div>
        <div class="summary-box" style="background:#fef9e7; border:2px solid #f39c12;">
          <div class="tiny-label" style="font-weight:700; color:#d68910;">INSIDE PACKS</div>
          <h5 class="mb-0 fw-bold" style="color:#d68910;">${d.insidePack.toLocaleString("en-US")}</h5>
        </div>
        <div class="summary-box" style="background:#fdedec; border:2px solid #e74c3c;">
          <div class="tiny-label" style="font-weight:700; color:#c0392b;">COVER PACKS</div>
          <h5 class="mb-0 fw-bold" style="color:#c0392b;">${d.coverPackNeed.toLocaleString("en-US")}</h5>
        </div>
        <div class="summary-box" style="background:#e8f4f8; border:2px solid #16a085;">
          <div class="tiny-label" style="font-weight:700; color:#16a085;">CTP PLATES</div>
          <h5 class="mb-0 fw-bold" style="color:#16a085;">${dispPlates.toLocaleString("en-US")}</h5>
        </div>
      </div>`;
  }

  // 3. Paper Pack Badges
  if (showMaterials) {
    html += `
      <div class="material-badges-row mb-3">
        <span class="badge bg-primary" style="background:#2980b9 !important;"><i class="fa-solid fa-layer-group"></i> WF: ${(d.wfSheets || 0).toLocaleString("en-US")} sh (${(d.wfPacks || 0).toLocaleString("en-US")} pk @${d.wfPackSize || 500})</span>
        <span class="badge bg-secondary" style="background:#7f8c8d !important;"><i class="fa-solid fa-layer-group"></i> MT: ${(d.mtSheets || 0).toLocaleString("en-US")} sh (${(d.mtPacks || 0).toLocaleString("en-US")} pk @${d.mtPackSize || 250})</span>
        <span class="badge bg-dark" style="background:#2c3e50 !important;"><i class="fa-solid fa-layer-group"></i> GL: ${(d.glSheets || 0).toLocaleString("en-US")} sh (${(d.glPacks || 0).toLocaleString("en-US")} pk @${d.glPackSize || 100})</span>
        <span class="badge bg-danger" style="background:#c0392b !important;"><i class="fa-solid fa-box"></i> Cover: ${(d.coverSheetsA3 || 0).toLocaleString("en-US")} A3 (${(d.coverPackNeed || 0).toLocaleString("en-US")} pk @${d.coverPackSize || 100})</span>
      </div>`;
  }

  // 4. Job Breakdown Table
  if (showBreakdown) {
    html += `
      <div class="card p-2 mb-3 border">
        <div class="d-flex justify-content-between align-items-center mb-2 flex-wrap gap-2">
          <h6 class="fw-bold mb-0 small"><i class="fa-solid fa-list-check text-primary"></i> Job Breakdown (1 Child = 8pg BOTH sides)</h6>
          <div class="d-flex gap-1 align-items-center">
            <span class="badge bg-primary">${d.formsPerDay || 10} Child/Day</span>
            <span class="badge bg-secondary">${dispDays} Work Days</span>
            <span class="badge bg-success">${d.finishDate}</span>
          </div>
        </div>
        <div class="table-responsive" style="overflow:visible;">
          <table class="table table-bordered table-sm text-center align-middle mb-0 small">
            <thead>
              <tr style="background:#f0f7ff; color:#1e3a8a;">
                <th style="width:30px;">#</th>
                <th class="text-start">Title</th>
                <th>Pg</th>
                <th style="color:#2980b9;">Qty</th>
                <th>Mat</th>
                <th>Bind</th>
                <th style="color:#8e44ad;">Child</th>
                <th style="color:#d35400;">Impress.</th>
                <th style="color:#27ae60;">A2 Sheets</th>
                <th style="color:#c0392b;">Plates</th>
                <th style="background:#2980b9; color:#fff;">Days</th>
              </tr>
            </thead>
            <tbody>`;
    if (booksToDisplay.length === 0) {
      html += `<tr><td colspan="11" class="text-muted py-2">No books selected</td></tr>`;
    } else {
      booksToDisplay.forEach(b => {
        html += `
          <tr>
            <td>${b.num}</td>
            <td class="text-start fw-semibold">${b.title}</td>
            <td>${b.pages}</td>
            <td style="color:#2980b9; font-weight:700;">${b.qty}</td>
            <td><span class="badge bg-light text-dark border">${b.material}</span></td>
            <td><span class="badge ${b.binding === 'Perfect' ? 'bg-primary' : 'bg-info'}">${b.binding}</span></td>
            <td style="color:#8e44ad; font-weight:700;">${b.forms}</td>
            <td style="color:#d35400;">${b.impressions}</td>
            <td style="color:#27ae60; font-weight:700;">${b.a2}</td>
            <td style="color:#c0392b;">${b.plates}</td>
            <td style="background:#eaf2f8; font-weight:700; color:#2980b9;">${b.days}</td>
          </tr>`;
      });
    }
    html += `
            </tbody>
            <tfoot>
              <tr style="background:#2c3e50; color:#fff; font-weight:700;">
                <td colspan="2" class="text-end">TOTAL</td>
                <td>—</td>
                <td>${dispQty.toLocaleString("en-US")}</td>
                <td>—</td>
                <td>—</td>
                <td>${dispForms.toLocaleString("en-US")}</td>
                <td>${dispImpressions.toLocaleString("en-US")}</td>
                <td style="color:#55efc4;">${dispA2.toLocaleString("en-US")}</td>
                <td>${dispPlates.toLocaleString("en-US")}</td>
                <td style="background:#2980b9;">${dispDays.toLocaleString("en-US")}</td>
              </tr>
            </tfoot>
          </table>
        </div>
      </div>`;
  }

  // 5. Consumables
  if (showConsumables) {
    html += `
      <div class="card p-2 mb-3 border">
        <h6 class="fw-bold mb-2 small"><i class="fa-solid fa-toolbox text-primary"></i> Consumables Requirement</h6>
        <div class="table-responsive" style="overflow:visible;">
          <table class="table table-bordered table-sm text-center align-middle mb-0 small">
            <thead>
              <tr style="background:#f0f7ff;">
                <th style="color:#e74c3c;">Ink (CMYK)</th>
                <th style="color:#8e44ad;">Alu (bottle)</th>
                <th style="color:#2980b9;">Powder</th>
                <th style="color:#27ae60;">Blanket</th>
                <th style="color:#d35400;">Glue</th>
                <th style="color:#f39c12;">Lam</th>
                <th style="color:#16a085;">Staple</th>
              </tr>
            </thead>
            <tbody>
              <tr>
                <td class="fw-bold fs-6">${(d.ink || 0).toLocaleString("en-US")}</td>
                <td class="fw-bold fs-6">${(d.alu || 0).toLocaleString("en-US")}</td>
                <td class="fw-bold fs-6">${(d.powder || 0).toLocaleString("en-US")}</td>
                <td class="fw-bold fs-6">${(d.blanket || 0).toLocaleString("en-US")}</td>
                <td class="fw-bold fs-6">${(d.totalGlue || 0).toLocaleString("en-US")}</td>
                <td class="fw-bold fs-6">${(d.totalLaminate || 0).toLocaleString("en-US")}</td>
                <td class="fw-bold fs-6">${(d.totalStaple || 0).toLocaleString("en-US")}</td>
              </tr>
            </tbody>
          </table>
        </div>
      </div>`;
  }

  // 6. Signatures
  if (showSignatures) {
    html += `
      <div class="print-signature-block">
        <div class="print-signature-box">Prepared By</div>
        <div class="print-signature-box">Production Mgr</div>
        <div class="print-signature-box">Press Operator</div>
        <div class="print-signature-box">QA / Delivery</div>
      </div>`;
  }

  $("#printArea").html(html);
}

// Print customized ticket
$("#doPrint").on("click", function () {
  const content = document.getElementById("printArea").innerHTML;
  const win = window.open("", "_blank");
  win.document.write(`<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>Production Job Ticket</title>
  <link href="https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;500;600;700;800&display=swap" rel="stylesheet">
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet">
  <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.2/css/all.min.css" rel="stylesheet">
  <style>
    @page {
      size: A4 portrait;
      margin: 6mm 8mm 6mm 8mm;
    }
    * {
      -webkit-print-color-adjust: exact !important;
      print-color-adjust: exact !important;
      color-adjust: exact !important;
      box-sizing: border-box;
    }
    body {
      font-family: 'Poppins', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
      background: #fff;
      color: #0f172a;
      margin: 0;
      padding: 4mm 6mm;
      font-size: 8.5px;
    }
    h4 { font-size: 14px; font-weight: 800; margin: 0; }
    h5 { font-size: 12px; font-weight: 800; margin: 0; }
    h6 { font-size: 9px; font-weight: 700; margin: 0 0 3px 0; }
    .prod-summary-grid {
      display: flex !important;
      flex-direction: row !important;
      flex-wrap: nowrap !important;
      gap: 4px !important;
      width: 100% !important;
      margin-bottom: 6px !important;
    }
    .prod-summary-grid > .summary-box {
      flex: 1 1 0 !important;
      width: 0 !important;
      min-width: 0 !important;
      padding: 4px 2px !important;
      border-radius: 5px !important;
      text-align: center !important;
      display: flex !important;
      flex-direction: column !important;
      justify-content: center !important;
      align-items: center !important;
      min-height: 48px !important;
      box-shadow: none !important;
    }
    .prod-summary-grid .tiny-label {
      font-size: 6.5px !important;
      font-weight: 700 !important;
      letter-spacing: 0.2px !important;
      margin-bottom: 2px !important;
      text-transform: uppercase !important;
      white-space: nowrap !important;
      overflow: hidden !important;
      text-overflow: ellipsis !important;
      max-width: 100% !important;
    }
    .prod-summary-grid h5 {
      font-size: 11.5px !important;
      font-weight: 800 !important;
      margin: 0 !important;
      line-height: 1.1 !important;
    }
    .material-badges-row {
      display: flex !important;
      flex-direction: row !important;
      flex-wrap: nowrap !important;
      justify-content: space-between !important;
      gap: 4px !important;
      width: 100% !important;
      margin-bottom: 6px !important;
    }
    .material-badges-row .badge {
      flex: 1 1 0 !important;
      font-size: 7.5px !important;
      padding: 3px 4px !important;
      text-align: center !important;
      border-radius: 4px !important;
      white-space: nowrap !important;
      border: 1px solid rgba(0,0,0,0.08) !important;
    }
    .card {
      border: 1px solid #cbd5e1 !important;
      border-radius: 5px !important;
      padding: 5px !important;
      margin-bottom: 6px !important;
      box-shadow: none !important;
      page-break-inside: avoid !important;
      break-inside: avoid !important;
    }
    .table {
      width: 100% !important;
      border-collapse: collapse !important;
      margin-bottom: 0 !important;
      font-size: 8px !important;
    }
    .table th {
      padding: 2.5px 2px !important;
      font-size: 7.5px !important;
      font-weight: 700 !important;
      vertical-align: middle !important;
      background-color: #f1f5f9 !important;
      border-color: #cbd5e1 !important;
      line-height: 1.1 !important;
    }
    .table td {
      padding: 2.5px 2px !important;
      font-size: 8px !important;
      vertical-align: middle !important;
      border-color: #cbd5e1 !important;
      line-height: 1.1 !important;
    }
    .badge {
      font-size: 7px !important;
      padding: 1.5px 4px !important;
    }
    .print-signature-block {
      display: flex !important;
      flex-direction: row !important;
      justify-content: space-between !important;
      gap: 12px !important;
      margin-top: 8px !important;
      padding-top: 4px !important;
      border-top: 1px dashed #cbd5e1 !important;
      font-size: 8px !important;
      page-break-inside: avoid !important;
      break-inside: avoid !important;
    }
    .print-signature-box {
      flex: 1 1 0 !important;
      text-align: center !important;
      border-top: 1px solid #334155 !important;
      margin-top: 22px !important;
      padding-top: 2px !important;
      font-weight: 600 !important;
      color: #334155 !important;
      font-size: 7.5px !important;
    }
    @media print {
      body { padding: 0 !important; margin: 0 !important; }
    }
  </style>
</head>
<body>
  ${content}
</body>
</html>`);
  win.document.close();
  setTimeout(() => {
    win.focus();
    win.print();
  }, 400);
});

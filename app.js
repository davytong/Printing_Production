// ====================================================
// PRINTING PRODUCTION SUITE - v2.0
// Multi-device responsive, multi-format import
// ====================================================

// ===== Shared Utilities =====
// Production is planned in half-day blocks, always rounded up.
const roundUpHalfDay = (days) => Math.ceil(days * 2) / 2;
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
    $(`#costBinding${i + 1}`).val(item.paper === "A3" ? "Staple" : "Perfect");
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

function costBookRow(i) {
  return `<div class="book-row row align-items-center g-1 g-md-2" id="costCard${i}">
    <div class="col-12 col-sm-6 col-md-3 mb-1"><input id="costTitle${i}" class="form-control form-control-sm calc-input" placeholder="Book Title" aria-label="Book Title"></div>
    <div class="col-6 col-sm-3 col-md-2 mb-1"><input id="costPages${i}" class="form-control form-control-sm calc-input" type="number" min="1" step="1" placeholder="Pages" aria-label="Pages" inputmode="numeric"></div>
    <div class="col-6 col-sm-3 col-md-2 mb-1"><input id="costQty${i}" class="form-control form-control-sm calc-input" type="number" min="1" step="1" placeholder="Qty" aria-label="Quantity" inputmode="numeric"></div>
    <div class="col-4 col-sm-3 col-md-1 mb-1"><select id="costPaper${i}" class="form-select form-select-sm calc-input" aria-label="Paper Size"><option>A4</option><option>A3</option></select></div>
    <div class="col-3 col-sm-3 col-md-1 mb-1"><select id="costMaterial${i}" class="form-select form-select-sm calc-input" aria-label="Material"><option>WF</option><option>MT</option></select></div>
    <div class="col-3 col-sm-4 col-md-2 mb-1"><select id="costBinding${i}" class="form-select form-select-sm calc-input" aria-label="Binding"><option>Perfect</option><option>Staple</option></select></div>
    <div class="col-2 col-sm-2 col-md-1 mb-1 text-end"><button class="btn btn-outline-danger btn-sm removeCostRow" title="Remove" aria-label="Remove book"><i class="fa fa-trash"></i></button></div>
  </div>`;
}

function createBookTable(n) {
  n = Math.max(1, parseInt(n) || 1);
  $("#numBooks").val(n);
  $("#bookTableContainer").html(Array.from({length:n}, (_, i) => costBookRow(i + 1)).join(""));
}

$("#addCostRow").on("click", function () {
  const n = document.querySelectorAll("#bookTableContainer .book-row").length + 1;
  $("#bookTableContainer").append(costBookRow(n));
  $("#numBooks").val(n);
});
$("#bookTableContainer").on("click", ".removeCostRow", function () {
  const rows = document.querySelectorAll("#bookTableContainer .book-row");
  if (rows.length === 1) {
    Swal.fire("Info", "At least one book row is required", "info");
    return;
  }
  $(this).closest(".book-row").remove();
  document.querySelectorAll("#bookTableContainer .book-row").forEach((row, i) => {
    row.id = `costCard${i + 1}`;
    row.querySelectorAll("input, select").forEach(input => { input.id = input.id.replace(/\d+$/, String(i + 1)); });
  });
  $("#numBooks").val(rows.length - 1);
  calculateCostBooks(false);
});
createBookTable(1);

function calculateCostBooks(saveHistory = false) {
  const extraInput = document.getElementById("costExtraDays");
  const extraDays = Number(extraInput.value || 0);
  const validExtra = Number.isFinite(extraDays) && extraDays >= 0 && Number.isInteger(extraDays * 2);
  extraInput.setCustomValidity(validExtra ? "" : "Use half-day steps: 0, 0.5, 1, 1.5...");
  if (!validExtra) { extraInput.reportValidity(); $("#costPrint, #btnExportExcel, #resultArea").hide(); window._costData = null; return; }
  const n = parseInt($("#numBooks").val());
  if (!n || n < 1) { $("#resultArea").hide(); $("#btnExportExcel").hide(); return; }

  $("#resultBody,#packBody,#summaryCards").html("");

  let totalBooksQty = 0, totalCost = 0, totalCoverSheets = 0, totalDigitalSheets = 0;
  let sheetsWF = 0, sheetsMT = 0, sheetsA3 = 0;
  let bookRows = [];

  for (let i = 1; i <= n; i++) {
    const title = $(`#costTitle${i}`).val() || `Book ${i}`;
    const pages = parseInt($(`#costPages${i}`).val()) || 0;
    const qty = parseInt($(`#costQty${i}`).val()) || 0;
    const paper = $(`#costPaper${i}`).val();
    const material = $(`#costMaterial${i}`).val();
    const binding = $(`#costBinding${i}`).val() || (paper === "A4" ? "Perfect" : "Staple");

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
    totalDigitalSheets += totalSheetsBook; // Count physical sheets for either paper size.

    if (paper === "A4" && material === "WF") sheetsWF += totalSheetsBook;
    if (paper === "A4" && material === "MT") sheetsMT += totalSheetsBook;
    if (paper === "A3") sheetsA3 += totalSheetsBook;

    bookRows.push({ title, pages, qty, paper, material, binding, pricePerBook, sheets: totalSheetsBook, idx: i });
  }

  bookRows.forEach((b, i) => {
    $(`#costCard${b.idx}`).toggleClass("expensive", b.pricePerBook > 50);
    $("#resultBody").append(`<tr>
      <td>${i + 1}</td><td class="text-start">${sanitize(b.title)}</td><td>${fmtInt(b.pages)}</td><td>${fmtInt(b.qty)}</td>
      <td>${b.paper}</td><td>${b.material}</td><td>${b.binding}</td>
      <td style="color:#27ae60;font-weight:700">${fmtInt(b.sheets)}</td><td>${fmtPrice(b.pricePerBook)}</td><td style="background:#eaf6ff;color:#2980b9;font-weight:700">${roundUpHalfDay(b.sheets / getCostCapacity())}</td>
    </tr>`);
  });

  const packsWF = sheetsWF > 0 ? Math.ceil(sheetsWF / COST.PACK_WF_A4) : 0;
  const packsMT = sheetsMT > 0 ? Math.ceil(sheetsMT / COST.PACK_MT_A4) : 0;
  const packsA3 = sheetsA3 > 0 ? Math.ceil(sheetsA3 / COST.PACK_A3) : 0;
  const coverPacks = totalCoverSheets > 0 ? Math.ceil(totalCoverSheets / COST.COVER_PACK) : 0;

  const capacity = getCostCapacity();
  const totalDays = totalDigitalSheets > 0 && capacity > 0 ? roundUpHalfDay(totalDigitalSheets / capacity) + extraDays : 0;
  $("#costTotalDays").text(totalDays);

  const paperGroups = digitalPaperGroups(bookRows);
  $("#packBody").html(paperGroups.filter(g => g.sheets > 0).map(g => `<tr><td>${g.label}</td><td>${fmtInt(g.sheets)}</td><td>${fmtInt(g.packs)}</td></tr>`).join(""));

  // Use the same summary boxes as offset production.
  const cards = [
    { title: "Books", value: bookRows.length, bg: "#f0f0f0", color: "#7f8c8d" },
    { title: "Total Qty", value: fmtInt(totalBooksQty), bg: "#eaf6ff", color: "#2980b9" },
    { title: "Inside Sheets", value: fmtInt(totalDigitalSheets), bg: "#e8fdf5", color: "#27ae60" },
    { title: "Inside Packs", value: fmtInt(paperGroups.filter(g => g.label !== "Cover").reduce((n,g) => n + g.packs, 0)), bg: "#fef9e7", color: "#f39c12" },
    { title: "Days", value: fmtInt(totalDays), bg: "#f5eef8", color: "#8e44ad" },
    { title: "Total Cost", value: fmtPrice(totalCost), bg: "#fdedec", color: "#e74c3c" },
    { title: "Cover Packs", value: fmtInt(coverPacks), bg: "#fef9e7", color: "#f39c12" }
  ];
  $("#summaryCards").html(cards.map(c => `<div class="col-4 col-md-3 col-lg"><div class="summary-box" style="background:${c.bg}; border:2px solid ${c.color};"><div class="tiny-label text-uppercase">${c.title}</div><h5 class="mb-0 fw-bold" style="color:${c.color};">${c.value}</h5></div></div>`).join(""));

  const startDate = document.getElementById("costStartDate").value;
  const holidays = document.getElementById("costHolidays").value.split(/[,\s]+/).filter(Boolean);
  const schedule = calcFinishDate(totalDays, startDate, holidays);
  window._costData = {books: bookRows, capacity, extraDays, startDate, holidays, totalDays, finishDate: schedule.finishDate};
  $("#costSpeedBadge").text(`${fmtInt(capacity)} Sheets/Day`);
  $("#costFinishDate").text(`${startDate} → ${schedule.finishDate}`);
  $("#costMaterialBreakdown").html(digitalPaperBadges(bookRows));
  $("#costDetailFoot").html(`<tr class="table-light fw-bold"><td colspan="3">TOTAL</td><td>${fmtInt(totalBooksQty)}</td><td colspan="3"></td><td>${fmtInt(totalDigitalSheets)}</td><td>${fmtPrice(totalCost)}</td><td>${totalDays}</td></tr>`);
  $("#resultArea").show();
  $("#btnExportExcel, #costPrint").toggle(bookRows.length > 0);

  // Store bookRows globally for export
  window._costBookRows = bookRows;
  window._costTotalCost = totalCost;

  if (saveHistory) saveCostHistory(n, totalBooksQty, totalCost, totalDays, bookRows);
}

function saveCostHistory(n, totalBooksQty, totalCost, totalDays, bookRows) {
  const hash = btoa(bookRows.map((b) => `${b.title}-${b.pages}-${b.qty}-${b.paper}-${b.material}-${b.binding}`).join("|"));
  if (hash === lastJobHash) return;
  lastJobHash = hash;

  const entry = {
    timestamp: new Date().toLocaleString("en-US"),
    books: n,
    totalBooksQty,
    totalCost,
    totalDays,
    bookRows: bookRows.map((b) => ({ title: b.title, pages: b.pages, qty: b.qty, paper: b.paper, material: b.material, binding: b.binding }))
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
    $(`#costBinding${i + 1}`).val(b.binding || (b.paper === "A3" ? "Staple" : "Perfect"));
  });
  calculateCostBooks(false);
  $("#historyArea").hide();
});

$(document).on("input change", ".calc-input, #costDailyCapacity, #costDailySlider, #costExtraDays, #costStartDate, #costHolidays", function () {
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
  data.push(["Work Days", window._costData.totalDays]);
  data.push(["Start Date", window._costData.startDate]);
  data.push(["Finish Date", window._costData.finishDate]);
  data.push(["#", "Title", "Pages", "Qty", "Paper", "Material", "Binding", "Price/Book", "Sheets", "Days"]);
  bookRows.forEach((b, i) => {
    data.push([i + 1, b.title, b.pages, b.qty, b.paper, b.material, b.binding, Number(b.pricePerBook.toFixed(2)), b.sheets, roundUpHalfDay(b.sheets / window._costData.capacity)]);
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
function calcFinishDate(workDays, startStr = document.getElementById("prodStartDate").value, holidays = prodHolidays) {
  const start = startStr ? new Date(startStr) : new Date();
  let current = new Date(start);
  let counted = 0;

  while (counted < workDays) {
    current.setDate(current.getDate() + 1);
    const day = current.getDay(); // 0=Sunday
    const dateStr = current.toISOString().split("T")[0];
    // Skip Sunday and holidays
    if (day === 0) continue;
    if (holidays.includes(dateStr)) continue;
    counted++;
  }

  return {
    finishDate: current.toISOString().split("T")[0],
    calendarDays: Math.round((current - start) / 86400000)
  };
}

// --- Production Calculate ---
function runProductionCalc() {
  const extraDaysInput = document.getElementById("prodExtraDays");
  const extraDays = extraDaysInput.value === "" ? 0 : Number(extraDaysInput.value);
  if (!Number.isFinite(extraDays) || extraDays < 0 || !Number.isInteger(extraDays * 2)) {
    extraDaysInput.setCustomValidity("Use half-day steps: 0, 0.5, 1, 1.5, 2, 2.5...");
    extraDaysInput.reportValidity();
    $("#prodPrint, #prodExportExcel").hide();
    window._prodData = null;
    return;
  }
  extraDaysInput.setCustomValidity("");
  const waste = getVal("prodWaste", 0) / 100;
  const formsPerDay = getVal("prodFormsPerDay", 10); // forms that can be printed per day (both sides)
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
    const originalPages = pages;
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

    }
    if (binding === "Staple" && pages % 4 !== 0) {
      const adjusted = Math.ceil(pages / 4) * 4;
      notices.push(`[${title}]: ${pages}pg → ${adjusted}pg (+${adjusted - pages} blank pages for Staple Binding)`);
      pages = adjusted;

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

    // CTP Plates: 1 Child = 8 plates (4/4 both sides A2)
    const plateCount = forms * 8;

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

    // Days calculation
    const daysForBook = roundUpHalfDay(forms / formsPerDay);

    totalForms += forms;
    totalQty += qty;

    processedCount++;
    const tr = document.createElement("tr");
    tr.dataset.originalPages = originalPages;
    tr.dataset.finalPages = pages;
    tr.dataset.addedPages = pages - originalPages;
    tr.innerHTML = `
      <td>${processedCount}</td>
      <td class="text-start text-truncate" style="max-width:180px"><strong>${title}</strong></td>
      <td>${pages.toLocaleString("en-US")}${pages > originalPages ? `<br><small class="text-warning-emphasis">${originalPages} original + ${pages - originalPages} blank</small>` : ""}</td>
      <td style="color:#2980b9; font-weight:600;">${qty.toLocaleString("en-US")}</td>
      <td><span class="badge bg-secondary">${material}</span></td>
      <td><span class="badge ${binding === "Perfect" ? "bg-primary" : "bg-info"}">${binding}</span></td>
      <td style="color:#8e44ad;">${forms.toLocaleString("en-US")}</td>
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

  $("#prodPageAdjustments").html(notices.length
    ? `<strong>Blank pages added for binding:</strong><br>${notices.join("<br>")}<br><small>Original page inputs are kept. Final pages and added blanks appear on the printed job ticket.</small>`
    : "").toggle(notices.length > 0);

  if (processedCount === 0) {
    detailTable.innerHTML = '<tr><td colspan="10" class="text-muted">Please enter book data above</td></tr>';
    $("#prodPrint").hide();
    $("#prodDetailFoot").hide();
    return;
  }

  // Show footer totals
  totalDaysAll = roundUpHalfDay(totalForms / formsPerDay);
  
  // Add extra days (buffer/maintenance/other)
  totalDaysAll += extraDays;
  
  // Calculate finish date (skip Sundays + holidays)
  const schedule = calcFinishDate(totalDaysAll);
  
  $("#prodDetailFoot").show();
  $("#prodTotalQty2").text(totalQty.toLocaleString("en-US"));
  $("#prodTotalForms").text(totalForms.toLocaleString("en-US"));
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
    totalForms, totalInsideA2, totalCTP, totalDaysAll,
    insidePack, coverPackNeed, ink, alu, powder, blanket,
    totalGlue, totalLaminate, totalStaple,
    materialSheets, wfPacks, mtPacks, glPacks,
    wfPackCap, mtPackCap, glPackCap, coverPack,
    formsPerDay, extraDays,
    wfSheets: materialSheets.WF, mtSheets: materialSheets.MT, glSheets: materialSheets.GL,
    wfPackSize: wfPackCap, mtPackSize: mtPackCap, glPackSize: glPackCap,
    coverSheetsA3: coverA3WithWaste, coverPackSize: coverPack,
    startDate,
    finishDate: schedule.finishDate, calendarDays: schedule.calendarDays,
    holidays: prodHolidays.length
  };
  // Collect book details from table
  detailTable.querySelectorAll("tr").forEach((tr) => {
    const cells = tr.querySelectorAll("td");
    if (cells.length >= 10) {
      window._prodData.books.push({
        num: cells[0].textContent,
        title: cells[1].textContent,
        pages: Number(tr.dataset.finalPages).toLocaleString("en-US"),
        originalPages: Number(tr.dataset.originalPages),
        finalPages: Number(tr.dataset.finalPages),
        addedPages: Number(tr.dataset.addedPages),
        qty: cells[3].textContent,
        material: cells[4].textContent,
        binding: cells[5].textContent,
        forms: cells[6].textContent,
        a2: cells[7].textContent,
        plates: cells[8].textContent,
        days: cells[9].textContent,
        rawQty: parseInt(cells[3].textContent.replace(/,/g, "")) || 0,
        rawForms: parseInt(cells[6].textContent.replace(/,/g, "")) || 0,
        rawA2: parseInt(cells[7].textContent.replace(/,/g, "")) || 0,
        rawPlates: parseInt(cells[8].textContent.replace(/,/g, "")) || 0,
        rawDays: parseFloat(cells[9].textContent.replace(/,/g, "")) || 0
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
$(document).on("input change keyup", "#bookContainer input, #bookContainer select, #prodSettings input, #prodFormsPerDay, #prodWaste, #prodBlanketLife, #prodWfPack, #prodMtPack, #prodGlPack, #prodCoverPack, #prodInkCap, #prodAluCap, #prodPowderCap, #prodGlueCap, #prodLamCap, #prodStapleCap, #prodExtraDays, #prodStartDate", function () {
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
  data.push(["#", "Title", "Pages", "Qty", "Material", "Binding", "Child", "A2 Sheets", "Plates", "Days", "Original Pages", "Blank Pages Added"]);
  d.books.forEach((b) => {
    data.push([b.num, b.title, b.pages, b.qty, b.material, b.binding, b.forms, b.a2, b.plates, b.days, b.originalPages, b.addedPages]);
  });
  data.push(["", "TOTAL", "", d.totalQty, "", "", d.totalForms, d.totalInsideA2, d.totalCTP, d.totalDaysAll]);
  data.push([]);

  // Consumables
  data.push(["CONSUMABLES"]);
  data.push(["Ink", "Alu", "Powder", "Blanket", "Glue", "Laminate", "Staples"]);
  data.push([d.ink, d.alu, d.powder, d.blanket, d.totalGlue, d.totalLaminate, d.totalStaple]);

  const ws = XLSX.utils.aoa_to_sheet(data);
  ws["!cols"] = [
    { wch: 5 }, { wch: 22 }, { wch: 10 }, { wch: 10 }, { wch: 8 }, { wch: 10 },
    { wch: 8 }, { wch: 12 }, { wch: 8 }, { wch: 6 }
  ];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Production");
  XLSX.writeFile(wb, "production_" + new Date().toISOString().split("T")[0] + ".xlsx");
});

// ====================================================
// PRODUCTION: CUSTOM PRINT & SELECTION
// ====================================================
$("#prodPrint, #costPrint").on("click", function () {
  const digital = this.id === "costPrint";
  window._printMachine = digital ? "digital" : "offset";
  const d = digital ? window._costData : window._prodData;
  $("#optConsumables").closest(".col-6").toggle(!digital);
  $("label[for='optSummary']").text("Summary Cards");
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
        <label class="form-check-label small" for="printBook_${i}">${i + 1}. ${sanitize(b.title) || "Book " + (i + 1)} <span class="badge bg-light text-dark border ms-1">${digital ? fmtInt(b.sheets) + " Sheets" : b.forms + " Child"}</span></label>
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

// Rebuild preview when section checkboxes or print scale change
$(document).on("change", ".print-opt, .print-book-opt, #printScaleSelect", function () {
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
  if (window._printMachine === "digital") { buildDigitalPrintPreview(); return; }
  const d = window._prodData;
  if (!d) return;

  const scale = $("#printScaleSelect").val() || "xl";
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
  let dispPlates = 0;
  let dispDays = 0;

  if (isSubset) {
    booksToDisplay.forEach(b => {
      dispQty += b.rawQty || 0;
      dispForms += b.rawForms || 0;
      dispA2 += b.rawA2 || 0;
      dispPlates += b.rawPlates || 0;
      dispDays += b.rawDays || 0;
    });
  } else {
    dispBooks = d.totalBooks;
    dispQty = d.totalQty;
    dispForms = d.totalForms;
    dispA2 = d.totalInsideA2;
    dispPlates = d.totalCTP;
    dispDays = d.totalDaysAll;
  }

  let html = `<div class="print-container print-scale-${scale}">`;

  // 1. Header
  if (showHeader) {
    html += `
      <div class="print-header mb-3 pb-2 border-bottom">
        <div class="d-flex justify-content-between align-items-center mb-1">
          <div class="text-start">
            <h4 class="fw-bold mb-1 text-dark"><i class="fa-solid fa-industry text-primary me-1"></i> Production Job Ticket</h4>
            <div class="d-flex gap-2 mt-1">
              <span class="badge bg-dark px-3 py-1">${dispBooks} Job${dispBooks > 1 ? "s" : ""}</span>
              <span class="badge bg-primary px-3 py-1">${dispQty.toLocaleString("en-US")} Total Qty</span>
              <span class="badge bg-purple px-3 py-1" style="background:#8e44ad; color:#fff;">${dispForms.toLocaleString("en-US")} Children</span>
            </div>
          </div>
          <div class="text-end" style="font-size: 13.5px; color:#475569;">
            <div><b>Start Date:</b> <span class="badge bg-primary px-2 py-1" style="font-size: 13px;">${d.startDate || document.getElementById("prodStartDate")?.value || new Date().toISOString().split("T")[0]}</span></div>
            <div class="mt-1"><b>Target Finish:</b> <span class="badge bg-success px-2 py-1" style="font-size: 13px;">${d.finishDate}</span></div>
          </div>
        </div>
        <div class="meta-bar d-flex justify-content-center gap-4 text-muted flex-wrap mt-2 pt-2 border-top" style="font-size: 13px;">
          <span><i class="fa-solid fa-gauge-high text-primary"></i> <b>Speed:</b> ${d.formsPerDay || 10} Child/Day</span>
          <span><i class="fa-solid fa-calendar-day text-secondary"></i> <b>Work Days:</b> ${dispDays} Days</span>
          <span><i class="fa-solid fa-play text-primary"></i> <b>Start:</b> ${d.startDate || document.getElementById("prodStartDate")?.value || new Date().toISOString().split("T")[0]}</span>
          <span><i class="fa-solid fa-flag-checkered text-success"></i> <b>Finish:</b> ${d.finishDate}</span>
        </div>
      </div>`;
  }

  // 2. 7 Summary Cards
  if (showSummary) {
    html += `
      <div class="prod-summary-grid mb-3">
        <div class="summary-box" style="background:#f8fafc; border:2.5px solid #64748b;">
          <div class="tiny-label" style="font-weight:800; color:#475569;">BOOKS</div>
          <h5 class="mb-0 fw-bold" style="color:#1e293b;">${dispBooks.toLocaleString("en-US")}</h5>
        </div>
        <div class="summary-box" style="background:#eff6ff; border:2.5px solid #3b82f6;">
          <div class="tiny-label" style="font-weight:800; color:#1d4ed8;">TOTAL QTY</div>
          <h5 class="mb-0 fw-bold" style="color:#1e40af;">${dispQty.toLocaleString("en-US")}</h5>
        </div>
        <div class="summary-box" style="background:#faf5ff; border:2.5px solid #a855f7;">
          <div class="tiny-label" style="font-weight:800; color:#7e22ce;">TOTAL CHILDREN</div>
          <h5 class="mb-0 fw-bold" style="color:#6b21a8;">${dispForms.toLocaleString("en-US")}</h5>
        </div>
        <div class="summary-box" style="background:#f0fdf4; border:2.5px solid #22c55e;">
          <div class="tiny-label" style="font-weight:800; color:#15803d;">A2 SHEETS</div>
          <h5 class="mb-0 fw-bold" style="color:#166534;">${dispA2.toLocaleString("en-US")}</h5>
        </div>
        <div class="summary-box" style="background:#fffbeb; border:2.5px solid #f59e0b;">
          <div class="tiny-label" style="font-weight:800; color:#b45309;">INSIDE PACKS</div>
          <h5 class="mb-0 fw-bold" style="color:#92400e;">${d.insidePack.toLocaleString("en-US")}</h5>
        </div>
        <div class="summary-box" style="background:#fef2f2; border:2.5px solid #ef4444;">
          <div class="tiny-label" style="font-weight:800; color:#b91c1c;">COVER PACKS</div>
          <h5 class="mb-0 fw-bold" style="color:#991b1b;">${d.coverPackNeed.toLocaleString("en-US")}</h5>
        </div>
        <div class="summary-box" style="background:#f0fdfa; border:2.5px solid #14b8a6;">
          <div class="tiny-label" style="font-weight:800; color:#0f766e;">CTP PLATES</div>
          <h5 class="mb-0 fw-bold" style="color:#115e59;">${dispPlates.toLocaleString("en-US")}</h5>
        </div>
      </div>`;
  }

  // 3. Paper Pack Badges
  if (showMaterials) {
    html += `
      <div class="material-badges-row mb-3">
        <span class="badge" style="background:#2563eb !important; color:#fff;"><i class="fa-solid fa-layer-group"></i> WF: ${(d.wfSheets || 0).toLocaleString("en-US")} sh (${(d.wfPacks || 0).toLocaleString("en-US")} pk @${d.wfPackSize || 500})</span>
        <span class="badge" style="background:#64748b !important; color:#fff;"><i class="fa-solid fa-layer-group"></i> MT: ${(d.mtSheets || 0).toLocaleString("en-US")} sh (${(d.mtPacks || 0).toLocaleString("en-US")} pk @${d.mtPackSize || 250})</span>
        <span class="badge" style="background:#334155 !important; color:#fff;"><i class="fa-solid fa-layer-group"></i> GL: ${(d.glSheets || 0).toLocaleString("en-US")} sh (${(d.glPacks || 0).toLocaleString("en-US")} pk @${d.glPackSize || 100})</span>
        <span class="badge" style="background:#dc2626 !important; color:#fff;"><i class="fa-solid fa-box"></i> Cover: ${(d.coverSheetsA3 || 0).toLocaleString("en-US")} A3 (${(d.coverPackNeed || 0).toLocaleString("en-US")} pk @${d.coverPackSize || 100})</span>
      </div>`;
  }

  // 4. Job Breakdown Table
  if (showBreakdown) {
    html += `
      <div class="card p-3 mb-3 border">
        <div class="d-flex justify-content-between align-items-center mb-2 flex-wrap gap-2">
          <h6 class="fw-bold mb-0"><i class="fa-solid fa-list-check text-primary me-1"></i> Job Breakdown <span class="text-muted fw-normal" style="font-size:12px;">(1 Child = 8pg BOTH sides)</span></h6>
          <div class="d-flex gap-2 align-items-center">
            <span class="badge bg-primary">${d.formsPerDay || 10} Child/Day</span>
            <span class="badge bg-secondary">${dispDays} Work Days</span>
            <span class="badge bg-success">${d.startDate || document.getElementById("prodStartDate")?.value || ""} → ${d.finishDate}</span>
          </div>
        </div>
        <div class="table-responsive" style="overflow:visible;">
          <table class="table table-bordered text-center align-middle mb-0">
            <thead>
              <tr style="background:#f1f5f9; color:#1e293b;">
                <th style="width:40px;">#</th>
                <th class="text-start">Title</th>
                <th>Final Pg</th>
                <th style="color:#2563eb;">Qty</th>
                <th>Mat</th>
                <th>Bind</th>
                <th style="color:#7e22ce;">Child</th>
                <th style="color:#16a34a;">A2 Sheets</th>
                <th style="color:#dc2626;">Plates</th>
                <th style="background:#2563eb; color:#fff;">Days</th>
              </tr>
            </thead>
            <tbody>`;
    if (booksToDisplay.length === 0) {
      html += `<tr><td colspan="10" class="text-muted py-2">No books selected</td></tr>`;
    } else {
      booksToDisplay.forEach(b => {
        html += `
          <tr>
            <td>${b.num}</td>
            <td class="text-start fw-semibold">${b.title}</td>
            <td>${b.pages}${b.addedPages ? `<br><small>${b.originalPages} original + <b>${b.addedPages} blank</b></small>` : ""}</td>
            <td style="color:#2563eb; font-weight:700;">${b.qty}</td>
            <td><span class="badge bg-light text-dark border">${b.material}</span></td>
            <td><span class="badge ${b.binding === 'Perfect' ? 'bg-primary' : 'bg-info'}">${b.binding}</span></td>
            <td style="color:#7e22ce; font-weight:700;">${b.forms}</td>
            <td style="color:#16a34a; font-weight:700;">${b.a2}</td>
            <td style="color:#dc2626;">${b.plates}</td>
            <td style="background:#eff6ff; font-weight:700; color:#2563eb;">${b.days}</td>
          </tr>`;
      });
    }
    html += `
            </tbody>
            <tfoot>
              <tr style="background:#1e293b; color:#fff; font-weight:700;">
                <td colspan="2" class="text-end">TOTAL</td>
                <td>—</td>
                <td>${dispQty.toLocaleString("en-US")}</td>
                <td>—</td>
                <td>—</td>
                <td>${dispForms.toLocaleString("en-US")}</td>
                <td style="color:#4ade80;">${dispA2.toLocaleString("en-US")}</td>
                <td>${dispPlates.toLocaleString("en-US")}</td>
                <td style="background:#2563eb;">${dispDays.toLocaleString("en-US")}</td>
              </tr>
            </tfoot>
          </table>
        </div>
      </div>`;
  }

  const adjustedBooks = booksToDisplay.filter(b => b.addedPages > 0);
  if (adjustedBooks.length) {
    html += `<div class="card p-3 mb-3 border"><h6 class="fw-bold">Blank Pages Added for Binding</h6><p class="small">Final pages include blank pages required by the selected binding. Confirm their placement in the artwork before printing.</p><table class="table table-bordered small"><thead><tr><th>Book</th><th>Original</th><th>Added blanks</th><th>Final</th></tr></thead><tbody>${adjustedBooks.map(b => `<tr><td>${b.title}</td><td>${b.originalPages}</td><td>+${b.addedPages}</td><td>${b.finalPages}</td></tr>`).join("")}</tbody></table></div>`;
  }

  // 5. Consumables
  if (showConsumables) {
    html += `
      <div class="card p-3 mb-3 border">
        <h6 class="fw-bold mb-2"><i class="fa-solid fa-toolbox text-primary me-1"></i> Consumables Requirement</h6>
        <div class="table-responsive" style="overflow:visible;">
          <table class="table table-bordered text-center align-middle mb-0 consumables-table">
            <thead>
              <tr style="background:#f1f5f9; color:#1e293b;">
                <th style="color:#dc2626;">Ink (CMYK)</th>
                <th style="color:#7e22ce;">Alu (bottle)</th>
                <th style="color:#2563eb;">Powder</th>
                <th style="color:#16a34a;">Blanket</th>
                <th style="color:#ea580c;">Glue</th>
                <th style="color:#d97706;">Lam</th>
                <th style="color:#0d9488;">Staple</th>
              </tr>
            </thead>
            <tbody>
              <tr>
                <td class="fw-bold" style="color:#dc2626;">${(d.ink || 0).toLocaleString("en-US")}</td>
                <td class="fw-bold" style="color:#7e22ce;">${(d.alu || 0).toLocaleString("en-US")}</td>
                <td class="fw-bold" style="color:#2563eb;">${(d.powder || 0).toLocaleString("en-US")}</td>
                <td class="fw-bold" style="color:#16a34a;">${(d.blanket || 0).toLocaleString("en-US")}</td>
                <td class="fw-bold" style="color:#ea580c;">${(d.totalGlue || 0).toLocaleString("en-US")}</td>
                <td class="fw-bold" style="color:#d97706;">${(d.totalLaminate || 0).toLocaleString("en-US")}</td>
                <td class="fw-bold" style="color:#0d9488;">${(d.totalStaple || 0).toLocaleString("en-US")}</td>
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

  html += `</div>`; // end print-container

  $("#printArea").html(html);
}

// Print customized ticket
$("#doPrint").on("click", function () {
  const content = document.getElementById("printArea").innerHTML;
  const scale = $("#printScaleSelect").val() || "xl";
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
      margin: 10mm 12mm 10mm 12mm;
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
      padding: 6mm 8mm;
      font-size: 13.5px;
      line-height: 1.45;
    }
    .print-header {
      margin-bottom: 16px !important;
      padding-bottom: 10px !important;
      border-bottom: 2px solid #e2e8f0 !important;
    }
    .print-header h4 {
      font-size: 26px !important;
      font-weight: 800 !important;
      margin: 0 0 6px 0 !important;
      color: #0f172a !important;
    }
    .print-header .badge {
      font-size: 13px !important;
      padding: 5px 12px !important;
      font-weight: 600 !important;
    }
    .print-header .meta-bar {
      font-size: 13px !important;
      padding: 8px 16px !important;
      background: #f8fafc !important;
      border: 1px solid #e2e8f0 !important;
      border-radius: 6px !important;
      margin-top: 8px !important;
    }
    .prod-summary-grid {
      display: flex !important;
      flex-direction: row !important;
      flex-wrap: nowrap !important;
      gap: 8px !important;
      width: 100% !important;
      margin-bottom: 16px !important;
    }
    .prod-summary-grid > .summary-box {
      flex: 1 1 0 !important;
      width: 0 !important;
      min-width: 0 !important;
      padding: 10px 4px !important;
      border-radius: 8px !important;
      text-align: center !important;
      display: flex !important;
      flex-direction: column !important;
      justify-content: center !important;
      align-items: center !important;
      min-height: 84px !important;
      box-shadow: none !important;
    }
    .prod-summary-grid .tiny-label {
      font-size: 10.5px !important;
      font-weight: 800 !important;
      letter-spacing: 0.2px !important;
      margin-bottom: 4px !important;
      text-transform: uppercase !important;
      white-space: normal !important;
      word-break: normal !important;
      overflow: visible !important;
      text-overflow: clip !important;
      max-width: 100% !important;
      line-height: 1.15 !important;
      min-height: 26px !important;
      display: flex !important;
      align-items: center !important;
      justify-content: center !important;
      text-align: center !important;
    }
    .prod-summary-grid h5 {
      font-size: 24px !important;
      font-weight: 800 !important;
      margin: 0 !important;
      line-height: 1.1 !important;
    }
    .material-badges-row {
      display: flex !important;
      flex-direction: row !important;
      flex-wrap: nowrap !important;
      justify-content: space-between !important;
      gap: 8px !important;
      width: 100% !important;
      margin-bottom: 16px !important;
    }
    .material-badges-row .badge {
      flex: 1 1 0 !important;
      font-size: 12.5px !important;
      padding: 9px 12px !important;
      text-align: center !important;
      border-radius: 8px !important;
      white-space: nowrap !important;
      font-weight: 600 !important;
      border: 1px solid rgba(0,0,0,0.08) !important;
    }
    .card {
      border: 2px solid #cbd5e1 !important;
      border-radius: 10px !important;
      padding: 16px 18px !important;
      margin-bottom: 18px !important;
      box-shadow: none !important;
      page-break-inside: avoid !important;
      break-inside: avoid !important;
    }
    .card h6 {
      font-size: 15px !important;
      font-weight: 700 !important;
      margin: 0 !important;
      color: #0f172a !important;
    }
    .card .badge {
      font-size: 12px !important;
      padding: 4px 10px !important;
      font-weight: 600 !important;
    }
    .table {
      width: 100% !important;
      border-collapse: collapse !important;
      margin-bottom: 0 !important;
      font-size: 13.5px !important;
    }
    .table th {
      padding: 11px 8px !important;
      font-size: 13px !important;
      font-weight: 700 !important;
      vertical-align: middle !important;
      background-color: #f1f5f9 !important;
      border-color: #cbd5e1 !important;
      color: #1e293b !important;
      line-height: 1.2 !important;
    }
    .table td {
      padding: 11px 8px !important;
      font-size: 13.5px !important;
      vertical-align: middle !important;
      border-color: #cbd5e1 !important;
      color: #0f172a !important;
      line-height: 1.35 !important;
    }
    .table td .badge {
      font-size: 12px !important;
      padding: 4px 10px !important;
    }
    .table tfoot tr td {
      font-size: 14px !important;
      padding: 11px 8px !important;
      font-weight: 800 !important;
    }
    .consumables-table th {
      font-size: 13px !important;
      padding: 10px 8px !important;
    }
    .consumables-table td {
      font-size: 22px !important;
      padding: 14px 8px !important;
      font-weight: 800 !important;
    }
    .print-signature-block {
      display: flex !important;
      flex-direction: row !important;
      justify-content: space-between !important;
      gap: 20px !important;
      margin-top: 35px !important;
      padding-top: 16px !important;
      border-top: 2px dashed #94a3b8 !important;
      font-size: 13.5px !important;
      page-break-inside: avoid !important;
      break-inside: avoid !important;
    }
    .print-signature-box {
      flex: 1 1 0 !important;
      text-align: center !important;
      border-top: 2px solid #1e293b !important;
      margin-top: 75px !important;
      padding-top: 8px !important;
      font-weight: 700 !important;
      color: #0f172a !important;
      font-size: 13.5px !important;
    }

    /* Scaling modifiers */
    body.print-scale-large { font-size: 12.5px; }
    .print-scale-large .print-header h4 { font-size: 22px !important; }
    .print-scale-large .prod-summary-grid > .summary-box { min-height: 74px !important; padding: 8px 4px !important; }
    .print-scale-large .prod-summary-grid .tiny-label { font-size: 9.5px !important; min-height: 23px !important; }
    .print-scale-large .prod-summary-grid h5 { font-size: 20px !important; }
    .print-scale-large .material-badges-row .badge { font-size: 11.5px !important; padding: 7px 10px !important; }
    .print-scale-large .table { font-size: 12px !important; }
    .print-scale-large .table th, .print-scale-large .table td { padding: 8px 6px !important; }
    .print-scale-large .consumables-table td { font-size: 18px !important; padding: 10px 6px !important; }
    .print-scale-large .print-signature-box { margin-top: 55px !important; font-size: 12px !important; }

    body.print-scale-medium { font-size: 11px; }
    .print-scale-medium .print-header h4 { font-size: 18px !important; }
    .print-scale-medium .prod-summary-grid > .summary-box { min-height: 64px !important; padding: 6px 3px !important; }
    .print-scale-medium .prod-summary-grid .tiny-label { font-size: 8.5px !important; min-height: 21px !important; }
    .print-scale-medium .prod-summary-grid h5 { font-size: 17px !important; }
    .print-scale-medium .material-badges-row .badge { font-size: 10.5px !important; padding: 6px 8px !important; }
    .print-scale-medium .table { font-size: 11px !important; }
    .print-scale-medium .table th, .print-scale-medium .table td { padding: 6px 5px !important; }
    .print-scale-medium .consumables-table td { font-size: 15px !important; padding: 8px 5px !important; }
    .print-scale-medium .print-signature-box { margin-top: 40px !important; font-size: 11px !important; }

    body.print-scale-compact { font-size: 9.5px; }
    .print-scale-compact .print-header h4 { font-size: 15px !important; }
    .print-scale-compact .prod-summary-grid > .summary-box { min-height: 52px !important; padding: 4px 2px !important; }
    .print-scale-compact .prod-summary-grid .tiny-label { font-size: 7.5px !important; min-height: 18px !important; }
    .print-scale-compact .prod-summary-grid h5 { font-size: 14px !important; }
    .print-scale-compact .material-badges-row .badge { font-size: 9.5px !important; padding: 4px 6px !important; }
    .print-scale-compact .table { font-size: 9.5px !important; }
    .print-scale-compact .table th, .print-scale-compact .table td { padding: 4px 3px !important; }
    .print-scale-compact .consumables-table td { font-size: 13px !important; padding: 5px 3px !important; }
    .print-scale-compact .print-signature-box { margin-top: 25px !important; font-size: 9.5px !important; }

    @media print {
      body { padding: 0 !important; margin: 0 !important; }
    }
  </style>
</head>
<body class="print-scale-${scale}">
  ${content}
</body>
</html>`);
  win.document.close();
  setTimeout(() => {
    win.focus();
    win.print();
  }, 400);
});

// Digital and offset share the print selection dialog and ticket styling.
function digitalPaperGroups(books) {
  const groups = new Map();
  books.forEach(b => {
    const key = `${b.paper}-${b.material}`;
    const g = groups.get(key) || {label: key, sheets: 0, pack: b.paper === "A3" ? COST.PACK_A3 : b.material === "MT" ? COST.PACK_MT_A4 : COST.PACK_WF_A4};
    g.sheets += b.sheets;
    groups.set(key, g);
  });
  groups.set("Cover", {label:"Cover", sheets:books.reduce((n,b)=>n+b.qty,0), pack:COST.COVER_PACK});
  return [...groups.values()].map(g=>({...g, packs:Math.ceil(g.sheets/g.pack)}));
}
function digitalPaperBadges(books) {
  return digitalPaperGroups(books).map((g,i)=>`<span class="badge px-2 py-1" style="background:${["#2563eb","#64748b","#334155","#dc2626"][i%4]}">${g.label}: ${fmtInt(g.sheets)} sh (${fmtInt(g.packs)} pk @${fmtInt(g.pack)})</span>`).join("");
}
function buildDigitalPrintPreview() {
  const d = window._costData;
  if (!d) return;
  const indices = [];
  $(".print-book-opt:checked").each(function(){indices.push(Number($(this).val()));});
  const books = d.books.filter((_,i)=>indices.includes(i));
  const qty = books.reduce((n,b)=>n+b.qty,0);
  const sheets = books.reduce((n,b)=>n+b.sheets,0);
  const cost = books.reduce((n,b)=>n+b.pricePerBook*b.qty,0);
  const days = books.length ? roundUpHalfDay(sheets/d.capacity)+d.extraDays : 0;
  const schedule = calcFinishDate(days,d.startDate,d.holidays);
  const groups = digitalPaperGroups(books);
  const cards = [
    ["Books",books.length,"#64748b"], ["Total Qty",fmtInt(qty),"#2563eb"],
    ["Inside Sheets",fmtInt(sheets),"#16a34a"], ["Work Days",days,"#8e44ad"],
    ["Inside Packs",fmtInt(groups.filter(g=>g.label!=="Cover").reduce((n,g)=>n+g.packs,0)),"#d68910"],
    ["Cover Packs",Math.ceil(qty/COST.COVER_PACK),"#dc2626"], ["Total Cost",fmtPrice(cost),"#16a085"]
  ];
  const scale = $("#printScaleSelect").val() || "xl";
  let html = `<div class="print-container print-scale-${scale}">`;
  if ($("#optHeader").is(":checked")) html += `<div class="print-header mb-3 pb-2 border-bottom"><h4 class="fw-bold">Digital Production Job Ticket</h4><div class="meta-bar d-flex gap-3 flex-wrap"><span><b>Speed:</b> ${fmtInt(d.capacity)} Sheets/Day</span><span><b>Work Days:</b> ${days}</span><span><b>Extra Days:</b> ${books.length ? d.extraDays : 0}</span><span><b>Start:</b> ${d.startDate}</span><span><b>Finish:</b> ${schedule.finishDate}</span></div></div>`;
  if ($("#optSummary").is(":checked")) html += `<div class="prod-summary-grid mb-3">${cards.map(([label,value,color])=>`<div class="summary-box" style="background:#f8fafc;border:2px solid ${color}"><div class="tiny-label">${label}</div><h5 style="color:${color}" class="mb-0 fw-bold">${value}</h5></div>`).join("")}</div>`;
  if ($("#optMaterials").is(":checked")) html += `<div class="material-badges-row mb-3">${digitalPaperBadges(books)}</div>`;
  if ($("#optBreakdown").is(":checked")) html += `<div class="card p-3 mb-3 border"><div class="d-flex justify-content-between mb-2"><h6 class="fw-bold">Job Breakdown</h6><span class="badge bg-primary">${fmtInt(d.capacity)} Sheets/Day</span></div><table class="table table-bordered text-center align-middle"><thead><tr><th>#</th><th>Title</th><th>Pg</th><th>Qty</th><th>Paper</th><th>Mat</th><th>Bind</th><th>Sheets</th><th>Price/Book</th><th>Days</th></tr></thead><tbody>${books.map((b,i)=>`<tr><td>${i+1}</td><td class="text-start fw-semibold">${sanitize(b.title)}</td><td>${b.pages}</td><td>${fmtInt(b.qty)}</td><td>${b.paper}</td><td>${b.material}</td><td>${b.binding}</td><td style="color:#16a34a;font-weight:700">${fmtInt(b.sheets)}</td><td>${fmtPrice(b.pricePerBook)}</td><td style="background:#eff6ff;color:#2563eb;font-weight:700">${roundUpHalfDay(b.sheets/d.capacity)}</td></tr>`).join("") || '<tr><td colspan="10">No books selected</td></tr>'}</tbody><tfoot><tr class="table-light fw-bold"><td colspan="3">TOTAL</td><td>${fmtInt(qty)}</td><td colspan="3"></td><td>${fmtInt(sheets)}</td><td>${fmtPrice(cost)}</td><td>${days}</td></tr></tfoot></table></div>`;
  if ($("#optSignatures").is(":checked")) html += `<div class="print-signature-block"><div class="print-signature-box">Prepared By</div><div class="print-signature-box">Production Mgr</div><div class="print-signature-box">Press Operator</div><div class="print-signature-box">QA / Delivery</div></div>`;
  $("#printArea").html(html + "</div>");
}
document.getElementById("costStartDate").value = document.getElementById("prodStartDate").value;


const ADDITION_TEMPLATE_URL = "/mnt/data/Template for Addition.xlsx";
const MODIFICATION_TEMPLATE_URL = "/mnt/data/Modification Template.xlsx";
const DELETION_TEMPLATE_URL = "/mnt/data/Template for Deletion.xlsx";

const POSSIBLE_ACTION_COLUMN_NAMES = ["action", "Action", "ACTION", "Type", "Operation"];

const POSSIBLE_DATE_OF_LEAVING_NAMES = ["Date of Leaving", "Date Of Leaving", "date_of_leaving", "date of leaving", "Date of Leaving "];

/* ========== STATE ========== */
let additionHeaders = null;
let modificationHeaders = null;
let deletionHeaders = null;

/* Utility to read remote xlsx (arraybuffer) and return first sheet as JSON (and headers) */
async function fetchXlsxHeaders(url) {
  try {
    const res = await fetch(url);
    if (!res.ok) throw new Error("Fetch failed: " + res.status);
    const ab = await res.arrayBuffer();
    const wb = XLSX.read(ab, { type: "array" });
    const firstSheetName = wb.SheetNames[0];
    const ws = wb.Sheets[firstSheetName];
    // get header row only
    const range = XLSX.utils.decode_range(ws['!ref'] || "A1");
    const headerRow = [];
    for (let C = range.s.c; C <= range.e.c; ++C) {
      const cell = ws[XLSX.utils.encode_cell({ r: range.s.r, c: C })];
      headerRow.push(cell ? String(cell.v).trim() : "");
    }
    return headerRow;
  } catch (err) {
    console.warn("Could not fetch template", url, err);
    return null;
  }
}

/* When page loads, attempt to fetch internal templates to extract headers */
async function loadInternalTemplates() {
  const log = document.getElementById("log");
  log.innerText = "Loading internal templates...";
  const [a, m, d] = await Promise.all([
    fetchXlsxHeaders(ADDITION_TEMPLATE_URL),
    fetchXlsxHeaders(MODIFICATION_TEMPLATE_URL),
    fetchXlsxHeaders(DELETION_TEMPLATE_URL)
  ]);
  additionHeaders = a;
  modificationHeaders = m;
  deletionHeaders = d;

  if (additionHeaders && modificationHeaders && deletionHeaders) {
    log.innerText = "Internal templates loaded. Waiting for Original Template upload.";
    document.getElementById("process-btn").disabled = false;
  } else {
    // fallback: still allow process but inform user
    log.innerText = "Could not load one or more internal templates automatically. The tool will try to infer headers from the Original Template. If outputs are incorrect, please ensure internal template files exist at the configured paths.";
    document.getElementById("process-btn").disabled = false;
  }
}

/* Detect action column header name within headers (case-insensitive) */
function detectActionColumn(headers) {
  if (!headers) return null;
  for (const candidate of POSSIBLE_ACTION_COLUMN_NAMES) {
    const idx = headers.findIndex(h => h && h.trim().toLowerCase() === candidate.trim().toLowerCase());
    if (idx >= 0) return { name: headers[idx], index: idx };
  }
  // else fallback: try any column with single-letter values A/C/T by checking sample rows later
  return null;
}

/* Converts worksheet to array of objects using headers (first row) */
function sheetToObjects(ws) {
  return XLSX.utils.sheet_to_json(ws, { defval: "" });
}

/* Create a workbook (array) from headers and rows (rows are array of objects keyed by header names) */
function buildWorkbookFrom(headers, rows) {
  // ensure rows have all headers (in that exact order)
  const data = [headers];
  for (const r of rows) {
    const row = headers.map(h => (h in r ? r[h] : ""));
    data.push(row);
  }
  const ws = XLSX.utils.aoa_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
  return wb;
}

/* Save workbook as .xlsx and return blob URL for download anchor */
function workbookToBlobUrl(wb) {
  const wopts = { bookType: "xlsx", bookSST: false, type: "array" };
  const wbout = XLSX.write(wb, wopts);
  const blob = new Blob([wbout], { type: "application/octet-stream" });
  const url = URL.createObjectURL(blob);
  return url;
}

/* Main processing function */
async function processFile(file) {
  const log = document.getElementById("log");
  log.innerText = "Reading original file...";
  const ab = await file.arrayBuffer();
  const wb = XLSX.read(ab, { type: "array" });
  const firstSheet = wb.SheetNames[0];
  const ws = wb.Sheets[firstSheet];
  const allRows = sheetToObjects(ws); // array of objects keyed by header names

  if (!allRows || allRows.length === 0) {
    log.innerText = "No rows found in Original Template.";
    return;
  }

  // determine headers from uploaded original
  const originalHeaders = XLSX.utils.sheet_to_json(ws, { header: 1 })[0].map(h => String(h).trim());
  const actionColumnInfo = detectActionColumn(originalHeaders);
  let actionColumnName = actionColumnInfo ? actionColumnInfo.name : null;

  // If not detected, attempt to find which column contains A/C/T in sample rows
  if (!actionColumnName) {
    // check each column for presence of 'A','C','T' in rows
    for (let h of originalHeaders) {
      let found = false;
      for (const r of allRows.slice(0, 10)) {
        const val = (r[h] || "").toString().trim();
        if (/^A$/i.test(val) || /^C$/i.test(val) || /^T$/i.test(val)) { found = true; break; }
      }
      if (found) { actionColumnName = h; break; }
    }
  }

  if (!actionColumnName) {
    log.innerText = "Could not detect Action column automatically. Please ensure Original Template has an 'Action' column containing A, C, or T.";
    return;
  }

  log.innerText = `Detected Action column: "${actionColumnName}". Splitting rows...`;

  // Prepare arrays
  const rowsAdd = [];
  const rowsMod = [];
  const rowsDel = [];

  // For deletion Date of Leaving mapping
  // find if original has any of the date of leaving column names
  let originalDateOfLeavingName = null;
  for (const cand of POSSIBLE_DATE_OF_LEAVING_NAMES) {
    const idx = originalHeaders.findIndex(h => h && h.trim().toLowerCase() === cand.trim().toLowerCase());
    if (idx >= 0) { originalDateOfLeavingName = originalHeaders[idx]; break; }
  }

  // iterate rows and split
  for (const r of allRows) {
    const action = (r[actionColumnName] || "").toString().trim().toUpperCase();
    if (action === "A") rowsAdd.push(r);
    else if (action === "C") rowsMod.push(r);
    else if (action === "T") {
      // ensure Date of Leaving column exists (copy if present else keep blank)
      if (deletionHeaders && originalDateOfLeavingName) {
        // ensure r has the original column mapped to the deletion header name if different
        // We'll keep row as is; mapping to exact deletion headers is done when building workbook
      }
      rowsDel.push(r);
    }
    // ignore rows with empty or unexpected action
  }

  log.innerText = `Rows assigned - Add: ${rowsAdd.length}, Modify: ${rowsMod.length}, Delete: ${rowsDel.length}`;

  // If internal templates not loaded, fallback: infer headers for each template from original headers
  if (!additionHeaders) additionHeaders = originalHeaders.slice();
  if (!modificationHeaders) modificationHeaders = originalHeaders.slice();
  if (!deletionHeaders) {
    // ensure deletionHeaders include Date of Leaving column (add if missing)
    deletionHeaders = originalHeaders.slice();
    const hasDateLeaving = deletionHeaders.some(h => POSSIBLE_DATE_OF_LEAVING_NAMES.some(c=>c.trim().toLowerCase()===h.trim().toLowerCase()));
    if (!hasDateLeaving) deletionHeaders.push("Date of Leaving");
  } else {
    // ensure deletionHeaders contain Date of Leaving (if not present, add it at end)
    const delHas = deletionHeaders.some(h => POSSIBLE_DATE_OF_LEAVING_NAMES.some(c=>c.trim().toLowerCase()===h.trim().toLowerCase()) || h.trim().toLowerCase()==="date of leaving");
    if (!delHas) deletionHeaders.push("Date of Leaving");
  }

  // Helper to remap rows to exact headers: if header exists in original row use it, otherwise blank.
  function remapRowsToHeaders(rows, headers) {
    return rows.map(r => {
      const out = {};
      for (const h of headers) {
        // If original has same header exactly use it
        if (h in r) out[h] = r[h];
        else {
          // try to find original column with same name ignoring case
          const key = Object.keys(r).find(k => k && k.trim().toLowerCase() === h.trim().toLowerCase());
          if (key) out[h] = r[key];
          else {
            // special rule: if header is "Date of Leaving" and original had another variant, copy from original
            const dolKey = Object.keys(r).find(k => POSSIBLE_DATE_OF_LEAVING_NAMES.some(c=>c.trim().toLowerCase()===k.trim().toLowerCase()));
            if (dolKey && h.trim().toLowerCase()==="date of leaving") out[h] = r[dolKey];
            else out[h] = "";
          }
        }
      }
      return out;
    });
  }

  // Build workbooks
  const wbAdd = buildWorkbookFrom(additionHeaders, remapRowsToHeaders(rowsAdd, additionHeaders));
  const wbMod = buildWorkbookFrom(modificationHeaders, remapRowsToHeaders(rowsMod, modificationHeaders));
  const wbDel = buildWorkbookFrom(deletionHeaders, remapRowsToHeaders(rowsDel, deletionHeaders));

  // Convert to blob URLs
  const urlAdd = workbookToBlobUrl(wbAdd);
  const urlMod = workbookToBlobUrl(wbMod);
  const urlDel = workbookToBlobUrl(wbDel);

  // Wire download buttons
  const dlAdd = document.getElementById("dl-add");
  const dlMod = document.getElementById("dl-mod");
  const dlDel = document.getElementById("dl-del");
  dlAdd.href = urlAdd;
  dlMod.href = urlMod;
  dlDel.href = urlDel;

  document.getElementById("download-area").classList.remove("hidden");
  log.innerText += "\nDownload links ready.";
}

/* ========== DOM wiring ========== */
window.addEventListener("DOMContentLoaded", async () => {
  await loadInternalTemplates();

  const input = document.getElementById("file-input");
  const processBtn = document.getElementById("process-btn");
  let uploadedFile = null;

  input.addEventListener("change", (e) => {
    if (e.target.files && e.target.files[0]) {
      uploadedFile = e.target.files[0];
      document.getElementById("log").innerText = `Selected: ${uploadedFile.name}`;
      // enable process button
      processBtn.disabled = false;
    }
  });

  processBtn.addEventListener("click", async () => {
    if (!uploadedFile) {
      alert("Please select the Original Template Excel file first.");
      return;
    }
    processBtn.disabled = true;
    await processFile(uploadedFile);
    processBtn.disabled = false;
  });
});

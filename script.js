/* NO INTERNAL TEMPLATE FILES USED */
const ADDITION_TEMPLATE_URL = null;
const MODIFICATION_TEMPLATE_URL = null;
const DELETION_TEMPLATE_URL = null;

/* Possible column name variations */
const POSSIBLE_ACTION_COLUMN_NAMES = ["action", "Action", "ACTION", "Type", "Operation"];
const POSSIBLE_DATE_OF_LEAVING_NAMES = [
    "Date of Leaving", "Date Of Leaving", "date_of_leaving",
    "date of leaving", "Date of Leaving "
];

/* State */
let additionHeaders = null;
let modificationHeaders = null;
let deletionHeaders = null;

/* Detect action column */
function detectActionColumn(headers) {
    if (!headers) return null;
    for (const candidate of POSSIBLE_ACTION_COLUMN_NAMES) {
        const idx = headers.findIndex(
            h => h && h.trim().toLowerCase() === candidate.trim().toLowerCase()
        );
        if (idx >= 0) return { name: headers[idx], index: idx };
    }
    return null;
}

/* Sheet to objects */
function sheetToObjects(ws) {
    return XLSX.utils.sheet_to_json(ws, { defval: "" });
}

/* Build workbook */
function buildWorkbookFrom(headers, rows) {
    const data = [headers];
    for (const r of rows) {
        data.push(headers.map(h => (r[h] !== undefined ? r[h] : "")));
    }
    const ws = XLSX.utils.aoa_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
    return wb;
}

/* Convert workbook → blob url */
function workbookToBlobUrl(wb) {
    const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    const blob = new Blob([wbout], { type: "application/octet-stream" });
    return URL.createObjectURL(blob);
}

/* MAIN processor */
async function processFile(file) {
    const log = document.getElementById("log");
    log.innerText = "Reading original file...";

    const ab = await file.arrayBuffer();
    const wb = XLSX.read(ab, { type: "array" });
    const ws = wb.Sheets[wb.SheetNames[0]];

    const allRows = sheetToObjects(ws);
    if (!allRows || allRows.length === 0) {
        log.innerText = "No rows found in Original Template.";
        return;
    }

    const originalHeaders =
        XLSX.utils.sheet_to_json(ws, { header: 1 })[0].map(h => String(h).trim());

    /* Detect Action Column */
    const actionInfo = detectActionColumn(originalHeaders);
    let actionColumnName = actionInfo ? actionInfo.name : null;

    if (!actionColumnName) {
        for (let h of originalHeaders) {
            let found = false;
            for (const r of allRows.slice(0, 10)) {
                const val = (r[h] || "").toString().trim();
                if (/^A$/i.test(val) || /^C$/i.test(val) || /^T$/i.test(val)) {
                    found = true;
                    break;
                }
            }
            if (found) {
                actionColumnName = h;
                break;
            }
        }
    }

    if (!actionColumnName) {
        log.innerText = "Action column not found. Expected A / C / T.";
        return;
    }

    log.innerText = `Action column detected: ${actionColumnName} — Processing…`;

    const rowsAdd = [], rowsMod = [], rowsDel = [];

    /* find date-of-leaving column if exists */
    let originalDateOfLeavingName = null;
    for (const cand of POSSIBLE_DATE_OF_LEAVING_NAMES) {
        const idx = originalHeaders.findIndex(
            h => h && h.trim().toLowerCase() === cand.trim().toLowerCase()
        );
        if (idx >= 0) {
            originalDateOfLeavingName = originalHeaders[idx];
            break;
        }
    }

    /* Split rows */
    for (const r of allRows) {
        const action = (r[actionColumnName] || "").toString().trim().toUpperCase();

        if (action === "A") rowsAdd.push(r);
        else if (action === "C") rowsMod.push(r);
        else if (action === "T") rowsDel.push(r);
    }

    log.innerText =
        `Addition: ${rowsAdd.length}\n` +
        `Modification: ${rowsMod.length}\n` +
        `Deletion: ${rowsDel.length}`;

    /* Define headers from Original Template */
    additionHeaders = originalHeaders.slice();
    modificationHeaders = originalHeaders.slice();

    deletionHeaders = originalHeaders.slice();
    const hasDOL = deletionHeaders.some(h =>
        POSSIBLE_DATE_OF_LEAVING_NAMES.some(c => c.trim().toLowerCase() === h.trim().toLowerCase())
    );
    if (!hasDOL) deletionHeaders.push("Date of Leaving");

    /* Map Rows to headers */
    function remap(rows, headers) {
        return rows.map(r => {
            const out = {};
            for (const h of headers) {
                const key = Object.keys(r).find(
                    k => k.trim().toLowerCase() === h.trim().toLowerCase()
                );
                if (key) out[h] = r[key];
                else if (h.trim().toLowerCase() === "date of leaving" && originalDateOfLeavingName)
                    out[h] = r[originalDateOfLeavingName] || "";
                else out[h] = "";
            }
            return out;
        });
    }

    /* Create files */
    const urlAdd = workbookToBlobUrl(
        buildWorkbookFrom(additionHeaders, remap(rowsAdd, additionHeaders))
    );
    const urlMod = workbookToBlobUrl(
        buildWorkbookFrom(modificationHeaders, remap(rowsMod, modificationHeaders))
    );
    const urlDel = workbookToBlobUrl(
        buildWorkbookFrom(deletionHeaders, remap(rowsDel, deletionHeaders))
    );

    /* Attach download links */
    document.getElementById("dl-add").href = urlAdd;
    document.getElementById("dl-mod").href = urlMod;
    document.getElementById("dl-del").href = urlDel;

    document.getElementById("download-area").classList.remove("hidden");
    log.innerText += "\nDownloads ready.";
}

/* DOM READY */
window.addEventListener("DOMContentLoaded", () => {
    document.getElementById("process-btn").disabled = false;
    document.getElementById("log").innerText = "Ready. Upload Original Template.";

    const input = document.getElementById("file-input");
    const processBtn = document.getElementById("process-btn");

    let uploadedFile = null;

    input.addEventListener("change", (e) => {
        uploadedFile = e.target.files[0];
        document.getElementById("log").innerText = `Selected: ${uploadedFile.name}`;
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

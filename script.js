// INTERNAL TEMPLATE HEADERS
const additionColumns = [
    "Member Name", "Email", "Phone", "Date of Birth", "Gender",
    "Date of Joining", "Employee ID", "Relationship", "Designation",
    "GHI Sum Insured", "GHI ID", "GHI Risk Inception Date", "GHI Risk End Date"
];

const modificationColumns = [
    "Status", "Unique ID", "Employee ID", "Member Name", "Email ID",
    "Phone Number", "Date of birth", "Gender", "Relationship",
    "Date of Joining", "Date of Leaving", "Designation",
    "GHI Sum Insured", "GHI ID", "GHI Claimed?", "GHI Risk Inception Date",
    "GHI Risk End Date", "GHI Premium", "GHI Refund", "Endorsement date"
];

const deletionColumns = modificationColumns;

let additionData = [], modificationData = [], deletionData = [];

document.getElementById("fileInput").addEventListener("change", (e) => {
    const file = e.target.files[0];
    if (!file) return;

    document.getElementById("loadingText").style.display = "block";

    setTimeout(() => processFile(file), 50);
});

function processFile(file) {
    const reader = new FileReader();

    reader.onload = function (evt) {
        const workbook = XLSX.read(evt.target.result, {
            type: "array",
            cellDates: true,
            raw: false
        });

        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });

        processData(rows);
    };

    reader.readAsArrayBuffer(file);
}

function processData(rows) {
    additionData = [];
    modificationData = [];
    deletionData = [];

    rows.forEach((row) => {
        const action = (row["Action"] || "").trim().toUpperCase();

        if (action === "A") additionData.push(map(row, additionColumns));
        else if (action === "C") modificationData.push(map(row, modificationColumns));
        else if (action === "T") {
            const d = map(row, deletionColumns);
            if (!d["Date of Leaving"]) d["Date of Leaving"] = "";
            deletionData.push(d);
        }
    });

    enableButtons();
    document.getElementById("loadingText").style.display = "none";
    alert("File processed successfully ✔");
}

function map(row, templateCols) {
    let obj = {};
    templateCols.forEach(col => obj[col] = row[col] ?? "");
    return obj;
}

function enableButtons() {
    btnAddition.disabled = false;
    btnModification.disabled = false;
    btnDeletion.disabled = false;

    btnAddition.onclick = () => download(additionData, "Template_for_Addition.xlsx");
    btnModification.onclick = () => download(modificationData, "Modification_Template.xlsx");
    btnDeletion.onclick = () => download(deletionData, "Template_for_Deletion.xlsx");
}

function download(data, filename) {
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
    XLSX.writeFile(wb, filename);
}

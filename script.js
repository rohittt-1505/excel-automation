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

        // ⭐ VALIDATION - Action column must exist
        if (!rows[0] || !("Action" in rows[0])) {
            document.getElementById("loadingText").style.display = "none";
            alert("❌ Error: 'Action' column is missing in the uploaded Excel file.");
            return;
        }

        processData(rows);
    };

    reader.readAsArrayBuffer(file);
}

function processData(rows) {
    additionData = [];
    modificationData = [];
    deletionData = [];

    for (let row of rows) {

        const actionRaw = row["Action"];
        const action = (actionRaw || "").trim().toUpperCase();

        // ⭐ VALIDATION - Action is empty
        if (!actionRaw || action === "") {
            alert("❌ Error: Some rows have empty 'Action' value. Please check your file.");
            document.getElementById("loadingText").style.display = "none";
            return;
        }

        // ⭐ VALIDATION - Invalid Action code
        if (!["A", "C", "T"].includes(action)) {
            alert(`❌ Error: Invalid Action value '${actionRaw}' found. Allowed values: A, C, T.`);
            document.getElementById("loadingText").style.display = "none";
            return;
        }

        // Process valid actions
        if (action === "A") additionData.push(map(row, additionColumns));
        else if (action === "C") modificationData.push(map(row, modificationColumns));
        else if (action === "T") {
            const d = map(row, deletionColumns);
            if (!d["Date of Leaving"]) d["Date of Leaving"] = "";
            deletionData.push(d);
        }
    }

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

    // Convert workbook to array buffer
    const wbout = XLSX.write(wb, {
        bookType: "xlsx",
        type: "array"
    });

    // Create file blob (Netlify + iPhone safe)
    const blob = new Blob([wbout], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    });

    // Create temporary download link
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = filename;
    document.body.appendChild(link);
    link.click();

    // Cleanup
    document.body.removeChild(link);
    URL.revokeObjectURL(link.href);
}


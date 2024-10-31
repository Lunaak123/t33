let data = []; // This holds the initial Excel data
let filteredData = []; // This holds the filtered data after user operations

// Function to load and display the Excel sheet initially
async function loadExcelSheet(fileUrl) {
    try {
        const response = await fetch(fileUrl);
        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(new Uint8Array(arrayBuffer), { type: 'array' });
        const sheetName = workbook.SheetNames[0]; // Load the first sheet
        const sheet = workbook.Sheets[sheetName];

        data = XLSX.utils.sheet_to_json(sheet, { defval: null });
        filteredData = [...data];

        // Initially display the full sheet
        displaySheet(filteredData);
    } catch (error) {
        console.error("Error loading Excel sheet:", error);
    }
}

// Function to display the Excel sheet as an HTML table
function displaySheet(sheetData) {
    const sheetContentDiv = document.getElementById('sheet-content');
    sheetContentDiv.innerHTML = ''; // Clear existing content

    if (sheetData.length === 0) {
        sheetContentDiv.innerHTML = '<p>No data available</p>';
        return;
    }

    const table = document.createElement('table');

    // Create table headers
    const headerRow = document.createElement('tr');
    Object.keys(sheetData[0]).forEach(header => {
        const th = document.createElement('th');
        th.textContent = header;
        headerRow.appendChild(th);
    });
    table.appendChild(headerRow);

    // Create table rows
    sheetData.forEach(row => {
        const tr = document.createElement('tr');
        Object.values(row).forEach(cell => {
            const td = document.createElement('td');
            td.textContent = cell;
            tr.appendChild(td);
        });
        table.appendChild(tr);
    });

    sheetContentDiv.appendChild(table);
}

// Function to apply selected operations and filter data
function applyOperation() {
    const primaryColumn = document.getElementById('primary-column').value.toUpperCase();
    const rowFrom = parseInt(document.getElementById('row-from').value) || 0;
    const rowTo = parseInt(document.getElementById('row-to').value) || rowFrom;
    const colFrom = document.getElementById('col-from').value.toUpperCase();
    const colTo = document.getElementById('col-to').value.toUpperCase();
    const operationType = document.getElementById('operation-type').value;
    const operation = document.getElementById('operation').value;

    filteredData = data.filter(row => {
        const primaryValue = row[primaryColumn];

        // Check row range
        const rowIndex = data.indexOf(row);
        const isInRowRange = rowIndex >= rowFrom && rowIndex <= rowTo;

        // Check column range
        const isInColRange = (colFrom === '' || row[colFrom] !== undefined) &&
                             (colTo === '' || row[colTo] !== undefined);

        // Perform null/not-null operation
        const isNullOperation = operation === 'null' ? primaryValue === null : primaryValue !== null;

        return isInRowRange && isInColRange && isNullOperation;
    });

    // Highlight the selected range
    highlightRows(rowFrom, rowTo);
    displaySheet(filteredData);
}

// Function to highlight rows based on selected range
function highlightRows(from, to) {
    const rows = document.querySelectorAll('#sheet-content tr');
    rows.forEach((row, index) => {
        if (index >= from && index <= to) {
            row.classList.add('highlighted');
        } else {
            row.classList.remove('highlighted');
        }
    });
}

// Function to handle download of filtered data
function downloadFilteredData() {
    const filename = document.getElementById('filename').value || 'filtered_data';
    const fileFormat = document.getElementById('file-format').value;

    if (fileFormat === 'xlsx') {
        const worksheet = XLSX.utils.json_to_sheet(filteredData);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, 'FilteredData');
        XLSX.writeFile(workbook, `${filename}.xlsx`);
    } else {
        const csvContent = XLSX.utils.sheet_to_csv(XLSX.utils.json_to_sheet(filteredData));
        const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
        const link = document.createElement("a");
        link.href = URL.createObjectURL(blob);
        link.setAttribute("download", `${filename}.csv`);
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    }
}

// Event Listeners
document.getElementById('apply-operation').addEventListener('click', applyOperation);
document.getElementById('download-button').addEventListener('click', () => {
    document.getElementById('download-modal').style.display = 'flex';
});
document.getElementById('confirm-download').addEventListener('click', downloadFilteredData);
document.getElementById('close-modal').addEventListener('click', () => {
    document.getElementById('download-modal').style.display = 'none';
});

// Load initial Excel file (replace with your file URL)
const excelFileUrl = 'YOUR_EXCEL_FILE_URL_HERE';
loadExcelSheet(excelFileUrl);

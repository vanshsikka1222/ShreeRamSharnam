const excelFile = document.getElementById("excel-file");
const tableContainer = document.getElementById("table-container");
const filterContainer = document.getElementById("filter-container");
const downloadBtn = document.getElementById("download-pdf");

let globalHeaders = [];
let globalRows = [];
let currentFiltered = [];

excelFile.addEventListener("change", (event) => {
  const file = event.target.files[0];
  const reader = new FileReader();

  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: "array" });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    globalHeaders = jsonData[1];
    globalRows = jsonData.slice(2);
    currentFiltered = [...globalRows];

    renderFilters();
    renderTable(globalHeaders, currentFiltered);
  };

  reader.readAsArrayBuffer(file);
});

function renderFilters() {
  filterContainer.innerHTML = "";

  globalHeaders.forEach((header, colIndex) => {
    const uniqueValues = [...new Set(globalRows.map(row => row[colIndex]))].filter(v => v !== undefined);

    const select = document.createElement("select");
    select.setAttribute("data-column", colIndex);

    const label = document.createElement("label");
    label.textContent = header;

    const box = document.createElement("div");
    box.className = "filter-box";

    select.innerHTML = `<option value="">All</option>` + 
      uniqueValues.map(value => `<option value="${value}">${value}</option>`).join("");

    select.addEventListener("change", filterTable);

    box.appendChild(label);
    box.appendChild(select);
    filterContainer.appendChild(box);
  });
}

function filterTable() {
  const filters = Array.from(document.querySelectorAll("#filter-container select"))
    .map(select => ({ column: parseInt(select.dataset.column), value: select.value }));

  currentFiltered = globalRows.filter(row =>
    filters.every(f => f.value === "" || row[f.column]?.toString() === f.value)
  );

  renderTable(globalHeaders, currentFiltered);
}

function renderTable(headers, rows) {
  let html = "<table><thead><tr>";
  headers.forEach(header => html += `<th>${header}</th>`);
  html += "</tr></thead><tbody>";

  rows.forEach(row => {
    html += "<tr>";
    headers.forEach((_, i) => {
      html += `<td>${row[i] !== undefined ? row[i] : ""}</td>`;
    });
    html += "</tr>";
  });

  html += "</tbody></table>";
  tableContainer.innerHTML = html;
}

// 📥 Download currentFiltered data as PDF
downloadBtn.addEventListener("click", async () => {
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF();
  const pageWidth = doc.internal.pageSize.getWidth();
  const colWidth = pageWidth / globalHeaders.length;

  doc.setFontSize(12);
  doc.text("Filtered Excel Data", 10, 10);

  // Headers
  globalHeaders.forEach((header, i) => {
    doc.text(header, 10 + i * colWidth, 20);
  });

  // Data
  currentFiltered.forEach((row, rIdx) => {
    row.forEach((cell, cIdx) => {
      doc.text(cell ? String(cell) : "", 10 + cIdx * colWidth, 30 + rIdx * 10);
    });
  });

  doc.save("filtered-data.pdf");
});

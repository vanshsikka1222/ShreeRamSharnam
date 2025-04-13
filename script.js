const excelFile = document.getElementById("excel-file");
const tableContainer = document.getElementById("table-container");

excelFile.addEventListener("change", (event) => {
  const file = event.target.files[0];
  const reader = new FileReader();

  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: "array" });

    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    // Skipping first row (MASTER FILE)
    const headers = jsonData[1];          // Actual headers in row 2
    const rows = jsonData.slice(2);       // Data from row 3 onwards

    renderTable(headers, rows);
  };

  reader.readAsArrayBuffer(file);
});

function renderTable(headers, rows) {
  let html = "<table><thead><tr>";
  headers.forEach(header => {
    html += `<th>${header}</th>`;
  });
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

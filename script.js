const excelFile = document.getElementById("excel-file");
const tableContainer = document.getElementById("table-container");
const filterControls = document.getElementById("filter-controls");
const categorySelect = document.getElementById("category-select");
const searchInput = document.getElementById("search-input");

let globalHeaders = [];
let globalRows = [];

excelFile.addEventListener("change", (event) => {
  const file = event.target.files[0];
  const reader = new FileReader();

  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: "array" });

    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    const headers = jsonData[1]; // headers in 2nd row
    const rows = jsonData.slice(2); // data from 3rd row

    globalHeaders = headers;
    globalRows = rows;

    renderTable(globalHeaders, globalRows);
    populateCategoryDropdown(globalHeaders);
    filterControls.style.display = "block";
  };

  reader.readAsArrayBuffer(file);
});

function populateCategoryDropdown(headers) {
  categorySelect.innerHTML = headers
    .map((header, index) => `<option value="${index}">${header}</option>`)
    .join('');
}

searchInput.addEventListener("input", () => {
  const colIndex = parseInt(categorySelect.value);
  const query = searchInput.value.toLowerCase();

  const filteredRows = globalRows.filter(row => {
    const cell = row[colIndex];
    return cell && cell.toString().toLowerCase().includes(query);
  });

  renderTable(globalHeaders, filteredRows);
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

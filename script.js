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

  // 🟠 Incharge Filter
  const inchargeBox = document.createElement("div");
  inchargeBox.className = "filter-box";
  inchargeBox.innerHTML = `
    <label>
      <input type="checkbox" id="incharge-only" />
      Show Only Incharges
    </label>
  `;
  filterContainer.appendChild(inchargeBox);
  document.getElementById("incharge-only").addEventListener("change", filterTable);

  // Existing dropdown filters
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

  const showOnlyIncharges = document.getElementById("incharge-only")?.checked;

  currentFiltered = globalRows.filter(row => {
    const matchesFilters = filters.every(f => f.value === "" || String(row[f.column]) === f.value);
    const nameColumn = row.find(cell => typeof cell === "string" && cell.includes("(Incharge)"));
    const isIncharge = Boolean(nameColumn);
    return matchesFilters && (!showOnlyIncharges || isIncharge);
  });

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

// Utility to load scripts dynamically
function loadScript(url) {
  return new Promise((resolve, reject) => {
    if (document.querySelector(`script[src="${url}"]`)) {
      resolve(); return;
    }
    const script = document.createElement('script');
    script.src = url;
    script.onload = resolve;
    script.onerror = reject;
    document.head.appendChild(script);
  });
}

// PDF Generation
async function createPaginatedPDF(headers, rows) {
  if (typeof window.jspdf === 'undefined') {
    await loadScript('https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js');
  }
  if (typeof html2canvas === 'undefined') {
    await loadScript('https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js');
  }

  const { jsPDF } = window.jspdf;
  const orientation = headers.length > 5 ? 'landscape' : 'portrait';
  const doc = new jsPDF({ orientation, unit: 'mm', format: 'a4' });

  const pageWidth = doc.internal.pageSize.getWidth();
  const pageHeight = doc.internal.pageSize.getHeight();
  const rowHeight = 12;
  const headerHeight = 15;
  const contentHeight = pageHeight - 30;
  const maxRowsPerPage = Math.floor((contentHeight - headerHeight) / rowHeight);
  const totalPages = Math.ceil(rows.length / maxRowsPerPage);

  for (let pageNum = 0; pageNum < totalPages; pageNum++) {
    if (pageNum > 0) doc.addPage();

    const pageContainer = document.createElement('div');
    pageContainer.style.width = orientation === 'landscape' ? '277mm' : '210mm';
    pageContainer.style.padding = '0';
    pageContainer.style.backgroundColor = 'white';
    pageContainer.style.position = 'absolute';
    pageContainer.style.left = '-9999px';
    pageContainer.style.top = '0';
    pageContainer.style.fontFamily = "'Noto Sans Devanagari', sans-serif";
    document.body.appendChild(pageContainer);

    const startRow = pageNum * maxRowsPerPage;
    const endRow = Math.min(startRow + maxRowsPerPage, rows.length);
    const pageRows = rows.slice(startRow, endRow);

    let tableHTML = `
      <div style="padding: 10mm; width: 100%;">
        ${pageNum === 0 ? '<h2 style="text-align: center; margin-bottom: 10mm;">Filtered Excel Data</h2>' : ''}
        <table style="width: 100%; border-collapse: collapse; font-size: 12px;">
          <thead><tr>`;

    headers.forEach(header => {
      tableHTML += `<th style="padding: 2mm; background-color: #007bff; color: white; border: 1px solid #ddd;">${header}</th>`;
    });

    tableHTML += `</tr></thead><tbody>`;
    pageRows.forEach((row, idx) => {
      const bgColor = idx % 2 === 0 ? '#ffffff' : '#f9f9f9';
      tableHTML += `<tr style="background-color: ${bgColor};">`;
      headers.forEach((_, i) => {
        tableHTML += `<td style="padding: 2mm; border: 1px solid #ddd;">${row[i] !== undefined ? row[i] : ""}</td>`;
      });
      tableHTML += `</tr>`;
    });

    tableHTML += `</tbody></table>
        <div style="text-align: center; margin-top: 5mm; font-size: 10px;">Page ${pageNum + 1} of ${totalPages}</div>
      </div>`;

    pageContainer.innerHTML = tableHTML;

    try {
      const canvas = await html2canvas(pageContainer, { scale: 2, useCORS: true });
      const imgData = canvas.toDataURL('image/jpeg', 1.0);
      doc.addImage(imgData, 'JPEG', 0, 0, pageWidth, pageHeight);
    } catch (error) {
      console.error("Error rendering page", pageNum, error);
    } finally {
      document.body.removeChild(pageContainer);
    }
  }

  doc.save("filtered-data.pdf");
}

downloadBtn.addEventListener("click", async function () {
  if (currentFiltered.length === 0) {
    alert("No data to export to PDF!");
    return;
  }

  const loadingMsg = document.createElement('div');
  loadingMsg.className = 'loading-msg';
  loadingMsg.innerHTML = '<div class="spinner"></div><div>Generating PDF...</div>';
  document.body.appendChild(loadingMsg);

  try {
    await createPaginatedPDF(globalHeaders, currentFiltered);
  } catch (error) {
    console.error('Error generating PDF:', error);
    alert('There was an error generating the PDF: ' + error.message);
  } finally {
    document.body.removeChild(loadingMsg);
  }
});

// Excel Download
document.getElementById("download-excel").addEventListener("click", () => {
  if (!currentFiltered || currentFiltered.length === 0) {
    alert("No data to export!");
    return;
  }

  // Compose the export data manually with:
  // Row 1: master title (e.g. "MASTER FILE")
  // Row 2: actual headers
  // Row 3+: filtered rows

  const masterTitleRow = [ ["MASTER FILE"] ]; // or dynamic if needed
  const headerRow = [ globalHeaders ];
  const dataRows = currentFiltered.map(row =>
    globalHeaders.map((_, i) => row[i] !== undefined ? row[i] : "")
  );

  const fullData = [...masterTitleRow, ...headerRow, ...dataRows];

  const worksheet = XLSX.utils.aoa_to_sheet(fullData);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Filtered Data");

  XLSX.writeFile(workbook, "filtered-data.xlsx");
});

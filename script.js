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

    if (jsonData.length < 2) {
      alert("Invalid Excel Format. First row should be empty, second row as headers.");
      return;
    }

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
    filters.every(f => f.value === "" || String(row[f.column]) === f.value)
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

// Function to dynamically load scripts
function loadScript(url) {
  return new Promise((resolve, reject) => {
    if (document.querySelector(`script[src="${url}"]`)) {
      resolve(); // Script already loaded
      return;
    }
    const script = document.createElement('script');
    script.src = url;
    script.onload = resolve;
    script.onerror = reject;
    document.head.appendChild(script);
  });
}

// Function to create paginated PDF
async function createPaginatedPDF(headers, rows) {
  // Make sure jsPDF is loaded
  if (typeof window.jspdf === 'undefined') {
    await loadScript('https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js');
  }
  
  // Make sure html2canvas is loaded
  if (typeof html2canvas === 'undefined') {
    await loadScript('https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js');
  }

  // Create PDF
  const { jsPDF } = window.jspdf;
  const orientation = headers.length > 5 ? 'landscape' : 'portrait';
  const doc = new jsPDF({
    orientation: orientation,
    unit: 'mm',
    format: 'a4'
  });

  // Calculate page dimensions
  const pageWidth = doc.internal.pageSize.getWidth();
  const pageHeight = doc.internal.pageSize.getHeight();
  
  // Define rows per page based on orientation
  const rowHeight = 12; // mm
  const headerHeight = 15; // mm
  const topMargin = 15; // mm
  const bottomMargin = 15; // mm
  
  const contentHeight = pageHeight - topMargin - bottomMargin;
  const maxRowsPerPage = Math.floor((contentHeight - headerHeight) / rowHeight);
  
  // Calculate total pages needed
  const totalPages = Math.ceil(rows.length / maxRowsPerPage);
  
  // Generate each page
  for (let pageNum = 0; pageNum < totalPages; pageNum++) {
    // Add page after first one
    if (pageNum > 0) {
      doc.addPage();
    }
    
    // Create container for this page
    const pageContainer = document.createElement('div');
    pageContainer.style.width = orientation === 'landscape' ? '277mm' : '210mm'; // A4 dimensions
    pageContainer.style.padding = '0';
    pageContainer.style.backgroundColor = 'white';
    pageContainer.style.position = 'absolute';
    pageContainer.style.left = '-9999px';
    pageContainer.style.top = '0';
    pageContainer.style.fontFamily = "'Noto Sans Devanagari', sans-serif";
    document.body.appendChild(pageContainer);
    
    // Calculate start and end row for this page
    const startRow = pageNum * maxRowsPerPage;
    const endRow = Math.min(startRow + maxRowsPerPage, rows.length);
    const pageRows = rows.slice(startRow, endRow);
    
    // Build table for this page
    let tableHTML = `
      <div style="padding: 10mm; width: 100%;">
        ${pageNum === 0 ? '<h2 style="text-align: center; margin-bottom: 10mm; font-size: 18px;">Filtered Excel Data</h2>' : ''}
        <table style="width: 100%; border-collapse: collapse; font-size: 12px;">
          <thead>
            <tr>`;
    
    // Add headers
    headers.forEach(header => {
      tableHTML += `<th style="padding: 2mm; background-color: #007bff; color: white; border: 1px solid #ddd; text-align: left;">${header}</th>`;
    });
    
    tableHTML += `</tr></thead><tbody>`;
    
    // Add rows for this page
    pageRows.forEach((row, idx) => {
      const bgColor = idx % 2 === 0 ? '#ffffff' : '#f9f9f9';
      tableHTML += `<tr style="background-color: ${bgColor};">`;
      
      headers.forEach((_, i) => {
        tableHTML += `<td style="padding: 2mm; border: 1px solid #ddd;">${row[i] !== undefined ? row[i] : ""}</td>`;
      });
      
      tableHTML += `</tr>`;
    });
    
    tableHTML += `</tbody></table>`;
    tableHTML += `<div style="text-align: center; margin-top: 5mm; font-size: 10px;">Page ${pageNum + 1} of ${totalPages}</div>`;
    tableHTML += `</div>`;
    
    pageContainer.innerHTML = tableHTML;
    
    // Convert page to image
    try {
      const canvas = await html2canvas(pageContainer, {
        scale: 2, // Higher resolution
        useCORS: true,
        logging: false,
        backgroundColor: '#ffffff'
      });
      
      // Add image to PDF
      const imgData = canvas.toDataURL('image/jpeg', 1.0);
      const imgWidth = pageWidth;
      const imgHeight = pageHeight;
      
      doc.addImage(imgData, 'JPEG', 0, 0, imgWidth, imgHeight);
      
      // Clean up
      document.body.removeChild(pageContainer);
      
    } catch (error) {
      console.error("Error rendering page", pageNum, error);
      document.body.removeChild(pageContainer);
    }
  }
  
  // Save the PDF
  doc.save("filtered-data.pdf");
}

// PDF download using paginated approach
downloadBtn.addEventListener("click", async function() {
  if (currentFiltered.length === 0) {
    alert("No data to export to PDF!");
    return;
  }

  // Show loading message
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
    // Remove loading message
    document.body.removeChild(loadingMsg);
  }
});

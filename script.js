const excelFile = document.getElementById("excel-file");
const tableContainer = document.getElementById("table-container");
const filterContainer = document.getElementById("filter-container");
const downloadBtn = document.getElementById("download-pdf");
const printByBusBtn = document.getElementById("download-bus");
const printByBackBusBtn = document.getElementById("download-bus-back");
const printByTrainBtn = document.getElementById("download-train");
const printByBacktrainBtn = document.getElementById("download-train-back");
const printByOwnBtn = document.getElementById("download-Own");
const printByBackOwnBtn = document.getElementById("download-Own-back");

let globalHeaders = [];
let globalRows = [];
let currentFiltered = [];

excelFile.addEventListener("change", (event) => {
  const file = event.target.files[0];
  const reader = new FileReader();

  reader.onload = function (e) {
    debugger;
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: "array" });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
localStorage.setItem("ExcelData",JSON.stringify(jsonData));
    globalHeaders = jsonData[1];
    globalRows = jsonData.slice(2);
    currentFiltered = [...globalRows];

    renderFilters();
    //saveExcel();
    
    renderTable(globalHeaders, currentFiltered);
//     //var f=new File(jsonData, "application/ms-excel", "ReportFile.xls");
//     var blob = new Blob([jsonData], { type: 'application/ms-excel' });
// var downloadUrl=URL.createObjectURL(blob);
// var a=document.createElement("a");
// a.href=downloadUrl;
// a.download="Reportfile.xls";
//     document.body.appendChild(a);
//     a.click();
  };

  reader.readAsArrayBuffer(file);
});
//Convert to binary Data
function s2ab(s) { 
  var buf = new ArrayBuffer(s.length); //convert s to arrayBuffer
  var view = new Uint8Array(buf);  //create uint8array as viewer
  for (var i=0; i<s.length; i++) view[i] = s.charCodeAt(i) & 0xFF;      //convert to octet
  return buf;    
  }
function saveExcel()
{
  debugger;
//$("#toExcel").click(function(){
//var FileSaver = require('file-saver');

var wb = XLSX.utils.book_new();

wb.Props = {
            Title: "SheetJS Tutorial",
            Subject: "Test",
            Author: "Red Stapler",
            CreatedDate: new Date(2017,12,19)
    };

 wb.SheetNames.push("Test Sheet");
 var ws_data = [['hello' , 'world']];  //a row with 2 columns
 var ws = XLSX.utils.aoa_to_sheet(ws_data);
 wb.Sheets["Test Sheet"] = ws;

 //Exporting the Workbook for Downloading
 var wbout = XLSX.write(wb, {bookType:'xlsx',  type: 'binary'});

const data=new Blob([s2ab(wbout)],{type:"application/octet-stream"});
var filename = "data/Ramshrnam";
saveAs(data,   filename + '.xlsx');
alert('saved');

}
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
  debugger;
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

function filterTableByBus(globalRowsdata) {
  debugger;
  const filters =[{column:1,value:"बस द्वारा"
  }];

  const showOnlyIncharges = document.getElementById("incharge-only")?.checked;

  currentFiltered = globalRowsdata.filter(row => {
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

function renderownCardPrint( rows) {
  let html = '<table  width="100%" style="border-collapse:collapse; border-color:rgba(0,0,0,0.5); border-width:1px;  border:none;">';
  
  let i=0;

  rows.forEach(row => {
    debugger;
    if(i%2==0){
      html += '<tr>';
    }
    html += '<td style="width: 50%;">';
     html += '<div style="height: 6cm; overflow:hidden;">';
             html += '<table  width="100%" style="border-collapse:collapse; border-color:rgba(0,0,0,0.5); border-width:1px;  border:none;">';
                      html += '  <tr class="carduptr" style="background-color: #e2efd9!important;">';
                          
                         
                           html += ' <td style="vertical-align:middle; height:50px; padding-left: 5px;" width="50%">';
                             html += ' <span class="zonetxt">जोन न:</span> <span class="zoneval"><u>'+row[14]+'</u></span>';
                                       html += '             </td>';
                                                    html += '<td style="text-align: right; vertical-align:middle; height:50px; padding-right: 15px;" width="50%">';
                                                      html += '<span class="zonetxt">हाउस  ी डी  न:</span> <span class="zoneval"><u>'+row[13]+'</u></span>';
                                                                            html += '</td>';
                        html += '</tr>';
                        html += '<tr class="carduptr" style="background-color: #e2efd9!important;">';
                          
                         
                         html += ' <td style="vertical-align:middle; height:60px; text-align: center; padding-left: 5px;" width="100%" colspan="2">';
                            html += '<h3>श्री राम शरणम सभा रजि.: (पानीपत)</h3> ';
                              html += '<span class="addrsstxt">185, सिविल लाइन, जालंधर 0181 2453185</span>';
                                                  html += '</td>';
        
                                                 
                      html += '</tr>';
                      
                    
                  html += '<tr class="carduptr" style="background-color: #e2efd9!important;">';
                          
                         
                    html += '<td style="vertical-align:middle; text-align: left; padding-left: 5px;" height="150px" width="100%" colspan="2">';
                      html += '<span class="addrsstxt">क्रम संख्या:- </span><span class="addrssval"><u>'+row[3]+'</u></span> ';
                      html += '&emsp;';
                      html += '<span class="addrsstxt">नाम:- </span><span class="addrssval"><u>'+row[4]+' - '+row[7]+'</u></span>';
                      html += '<br>';
                      html += '<span class="addrsstxt">पानीपत पहुँचने का समय :  </span><span class="addrssval"><u>09.07.25</u></span><span class="addrsstxt"> को दोपहर </span><span class="addrssval"><u>1:00</u></span><span class="addrsstxt"> बजे तक | </span>';
                      html += '<br>';
                     html += '<div style="display: flex;"> <div style="float: left; width: 40%;"><div><span class="addrsstxt">पानीपत में ढहरने का स्थान:- </span></div></div>';
                      html += '<div style="float: left;"><div><span class="addrssval"><u>'+row[8]+'</u></span></div></div>';                       
                    html += '</div><br/>';
                    html += '</td>';
                                            
                                           
                html += '</tr>';
                html += '</table>';
                html += '</div>';
                html += '<div style="height: 2.7cm; overflow:hidden;">';
              html += '<table  width="100%" style="border-collapse:collapse; border-color:rgba(0,0,0,0.5); border-width:1px; border:none;">';
                html += '<tr>';
                          
                         
                    html += '<td style="vertical-align:middle; text-align: left; padding-left: 5px;" width="100%" colspan="2">';
                     html += ' <p class="lunchdiv"><u>भोजन कूपन</u> </p> ';
                      html += '<div style="display: flex;">';
                        html += '<div style="float: left; width: 35%;">';
                            html += '<span class="lunchdet">तिथि:</span>';
                        html += '</div>';
                        html += '<div style="float: left; ">';
                            html += '<span class="lunchdet">10 जुलाई, 2025</span>';
                        html += '</div>';
                      html += '</div>';
                      html += '<div style="display: flex;">';
                        html += '<div style="float: left; width: 35%;">';
                            html += '<span class="lunchdet">समय :</span>';
                        html += '</div>';
                        html += '<div style="float: left; ">';
                            html += '<span class="lunchdet">प्रातः 11:30 बजे से दोपहर 12:30 बजे तक | </span>';
                        html += '</div>';
                      html += '</div>';
                      html += '<div style="display: flex;">';
                        html += '<div style="float: left; width: 35%;">';
                            html += '<span class="lunchdet">स्थान :</span>';
                        html += '</div>';
                        html += '<div style="float: left; ">';
                            html += '<span class="lunchdet">आर्य समाज मंदिर, मॉडल टाउन, पानीपत |</span>';
                        html += '</div>';
                      html += '</div><br/>';
                    html += '</td>';
                                            
                                           
                html += '</tr>';
                    html += '</table>';
                html += '</td>';
                
                if(i%2==1){
                  html += '</tr>';
                }
                if(i%2==1&&i%8==7){
                  html += '<tr>';
                  html += '<td style="min-height:2cm" colspan="2"><br/><br/><br/><br/><br/><br/><br/><br/><br/><br/><br/><br/><br/><br/><td>';
                  html += '</tr>';
                }
                i++;
  });

  html += "</table>";
  return html;
}


function renderbackownCardPrint( rows) {
  let html = '<table  width="100%" style="border-collapse:collapse; border-color:rgba(0,0,0,0.5); border-width:1px;margin-top:-0.3cm; border:none;">';
  
  let i=0;

  rows.forEach(row => {
    
    if(i%2==0){
      html += '<tr>';
    }
    html += '<td style="width: 50%;">';
     html += '<div style="height: 6cm; overflow:hidden;">';
     html += '<table  width="100%" style="border-collapse:collapse; border-color:rgba(0,0,0,0.5); border-width:1px; border:none;">';
     html += '<tr class="carduptr" style="background-color: #e2efd9!important;">';
       
      
         html += '<td style="vertical-align:left; height:50px; padding-left: 5px;">';
           html += '<span class="zonetxt"><u>ध्यान देने योग्य जरूरी बातें :-</u></span>';
                                 html += '</td>';
                                 
     html += '</tr>';
     html += '<tr class="carduptr" style="background-color: #e2efd9!important;">';
       
      
       html += '<td style="vertical-align:middle; height:60px; text-align: left; padding-left: 5px;" width="100%">';
         
           html += '<p class="addrsstxt">1. कृपया आप दिनांक: 09.07.25 दोपहर <u>1:00</u> तक पानीपत में ढहरने के स्थान पर पहुंचना सुनिश्चित बनाएं | </p>';
           
           html += '<p class="addrsstxt">2. कृपया अनुसाशन बनाये रखें |</p>';
           html += '<p class="addrsstxt">3. आवश्यकता पड़ने पर मो. संख्या 9988337689 या 9872455886 पर सम्पर्क करें |</p>';
           html += '<p class="addrsstxt">4. कृपया सुनिश्चित कर ले कि पानीपत में ढहरने के स्थान व आश्रम में आपका मोबाइल बंद है |</p>';
           html += '<br/><br/>';
                               html += '</td>';

                              
   html += '</tr>';

html += '</table>';
                html += '</div>';
                html += '<div style="height: 2.7cm; overflow:hidden;">';
              html += '<table  width="100%" style="border-collapse:collapse; border-color:rgba(0,0,0,0.5); border-width:1px; border:none;">';
                html += '<tr>';
                          
                         
                    html += '<td style="vertical-align:middle; text-align: left; padding-left: 5px;" width="100%" colspan="2">';
                      html += '<p class="lunchdiv" style="margin-top: .2cm;"><u>विनती</u> </p>'; 
                      html += '<div style="display: flex;">';
                        html += '<div style="float: left; width: 100%;min-height: 2cm; text-align: center;">';
                            html += '<span class="lunchdet"><b>कृपया समय एवं अनुसाशन का विशेष ध्यान रखें |</b></span>';
                        html += '</div>';
                      html += '</div>';
                      
                      
                    html += '</td>';
                                            
                                           
                html += '</tr>';
                    html += '</table>';
                html += '</td>';
                
                if(i%2==1){
                  html += '</tr>';
                }
                if(i%2==1&&i%8==7){
                  html += '<tr>';
                  html += '<td style="min-height:2cm" colspan="2"><br/><br/><br/><br/><br/><br/><br/><br/><br/><br/><br/><br/><br/><br/><td>';
                  html += '</tr>';
                }
                i++;
  });

  html += "</table>";
  return html;
}




function rendertrainCardPrint( rows) {
  let html = '<table  width="100%" style="border-collapse:collapse; border-color:rgba(0,0,0,0.5); border-width:1px;  border:none;">';
  
  let i=0;

  rows.forEach(row => {
    debugger;
    if(i%2==0){
      html += '<tr>';
    }
    html += '<td style="width: 50%;">';
     html += '<div style="height: 6cm; overflow:hidden;">';
             html += '<table  width="100%" style="border-collapse:collapse; border-color:rgba(0,0,0,0.5); border-width:1px;  border:none;">';
                      html += '  <tr class="carduptr" style="background-color: #fbe4d5!important;">';
                          
                         
                           html += ' <td style="vertical-align:middle; height:50px; padding-left: 5px;" width="50%">';
                             html += ' <span class="zonetxt">जोन न:</span> <span class="zoneval"><u>'+row[14]+'</u></span>';
                                       html += '             </td>';
                                                    html += '<td style="text-align: right; vertical-align:middle; height:50px; padding-right: 15px;" width="50%">';
                                                      html += '<span class="zonetxt">हाउस  ी डी  न:</span> <span class="zoneval"><u>'+row[13]+'</u></span>';
                                                                            html += '</td>';
                        html += '</tr>';
                        html += '<tr class="carduptr" style="background-color: #fbe4d5!important;">';
                          
                         
                         html += ' <td style="vertical-align:middle; height:60px; text-align: center; padding-left: 5px;" width="100%" colspan="2">';
                            html += '<h3>श्री राम शरणम सभा रजि.: (पानीपत)</h3> ';
                              html += '<span class="addrsstxt">185, सिविल लाइन, जालंधर 0181 2453185</span>';
                                                  html += '</td>';
        
                                                 
                      html += '</tr>';
                      
                    
                  html += '<tr class="carduptr" style="background-color: #fbe4d5!important;">';
                          
                         
                    html += '<td style="vertical-align:middle; text-align: left; padding-left: 5px;" height="150px" width="100%" colspan="2">';
                      html += '<span class="addrsstxt">क्रम संख्या:- </span><span class="addrssval"><u>'+row[3]+'</u></span> ';
                      html += '&emsp;';
                      html += '<span class="addrsstxt">नाम:- </span><span class="addrssval"><u>'+row[4]+' - '+row[7]+'</u></span>';
                      html += '<br>';
                      html += '<span class="addrsstxt">जालंधर से जाने का ट्रैन नं:  </span><span class="addrssval"><u>'+row[17]+'</u></span><span class="addrsstxt"> कोच नं: </span><span class="addrssval"><u>'+row[9]+'</u></span><span class="addrsstxt"> सीट नं </span><span class="addrssval"><u>'+row[10]+'</u></span>';
                      html += '<br>';
                      html += '<span class="addrsstxt">पानीपत से आने का ट्रैन नं:  </span><span class="addrssval"><u>'+row[18]+'</u></span><span class="addrsstxt"> कोच नं: </span><span class="addrssval"><u>'+row[11]+'</u></span><span class="addrsstxt"> सीट नं </span><span class="addrssval"><u>'+row[12]+'</u></span>';
                      html += '<br>';
                     html += '<div style="display: flex;"> <div style="float: left; width: 40%;"><div><span class="addrsstxt">पानीपत में ढहरने का स्थान:- </span></div></div>';
                      html += '<div style="float: left;"><div><span class="addrssval"><u>'+row[8]+'</u></span></div></div>';                       
                    html += '</div><br/>';
                    html += '</td>';
                                            
                                           
                html += '</tr>';
                html += '</table>';
                html += '</div>';
                html += '<div style="height: 2.7cm; overflow:hidden;">';
              html += '<table  width="100%" style="border-collapse:collapse; border-color:rgba(0,0,0,0.5); border-width:1px; border:none;">';
                html += '<tr>';
                          
                         
                    html += '<td style="vertical-align:middle; text-align: left; padding-left: 5px;" width="100%" colspan="2">';
                     html += ' <p class="lunchdiv"><u>भोजन कूपन</u> </p> ';
                      html += '<div style="display: flex;">';
                        html += '<div style="float: left; width: 35%;">';
                            html += '<span class="lunchdet">तिथि:</span>';
                        html += '</div>';
                        html += '<div style="float: left; ">';
                            html += '<span class="lunchdet">10 जुलाई, 2025</span>';
                        html += '</div>';
                      html += '</div>';
                      html += '<div style="display: flex;">';
                        html += '<div style="float: left; width: 35%;">';
                            html += '<span class="lunchdet">समय :</span>';
                        html += '</div>';
                        html += '<div style="float: left; ">';
                            html += '<span class="lunchdet">प्रातः 11:30 बजे से दोपहर 12:30 बजे तक | </span>';
                        html += '</div>';
                      html += '</div>';
                      html += '<div style="display: flex;">';
                        html += '<div style="float: left; width: 35%;">';
                            html += '<span class="lunchdet">स्थान :</span>';
                        html += '</div>';
                        html += '<div style="float: left; ">';
                            html += '<span class="lunchdet">आर्य समाज मंदिर, मॉडल टाउन, पानीपत |</span>';
                        html += '</div>';
                      html += '</div><br/>';
                    html += '</td>';
                                            
                                           
                html += '</tr>';
                    html += '</table>';
                html += '</td>';
                
                if(i%2==1){
                  html += '</tr>';
                }
                if(i%2==1&&i%8==7){
                  html += '<tr>';
                  html += '<td style="min-height:2cm" colspan="2"><br/><br/><br/><br/><br/><br/><br/><br/><br/><br/><br/><br/><br/><br/><td>';
                  html += '</tr>';
                }
                i++;
  });

  html += "</table>";
  return html;
}


function renderbacktrainCardPrint( rows) {
  let html = '<table  width="100%" style="border-collapse:collapse; border-color:rgba(0,0,0,0.5); border-width:1px;margin-top:-0.3cm; border:none;">';
  
  let i=0;

  rows.forEach(row => {
    
    if(i%2==0){
      html += '<tr>';
    }
    html += '<td style="width: 50%;">';
     html += '<div style="height: 6cm; overflow:hidden;">';
     html += '<table  width="100%" style="border-collapse:collapse; border-color:rgba(0,0,0,0.5); border-width:1px; border:none;">';
     html += '<tr class="carduptr" style="background-color: #fbe4d5!important;">';
       
      
         html += '<td style="vertical-align:left; height:50px; padding-left: 5px;">';
           html += '<span class="zonetxt"><u>ध्यान देने योग्य जरूरी बातें :-</u></span>';
                                 html += '</td>';
                                 
     html += '</tr>';
     html += '<tr class="carduptr" style="background-color: #fbe4d5!important;">';
       
      
       html += '<td style="vertical-align:middle; height:60px; text-align: left; padding-left: 5px;" width="100%">';
         
           html += '<p class="addrsstxt">1. कृपया आप दिनांक: 09.07.25 सुबह 6:50 तक जालंधर सिटी रेलवे स्टेशन, <u>प्लेटफार्म नं: 2</u> पर पहुंचे | </p>';
           html += '<p class="addrsstxt">2. कृपया आप वापिसी पर दिनांक 10.07.25 दोपहर <u>2:30</u> तक पानीपत रेलवे स्टेशन पर पहुंचे |</p>';
           html += '<p class="addrsstxt">3. कृपया अपना कोच व सीट नं. देख कर बेठेें |</p>';
           html += '<p class="addrsstxt">4. कृपया अनुसाशन बनाये रखें |</p>';
           html += '<p class="addrsstxt">5. आवश्यकता पड़ने पर मो. संख्या 9988337689 या 9872455886 पर सम्पर्क करें |</p>';
           html += '<p class="addrsstxt">6. कृपया सुनिश्चित कर ले कि पानीपत में ढहरने के स्थान व आश्रम में आपका मोबाइल बंद है |</p>';
           html += '<br/><br/>';
                               html += '</td>';

                              
   html += '</tr>';

html += '</table>';
                html += '</div>';
                html += '<div style="height: 2.7cm; overflow:hidden;">';
              html += '<table  width="100%" style="border-collapse:collapse; border-color:rgba(0,0,0,0.5); border-width:1px; border:none;">';
                html += '<tr>';
                          
                         
                    html += '<td style="vertical-align:middle; text-align: left; padding-left: 5px;" width="100%" colspan="2">';
                      html += '<p class="lunchdiv" style="margin-top: .2cm;"><u>विनती</u> </p>'; 
                      html += '<div style="display: flex;">';
                        html += '<div style="float: left; width: 100%;min-height: 2cm; text-align: center;">';
                            html += '<span class="lunchdet"><b>कृपया समय एवं अनुसाशन का विशेष ध्यान रखें |</b></span>';
                        html += '</div>';
                      html += '</div>';
                      
                      
                    html += '</td>';
                                            
                                           
                html += '</tr>';
                    html += '</table>';
                html += '</td>';
                
                if(i%2==1){
                  html += '</tr>';
                }
                if(i%2==1&&i%8==7){
                  html += '<tr>';
                  html += '<td style="min-height:2cm" colspan="2"><br/><br/><br/><br/><br/><br/><br/><br/><br/><br/><br/><br/><br/><br/><td>';
                  html += '</tr>';
                }
                i++;
  });

  html += "</table>";
  return html;
}

function renderCardPrint( rows) {
  let html = '<table  width="100%" style="border-collapse:collapse; border-color:rgba(0,0,0,0.5); border-width:1px;  border:none;">';
  
  let i=0;

  rows.forEach(row => {
    debugger;
    if(i%2==0){
      html += '<tr>';
    }
    html += '<td style="width: 50%;">';
     html += '<div style="height: 6cm; overflow:hidden;">';
             html += '<table  width="100%" style="border-collapse:collapse; border-color:rgba(0,0,0,0.5); border-width:1px;  border:none;">';
                      html += '  <tr class="carduptr" style="background-color: #ccffff!important;">';
                          
                         
                           html += ' <td style="vertical-align:middle; height:50px; padding-left: 5px;" width="50%">';
                             html += ' <span class="zonetxt">जोन न:</span> <span class="zoneval"><u>'+row[14]+'</u></span>';
                                       html += '             </td>';
                                                    html += '<td style="text-align: right; vertical-align:middle; height:50px; padding-right: 15px;" width="50%">';
                                                      html += '<span class="zonetxt">हाउस  ी डी  न:</span> <span class="zoneval"><u>'+row[13]+'</u></span>';
                                                                            html += '</td>';
                        html += '</tr>';
                        html += '<tr class="carduptr">';
                          
                         
                         html += ' <td style="vertical-align:middle; height:50px; text-align: center; padding-left: 5px;" width="100%" colspan="2">';
                            html += '<h3>श्री राम शरणम सभा रजि.: (पानीपत)</h3> ';
                              html += '<span class="addrsstxt">185, सिविल लाइन, जालंधर 0181 2453185</span>';
                                                  html += '</td>';
        
                                                 
                      html += '</tr>';
                      html += '<tr class="carduptr">';
                          
                         
                        html += '<td style="vertical-align:middle; height:50px; text-align: center; padding-left: 5px;" width="100%" colspan="2">';
                          html += '<h3><u>बस न: '+row[9]+'</u></h3>'; 
                            
                                   html += '             </td>';
                                                
                                               
                    html += '</tr>';
                    
                  html += '<tr class="carduptr">';
                          
                         
                    html += '<td style="vertical-align:middle; text-align: left; padding-left: 5px;" height="150px" width="100%" colspan="2">';
                      html += '<span class="addrsstxt">क्रम संख्या:- </span><span class="addrssval"><u>'+row[3]+'</u></span> ';
                      html += '&emsp;';
                      html += '<span class="addrsstxt">नाम:- </span><span class="addrssval"><u>'+row[4]+' - '+row[7]+'</u></span>';
                      html += '<br>';
                     
                     html += '<div style="display: flex;"> <div style="float: left; width: 40%;"><div><span class="addrsstxt">पानीपत में ढहरने का स्थान:- </span></div></div>';
                      html += '<div style="float: left;"><div><span class="addrssval"><u>'+row[8]+'</u></span></div></div>';                       
                    html += '</div><br/>';
                    html += '</td>';
                                            
                                           
                html += '</tr>';
                html += '</table>';
                html += '</div>';
                html += '<div style="height: 2.7cm; overflow:hidden;">';
              html += '<table  width="100%" style="border-collapse:collapse; border-color:rgba(0,0,0,0.5); border-width:1px; border:none;">';
                html += '<tr>';
                          
                         
                    html += '<td style="vertical-align:middle; text-align: left; padding-left: 5px;" width="100%" colspan="2">';
                     html += ' <p class="lunchdiv"><u>भोजन कूपन</u> </p> ';
                      html += '<div style="display: flex;">';
                        html += '<div style="float: left; width: 35%;">';
                            html += '<span class="lunchdet">तिथि:</span>';
                        html += '</div>';
                        html += '<div style="float: left; ">';
                            html += '<span class="lunchdet">10 जुलाई, 2025</span>';
                        html += '</div>';
                      html += '</div>';
                      html += '<div style="display: flex;">';
                        html += '<div style="float: left; width: 35%;">';
                            html += '<span class="lunchdet">समय :</span>';
                        html += '</div>';
                        html += '<div style="float: left; ">';
                            html += '<span class="lunchdet">प्रातः 11:30 बजे से दोपहर 12:30 बजे तक | </span>';
                        html += '</div>';
                      html += '</div>';
                      html += '<div style="display: flex;">';
                        html += '<div style="float: left; width: 35%;">';
                            html += '<span class="lunchdet">स्थान :</span>';
                        html += '</div>';
                        html += '<div style="float: left; ">';
                            html += '<span class="lunchdet">आर्य समाज मंदिर, मॉडल टाउन, पानीपत |</span>';
                        html += '</div>';
                      html += '</div><br/>';
                    html += '</td>';
                                            
                                           
                html += '</tr>';
                    html += '</table>';
                html += '</td>';
                
                if(i%2==1){
                  html += '</tr>';
                }
                if(i%2==1&&i%8==7){
                  html += '<tr>';
                  html += '<td style="min-height:2cm" colspan="2"><br/><br/><br/><br/><br/><br/><br/><br/><br/><br/><br/><br/><br/><br/><td>';
                  html += '</tr>';
                }
                i++;
  });

  html += "</table>";
  return html;
}



function renderbackbusCardPrint( rows) {
  let html = '<table  width="100%" style="border-collapse:collapse; border-color:rgba(0,0,0,0.5); border-width:1px;margin-top:-0.3cm; border:none;">';
  
  let i=0;

  rows.forEach(row => {
    
    if(i%2==0){
      html += '<tr>';
    }
    html += '<td style="width: 50%;">';
     html += '<div style="height: 6cm; overflow:hidden;">';
     html += '<table  width="100%" style="border-collapse:collapse; border-color:rgba(0,0,0,0.5); border-width:1px; border:none;">';
     html += '<tr class="carduptr" style="background-color: #ccffff!important;">';
       
      
         html += '<td style="vertical-align:left; height:50px; padding-left: 5px;">';
           html += '<span class="zonetxt"><u>ध्यान देने योग्य जरूरी बातें :-</u></span>';
                                 html += '</td>';
                                 
     html += '</tr>';
     html += '<tr class="carduptr">';
       
      
       html += '<td style="vertical-align:middle; height:60px; text-align: left; padding-left: 5px;" width="100%">';
         
           html += '<p class="addrsstxt">1. कृपया आप दिनांक: 09.07.25 सुबह 5:30 तक श्री राम शरणम् जालंधर पहुँच जाएं </p>';
           html += '<p class="addrsstxt">2. वापिसी पर बस दिनांक 10.07.25 को दोपहर भोजन उपरान्त 1:00 बजे भाटिया भवन पानीपत से चलेगी |</p>';
           html += '<p class="addrsstxt">3. कृपया अनुसाशन बनाये रखें |</p>';
           html += '<p class="addrsstxt">4. आवश्यकता पड़ने पर मो. संख्या 9988337689 या 9872455886 पर सम्पर्क करें |</p>';
           html += '<p class="addrsstxt">5. कृपया सुनिश्चित कर ले कि पानीपत में ढहरने के स्थान व आश्रम में आपका मोबाइल बंद है |</p>';
           html += '<br/><br/>';
                               html += '</td>';

                              
   html += '</tr>';

html += '</table>';
                html += '</div>';
                html += '<div style="height: 2.7cm; overflow:hidden;">';
              html += '<table  width="100%" style="border-collapse:collapse; border-color:rgba(0,0,0,0.5); border-width:1px; border:none;">';
                html += '<tr>';
                          
                         
                    html += '<td style="vertical-align:middle; text-align: left; padding-left: 5px;" width="100%" colspan="2">';
                      html += '<p class="lunchdiv" style="margin-top: .2cm;"><u>विनती</u> </p>'; 
                      html += '<div style="display: flex;">';
                        html += '<div style="float: left; width: 100%;min-height: 2cm; text-align: center;">';
                            html += '<span class="lunchdet"><b>कृपया समय एवं अनुसाशन का विशेष ध्यान रखें |</b></span>';
                        html += '</div>';
                      html += '</div>';
                      
                      
                    html += '</td>';
                                            
                                           
                html += '</tr>';
                    html += '</table>';
                html += '</td>';
                
                if(i%2==1){
                  html += '</tr>';
                }
                if(i%2==1&&i%8==7){
                  html += '<tr>';
                  html += '<td style="min-height:2cm" colspan="2"><br/><br/><br/><br/><br/><br/><br/><br/><br/><br/><br/><br/><br/><br/><td>';
                  html += '</tr>';
                }
                i++;
  });

  html += "</table>";
  return html;
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

printByOwnBtn.addEventListener("click", async function () {
  if (currentFiltered.length === 0) {
    alert("No data of By Own!");
    return;
  }

  try {
    //await filterTableByBus();
    await window.open('./ownfrontcardprint.html', '_blank');
  } catch (error) {
    console.error('Error generating PDF:', error);
    alert('There was an error generating the PDF: ' + error.message);
  } finally {
    document.body.removeChild(loadingMsg);
  }
});
printByBackOwnBtn.addEventListener("click", async function () {
  if (currentFiltered.length === 0) {
    alert("No data of By Own!");
    return;
  }

  try {
    //await filterTableByBus();
    await window.open('./ownbackcardprint.html', '_blank');
  } catch (error) {
    console.error('Error generating PDF:', error);
    alert('There was an error generating the PDF: ' + error.message);
  } finally {
    document.body.removeChild(loadingMsg);
  }
});

printByTrainBtn.addEventListener("click", async function () {
  if (currentFiltered.length === 0) {
    alert("No data of By Train!");
    return;
  }

  try {
    //await filterTableByBus();
    await window.open('./trainfrontcardprint.html', '_blank');
  } catch (error) {
    console.error('Error generating PDF:', error);
    alert('There was an error generating the PDF: ' + error.message);
  } finally {
    document.body.removeChild(loadingMsg);
  }
});
printByBusBtn.addEventListener("click", async function () {
  if (currentFiltered.length === 0) {
    alert("No data of By Bus!");
    return;
  }

  try {
    //await filterTableByBus();
    await window.open('./busfrontcardprint.html', '_blank');
  } catch (error) {
    console.error('Error generating PDF:', error);
    alert('There was an error generating the PDF: ' + error.message);
  } finally {
    document.body.removeChild(loadingMsg);
  }
});

printByBacktrainBtn.addEventListener("click", async function () {
  if (currentFiltered.length === 0) {
    alert("No data of By Train!");
    return;
  }

  try {
    //await filterTableByBus();
    await window.open('./trainbackcardprint.html', '_blank');
  } catch (error) {
    console.error('Error generating PDF:', error);
    alert('There was an error generating the PDF: ' + error.message);
  } finally {
    document.body.removeChild(loadingMsg);
  }
});

printByBackBusBtn.addEventListener("click", async function () {
  if (currentFiltered.length === 0) {
    alert("No data of By Bus!");
    return;
  }

  try {
    //await filterTableByBus();
    await window.open('./busbackcardprint.html', '_blank');
  } catch (error) {
    console.error('Error generating PDF:', error);
    alert('There was an error generating the PDF: ' + error.message);
  } finally {
    document.body.removeChild(loadingMsg);
  }
});

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
debugger;
  const masterTitleRow = [ ["MASTER FILE"] ]; // or dynamic if needed
  const headerRow = [ globalHeaders ];
  const dataRows = currentFiltered.map(row =>
    globalHeaders.map((_, i) => row[i] !== undefined ? row[i] : "")
  );
  const path = "/";

  // You may need to use relative path in join function depending upon the working file location
  const filePath = "/data/test.xlsx";
  const fullData = [...masterTitleRow, ...headerRow, ...dataRows];

  const worksheet = XLSX.utils.aoa_to_sheet(fullData);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Filtered Data");

  //XLSX.writeFile(workbook, "filtered-data.xlsx");
  XLSX.writeFile(workbook, filePath, {
    bookType: 'xlsx',
    type: 'file'
});
});

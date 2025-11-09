// Run: node create-excel.js
const XLSX = require("xlsx");
const fs = require("fs");

// 1) Your data: array of objects or arrays both work
const rows = [
  { Name: "Aashrey", Age: 28, City: "Delhi" },
  { Name: "Node",   Age: 34, City: "Mumbai" },
  { Name: "JavaScript",  Age: 25, City: "Pune" },
];

// 2) Convert data to a worksheet
const ws = XLSX.utils.json_to_sheet(rows);

// 3) Create a workbook and append the worksheet
const wb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb, ws, "People");

// 4) Optional: set column widths
ws['!cols'] = [{ wch: 12 }, { wch: 6 }, { wch: 12 }];

// 5) Write file
const outFile = "jsonToExcel.xlsx"; // Excel file
XLSX.writeFile(wb, outFile);

// Verify file exists
if (fs.existsSync(outFile)) {
  console.log(`Created ${outFile}`);
}

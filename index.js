const XLSX = require("xlsx");

// Load excel file with addresses
const workbook = XLSX.readFile("SolarApi Locations.xlsx");

//access the first sheet
const sheetName = workbook.SheetNames[0];
const sheet = workbook.Sheets[sheetName];

// Convert to JSON
const data = XLSX.utils.sheet_to_json(sheet);

data.forEach((row) => {
  console.log(row.Name + " " + row.Address + "\n");
});

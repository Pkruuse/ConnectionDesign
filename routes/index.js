var express = require('express');
var ExcelJS = require('exceljs');
var path = require('path');

var router = express.Router();

let excel_data = [];

async function loadSpreadsheet() {
  // adjust __dirname if this file lives elsewhere
  const filePath = path.resolve(__dirname, '../public', 'spreadsheets', 'connections_calculations.xlsm');

  const workbook = new ExcelJS.Workbook();
  try {
    // reads .xlsx/.xlsm as if it were .xlsx
    await workbook.xlsx.readFile(filePath);
    // console.log('✅ Loaded workbook:', workbook.creator, workbook.created);

    // example: grab the first worksheet and log A1
    const sheet = workbook.getWorksheet(1);
    // console.log('A1 value =', sheet.getCell('Propiedades').value);

    // you can also iterate rows
    // console.log('Imprimiendo worksheet');
    // console.log(sheet.getRow(6).getCell(2).value);
    return ({elements_data: [sheet.getRow(6).getCell(2).value, sheet.getRow(6).getCell(3).value, sheet.getRow(6).getCell(4).value]})
    // sheet.eachRow((row, rowNumber) => {
    //   console.log(`Row ${rowNumber}:`, row.values);
    // });
  } catch (err) {
    console.error('❌ Error reading file:', err);
  }
}

loadSpreadsheet()
  .then(data => { excel_data = data; })
  .catch(err => { console.error('Failed to load sheet A:', err); });


/* GET home page. */
router.get('/', function(req, res, next) {
  console.log(excel_data);
  res.render('index', { title: 'ConnectionDesign', excel_data: excel_data });
});

module.exports = router;

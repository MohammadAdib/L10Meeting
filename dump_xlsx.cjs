const ExcelJS = require('exceljs');
const path = require('path');

(async () => {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(path.join(__dirname, 'src', 'blank.xlsx'));

  wb.eachSheet((sheet, sheetId) => {
    console.log(`\n=== Sheet: "${sheet.name}" (id=${sheetId}, rows=${sheet.rowCount}, cols=${sheet.columnCount}) ===`);
    sheet.eachRow({ includeEmpty: false }, (row, rowNum) => {
      const cells = [];
      row.eachCell({ includeEmpty: false }, (cell, colNum) => {
        // Convert column number to letter(s)
        let col = '';
        let n = colNum;
        while (n > 0) {
          n--;
          col = String.fromCharCode(65 + (n % 26)) + col;
          n = Math.floor(n / 26);
        }
        let val = cell.value;
        // Handle rich text, formula, etc.
        if (val && typeof val === 'object') {
          if (val.richText) {
            val = val.richText.map(r => r.text).join('');
          } else if (val.formula) {
            val = `[formula: ${val.formula}] result=${val.result}`;
          } else if (val.sharedFormula) {
            val = `[shared: ${val.sharedFormula}] result=${val.result}`;
          } else if (val instanceof Date) {
            val = val.toISOString();
          } else {
            val = JSON.stringify(val);
          }
        }
        const s = String(val);
        cells.push(`${col}=${s.length > 50 ? s.substring(0, 50) + '...' : s}`);
      });
      if (cells.length > 0) {
        console.log(`  Row ${rowNum}: ${cells.join(' | ')}`);
      }
    });
  });
})();

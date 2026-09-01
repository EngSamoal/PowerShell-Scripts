/**
 * excel.js — thin wrapper around the locally bundled SheetJS library
 * (lib/xlsx.full.min.js) for Excel import/export, reused by the
 * products, sales, and reports screens.
 */

/** Triggers a browser download of rows (array of arrays) as an .xlsx file. */
function exportRowsToExcel(filename, sheetName, headers, rows) {
  const worksheet = XLSX.utils.aoa_to_sheet([headers, ...rows]);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
  XLSX.writeFile(workbook, filename);
}

/** Reads the first sheet of an uploaded .xlsx/.xls/.csv File and returns an array of row objects keyed by header text. */
function readExcelFileAsObjects(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const workbook = XLSX.read(event.target.result, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[firstSheetName];
        const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
        resolve(rows);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = () => reject(reader.error);
    reader.readAsArrayBuffer(file);
  });
}

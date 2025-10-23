const XLSX = require("xlsx");

// Chuyển Excel serial date sang JS Date
function excelDateToJSDate(value) {
  if (!value) return null;
  if (typeof value === "number") return new Date((value - 25569) * 86400 * 1000);
  if (typeof value === "string" && !isNaN(Date.parse(value))) return new Date(value);
  if (value instanceof Date) return value;
  return null;
}

// Đọc sheet Excel theo tên
function readExcelSheet(filePath, preferredSheetName, sheetIndex = 0) {
  const workbook = XLSX.readFile(filePath);
  const sheetName = workbook.SheetNames.find((s) => s === preferredSheetName) || workbook.SheetNames[sheetIndex];
  const sheet = workbook.Sheets[sheetName];
  const data = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: true });
  return { data, sheetName };
}

module.exports = { excelDateToJSDate, readExcelSheet };

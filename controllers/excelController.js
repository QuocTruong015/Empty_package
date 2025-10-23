const path = require("path");
const fs = require("fs");
const { readExcelSheet } = require("../utils/excelUtils");
const { processEmptyPackage } = require("../services/emptyPackageService");

async function uploadExcel(req, res) {
  try {
    const month = parseInt(req.query.month);
    const year = parseInt(req.query.year);

    if (!req.file) return res.status(400).json({ error: "Vui lòng upload 1 file Excel!" });
    if (!month || !year) return res.status(400).json({ error: "Vui lòng nhập ?month=...&year=..." });

    const filePath = path.join(__dirname, "..", req.file.path);
    const { data, sheetName } = readExcelSheet(filePath, "Empty Package", 8);

    const finalData = processEmptyPackage(data, month, year);

    fs.unlinkSync(filePath);

    res.json({
      sheetName,
      month,
      year,
      totalSellers: finalData.length,
      data: finalData,
    });
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: "Đọc file Excel thất bại!" });
  }
}

module.exports = { uploadExcel };

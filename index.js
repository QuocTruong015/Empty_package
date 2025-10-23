const express = require("express");
const multer = require("multer");
const XLSX = require("xlsx");
const path = require("path");
const fs = require("fs");

const app = express();
const PORT = 3000;

// Cấu hình multer để upload file tạm thời
const upload = multer({ dest: "uploads/" });

// Hàm chuyển Excel serial date sang JS Date
function excelDateToJSDate(value) {
  if (!value) return null;

  if (typeof value === "number") {
    return new Date((value - 25569) * 86400 * 1000);
  }

  if (typeof value === "string" && !isNaN(Date.parse(value))) {
    return new Date(value);
  }

  if (value instanceof Date) {
    return value;
  }

  return null;
}

// API upload Excel và tính toán
app.post("/upload-excel", upload.single("file"), (req, res) => {
  try {
    // Lấy tháng và năm từ query
    const month = parseInt(req.query.month);
    const year = parseInt(req.query.year);

    if (!req.file) {
      return res.status(400).json({ error: "Vui lòng upload 1 file Excel!" });
    }
    if (!month || !year) {
      return res
        .status(400)
        .json({ error: "Vui lòng nhập ?month=...&year=... trong URL" });
    }

    const filePath = path.join(__dirname, req.file.path);
    const workbook = XLSX.readFile(filePath);

    const sheetName = workbook.SheetNames[8];
    const sheet =
      workbook.Sheets["Empty Package"] || workbook.Sheets[sheetName];

    const data = XLSX.utils.sheet_to_json(sheet, {
      defval: "",
      raw: true,
    });

    // Lọc theo tháng và năm
    const filtered = data.filter((row) => {
      const date = excelDateToJSDate(row.Date);
      if (!date) return false;
      return date.getMonth() + 1 === month && date.getFullYear() === year;
    });

    // Gom nhóm theo Seller, tính profit từng dòng rồi cộng dồn
    const result = {};
    filtered.forEach((row) => {
      const seller = row.Seller?.trim() || "Unknown";
      const rev = parseFloat(row.Rev) || 0;
      const cost = parseFloat(row.Cost) || 0;

      // ✅ Logic tính profit đúng
      let profit = 0;
      if (rev === 1.5) {
        profit = cost * 0.3;
      } else {
        profit = (rev - cost) + rev * 0.3;
      }

      if (!result[seller]) {
        result[seller] = { Seller: seller, TotalRev: 0, TotalProfit: 0 };
      }

      result[seller].TotalRev += rev;
      result[seller].TotalProfit += profit;
    });

    const finalData = Object.values(result).map((seller) => ({
      Seller: seller.Seller,
      TotalRev: +seller.TotalRev.toFixed(2),
      TotalProfit: +seller.TotalProfit.toFixed(2),
    }));

    // Xóa file tạm
    fs.unlinkSync(filePath);

    res.json({
      month,
      year,
      totalSellers: finalData.length,
      data: finalData,
    });
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: "Đọc file Excel thất bại!" });
  }
});

// Trang chủ test nhanh
app.get("/", (req, res) => {
  res.send(`
    <h2>Upload file Excel để đọc và tính Profit</h2>
    <p>Truyền tháng và năm trên URL: <code>?month=5&year=2025</code></p>
    <form action="/upload-excel?month=5&year=2025" method="post" enctype="multipart/form-data">
      <input type="file" name="file" />
      <button type="submit">Tải lên</button>
    </form>
  `);
});

// Khởi động server
app.listen(PORT, () => {
  console.log(`🚀 Server đang chạy tại http://localhost:${PORT}`);
});

const express = require("express");
const multer = require("multer");
const { uploadEmptyPackage, uploadBuyingLabel } = require("../controllers/excelController");

const router = express.Router();
const upload = multer({ dest: "uploads/" });

router.post("/upload-excel/empty-package", upload.single("file"), uploadEmptyPackage);
router.post("/upload-excel/buying-label", upload.single("file"), uploadBuyingLabel);

module.exports = router;

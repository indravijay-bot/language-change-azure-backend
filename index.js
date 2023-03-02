const express = require("express");
const multer = require("multer");
const XLSX = require("xlsx");
// const json2xls = require("json2xls");

const path = require("path");
const bodyParser = require("body-parser");

const cors = require("cors");
const app = express();

app.use(cors({ origin: "http://localhost:3000" }));
const upload = multer();
app.use(bodyParser.urlencoded({ extended: true }));
app.use(express.json());
var storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, "./public/uploads");
  },
  filename: (req, file, cb) => {
    cb(null, file.fieldname + ".xlsx");
  },
});
var uploadFile = multer({ storage: storage }).single("user_file");

app.post("/upload", uploadFile, (req, res) => {
  try {
    const wb = XLSX.readFile("./public/uploads/user_file.xlsx");
    const ws = wb.Sheets["Sheet1"];
    const data = XLSX.utils.sheet_to_json(ws);
    res.send({ status: 200, success: true, data: data });
  } catch (error) {
    res.send({ status: 404, success: false, error: error.message });
  }
});
app.post("/export", (req, res) => {
  const data = req.body;

  const workbook = XLSX.utils.book_new();
  const worksheet = XLSX.utils.json_to_sheet(data);

  XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");

  const buffer = XLSX.write(workbook, { type: "buffer", bookType: "xlsx" });

  res.setHeader(
    "Content-Type",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  );
  res.setHeader("Content-Disposition", "attachment; filename=data.xlsx");

  res.send(buffer);
});

// Start the server
app.listen(3001, () => {
  console.log("Server is listening on port 3001");
});

const fs = require("fs");
const multer = require("multer");
const express = require("express");

const dotenv = require("dotenv");
dotenv.config();

var MongoClient = require("mongodb").MongoClient; // it's not a function
const URL = process.env.URL;

console.log(process.env);
const excelToJson = require("convert-excel-to-json");

const app = express();

global.__basedir = __dirname;

const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, __basedir + "/uploads/");
  },
  filename: (req, file, cb) => {
    cb(null, file.fieldname + "-" + Date.now() + "-" + file.originalname);
  },
});

const upload = multer({ storage: storage });

app.post("/api/uploadfile", upload.single("uploadfile"), (req, res) => {
  importExcel2Mongo(__basedir + "/uploads/" + req.file.filename);
  res.json({
    msg: "file upload succesfully",
    file: req.file,
  });
});

function importExcel2Mongo(filePath) {
  console.log(filePath);

  // -> Read Excel File to Json Data
  const excelData = excelToJson({
    sourceFile: filePath,
    sheets: [
      {
        // Excel Sheet Name
        name: "Sheet1",

        // Header Row -> be skipped and will not be present at our result object.
        header: {
          rows: 1,
        },

        // Mapping columns to keys
        columnToKey: {
          A: "roll",
          B: "name",
          C: "address",
          D: "age",
        },
      },
    ],
  });

  // -> Log Excel Data to Console
  console.log(excelData);

  // Insert Json-Object to MongoDB
  MongoClient.connect(URL, { useNewUrlParser: true }, (err, db) => {
    if (err) throw err;
    let dbo = db.db("TKF");
    dbo.collection("tempCustomer").insertMany(excelData.Sheet1, (err, res) => {
      if (err) throw err;
      console.log("Number of documents inserted: " + res.insertedCount);
      /**
                Number of documents inserted: 5
            */
      db.close();
    });
  });

  fs.unlinkSync(filePath);
}

// Create a Server
let server = app.listen(process.env.PORT, function () {
  let host = server.address().address;
  let port = server.address().port;

  console.log("App listening at http://%s:%s", host, port);
});

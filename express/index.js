import express from "express";
import * as ExcelTable from "mr-excel";
import * as example from "./example.js";
import * as sharedExamples from "shared-example"; 
import fs from 'fs'
import morgan from "morgan";
import winston from "winston";
import { hasImage } from "./utils.js";
import path from "path";
import { fileURLToPath } from 'url';
import { dirname } from 'path';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

const logger = winston.createLogger({
  level: 'info',
  format: winston.format.combine(
    winston.format.timestamp(),
    winston.format.json()
  ),
  transports: [
    new winston.transports.Console(),
    new winston.transports.File({ filename:path.join(__dirname, 'logs/error.log'), level: 'error' }),
    new winston.transports.File({ filename: path.join(__dirname, 'logs/app.log') }),
  ],
});

const app = express();


app.use(morgan('dev'));

app.get("/list", function (req, response) {
  response.json(Object.keys(sharedExamples).filter(name => name.startsWith("ex")))
})

app.get("/", async (req, response) => {
  response.setHeader(
    "Content-Disposition",
    "attachment; filename=" + "Report.xlsx"
  );
  const result = await ExcelTable.generateExcel({
    ...example.ex1(),
    backend: true,
    generateType: "nodebuffer",
  });
  response.send(result);
});

app.get("/ex2/:id", async (req, response) => {
  let exampleName = "ex1";
  let ex;
  if (typeof req.params.id !== "undefined" && sharedExamples["ex" + req.params.id]) {
    exampleName = "ex" + req.params.id
    ex = sharedExamples[exampleName];
  } else {
    ex = sharedExamples["ex1"];
  }
  const data = {
    ...ex(),
    backend: true,
    generateType: "nodebuffer",
  };
  if(hasImage(data)) {
    data.fetch = example.callApiWithQueue
    logger.info(`${exampleName}: contains image`);
  } 
  fs.writeFileSync(path.join(__dirname, 'data.json'), JSON.stringify(data))
  const result = await ExcelTable.generateExcel(data);
  logger.info(`${exampleName}: file generated successfully`);
  if (data.filename) {
    response.setHeader('Content-Disposition', `attachment; filename="${data.filename}"`);
  } else {
    response.setHeader(
      "Content-Disposition",
      "attachment; filename=" + "Report.xlsx"
    );
  }
  response.end(result);
});
app.get("/ex/:id", async (req, response) => {
  let ex;
  if (typeof req.params.id !== "undefined" && example["ex" + req.params.id]) {
    ex = example["ex" + req.params.id];
  } else {
    ex = example["ex1"];
  }
  response.setHeader(
    "Content-Disposition",
    "attachment; filename=" + "Report.xlsx"
  );
  const data = {
    ...ex(),
    backend: true,
    generateType: "nodebuffer",
  };
  fs.writeFileSync(path.join(__dirname, 'data.json'), JSON.stringify(data))
  const result = await ExcelTable.generateExcel(data);
  response.send(result);
});
app.get("/read/:id", async (req, response) => {
  let paths;
  if (req.params.id == 'x') {
    paths =
      "https://github.com/mohammadrezaeicode/mr-excel-page-repo/blob/main/public/x.xlsx?raw=true";
  } else {
    paths =
      "https://github.com/mohammadrezaeicode/mr-excel-page-repo/blob/main/public/y.xlsx?raw=true";
  }
  response.setHeader("Content-Type", "application/json");
  const result = await ExcelTable.extractExcelData(paths, true, example.callApi);
  response.send({ ...result, sheetName: Array.from(result.sheetName) });
});
app.listen(3000, function () {
  logger.info("run on 3000");
});

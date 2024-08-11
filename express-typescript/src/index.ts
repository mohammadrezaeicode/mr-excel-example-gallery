import  express from "express";

import * as example from "./example";
import { generateExcel,DataModel } from "mr-excel";
import fs from 'fs'
const app = express();
app.get("/", (req, res) => {
  res.send("Hello, TypeScript Express!");
});
app.get("/ex/:id", async (req, response) => {
  let ex;
  if (typeof req.params.id !== "undefined" && example[("ex" + req.params.id) as keyof object]) {
    ex = example[("ex" + req.params.id) as keyof object];
  } else {
    ex = example["ex1"];
  }
  response.setHeader(
    "Content-Disposition",
    "attachment; filename=" + "Report.xlsx"
  );
  const data: DataModel.ExcelTable = {
    ...ex(),
    backend: true,
    generateType: "nodebuffer",
  };
  fs.writeFileSync("data.json", JSON.stringify(data));
  const result = await generateExcel(data);
  response.send(result);
});
const port = process.env.PORT || 3001;
app.listen(port, () => {
  console.log(`Server running at http://localhost:${port}`);
});
import fs from "fs";
import ExcelJS from "exceljs";

// 创建一个Excel工作簿对象
const workbook = new ExcelJS.Workbook();
// 创建一个工作表
const worksheet = workbook.addWorksheet("INN-tib");

const fileContent = fs.readFileSync("pls.txt", "utf8");

//const words = fileContent.match(/[a-zA-Z]+/g);
//const words = fileContent.match(/\b[a-zA-Z]*tinib\b/g);
const words =[...new Set(fileContent.match(/\b\S*tinib\b/g))];


console.log(words);
let lines: string[];
if (words) {
  words.forEach(function (word) {
    worksheet.addRow([word]);
  });
}

// 保存工作簿为Excel文件
workbook.xlsx
  .writeFile("INN-tib-Lists.xlsx")
  .then(function () {
    console.log("Excel文件已成功保存！");
  })
  .catch(function (error) {
    console.log("保存Excel文件时发生错误：", error);
  });

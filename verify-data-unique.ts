import XLSX from "xlsx";
const workbook = XLSX.readFile("INN-tinib-Lists.xlsx");
const worksheet = workbook.Sheets[workbook.SheetNames[0]];

// 获取所有单元格的值
// rome-ignore lint/suspicious/noExplicitAny: <explanation>
const data: any = XLSX.utils.sheet_to_json(worksheet, {
  header: 1,
  blankrows: false,
});

// 单词
//const word = "ldhwdt";
const words = [
  // "rasem",
  // "racon",
  // "ire",
  // "conpel",
  // "susalt",
  // "sutal",
  // "licon",
  // "lipol",
  // "selco",
  "lifcon"
];
// 比较单词和 Excel 表格中的数据
words.forEach((word, index) => {
  const log = `正在匹配第${index + 1}个名称{ ${word}tinib }，重复数据如下：`;
  console.log(log);
  for (let i = 1; i < data.length; i++) {
    const value = data[i][0].toString();

    for (let j = 0; j < word.length - 1; j++) {
      const pair = word.substring(j, j + 2);

      if (value.includes(pair)) {
        console.log(value);
        break;
      }
    }
  }
});

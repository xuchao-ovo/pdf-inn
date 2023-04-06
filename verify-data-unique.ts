import XLSX from 'xlsx';
const workbook = XLSX.readFile('INN-tib-Lists.xlsx');
const worksheet = workbook.Sheets[workbook.SheetNames[0]];

// 获取所有单元格的值
// rome-ignore lint/suspicious/noExplicitAny: <explanation>
const  data:any = XLSX.utils.sheet_to_json(worksheet, { header: 1, blankrows: false });

// 单词
const word = 'ldhwdt';

// 比较单词和 Excel 表格中的数据
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

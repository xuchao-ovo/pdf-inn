import XLSX from "xlsx";

// 读取 Excel 文件
const workbook = XLSX.readFile("INN-tib-Lists.xlsx");
const worksheet = workbook.Sheets[workbook.SheetNames[0]];

// 获取所有单元格的值
const data = XLSX.utils.sheet_to_json(worksheet, {
  header: 1,
  blankrows: false,
});

function generateUniqueStrings(
  length: number,
  count: number,
  // rome-ignore lint/suspicious/noExplicitAny: <explanation>
  data: string | any[]
) {
  const characters = "abcdefghijklmnopqrstuvwxyz";
  const uniqueStrings: string[] = [];

  while (uniqueStrings.length < count) {
    let uniqueString = "";

    for (let i = 0; i < length; i++) {
      const randomIndex = Math.floor(Math.random() * characters.length);
      uniqueString += characters.charAt(randomIndex);
    }

    let overlap = false;

    for (let i = 1; i < data.length; i++) {
      const value = data[i][0].toString();

      for (let j = 0; j < value.length - 1; j++) {
        const pair = value.substring(j, j + 2);

        if (uniqueString.includes(pair)) {
          overlap = true;
          break;
        }
      }

      if (overlap) {
        break;
      }
    }

    if (!overlap) {
      uniqueStrings.push(uniqueString);
    }
  }

  return uniqueStrings;
}

const uniqueStrings = generateUniqueStrings(4, 100, data);

// // 创建工作簿
// const workbook2 = XLSX.utils.book_new();

// // 创建工作表
// const worksheet2 = XLSX.utils.aoa_to_sheet(uniqueStrings.map(str => [str]));

// // 将工作表添加到工作簿
// XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');

// // 将工作簿保存到文件
// XLSX.writeFile(workbook, 'output.xlsx');

console.log(uniqueStrings);

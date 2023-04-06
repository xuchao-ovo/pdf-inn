import XLSX from "xlsx";
import crypto from "crypto";

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

const uniqueStrings = generateUniqueStrings(4, 10, data);
console.log(uniqueStrings);

import ExcelJS from 'exceljs';
import fs from 'fs';
import path from 'path';

const soursDir = path.join('C:/Users/User.SPANKY/Documents', 'work');
const imgSourceDir = path.join(soursDir, 'img');
const sourceTablePath = path.join(soursDir, 'source_table.xlsx');
const resultTablePath = path.join(soursDir, 'result_table.xlsx');

async function findPricesFiles(dir) {
  const filesObj = {};

  function searchInDir(currentDir, folderName) {
    const files = fs.readdirSync(currentDir, { withFileTypes: true });

    for (const file of files) {
      const fullPath = path.join(currentDir, file.name);

      if (file.isDirectory()) {
        searchInDir(fullPath, file.name);
      } else if (
        file.isFile() &&
        file.name.toLowerCase().includes('prices_new')
      ) {
        filesObj[folderName] = {
          folderName: folderName,
          filePath: fullPath,
        };
      }
    }
  }

  const dirs = fs
    .readdirSync(dir, { withFileTypes: true })
    .filter((file) => file.isDirectory());
  for (const directory of dirs) {
    searchInDir(path.join(dir, directory.name), directory.name);
  }

  return filesObj;
}

async function extractPricesData(pricesFiles) {
  const result = [];

  for (const file of Object.values(pricesFiles)) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(file.filePath);
    const sheet = workbook.worksheets[0];
    const data = [];

    sheet.eachRow((row, rowNumber) => {
      const rowData = {};
      row.eachCell((cell, colNumber) => {
        rowData[sheet.getRow(1).getCell(colNumber).value] = cell.value;
      });
      data.push(rowData);
    });

    const folderPrices = {
      folderName: file.folderName,
      prices: {},
    };

    for (const row of data) {
      const name =
        row['название'] ||
        row['Название'] ||
        row['наименование'] ||
        row['Наименование'];
      const price = row['цена'] || row['Цена'];

      if (name && price) {
        folderPrices.prices[name] = price;
      }
    }

    result.push(folderPrices);
  }
  return result;
}

async function insertImages(workbook, sheet, headers, imgSourceDir) {
  const imageColumn1 = headers.findIndex((header) => header === 'Картинка') + 1;
  const imageColumn2 =
    headers.findIndex((header) => header === 'Картинка 2') + 1;

  sheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;

    const productName = row.getCell(1).value;
    const files = fs
      .readdirSync(imgSourceDir)
      .filter((file) =>
        file.toLowerCase().startsWith(productName.toLowerCase())
      );

    if (files.length > 0) {
      const imagePath1 = path.join(imgSourceDir, files[0]);
      const image1 = workbook.addImage({
        filename: imagePath1,
        extension: path.extname(imagePath1).substring(1),
      });

      sheet.addImage(image1, {
        tl: { col: imageColumn1 - 1, row: rowNumber - 1 },
        br: { col: imageColumn1, row: rowNumber },
      });

      if (files.length > 1 && imageColumn2) {
        const imagePath2 = path.join(imgSourceDir, files[1]);
        const image2 = workbook.addImage({
          filename: imagePath2,
          extension: path.extname(imagePath2).substring(1),
        });

        sheet.addImage(image2, {
          tl: { col: imageColumn2 - 1, row: rowNumber - 1 },
          br: { col: imageColumn2, row: rowNumber },
        });
      }
    }
  });
}

async function writeInResultTable(sourceTablePath, resultTablePath) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(sourceTablePath);
  const sheet = workbook.worksheets[0];

  const newSheet = workbook.addWorksheet('TempSheet');
  sheet.eachRow((row, rowNumber) => {
    const newRow = newSheet.getRow(rowNumber);
    row.eachCell((cell, colNumber) => {
      const newCell = newRow.getCell(colNumber);
      newCell.style = { ...cell.style };
      newCell.value = cell.value;
    });
  });

  const headers = [];
  const headerRow = newSheet.getRow(1);
  headerRow.eachCell((cell, colNumber) => {
    headers[colNumber - 1] = cell.value;
  });

  const suppliersData = await extractPricesData(
    await findPricesFiles(soursDir)
  );

  // Обновляем цены
  newSheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    const productName = row.getCell(1).value;

    for (const supplier of suppliersData) {
      const supplierColumnIndex = headers.findIndex(
        (header) => header === supplier.folderName
      );
      if (supplierColumnIndex !== -1) {
        const price = supplier.prices[productName];
        if (price !== undefined) {
          const cell = row.getCell(supplierColumnIndex + 1);
          cell.value = Number(price);
        }
      }
    }
  });
  // После создания нового листа и до записи файла
  headers.forEach((header, index) => {
    const column = newSheet.getColumn(index + 1);
    column.width = 15;
  });
  const komment = headers.findIndex((header) => header === 'Комментарий') + 1;
  newSheet.getColumn(komment).width = 30;
  // Вставляем изображения
  await insertImages(workbook, newSheet, headers, imgSourceDir);

  workbook.removeWorksheet(sheet.id);
  newSheet.name = sheet.name;

  await workbook.xlsx.writeFile(resultTablePath);
}

// Вызов функции с двумя путями
(async () => {
  await writeInResultTable(sourceTablePath, resultTablePath);
})();

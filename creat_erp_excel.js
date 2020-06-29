const Excel = require("exceljs");
const _ = require("lodash");
const dayjs = require("dayjs");


(async function () {
  const workbook = new Excel.Workbook();
  const writeBook = new Excel.Workbook();
  await workbook.xlsx.readFile("./from.xlsx");
  let worksheet1 = workbook.getWorksheet(1);
  await workbook.xlsx.readFile("./contrast.xlsx");
  let worksheet2 = workbook.getWorksheet(1);
  await writeBook.xlsx.readFile("./model.xlsx")
  let worksheet3 = writeBook.getWorksheet(1);
  let header = [
    "品号",
    "客户品号",
    "交易单位",
    "交易数量",
    "税率",
    "含税",
    "单价",
    "赠品",
    "预交货日期",
  ];
  header = header.map((item) => ({ name: item }));
  new Promise((resolve, reject) => {
    let contrast_map_us = {};
    worksheet2.eachRow(function (row, rowNumber) {
      if (rowNumber === 1) {
        return;
      }
      const row_value = row.values;
      row_value.shift();
      contrast_map_us[row_value[1]] = row_value;
    });
    resolve(contrast_map_us);
  }).then(async (res) => {
    let array = []
    worksheet1.eachRow(function (row, rowNumber) {
      if (rowNumber === 1) return;
      const values = row.values;
      values.shift()
      const date = dayjs(values[values.length - 2]).year(2021).format("YYYY/MM/DD");
      const us_num = res[values[0]][2]
      const newRow = [
        us_num,
        "",
        "PCS",
        parseInt(values[values.length - 1]),
        0.13,
        "T",
        parseFloat(values[1]),
        "F",
        date,
      ];
      array.push(newRow)
    });
    for(let arr of array){
      worksheet3.addRow(arr)
    }
    await writeBook.xlsx.writeFile("导入ERP.xlsx");
  });
})();
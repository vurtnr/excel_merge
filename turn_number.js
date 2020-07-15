const Excel = require("exceljs");
const {  turnCapacitance,c1,c2,c4,c5 } = require("./config");

(async function () {
  const workbook = new Excel.Workbook();
  await workbook.xlsx.readFile("./old_number.xlsx");
  let worksheet = workbook.getWorksheet(1);
  const writeWorkbook = new Excel.Workbook();
  await writeWorkbook.xlsx.readFile("./turn_model.xlsx");
  let workSheetModel = writeWorkbook.getWorksheet(1);


  new Promise((resolve,reject) => {
    const final_array = [];
    worksheet.eachRow((row, rowNumber) => {
      let values = row.values;
      values.shift();
      if (rowNumber === 1) return;
      values[4] = values[4];
      let arr = values[4].split(" ");
      let letter = arr[2].substr(-2, 1).toLowerCase();
      arr[2] = arr[2].slice(0, -2) + letter + "F";
      arr[1] = arr[1].replace("O", "0");
      let new_model = new Array(5);
      new_model[0] = c1[arr[0]];
      new_model[1] = c2[arr[1]];
      new_model[2] = turnCapacitance(arr[2]);
      new_model[3] = c4[arr[3]];
      new_model[4] = c5[arr[4].toUpperCase()];
      new_model = "MCS" + new_model.join("") + "RB00";
      arr[1] = ["NP0", "C0G"].includes(arr[1]) ? "NP0/C0G" : arr[1];
      values[4] = arr.join(" ");
      let new_spec =
        "C" +
        values[4].slice(0, values[4].indexOf("(")) +
        " " +
        values[4].slice(values[4].indexOf("(")) +
        " 卷装 通用";
      row.getCell("A").value = new_model;
      row.getCell("BJ").value = new_spec;
      values = row.values;
      values.shift();
      values.splice(1, 2);
      values.splice(2, 3);
      final_array.push(values);
    });
    resolve(final_array)
  }).then(async res => {
    const header = []
    const cn_header = []
    let count = 0
    while(count < res.length){
      workSheetModel.addRow(res[count]);
      count++
    }
    
    await writeWorkbook.xlsx.writeFile("贴片电容导入ERP.xlsx");

  })
})();
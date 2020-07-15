const Excel = require("exceljs");
const _ = require("lodash");

(async function () {
  const workbook = new Excel.Workbook();
  await workbook.xlsx.readFile("./years_data.xlsx");
  let worksheet1 = workbook.getWorksheet(1);
  let all_array = [];
  worksheet1.eachRow((row, rowNumber) => {
    let values = row.values;
    values.shift();
    if (rowNumber === 1) return;
    let year = values[0].substring(0, 4);
    let month = values[0].substring(5, 7);

    let arr = [values[2], parseInt(year), parseInt(month), parseInt(values[5])];
    all_array = [...all_array, arr];
  });
  let obj = turnArrayToObject(all_array);
  let all_years_data = [];
  Object.keys(obj).map((key) => {
    let temp = turnArrayToObject(obj[key]);
    Object.keys(temp).forEach((item) => {
      let temp_list = mergeCounts(temp[item]).sort((a, b) => a[0] - b[0]);
      let newArray = new Array(12);
      for (let tem of temp_list) {
        newArray[tem[0] - 1] = tem[1];
      }
      let final_arr = [key, item, ...newArray];
      all_years_data.push(final_arr);
    });
  });
  let header = [
    "产品型号",
    "年份",
    "1月",
    "2月",
    "3月",
    "4月",
    "5月",
    "6月",
    "7月",
    "8月",
    "9月",
    "10月",
    "11月",
    "12月",
  ];
  header = header.map(i => ({name:i}))
  const writeWorkbook = new Excel.Workbook();
  const workSheetOne = writeWorkbook.addWorksheet("数据一览", {
    properties: {
      defaultColWidth: 15,
    },
  });
  workSheetOne.addTable({
    name: "MyTable",
    ref: "A1",
    headerRow: true,
    columns: header,
    rows: all_years_data,
  });
  await writeWorkbook.xlsx.writeFile("亮迪订单数据一览表.xlsx");
})();

function turnArrayToObject(original) {
  let tmp_original = _.cloneDeep(original);
  let final_obj = {};
  for (let item of tmp_original) {
    let key = item[0];
    item.splice(0, 1);
    if (!final_obj[key]) {
      final_obj[key] = [];
      final_obj[key].push(item);
    } else {
      final_obj[key].push(item);
    }
  }
  return final_obj;
}

function mergeCounts(array) {
  let obj = {},
    list = [];
  for (let arr of array) {
    let key = arr[0];
    if (!obj[key]) {
      list.push(arr);
      obj[key] = arr;
    } else {
      for (let item of list) {
        if (item[0] === arr[0]) {
          item[1] = arr[1] + item[1];
        }
      }
    }
  }
  return list;
}

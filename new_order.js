const Excel = require("exceljs");
const _ = require("lodash");
const { cut_index } = require("./config");
const dayjs = require("dayjs");
(async function () {
  const workbook = new Excel.Workbook();
  await workbook.xlsx.readFile("./new_order.xlsx");
  let worksheet1 = workbook.getWorksheet(1);
  
  await workbook.xlsx.readFile("./contrast.xlsx");
  let worksheet3 = workbook.getWorksheet(1);

  let contrast = {};
  worksheet3.eachRow(function (row, rowNumber) {
    if (rowNumber === 1) return;
    const values = row.values;
    values.shift();
    contrast[values[2]] = values;
  });

  let new_array = []
  worksheet1.eachRow(function (row, rowNumber) {
    if (rowNumber === 1) return;
    const values = row.values;
    values.shift();
    let arr = []
    let item = contrast[values[2]]
    let client_num = item[1] 
    let specifications = item[0]+"("+item[1]+")"
    let type = item[item.length - 1];
    for(let i of cut_index){
      let value = values[i]
      if(i === 5)
        value = parseInt(value)
      if(i === 7)
        value = parseFloat(value)
      arr.push(value);
    }
    arr.splice(2, 0, client_num);
    arr.splice(6, 1, ...arr.splice(5, 1, arr[6]));
    arr.push(arr[6])
    arr.splice(5, 0, specifications);
    arr = arr.concat(['','']);
    arr.push(type);
    new_array.push(arr)
  });
  let new_obj = turnArrayToObject(new_array,3);
  
  let array = []
  let table_header = []
  await workbook.xlsx.readFile("./all_orders.xlsx");
  workbook.worksheets[0].eachRow((row,rowNumber) => {
    let values = row.values;
    values.shift();
    if(rowNumber === 1) {
      table_header = values;
      return
    }
    let obj = contrast[values[3]];
    let type = obj[obj.length -1];
    values.push(type)
    array = [...array,values]
  })
  let data_obj = turnArrayToObject(array);
  let data_array = []
  Object.keys(data_obj).forEach(key => {
    let key_obj = {}
    key_obj = turnArrayToObject(data_obj[key],3);
    Object.keys(key_obj).map(key => {
      let key_array_list = Object.values(key_obj[key]);
      if(new_obj[key]){
        let values = Object.values(new_obj[key])
        key_array_list = [...key_array_list,...values]
      }
     data_array = [...data_array,...key_array_list]
    })
    let blank_array = ["", "", "", "", "", "", "", "", "", ""];
    data_array = [...data_array, blank_array, blank_array, blank_array];
  })
  
  data_array.map(item => {
    item.pop();
    if(item.length === 10){
      item.push('')
    }
  })
  table_header = table_header.map((item) => ({ name: item }));
  const writeWorkbook = new Excel.Workbook();
  const workSheetOne = writeWorkbook.addWorksheet("未完结", {
    properties: {
      defaultColWidth: 25,
    },
  });
  workSheetOne.addTable({
    name: "MyTable",
    ref: "A1",
    headerRow: true,
    columns: table_header,
    rows: data_array,
  });
  await writeWorkbook.xlsx.writeFile("亮迪未交订单.xlsx");
  // let row_header = []
  // worksheet2.eachRow(function (row, rowNumber) {
  //   if (rowNumber === 1) {
  //     row_header = row.values;
  //     return;
  //   }
  // });
})();


function turnArrayToObject(original,idx) {
  let tmp_original = _.cloneDeep(original);
  let final_obj = {};
  for (let item of tmp_original) {
    let i = idx ? idx : item.length -1
    if (!final_obj[item[i]]) {
      final_obj[item[i]] = [];
      final_obj[item[i]].push(item);
    } else {
      final_obj[item[i]].push(item);
    }
  }
  return final_obj;
}
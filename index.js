const Excel = require("exceljs");
var _ = require('lodash');

(async function () {
  const workbook = new Excel.Workbook();

  await workbook.xlsx.readFile("./client.xlsx");
  let worksheet1 = workbook.getWorksheet(1);
  await workbook.xlsx.readFile("./对照表.xlsx");
  let worksheet2 = workbook.getWorksheet(1);
  await workbook.xlsx.readFile("./us.xlsx");
  let worksheet3 = workbook.getWorksheet(1);

  new Promise((resolve, reject) => {
    let contrast_map_client = {};
    let contrast_map_us = {};
    worksheet2.eachRow(function (row, rowNumber) {
      if (rowNumber === 1) {
        return;
      }
      const row_value = row.values;
      row_value.shift();
      contrast_map_client[row_value[1]] = row_value;
      contrast_map_us[row_value[2]] = row_value[row_value.length - 1];
    });
    resolve({ contrast_map_client, contrast_map_us });
  })
    .then((res) => {
      const { contrast_map_client, contrast_map_us } = res;
      let client_header = [];
      let client_values_positive = []; //正数
      let client_values_negative = []; //负数
      worksheet1.eachRow(function (row, rowNumber) {
        let values = row.values;
        values.shift();
        if (rowNumber === 1) {
          client_header = values;
          client_header[0] = "亮迪型号";
          client_header[1] = "亮迪规格";
          client_header[2] = "亮迪单价";
          client_header.splice(0, 0, "巨数型号");
          client_header.splice(4, 0, "巨数单价");
          return;
        }
        const sd = contrast_map_client[values[0]][2];
        const price = contrast_map_client[values[0]][3];
        values.splice(0, 0, sd);
        values.splice(4, 0, price);
        if (values[values.length - 1] > 0) {
          client_values_positive.push(values);
        } else {
          client_values_negative.push(values);
        }
      });
      return {
        client_header,
        client_values_positive,
        client_values_negative,
        contrast_map_us,
      };
    })
    .then((res) => {
      let {
        client_header,
        client_values_positive,
        client_values_negative,
        contrast_map_us,
      } = res;

      let us_header = [];
      let us_values_positive = [];
      let us_values_negative = [];

      worksheet3.eachRow(function (row, rowNumber) {
        let values = row.values;
        values.shift();
        if (rowNumber === 1) {
          us_header = values;
          return;
        }
        const price = contrast_map_us[values[0]] ?contrast_map_us[values[0]]:0;
        values.splice(2, 0, price);
        if (values[values.length - 1] > 0) {
          us_values_positive.push(values);
        } else {
          us_values_negative.push(values);
        }
      });
      let client_positive_tmp = _.cloneDeep(client_values_positive);
      let client_negative_tmp = _.cloneDeep(client_values_negative);
      let us_positive_tmp = _.cloneDeep(us_values_positive);
      let us_negative_tmp =  _.cloneDeep(us_values_negative)
      let final_header = client_header.concat(us_header);
      final_header = final_header.map((item) => ({ name: item }));

      let original_single_array = countSignArray(
        client_values_positive,
        client_values_negative,
        us_values_positive,
        us_values_negative
      );
      let original_merge_array = countMergeArray(
        client_positive_tmp,
        client_negative_tmp,
        us_positive_tmp,
        us_negative_tmp
      );
      createNewWorkbook(
        final_header,
        original_merge_array,
        original_single_array
      );
    });
})();

function mergeCount(original, start, price_idx) {
  let obj = {},
    array = [];
  for (let item of original) {
    let key = item[start] + "-" + item[price_idx];
    if (!obj[key]) {
      array.push(item);
      obj[key] = item;
    } else {
      for (let arr of array) {
        if (arr[start] === item[start] && arr[price_idx] === item[price_idx]) {
          arr[arr.length - 1] =
            parseInt(arr[arr.length - 1]) + parseInt(item[item.length - 1]);
        }
      }
    }
  }
  return array;
}

function turnArrayToObject(original) {
  let tmp_original = _.cloneDeep(original);
  let final_obj = {};
  for (let item of tmp_original) {
    item[item.length - 1] = parseInt(item[item.length - 1]);
    if (!final_obj[item[0]]) {
      final_obj[item[0]] = [];
      final_obj[item[0]].push(item);
    } else {
      final_obj[item[0]].push(item);
    }
  }
  Object.keys(final_obj).forEach((key) => {
    return final_obj[key].sort(function (a, b) {
      return a[a.length - 1] - b[b.length - 1];
    });
  });
  return final_obj;
}

function countMergeArray(
  client_values_positive,
  client_values_negative,
  us_values_positive,
  us_values_negative
) {
  let client_merge_positive = mergeCount(client_values_positive, 0, 3);
  let client_merge_negative = mergeCount(client_values_negative, 0, 3);
  let us_merge_positive = mergeCount(us_values_positive, 0, 2);
  let us_merge_negative = mergeCount(us_values_negative, 0, 2);
  let final_merge_client_values = client_merge_positive.concat(
    client_merge_negative
  );
  let final_merge_us_values = us_merge_positive.concat(us_merge_negative);

  let client_merge_group_map = turnArrayToObject(final_merge_client_values);
  let us_merge_group_map = turnArrayToObject(final_merge_us_values);

  let merge_array = countAllValues(client_merge_group_map, us_merge_group_map);
  return merge_array;
}
function countSignArray(
  client_values_positive,
  client_values_negative,
  us_values_positive,
  us_values_negative
) {
  let all_client_positive = turnArrayToObject(client_values_positive);
  let all_client_negative = turnArrayToObject(client_values_negative);
  let all_us_positive = turnArrayToObject(us_values_positive);
  let all_us_negative = turnArrayToObject(us_values_negative);
 
  let single_positive_array = signCountAllValues(
    all_client_positive,
    all_us_positive
  );
  let single_negative_array = signCountAllValues(
    all_client_negative,
    all_us_negative
  );
  let single_final_array = single_positive_array.concat(single_negative_array);
  let single_final_map = {};
  single_final_map = turnArrayToObject(single_final_array);
  let original_single_array = [];
  Object.values(single_final_map).forEach((array) => {
    for (let arr of array) {
      original_single_array.push(arr);
    }
  });
  return original_single_array;
}

function signCountAllValues(client_group_map, us_group_map) {
  let client_group_map_tmp = _.cloneDeep(client_group_map);
  let us_group_map_tmp = _.cloneDeep(us_group_map);
  let tableData = [];
  Object.keys(client_group_map_tmp).map((key) => {
    let client_array = client_group_map_tmp[key];
    let us_array = us_group_map_tmp[key];
    if (!us_array && client_array) {
      for (let i of client_array) {
        let array = i.concat([key, "", 0]);
        tableData.push(array);
      }
    } else {
      let tmp = us_array;
      for (let i in client_array) {
        let array = [];
        if (us_array[i]) {
          us_array[i].splice(2, 1);
          array = client_array[i].concat(us_array[i]);
          tmp.splice(i, 1);
        } else {
          array = client_array[i].concat([key, "", 0]);
        }
        tableData.push(array);
      }

      if (tmp.length > 0) {
        for (let i of tmp) {
          i.splice(2, 1);
          let client = [
            i[0],
            client_array[client_array.length - 1][1],
            client_array[client_array.length - 1][2],
            0,
            client_array[client_array.length - 1][3],
            0,
            ...i,
          ];
          tableData.push(client);
        }
      }
    }
  });
  return tableData;
}

function countAllValues(client_group_map, us_group_map) {
  let client_group_map_tmp = _.cloneDeep(client_group_map);
  let us_group_map_tmp = _.cloneDeep(us_group_map);
  let tableData = [];
  Object.keys(client_group_map_tmp).map((key) => {
    let client_array = client_group_map_tmp[key];
    let us_array = us_group_map_tmp[key];
    if (us_array && us_array.length > client_array.length) {
      client_array.map((client, idx) => {
        us_array.forEach((us, i) => {
          if (
            (client[3] === us[2] &&
              client[client.length - 1] > 0 &&
              us[us.length - 1] > 0) ||
            (client[client.length - 1] < 0 && us[us.length - 1] < 0)
          ) {
            us.splice(2, 1);
            client = client.concat(us);
            us_array.splice(i, 1);
            tableData.push(client);
          }
        });
      });
      for (let i of us_array) {
        i.splice(2, 1);
        let client = [
          i[0],
          client_array[client_array.length - 1][1],
          client_array[client_array.length - 1][2],
          0,
          client_array[client_array.length - 1][3],
          0,
          ...i,
        ];
        tableData.push(client);
      }
    } else {
      if (!us_array) return;
      client_array.map((client, idx) => {
        if (us_array.length === 0) {
          client = client.concat([key, "", 0]);
        } else {
          us_array.forEach((us, i) => {
            if (
              (client[3] === us[2] &&
                client[client.length - 1] > 0 &&
                us[us.length - 1] > 0) ||
              (client[client.length - 1] < 0 && us[us.length - 1] < 0)
            ) {
              us.splice(2, 1);
              client = client.concat(us);
              us_array.splice(i, 1);
            } else {
              client = client.concat([key, "", 0]);
            }
          });
        }
        tableData.push(client);
      });
    }
  });
  return tableData;
}

async function createNewWorkbook(tableHeader, merge_data, single_data) {
  const workbook = new Excel.Workbook();
  const workSheetOne = workbook.addWorksheet("合并", {
    properties: { defaultColWidth: 20 },
  });
  const workSheetTwo = workbook.addWorksheet("单行", {
    properties: { defaultColWidth: 20 },
  });
  workSheetOne.addTable({
    name: "MyTable",
    ref: "A1",
    headerRow: true,
    columns: tableHeader,
    rows: merge_data,
  });
  workSheetTwo.addTable({
    name: "MyTable",
    ref: "A1",
    headerRow: true,
    columns: tableHeader,
    rows: single_data,
  });
  workSheetOne.eachRow(function (row, rowNumber) {
    const values = row.values;
    values.shift();
    if (rowNumber === 1) return;
    if (values[5] !== values[values.length - 1]) {
      row.getCell(values.length).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFed1941" },
      };
      row.getCell(values.length).font = {
        color: { argb: "FFFFFFFB" },
      };
      row.getCell(6).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFed1941" },
      };
      row.getCell(6).font = {
        color: { argb: "FFFFFFFB" },
      };
    }
    if (values[3] !== values[4]) {
      row.getCell(4).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FF426ab3" },
      };
      row.getCell(4).font = {
        color: { argb: "FFFFFFFB" },
      };
      row.getCell(5).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FF426ab3" },
      };
      row.getCell(5).font = {
        color: { argb: "FFFFFFFB" },
      };
    }
  });
  workSheetTwo.eachRow(function (row, rowNumber) {
    const values = row.values;
    values.shift();
    if (rowNumber === 1) return;
    if (values[5] !== values[values.length - 1]) {
      row.getCell(values.length).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFed1941" },
      };
      row.getCell(values.length).font = {
        color: { argb: "FFFFFFFB" },
      };
      row.getCell(6).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFed1941" },
      };
      row.getCell(6).font = {
        color: { argb: "FFFFFFFB" },
      };
    }
    if (values[3] !== values[4]) {
      row.getCell(4).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FF426ab3" },
      };
      row.getCell(4).font = {
        color: { argb: "FFFFFFFB" },
      };
      row.getCell(5).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FF426ab3" },
      };
      row.getCell(5).font = {
        color: { argb: "FFFFFFFB" },
      };
    }
  });
  await workbook.xlsx.writeFile("对账表.xlsx");
}
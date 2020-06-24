const Excel = require("exceljs");
const _ = require("lodash");

let maps = {};

(async function () {
  const workbook = new Excel.Workbook();

  await workbook.xlsx.readFile("./client.xlsx");
  let worksheet1 = workbook.getWorksheet(1);
  await workbook.xlsx.readFile("./contrast.xlsx");
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
      contrast_map_us[row_value[2]] = row_value;
    });
    maps = contrast_map_us;
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
        // values.splice(0, 0, rowNumber);
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
        const price = contrast_map_us[values[0]]
          ? contrast_map_us[values[0]][3]
          : 0;
        values.splice(2, 0, price);
        // values.splice(0, 0, rowNumber);
        if (values[values.length - 1] > 0) {
          us_values_positive.push(values);
        } else {
          us_values_negative.push(values);
        }
      });
      let client_positive_tmp = _.cloneDeep(client_values_positive);
      let client_negative_tmp = _.cloneDeep(client_values_negative);
      let us_positive_tmp = _.cloneDeep(us_values_positive);
      let us_negative_tmp = _.cloneDeep(us_values_negative);
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

//转化数组为Object对象最后根据数量倒叙排序
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
      return b[b.length - 1] - a[a.length - 1];
    });
  });
  return final_obj;
}


// 合并sheet页计算
function countMergeArray(
  client_values_positive,
  client_values_negative,
  us_values_positive,
  us_values_negative
) {
    // 合并客户正数数组
    let client_merge_positive = mergeCount(client_values_positive, 0, 3);
    // 合并客户负数数组
    let client_merge_negative = mergeCount(client_values_negative, 0, 3);
    // 合并公司正数数组
    let us_merge_positive = mergeCount(us_values_positive, 0, 2);
    // 合并公司负数数组
    let us_merge_negative = mergeCount(us_values_negative, 0, 2);
    let client_merge_positive_map = turnArrayToObject(client_merge_positive);
    let client_merge_negative_map = turnArrayToObject(client_merge_negative);
    let us_merge_positive_map = turnArrayToObject(us_merge_positive);
    let us_merge_negative_map = turnArrayToObject(us_merge_negative);

    // 合并公司与客户的正数数组
    let merge_positive_array = signCountAllValues(
      client_merge_positive_map,
      us_merge_positive_map
    );
    // 合并公司与客户的负数数组
    let merge_negative_array = signCountAllValues(
      client_merge_negative_map,
      us_merge_negative_map
    );

    let merge_final_array = merge_positive_array.concat(merge_negative_array);
    let merge_final_map = {};
    /**
     * 因为数组是乱序
     * 偷懒没用任何算法
     * 直接先转换成object对象
     * 再通过Object.values转化成数组
     */
    merge_final_map = turnArrayToObject(merge_final_array);
    let original_merge_array = [];
    Object.values(merge_final_map).forEach((array) => {
      for (let arr of array) {
        original_merge_array.push(arr);
      }
    });
    return original_merge_array;
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

//核心算法
function signCountAllValues(client_group_map, us_group_map) {
  let client_group_map_tmp = _.cloneDeep(client_group_map); //深拷贝对象，不影响原有的数据
  let us_group_map_tmp = _.cloneDeep(us_group_map);
  let tableData = [];
  Object.keys(client_group_map_tmp).map((key) => { //根据对象key循环
    let client_array = client_group_map_tmp[key];
    // 当客户拥有对应型号数据而公司数据里没有时
    // 根据当前key值查出来的数据每条合并对应型号，对照表中的单价，数量为0
    let us_array = us_group_map_tmp[key];
    if (!us_array) {
      for (let i of client_array) {
        let array = i.concat([key, maps[key][3], 0]);
        tableData.push(array);
      }
    } else {
      // 两边数量相等
      if (client_array.length === us_array.length) {
        let client_array_back = _.cloneDeep(client_array);
        let us_array_back = _.cloneDeep(us_array);
        for (let i in client_array_back) {
          let array = [];
          if (us_array_back[i]) {
            us_array_back[i].splice(2, 1);
            array = client_array_back[i].concat(us_array_back[i]);
          } else {
            array = client_array_back[i].concat([key, maps[key][3], 0]);
          }
          tableData.push(array);
        }
      } else if (client_array.length > us_array.length) { //客户的订单数量大于公司的订单数量
        let client_array_back = _.cloneDeep(client_array);
        let us_array_back = _.cloneDeep(us_array);
        let array = [];
        for (let i = 0; i < us_array_back.length; i++) {
          us_array_back[i].splice(2, 1);
          array = client_array_back[i].concat(us_array_back[i]);
          tableData.push(array);
        }
        let cut_array = client_array_back.splice(us_array_back.length);
        for (let i of cut_array) {
          array = i.concat([key, maps[key][3], 0]);
          tableData.push(array);
        }
      } else if (client_array.length < us_array.length) { // 客户的订单数量小于公司的订单数量
        let client_array_back = _.cloneDeep(client_array);
        let us_array_back = _.cloneDeep(us_array);
        for (let i = 0; i < client_array_back.length; i++) {
          us_array_back[i].splice(2, 1);
          array = client_array_back[i].concat(us_array_back[i]);
          tableData.push(array);
        }
        let cut_array = us_array_back.splice(client_array_back.length);
        for (let i of cut_array) {
          i.splice(2, 1);
          let client = [
            i[0],
            maps[i[0]][1],
            maps[i[0]][0],
            0,
            maps[i[0]][3],
            0,
            ...i,
          ];
          tableData.push(client);
        }
      }
      delete us_group_map_tmp[key];
    }
    delete client_group_map_tmp[key];
  });


  /**
   * 完成筛选后还遗留下来的数据进行另外的操作
   */
  const left_client_map = Object.keys(client_group_map_tmp);
  const left_us_map = Object.keys(us_group_map_tmp);
  if (left_client_map.length > 0) {
    left_client_map.forEach((key) => {
      left_client_map[key].forEach((item) => {
        item.splice(0, 1);
        let arr = item.concat([key, maps[key][3], 0]);
        tableData.push(arr);
      });
    });
  }
  if (left_us_map.length > 0) {
    left_us_map.forEach((key) => {
      us_group_map_tmp[key].forEach((item) => {
        let backup_item = _.cloneDeep(item);
        backup_item.splice(2, 1);
        let client_number = maps[backup_item[0]][1];
        let client_specifications = maps[backup_item[0]][0];
        let client_price = maps[backup_item[0]][3];
        let arr = [
          backup_item[0],
          client_number,
          client_specifications,
          0,
          client_price,
          0,
          ...backup_item,
        ];
        tableData.push(arr);
      });
    });
  }
  return tableData;
}

async function createNewWorkbook(tableHeader, merge_data, single_data) {
  const workbook = new Excel.Workbook();
  const workSheetOne = workbook.addWorksheet("合并", {
    properties: {
      defaultColWidth: 20,
    },
  });
  const workSheetTwo = workbook.addWorksheet("单行", {
    properties: {
      defaultColWidth: 20,
    },
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
        fgColor: {
          argb: "FFed1941",
        },
      };
      row.getCell(values.length).font = {
        color: {
          argb: "FFFFFFFB",
        },
      };
      row.getCell(6).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: {
          argb: "FFed1941",
        },
      };
      row.getCell(6).font = {
        color: {
          argb: "FFFFFFFB",
        },
      };
    }
    if (values[3] !== values[4]) {
      row.getCell(4).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: {
          argb: "FF426ab3",
        },
      };
      row.getCell(4).font = {
        color: {
          argb: "FFFFFFFB",
        },
      };
      row.getCell(5).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: {
          argb: "FF426ab3",
        },
      };
      row.getCell(5).font = {
        color: {
          argb: "FFFFFFFB",
        },
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
        fgColor: {
          argb: "FFed1941",
        },
      };
      row.getCell(values.length).font = {
        color: {
          argb: "FFFFFFFB",
        },
      };
      row.getCell(6).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: {
          argb: "FFed1941",
        },
      };
      row.getCell(6).font = {
        color: {
          argb: "FFFFFFFB",
        },
      };
    }
    if (values[3] !== values[4]) {
      row.getCell(4).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: {
          argb: "FF426ab3",
        },
      };
      row.getCell(4).font = {
        color: {
          argb: "FFFFFFFB",
        },
      };
      row.getCell(5).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: {
          argb: "FF426ab3",
        },
      };
      row.getCell(5).font = {
        color: {
          argb: "FFFFFFFB",
        },
      };
    }
  });
  await workbook.xlsx.writeFile("对账表.xlsx");
}

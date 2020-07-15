module.exports = {
  includeColumn:['销售订单号','客户型号','本厂型号','规格','订单数量','销售下单'],
  cut_index: [0, 1, 2, 4, 5, 7],
  c1: {
    "0201(0603)": "A",
    "0402(1005)": "B",
    "0603(1608)": "C",
    "0805(2012)": "D",
    "1206(3216)": "E",
    "1210(3225)": "F",
    "1615": "G",
    "2835": "H",
    "5050": "J",
  },
  c2: {
    X5R: "D",
    X7R: "E",
    Y5V: "P",
    NP0: "K",
    NPO: "K",
    C0G: "K",
    COG: "K",
  },
  c4: {
    "6.3V": "A",
    "10V": "B",
    "16V": "C",
    "25V": "D",
    "35V": "E",
    "50V": "F",
    "63V": "G",
    "100V": "H",
    "250V": "J",
    "500V": "K",
    "630V": "L",
    "1000V": "M",
    "1600V": "N",
    "400V": "P",
  },
  c5: {
    "±0.1PF": "A",
    "±0.25PF": "B",
    "±0.5PF": "C",
    "±1%": "D",
    "±2%": "E",
    "±5%": "F",
    "±10%": "G",
    "±20%": "H",
    "+50%/-20%": "J",
    "+80%/-20%": "K",
  },
  turnCapacitance: (origin) => {
    const reg1 = /^([0-9]+|0)$/;
    const reg2 = /^([0-9]+|0)(\.(([0-9][1-9])|[1-9]{1,2}))$/;

    let num = 0;
    let unit = origin.substr(-2, 1).toLowerCase();
    let count = origin.slice(0, -2);
    switch (unit) {
      case "u":
        num = 10 ** 6;
        break;
      case "n":
        num = 10 ** 3;
        break;
      default:
        break;
    }
    let pfnum = "";
    if (unit === "p") {
      let number = parseFloat(count);
      if (number < 10) {
        if (reg1.test(number)) {
          pfnum = number.toString() + "P" + 0;
        } else {
          let arr = number.toString().split(".");
          pfnum = arr[0] + "P" + arr[1];
        }
      } else if (number < 100) {
        pfnum = number.toString() + 0;
      } else {
        pfnum = number.toString().substr(0, 2) + 1;
      }
    } else {
      const final = (num * count).toString();
      let end = final.slice(2).length;
      let start = final.substr(0, 2);
      pfnum = start + end;
    }
    return pfnum;
  },
};

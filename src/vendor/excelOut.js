/* eslint-disable */
require('script-loader!file-saver');
import XLSX from 'xlsx-style-luleiyu'

function generateArray(table) {
  var out = [];
  var rows = table.querySelectorAll('tr');
  var ranges = [];
  for (var R = 0; R < rows.length; ++R) {
    var outRow = [];
    var row = rows[R];
    var columns = row.querySelectorAll('td');
    for (var C = 0; C < columns.length; ++C) {
      var cell = columns[C];
      var colspan = cell.getAttribute('colspan');
      var rowspan = cell.getAttribute('rowspan');
      var cellValue = cell.innerText;
      if (cellValue !== "" && cellValue == +cellValue) cellValue = +cellValue;

      //Skip ranges
      ranges.forEach(function (range) {
        if (R >= range.s.r && R <= range.e.r && outRow.length >= range.s.c && outRow.length <= range.e.c) {
          for (var i = 0; i <= range.e.c - range.s.c; ++i) outRow.push(null);
        }
      });

      //Handle Row Span
      if (rowspan || colspan) {
        rowspan = rowspan || 1;
        colspan = colspan || 1;
        ranges.push({
          s: {
            r: R,
            c: outRow.length
          },
          e: {
            r: R + rowspan - 1,
            c: outRow.length + colspan - 1
          }
        });
      }
      ;

      //Handle Value
      outRow.push(cellValue !== "" ? cellValue : null);

      //Handle Colspan
      if (colspan)
        for (var k = 0; k < colspan - 1; ++k) outRow.push(null);
    }
    out.push(outRow);
  }
  return [out, ranges];
};

function datenum(v, date1904) {
  if (date1904) v += 1462;
  var epoch = Date.parse(v);
  return (epoch - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
}

function sheet_from_array_of_arrays(data, opts) {
  var ws = {};
  var range = {
    s: {
      c: 10000000,
      r: 10000000
    },
    e: {
      c: 0,
      r: 0
    }
  };

  for (var R = 0; R != data.length; ++R) {
    for (var C = 0; C != data[R].length; ++C) {
      if (range.s.r > R) range.s.r = R;
      if (range.s.c > C) range.s.c = C;
      if (range.e.r < R) range.e.r = R;
      if (range.e.c < C) range.e.c = C;
      var cell = {
        v: data[R][C]
      };
      if (cell.v == null) continue;
      var cell_ref = XLSX.utils.encode_cell({
        c: C,
        r: R
      });

      if (typeof cell.v === 'number') cell.t = 'n';
      else if (typeof cell.v === 'boolean') cell.t = 'b';
      else if (cell.v instanceof Date) {
        cell.t = 'n';
        cell.z = XLSX.SSF._table[14];
        cell.v = datenum(cell.v);
      } else cell.t = 's';

      ws[cell_ref] = cell;
    }
  }
  if (range.s.c < 10000000) ws['!ref'] = XLSX.utils.encode_range(range);
  return ws;
}

function Workbook() {
  if (!(this instanceof Workbook)) return new Workbook();
  this.SheetNames = [];
  this.Sheets = {};
}

function s2ab(s) {
  var buf = new ArrayBuffer(s.length);
  var view = new Uint8Array(buf);
  for (var i = 0; i != s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
  return buf;
}

export function export_table_to_excel(id) {
  var theTable = document.getElementById(id);
  var oo = generateArray(theTable);
  var ranges = oo[1];

  /* original data */
  var data = oo[0];
  var ws_name = "SheetJS";

  var wb = new Workbook(),
    ws = sheet_from_array_of_arrays(data);

  /* add ranges to worksheet */
  // ws['!cols'] = ['apple', 'banan'];
  ws['!merges'] = ranges;

  /* add worksheet to workbook */
  wb.SheetNames.push(ws_name);
  wb.Sheets[ws_name] = ws;

  var wbout = XLSX.write(wb, {
    bookType: 'xlsx',
    bookSST: false,
    type: 'binary'
  });
  wbout.then(res => {
    saveAs(new Blob([s2ab(res)], {
      type: "application/octet-stream"
    }), "test.xlsx")
  })
}

export function exportJsonToExcel({
   title,
   header,
   data,
   filename,
   merges = [],
   autoWidth = true,
   bookType = 'xlsx',
   myRowFont = '1',
   lastColFlag = false,
   multiHeader = [],
  } = {}) {
  /* original data */
  filename = filename || 'excel-list'
  data = [...data]
  data.unshift(header);
  if (merges.length > 0) { // 控制了是否合并单元格，是否出现表头
    data.unshift(title);
    for (let i = multiHeader.length - 1; i > -1; i--) {
      data.unshift(multiHeader[i])
    }
  }

  var ws_name = "SheetJS";
  var wb = new Workbook(),
    ws = sheet_from_array_of_arrays(data);

  if (merges.length > 0) {
    if (!ws['!merges']) ws['!merges'] = [];
    merges.forEach(item => {
      ws['!merges'].push(XLSX.utils.decode_range(item))
    })
  }

  if (autoWidth) {
    /*设置worksheet每列的最大宽度*/
    const colWidth = data.map(row => row.map(val => {
      /*先判断是否为null/undefined*/
      if (val == null) {
        return {
          'wch': 10
        };
      }
      /*再判断是否为中文*/
      else if (val.toString().charCodeAt(0) > 255) {
        return {
          'wch': val.toString().length * 2
        };
      } else {
        return {
          'wch': val.toString().length
        };
      }
    }))

    /*以第一行为初始值*/
    let result = colWidth[0];
    for (let i = 1; i < colWidth.length; i++) {
      for (let j = 0; j < colWidth[i].length; j++) {
        if (result[j]['wch'] < colWidth[i][j]['wch']) {
          result[j]['wch'] = colWidth[i][j]['wch'];
        }
      }
    }
    ws['!cols'] = result;
    //每一行的高度
    const colHeight = data.map(row => row.map(val => {
      /*先判断是否为null/undefined*/
      if (val == null) {
        return {
          'hpx': 40
        };
      }
      /*再判断是否为中文*/
      else if (val.toString().charCodeAt(0) > 255) {
        return {
          'hpx': 40
        };
      } else {
        return {
          'hpx': 40
        };
      }
    }))

    /*以第一行为初始值*/
    let result1 = [];
    for (let i = 0; i < colHeight.length; i++) {
      result1.push({
        hpx: 40
      })
    }

    ws['!rows'] = result1;
  }

  /* add worksheet to workbook */
  wb.SheetNames.push(ws_name);
  wb.Sheets[ws_name] = ws;
  var dataInfo = wb.Sheets[wb.SheetNames[0]];

  const borderAll = {  //单元格外侧框线
    top: {
      style: 'thin'
    },
    bottom: {
      style: 'thin'
    },
    left: {
      style: 'thin'
    },
    right: {
      style: 'thin'
    }
  };

  var arr = ["A1", "B1", "C1", "D1", "E1", "F1", "G1", "H1", "I1", "J1", "K1", "L1", "M1", "N1", "O1", "P1", "Q1", "R1", "S1", "T1", "U1", "V1", "W1", "X1", "Y1", "Z1"];

  if (merges.length > 0) {//暂时没用
    // 去掉标题边框
    arr.some(v => {
      let a = merges[0].split(':')
      if (v == a[1]) {
        dataInfo[v].s = {
          border: borderAll
        }
        return true;
      } else {
        dataInfo[v].s = {}
      }
    })
  }

  //设置主标题样式字体样式的方法
  function themeStyle (someCols,fontF,fontC,fontW,bgc) {
    dataInfo[someCols].s = {
      border: borderAll,
      font: {
        name: fontF,
        sz: 12,
        color: {rgb: fontC},
        bold: fontW,
        italic: false,
        underline: false
      },
      alignment: {
        horizontal: "center",
        vertical: "center"
      },
      fill: {
        fgColor: {rgb: bgc}
      },
    };
  }
  arr.forEach((item,index) => {
    if (item == 'A1') {//表头A1的样式
      themeStyle('A1','宋体','ffffff',true,'C0504D');
    } else {
      if (item.indexOf('1') > -1 && index <= (data[0].length-1)) { //表头除了A1 从B1开始整个一行的样式
        themeStyle(item,'微软雅黑','C00000',true,'8DB4E2');
        if (lastColFlag) {// 控制第一行最后一个格的字体
          if (index == (data[0].length-1)) {
            themeStyle(item,'微软雅黑','000000',true,'8DB4E2');
          }
        }
      }
    }
  });

  // 这是表头行的样式
  var tableTitleFontA = {//A2开始的这个A列的，所有样式
    border: borderAll,
    font: {
      name: '宋体',
      sz: 12,
      color: {rgb: "FF6600"},
      bold: false,
      italic: false,
      underline: false
    },
    alignment: {
      horizontal: "center",
      vertical: "center"
    },
    fill: {
      // fgColor: {rgb: "C0504D"},
    },
  },
  tableTitleFontB = {//B2纵向，横向，字体样式
    border: borderAll,
    font: {
      name: '微软雅黑',
      sz: 12,
      color: {rgb: "000000"},
      bold: false,
      italic: false,
      underline: false
    },
    alignment: {
      horizontal: "center",
      vertical: "center"
    },
    fill: {
      // fgColor: {rgb: "C0504D"},
    },
  };
  // 控制了除第一个行的其他行的样式
  for (var b in dataInfo) {
    if (b.indexOf('A') > -1 && b.substr(1) > 1) {//控制了除去A1的样式，其他所有包含A的列的样式
      dataInfo[b].s = tableTitleFontA;
    } else if (b.substr(1) > 1 && b.indexOf('A') == -1) { // 控制了从B2开始，横向，纵向的，所有的字体样式
      dataInfo[b].s = tableTitleFontB;
    }
  }

  console.log(dataInfo);

  var wbout = XLSX.write(wb, {
    bookType: bookType,
    bookSST: false,
    type: 'binary'
  });

  wbout.then(res => {
    saveAs(new Blob([s2ab(res)], {
      type: "application/octet-stream"
    }), `${filename}.${bookType}`);
  })
}

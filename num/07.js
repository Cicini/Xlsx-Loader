// 从网络上读取某个excel文件，url必须同域，否则报错
function readworkbook07FromRemoteFile(url, callback) {
  var xhr = new XMLHttpRequest();
  xhr.open("get", url, true);
  xhr.responseType = "arraybuffer";
  xhr.onload = function (e) {
    if (xhr.status == 200) {
      var data = new Uint8Array(xhr.response);
      var workbook07 = XLSX.read(data, { type: "array" });
      if (callback) callback(workbook07);
    }
  };
  xhr.send();
}

// 读取 excel文件
function outputworkbook07(workbook07) {
  var sheetNames = workbook07.SheetNames; // 工作表名称集合
  sheetNames.forEach((name) => {
    var worksheet = workbook07.Sheets[name]; // 只能通过工作表名称来获取指定工作表
    for (var key in worksheet) {
      // v是读取单元格的原始值
      console.log(key, key[0] === "!" ? worksheet[key] : worksheet[key].v);
    }
  });
}

function readworkbook07(workbook07) {
  var sheetNames = workbook07.SheetNames; // 工作表名称集合
  var worksheet = workbook07.Sheets[sheetNames[0]]; // 这里我们只读取第一张sheet
  var csv = XLSX.utils.sheet_to_csv(worksheet);
  document.getElementById("result07").innerHTML = csv2table(csv);

  // 合并单元格
  mergeTable(workbook07, true);
}

// 将csv转换成表格
function csv2table(csv) {
  var html = "<table class=excel>";
  var rows = csv.split("\n");
  rows.pop(); // 最后一行没用的
  rows.forEach(function (row, idx) {
    var columns = row.split(",");
    html += "<tr>";
    columns.forEach(function (column) {
      html += "<td>" + column + "</td>";
    });
    html += "</tr>";
  });
  html += "</table>";
  return html;
}

// 合并单元格
function mergeTable(workbook07) {
  let SheetNames = workbook07.SheetNames[0];
  let mergeInfo = workbook07.Sheets[SheetNames]["!merges"];
  let result = document.getElementById("result07");

  mergeInfo.forEach((item07) => {
    let start_r = item07.s.r;
    let end_r = item07.e.r;

    let start_c = item07.s.c;
    let end_c = item07.e.c;

    for (let i = start_r; i <= end_r; i++) {
      let row = document.querySelectorAll("#result07 tr")[i];
      for (let child = start_c; child <= end_c; child++) {
        if (child === start_c && i === start_r) {
          // 循环到就是第一个单元格，以这个单元格为开始进行合并
          row.children[child].classList.add("will_span");
          row.children[child].setAttribute("row", end_r - start_r + 1);
          row.children[child].setAttribute("col", end_c - start_c + 1);
          row.children[child].setAttribute("align", "center");
        } else {
          // 只做标记，不在这里删除
          row.children[child].classList.add("remove");
        }
      }
    }
  });

  // 移除对应的td
  document.querySelectorAll(".remove").forEach((item07) => {
    item07.parentNode.removeChild(item07);
  });

  // 为大的td设置跨单元格合并
  document.querySelectorAll(".will_span").forEach((item07) => {
    item07.classList.remove("will_span");
    item07.rowSpan = item07.getAttribute("row");
    item07.colSpan = item07.getAttribute("col");
  });
}

function loadRemoteFile(url) {
  readworkbook07FromRemoteFile(url, function (workbook07) {
    readworkbook07(workbook07);
  });
}

//自定义html标签
class XlsxRender07 extends HTMLElement {
  constructor() {
    super();
    loadRemoteFile(this.getAttribute("content"));
  }
}
customElements.define("xlsx-render07", XlsxRender07);

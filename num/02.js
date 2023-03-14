// 从网络上读取某个excel文件，url必须同域，否则报错
function readworkbook02FromRemoteFile(url, callback) {
  var xhr = new XMLHttpRequest();
  xhr.open("get", url, true);
  xhr.responseType = "arraybuffer";
  xhr.onload = function (e) {
    if (xhr.status == 200) {
      var data = new Uint8Array(xhr.response);
      var workbook02 = XLSX.read(data, { type: "array" });
      if (callback) callback(workbook02);
    }
  };
  xhr.send();
}

// 读取 excel文件
function outputworkbook02(workbook02) {
  var sheetNames = workbook02.SheetNames; // 工作表名称集合
  sheetNames.forEach((name) => {
    var worksheet = workbook02.Sheets[name]; // 只能通过工作表名称来获取指定工作表
    for (var key in worksheet) {
      // v是读取单元格的原始值
      console.log(key, key[0] === "!" ? worksheet[key] : worksheet[key].v);
    }
  });
}

function readworkbook02(workbook02) {
  var sheetNames = workbook02.SheetNames; // 工作表名称集合
  var worksheet = workbook02.Sheets[sheetNames[0]]; // 这里我们只读取第一张sheet
  var csv = XLSX.utils.sheet_to_csv(worksheet);
  document.getElementById("result02").innerHTML = csv2table(csv);

  // 合并单元格
  mergeTable(workbook02, true);
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
function mergeTable(workbook02) {
  let SheetNames = workbook02.SheetNames[0];
  let mergeInfo = workbook02.Sheets[SheetNames]["!merges"];
  let result = document.getElementById("result02");

  mergeInfo.forEach((item02) => {
    let start_r = item02.s.r;
    let end_r = item02.e.r;

    let start_c = item02.s.c;
    let end_c = item02.e.c;

    for (let i = start_r; i <= end_r; i++) {
      let row = document.querySelectorAll("#result02 tr")[i];
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
  document.querySelectorAll(".remove").forEach((item02) => {
    item02.parentNode.removeChild(item02);
  });

  // 为大的td设置跨单元格合并
  document.querySelectorAll(".will_span").forEach((item02) => {
    item02.classList.remove("will_span");
    item02.rowSpan = item02.getAttribute("row");
    item02.colSpan = item02.getAttribute("col");
  });
}

function loadRemoteFile(url) {
  readworkbook02FromRemoteFile(url, function (workbook02) {
    readworkbook02(workbook02);
  });
}

//自定义html标签
class XlsxRender02 extends HTMLElement {
  constructor() {
    super();
    loadRemoteFile(this.getAttribute("content"));
  }
}
customElements.define("xlsx-render02", XlsxRender02);

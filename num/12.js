// 从网络上读取某个excel文件，url必须同域，否则报错
function readworkbook12FromRemoteFile(url, callback) {
  var xhr = new XMLHttpRequest();
  xhr.open("get", url, true);
  xhr.responseType = "arraybuffer";
  xhr.onload = function (e) {
    if (xhr.status == 200) {
      var data = new Uint8Array(xhr.response);
      var workbook12 = XLSX.read(data, { type: "array" });
      if (callback) callback(workbook12);
    }
  };
  xhr.send();
}

// 读取 excel文件
function outputworkbook12(workbook12) {
  var sheetNames = workbook12.SheetNames; // 工作表名称集合
  sheetNames.forEach((name) => {
    var worksheet = workbook12.Sheets[name]; // 只能通过工作表名称来获取指定工作表
    for (var key in worksheet) {
      // v是读取单元格的原始值
      console.log(key, key[0] === "!" ? worksheet[key] : worksheet[key].v);
    }
  });
}

function readworkbook12(workbook12) {
  var sheetNames = workbook12.SheetNames; // 工作表名称集合
  var worksheet = workbook12.Sheets[sheetNames[0]]; // 这里我们只读取第一张sheet
  var csv = XLSX.utils.sheet_to_csv(worksheet);
  document.getElementById("result12").innerHTML = csv2table(csv);

  // 合并单元格
  mergeTable(workbook12, true);
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
function mergeTable(workbook12) {
  let SheetNames = workbook12.SheetNames[0];
  let mergeInfo = workbook12.Sheets[SheetNames]["!merges"];
  let result = document.getElementById("result12");

  mergeInfo.forEach((item12) => {
    let start_r = item12.s.r;
    let end_r = item12.e.r;

    let start_c = item12.s.c;
    let end_c = item12.e.c;

    for (let i = start_r; i <= end_r; i++) {
      let row = document.querySelectorAll("#result12 tr")[i];
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
  document.querySelectorAll(".remove").forEach((item12) => {
    item12.parentNode.removeChild(item12);
  });

  // 为大的td设置跨单元格合并
  document.querySelectorAll(".will_span").forEach((item12) => {
    item12.classList.remove("will_span");
    item12.rowSpan = item12.getAttribute("row");
    item12.colSpan = item12.getAttribute("col");
  });
}

function loadRemoteFile(url) {
  readworkbook12FromRemoteFile(url, function (workbook12) {
    readworkbook12(workbook12);
  });
}

//自定义html标签
class XlsxRender12 extends HTMLElement {
  constructor() {
    super();
    loadRemoteFile(this.getAttribute("content"));
  }
}
customElements.define("xlsx-render12", XlsxRender12);

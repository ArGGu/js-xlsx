import XLSX from "xlsx-style"
// 配置
const wopts = {
  bookType: "xlsx",
  bookSST: true,
  type: "binary",
  cellStyles: true
};
const cellStyle = {
  font: { sz: 10, name: "微软雅黑" },
  alignment: {
    horizontal: "center",
    wrapText: true,
    vertical: "center"
  }
}
function downloadExl(json, config) {
  if (config.title) json = config.title.concat(json)
  if (config.ignoreLine) for (let i = 0; i < config.ignoreLine; i++) json = config.title.concat(json)

  var tmpdata = json[0];
  var keyMap = []; //获取keys
  for (var k in tmpdata) {
    keyMap.push(k);
  }
  tmpdata = []; //用来保存转换好的json
  json
  .map((v, i) =>
    keyMap.map((k, j) =>
      Object.assign(
        {},
        {
          v: v[k],
          position: (j > 25 ? getCharCol(j) : String.fromCharCode(65 + j)) + (i + 1)
        }
      )
    )
  )
  .reduce((prev, next) => prev.concat(next))
  .forEach((v, i) => {
    // 在遍历时修改样式
    let s = cellStyle;
    // 判断是否是标题
    if (config.title && config.title.includes(v.v) && v.position.length == 2 && v.position[1] == "1") {
      s.fill = { fgColor: { rgb: "008000" } };
      s.border = { right: { style: "thin", color: { rgb: "000000" } } };
      // 判断是否是需要红色字体的标题字段
      if (config.redTitle.includes(v.v)) {
        s.font.color = { rgb: "ff0000" };
      }
    }
    tmpdata[v.position] = { v: v.v || "", s};
  });
  var outputPos = Object.keys(tmpdata); //设置区域,比如表格从A1到D10
  //设置每列对应的宽度
  if (config.colwidth) tmpdata["!cols"] = config.colwidth;
  if (config.colHeight) tmpdata['!rows'] = config.colHeight
  if (config.merges) tmpdata['!merges'] = config.merges
  if (config.header) {
    for ( let [key, val] of Object.entries(config.header)) {
      tmpdata[key] = val
      tmpdata[key].s = val.s ? val.s : cellStyle
    } 
  }
  console.log(tmpdata)
  var tmpWB = {
    SheetNames: ["mySheet"], //保存的表标题
    Sheets: {
      mySheet: Object.assign(
        {},
        tmpdata, //内容
        {
          "!ref": outputPos[0] + ":" + outputPos[outputPos.length - 1] //设置填充区域
        }
      )
    }
  };
  const tmpDown = new Blob(
    [
      s2ab(
        XLSX.write(
          tmpWB,
          {
            bookType: "xlsx",
            bookSST: false,
            type: "binary"
          } //这里的数据是用来定义导出的格式类型
        )
      )
    ],
    {
      type: ""
    }
  );
  // 数据处理完后传入下载
  saveAs( tmpDown, config.fileName + "." + (wopts.bookType == "biff2" ? "xls" : wopts.bookType));
}
// 获取26个英文字母用来表示excel的列
function getCharCol(n) {
  let s = "",
    m = 0;
  while (n > 0) {
    m = (n % 26) + 1;
    s = String.fromCharCode(m + 64) + s;
    n = (n - m) / 26;
  }
  return s;
}
function s2ab(s) {
  let buf
  if (typeof ArrayBuffer !== "undefined") {
    buf = new ArrayBuffer(s.length);
    var view = new Uint8Array(buf);
    for (let i = 0; i != s.length; ++i) view[i] = s.charCodeAt(i) & 0xff;
    return buf;
  } else {
    buf = new Array(s.length);
    for (let i = 0; i != s.length; ++i) buf[i] = s.charCodeAt(i) & 0xff;
    return buf;
  }
}
// 下载功能
function saveAs(obj, fileName) {
  var tmpa = document.createElement("a");
  tmpa.download = fileName || "未命名";
  // 兼容ie 火狐 下载文件
  if ("msSaveOrOpenBlob" in navigator) {
    window.navigator.msSaveOrOpenBlob(obj, fileName);
  } else if (window.navigator.userAgent.includes("Firefox")) {
    var a = document.createElement("a");
    a.href = URL.createObjectURL(obj);
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
  } else {
    tmpa.href = URL.createObjectURL(obj);
  }
  tmpa.click();
  setTimeout(function () {
    URL.revokeObjectURL(obj);
  }, 100);
}
export default downloadExl;
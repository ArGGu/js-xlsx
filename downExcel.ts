import XLSX from './xlsx'

// 配置
const wopts = {
  bookType: 'xlsx',
  bookSST: true,
  type: 'binary',
  cellStyles: true
}

// 通用单元格样式
const cellStyle = {
  font: { sz: 10, name: '微软雅黑' },
  alignment: {
    wrapText: true,
    vertical: 'center'
  }
}

// 通用标题样式
const titleStyle = {
  font: { sz: 10, name: '微软雅黑' },
  alignment: {
    wrapText: true,
    horizontal: 'center',
    vertical: 'center'
  }
}

let jsonList: any[] = []

// 生成合并单元格数据
const getMerges = (row: any, rowIndex: any, config: any) => {
  let len = config.title.length - 1
  let keys = Object.keys(row)
  if (!config.merges) config.merges = []
  for (let i = 0; i < keys.length - 1; i++) {
    if (+keys[i + 1] - +keys[i] > 1) {
      config.merges.push({
        s: {c: +keys[i], r: rowIndex},
        e: {c: +keys[i + 1] - 1, r: rowIndex}
      })
    }
  }
  let lastIndex = +keys[keys.length - 1]
  if (lastIndex < len) {
    config.merges.push({
      s: {c: +lastIndex, r: rowIndex},
      e: {c: +len, r: rowIndex}
    })
  }
}

const downExcel = (json: any, config: any, showError: any) => {
  jsonList = []
  // 判断是否为多表操作
  if (config.sheetConfigs) {
    for (let i in config.sheetConfigs) {
      formatSheetData(json[i], config.sheetConfigs[i], showError)
    }
  } else {
    formatSheetData(json, config, showError)
  }
  jsonToBolb(config)
}

// sheet页数据格式化
const formatSheetData = (json: any, config: any, showError: any) => {
  json = json.map((row: any, rowIndex: any) => {
    if (!Array.isArray(row)) {
      let arr = []
      for (let [key, val] of Object.entries(row)) {
        arr[+key] = val
      }
      getMerges(row, rowIndex, config)
      return arr
    } else return row
  })
  let tmpdata = config.title || json[0]
  let keyMap : any[]= [] // 获取keys
  for (let k in tmpdata) {
    keyMap.push(k)
  }
  if (showError) keyMap.push(keyMap.length)
  tmpdata = [] // 用来保存转换好的json
  json.map((v: any[], i: any) =>
    keyMap.map((k, j) => {
      let obj : any = {
        v: v[k],
        position: (j > 25 ? getCharCol(j) : String.fromCharCode(65 + j)) + (i + 1)
      }
      if (config.titleRange && config.titleRange.includes(i)) obj['s'] = titleStyle
      return Object.assign({}, obj)
    })
  )
    .reduce((prev: any, next: any) => prev.concat(next))
    .forEach((v: any, i: any) => {
      let content = v.v
      let cell = {v: content || '', s: v.s || cellStyle, t: 's'}
      if (content === 0 || content === '0') {
        cell.v = 0
        cell.t = 'n'
      }
      if (content && !isNaN(content) && String(content).length < 7) { cell.t = 'n' }
      // 在遍历时修改样式
      tmpdata[v.position] = cell
    })
  let outputPos = Object.keys(tmpdata) // 设置区域,比如表格从A1到D10
  // 设置每列对应的宽度
  if (config.colwidth) { tmpdata['!cols'] = config.colwidth }
  if (config.colHeight) { tmpdata['!rows'] = config.colHeight }
  if (config.merges) { tmpdata['!merges'] = config.merges }
  if (config.header) {
    let header: object = config.header
    for (let [key, val] of Object.entries(header)) {
      if (val['v']) tmpdata[key].v = val['v']
      if (!tmpdata[key]) tmpdata[key] = {}
      tmpdata[key].s = val['s'] ? val['s'] : cellStyle
    }
  }
  jsonList.push({tmpdata, sheetConfig: config, outputPos})
}

const jsonToBolb = (config: any) => {
  let SheetNames : any[] = []
  let Sheets : {[key: string]: any} = {}
  let tmpWB = {
    SheetNames : SheetNames, // 保存的表标题
    Sheets: Sheets // 设置填充区域
  }
  jsonList.forEach(({ tmpdata, sheetConfig, outputPos }) => {
    let sheetName = sheetConfig.sheetName || 'mySheet'
    tmpWB.SheetNames.push(sheetName)
    tmpWB.Sheets[sheetName] = Object.assign({}, tmpdata, {'!ref': outputPos[0] + ':' + outputPos[outputPos.length - 1]})
  })
  const tmpDown = new Blob([
    <BlobPart>s2ab(XLSX.write(tmpWB, {
      bookType: 'xlsx',
      bookSST: false,
      type: 'binary'
    }))], {type: ''})
  // 数据处理完后传入下载
  saveAs(tmpDown, config.fileName + '.' + (wopts.bookType === 'biff2' ? 'xls' : wopts.bookType))
}

// 获取26个英文字母用来表示excel的列
const getCharCol = (n:any) => {
  let s = ''
  let m = 0
  while (n > 0) {
    m = (n % 26) + 1
    s = String.fromCharCode(m + 64) + s
    n = (n - m) / 26
  }
  return s
}

const s2ab = (s: any) => {
  let buf = ArrayBuffer
    ? new Uint8Array(new ArrayBuffer(s.length))
    : new Array(s.length)

  for (let i = 0; i !== s.length; ++i) buf[i] = s.charCodeAt(i) & 0xff
  return buf
}

// 下载功能
const saveAs = (obj: any, fileName: any) => {
  let tmpa = document.createElement('a')
  tmpa.download = fileName || '未命名'
  // 兼容ie 火狐 下载文件
  if (navigator.msSaveOrOpenBlob) {
    window.navigator.msSaveOrOpenBlob(obj, fileName)
  } else if (window.navigator.userAgent.includes('Firefox')) {
    let a = document.createElement('a')
    a.href = URL.createObjectURL(obj)
    a.download = fileName
    document.body.appendChild(a)
    a.click()
  } else {
    tmpa.href = URL.createObjectURL(obj)
  }
  tmpa.click()
  setTimeout(function () {
    URL.revokeObjectURL(obj)
  }, 100)
}

export default downExcel

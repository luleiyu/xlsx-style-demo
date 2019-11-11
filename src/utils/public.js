import { exportJsonToExcel } from '@/vendor/excelOut'

export function handleExport (list, tHeader, filename, filterVal, titleList = [], mergeList = [], lastColFlag = false) {
  // const title = [''] // 标题
  const data = formatJson(filterVal, list)
  data.map(item => {
    item.map((i, index) => {
      if (!i) {
        item[index] = ''
      }
    })
  })
  // 空数组和非空数组，出来的表格是不一样的，都是可以自己定制的
  // const merges = ['A1:E1'] //合并单元格的参数，excel表格，分横向是字母A-Z，纵向是数字1-很多，所以A1就代表第一个格子
  // const merges = []
  exportJsonToExcel({
    title: titleList,
    header: tHeader,
    data,
    merges: mergeList,
    filename: filename,
    autoWidth: true,
    bookType: 'xlsx',
    myRowFont: '1',
    lastColFlag: lastColFlag
  })
}

function formatJson (filterVal, jsonData) {
  return jsonData.map(v => filterVal.map(j => v[j]))
}

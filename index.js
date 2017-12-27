const exel = require('excel4node')

const workbook = new exel.Workbook()
const reportOrder = workbook.addWorksheet('Report Order')

const numberFormat = {
  numberFormat: '#,###.##; (#,###.##); -'
}
const borderBlack = {
  style: 'thin',
  color: '000000'
}
const border = {
  top: borderBlack,
  left: borderBlack,
  right: borderBlack,
  bottom: borderBlack
}
const defaultCell = {
  alignment: {
    horizontal: 'center',
    vertical: 'center'
  },
  border
}

const headerCell = {
  ...defaultCell,
  fill: {
    type: 'pattern',
    patternType: 'solid',
    fgColor: 'C5D6EA'
  },
  font: {
    bold: true
  }
}
const subHeaderCell = {
  ...headerCell,
  font: {
    size: 9
  }
}

const headers = ['ไมโลแคน', 'โค๊กขวดแก้ว', 'น้ำปลา', 'Total']
const datas = {
  '05/01/2018': {
    'ไมโลแคน': {
      set: 1029,
      total: 36015
    },
    'โค๊กขวดแก้ว': {
      set: 1,
      total: 15
    },
    'น้ำปลา': {
      set: 19283,
      total: 366377
    }
  },
  '06/01/2018': {
    'ไมโลแคน': {
      set: 912,
      total: 31920
    },
    'โค๊กขวดแก้ว': {
      set: 22918,
      total: 343770
    },
    'น้ำปลา': {
      set: 21025,
      total: 399475
    }
  },
  '07/01/2018': {
    'ไมโลแคน': {
      set: 0,
      total: 0
    },
    'โค๊กขวดแก้ว': {
      set: 129,
      total: 1935
    },
    'น้ำปลา': {
      set: 1,
      total: 19
    }
  }
}

/**
 * Set Headers
 */
reportOrder.row(1).setHeight(20)
reportOrder.row(2).setHeight(30)
reportOrder.column(1).setWidth(20)
reportOrder.cell(1, 1, 2, 1, true).string('Date').style(headerCell)

/**
 * Generate headers
 */
let columnNo = 2 // start insert header in this column number
let columnPoints
headers.map(header => {
  const set = columnNo
  const total = (columnNo + 1)
  if (header !== 'Total') {
    columnPoints = {
      ...columnPoints,
      [header]: {
        set,
        total
      }
    }
  }
  reportOrder.cell(1, set, 1, total, true).string(header).style(headerCell)
  reportOrder.cell(2, set).string('Set').style(subHeaderCell)
  reportOrder.cell(2, total).string('Total (THB)').style(subHeaderCell)
  columnNo = (columnNo + 2)
})

let rowNo = 3 // start insert data in this row number
for (let date of Object.keys(datas)) {
  reportOrder.cell(rowNo, 1).string(date)
  let sumSet = 0
  let sumeTotal = 0
  for (let kind of Object.keys(datas[date])) {
    reportOrder.cell(rowNo, columnPoints[kind].set).number(datas[date][kind].set).style(numberFormat)
    reportOrder.cell(rowNo, columnPoints[kind].total).number(datas[date][kind].total).style(numberFormat)
    sumSet += datas[date][kind].set
    sumeTotal += datas[date][kind].total
  }
  reportOrder.cell(rowNo, (columnNo - 2)).number(sumSet).style(numberFormat)
  reportOrder.cell(rowNo, (columnNo - 1)).number(sumeTotal).style(numberFormat)
  rowNo++
}

workbook.write('reports/test.xlsx')
var Excel = require('exceljs');
const path = require('path')
const fs = require('fs')

var filename = path.join('merged' + new Date().toUTCString().replace(/[\s,:]/g, "_") + '.xlsx')
var options = {
    filename
};
var workbook = new Excel.stream.xlsx.WorkbookWriter(options);
var worksheet = workbook.addWorksheet('merged')

function run() {
    let filePath = 'e:\\\\assay\\data\\m\\capital-debt\\capital-debt.xlsx'
    filterMonth(filePath)
}
async function filterMonth(filePath) {

    var workbook1 = new Excel.Workbook();
    await workbook1.xlsx.readFile(filePath)
    let worksheet1 = workbook1.getWorksheet(1)
    worksheet1.eachRow((row, rowNumber) => {
        if (rowNumber < 4) {
            worksheet.addRow(row.values).commit()
        } else {
            let values = row.values
            let reg = /12-31/
            if (reg.test(row.getCell(2).value)) {
                worksheet.addRow(row.values).commit()
            }
        }
    })
    workbook.commit()
}
async function merge() {
    var workbook1 = new Excel.Workbook();
    await workbook1.xlsx.readFile('e:\\assay\\data\\m\\ipo-time\\IPO_Cobasic.xlsx')
    let worksheet1 = workbook1.getWorksheet(1)
    worksheet1.eachRow((row, rowNumber) => {
        let values = row.values
    })
}
module.exports = run
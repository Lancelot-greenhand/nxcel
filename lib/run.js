var Excel = require('exceljs');
const path = require('path')
const fs = require('fs')



function run() {
    let target = '/Users/losingyoung/assay/data/m'
    mergeTwoTypes(target)
    // mergeTwoType(target)  IPO and other
    // let filePath = '/Users/losingyoung/assay/data/m/total-capital-growth/merged.xlsx'
    // filterMonth(filePath)
}
async function mergeTwoTypes(target) {
    let first =  '/Users/losingyoung/assay/data/m/merged_capitaldebt.xlsx'
    let second = '/Users/losingyoung/assay/data/m/capital-total/filtered.xlsx'
    var filename = path.join(target, 'merged_' + new Date().toUTCString().replace(/[\s,:]/g, "_") + '.xlsx')
    console.log('filename', filename)
    var options = {
     filename
    };
    var workbook = new Excel.stream.xlsx.WorkbookWriter(options);
    var worksheet = workbook.addWorksheet('merged')
    
    var workbook1 = new Excel.Workbook();
    console.log('workbook1 start reading')
    await workbook1.xlsx.readFile(first)
    console.log('workbook1读取完毕')
    let worksheet1 = workbook1.getWorksheet(1)

    var workbook2 = new Excel.Workbook();
    console.log('workbook2 start reading')
    await workbook2.xlsx.readFile(second)
    console.log('workbook2读取完毕')
    let worksheet2 = workbook2.getWorksheet(1)
    
  let existingColumnL = worksheet1.columns.length
  let rowSize1 = worksheet1.rowCount
    worksheet1.eachRow((row, rowNumber) => {
        let values1 = row.values
       if (rowNumber < 4) {
          let row2 = worksheet2.getRow(rowNumber).values.slice(1)
          worksheet.addRow(values1.concat(row2)).commit()
       } else {
        // console.log('existingColumnL',existingColumnL)
        
           let valL = values1.length
        //    console.log('val1', valL)
           if (valL < existingColumnL) {
            values1 = values1.concat(Array(existingColumnL - valL).fill(null))
           }

        //    console.log('values1', values1)
           let id = row.getCell(1).value
           let date1 = row.getCell(2).value
           worksheet2.eachRow((row2, rowNumber2) => {
               let id2 = row2.getCell(1).value
               let date2 = row2.getCell(2).value
               if (id == id2 && date1==date2) {
                values1 = values1.concat(row2.values.slice(1))
                worksheet2.spliceRows(rowNumber2,1)

               }
           })
           worksheet.addRow(values1).commit()
       }
       console.log((rowNumber/rowSize1 * 100) + '%' )
    })
    worksheet2 = null
    worksheet1 = null
    workbook.commit()
}

async function mergeTwoType(target) {
    let first =  '/Users/losingyoung/assay/data/m/capital-debt/filtered.xlsx'
    let second = '/Users/losingyoung/assay/data/m/ipo-time/IPO_Cobasic.xlsx'
    var filename = path.join(target, 'merged_' + new Date().toUTCString().replace(/[\s,:]/g, "_") + '.xlsx')
    console.log('filename', filename)
    var options = {
     filename
    };
    var workbook = new Excel.stream.xlsx.WorkbookWriter(options);
    var worksheet = workbook.addWorksheet('merged')
    
    var workbook1 = new Excel.Workbook();
    console.log('workbook1 start reading')
    await workbook1.xlsx.readFile(first)
    console.log('workbook1读取完毕')
    let worksheet1 = workbook1.getWorksheet(1)

    var workbook2 = new Excel.Workbook();
    console.log('workbook2 start reading')
    await workbook2.xlsx.readFile(second)
    console.log('workbook2读取完毕')
    let worksheet2 = workbook2.getWorksheet(1)
    

    worksheet1.eachRow((row, rowNumber) => {
        let values1 = row.values
       if (rowNumber < 4) {
          let row2 = worksheet2.getRow(rowNumber).values.slice(1)
          worksheet.addRow(values1.concat(row2)).commit()
       } else {
           let id = row.getCell(1).value
           worksheet2.eachRow((row2, rowNumber2) => {
               let id2 = row2.getCell(1).value
               if (id == id2) {
                values1 = values1.concat(row2.values.slice(1))
                // worksheet2.spliceRows(rowNumber2,1)

               }
           })
           worksheet.addRow(values1).commit()
       }
    })
    worksheet2 = null
    worksheet1 = null
    workbook.commit()
}
async function filterMonth(filePath) {
    let regP = /(.*)[\/\\].*\.xlsx$/
    let ret = filePath.match(regP)
    var filename = path.join(ret[1], 'filtered_' + new Date().toUTCString().replace(/[\s,:]/g, "_") + '.xlsx')
    console.log('filename', filename)
    var options = {
     filename
    };
    var workbook = new Excel.stream.xlsx.WorkbookWriter(options);
    var worksheet = workbook.addWorksheet('merged')
    
    var workbook1 = new Excel.Workbook();
    console.log('workbook start reading')
    await workbook1.xlsx.readFile(filePath)
    console.log('workbook读取完毕')
    let worksheet1 = workbook1.getWorksheet(1)
    worksheet1.eachRow((row, rowNumber) => {
        if (rowNumber < 4) {
            worksheet.addRow(row.values).commit()
            console.log('less than 4', row.values)
        } else {
            let values = row.values
            let reg = /12-31/
            if (reg.test(row.getCell(2).value) && row.getCell(3).value === "A") {
                console.log('12-31 && A', row.values)
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
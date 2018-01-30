var Excel = require('exceljs');
const path = require('path')
const fs = require('fs')

function merge(opts) {
    const first = opts.first
    const second = opts.second
    // read from a file 
    var workbook1 = new Excel.Workbook();
    var workbook2 = new Excel.Workbook();
    var reg = /(.*)\/.*\.xlsx$/
    var firstPath = first.match(reg)
    var filename = opts.name || path.join(firstPath[1], 'merged_workbook_' + new Date().toDateString().replace(/\s+/g, "-") + '.xlsx')
    var options = {
        filename
    };
    var workbook = new Excel.stream.xlsx.WorkbookWriter(options);
    var worksheet = workbook.addWorksheet('merged')
    // worksheet.addRow([null,1,2]).commit()
    // workbook.commit().then(function () {
    //     console.log('workbook end')
    // });
    // return
    workbook1.xlsx.readFile(first)
        .then(function () {
            // console.log(workbook)
            var worksheet1 = workbook1.getWorksheet(1)
            console.log('第一个workbook读取成功')
            worksheet1.eachRow(function (row, rowNumber) {
                // console.log('Row ' + rowNumber + ' = ' + JSON.stringify(row.values));
                worksheet.addRow(row.values).commit()
            });
            console.log('写完第一个文件')
            worksheet1 = null
            workbook1 = null
 
            workbook2.xlsx.readFile(second).then(() => {
                console.log('第二个workbook读取成功')
                var worksheet2 = workbook2.getWorksheet(1)
                worksheet2.eachRow(function (row, rowNumber) {
                    // console.log('Row ' + rowNumber + ' = ' + JSON.stringify(row.values));
                    worksheet.addRow(row.values).commit()
                });
                console.log('写完第二个文件')
                worksheet2 = null
                workbook2 = null
                
                workbook.commit()
                    .then(function () {
                        console.log('workbook end')
                    });
            })

        }).catch(err => {
            console.log('err', err)
        });
}

module.exports = merge

/*
nxcel merge /Users/losingyoung/assay/data/m/total-capital-growth/FI_T8.xlsx /Users/losingyoung/assay/data/m/total-capital-growth/FI_T81.xlsx --name /Users/losingyoung/assay/data/m/total-capital-growth/merged.xlsx


*/
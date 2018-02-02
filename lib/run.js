var Excel = require('exceljs');
const path = require('path')
const fs = require('fs')
const cluster = require('cluster');
const numCPUs = require('os').cpus().length;

class Percent {
    constructor() {
        this.startTime = null
    }
    init() {
        this.startTime = new Date().getTime()
    }
    showPercent(cur, all) {
        let taskPercent = (cur / all * 100).toFixed(2)
        let now = new Date().getTime()
        let last = now - this.startTime
        let left = (last / taskPercent) * 100 - last
        console.log('已用时', last / 1000, 's ', '剩余时间', (left / 1000).toFixed(2), 's ', taskPercent + '%')
    }
}

function run() {
    clusterRun()

    // splitInto
    // let target = 'E:\\assay\\data\\x\\age-degree'
    // mergeManagers(target)
    // filterMangers(target)
    // let target = '/Users/losingyoung/assay/data/m'
    // mergeTwoTypes(target)
    // mergeTwoType(target)  IPO and other
    // let filePath = '/Users/losingyoung/assay/data/m/total-capital-growth/merged.xlsx'
    // filterMonth(filePath)
}
async function clusterRun() {
    if (cluster.isMaster) {
        let fileId = ["0", "1", "2", "3", "4"]
        let second = 'E:\\assay\\data\\x\\term\\TMT_Position_merged.xlsx'
        var workbook2 = new Excel.Workbook();
        console.log('workbook2 start reading')
        await workbook2.xlsx.readFile(second)
        let worksheet2 = workbook2.getWorksheet(1)
        console.log('workbook2读取完毕')
        let workrow2 = []
        worksheet2.eachRow((row2, rowNumber2) => {
            workrow2.push(row2.values)
        })

        for (let i = 0; i < numCPUs - 1; i++) {
            let filePath = `E:\\assay\\data\\x\\age-degree\\filtered${fileId[i]}.xlsx`
            let wk = cluster.fork();
            console.log('send')
            wk.send({ first: filePath, idx: i, workrow2 })
        }
        var numOfCompelete = 0

        Object.keys(cluster.workers).forEach(function(id) {
            cluster.workers[id].on('message', function(msg) {
                console.log(`[master] receive message from [worker ${id}]: ${msg}`);
                numOfCompelete++;
                if (numOfCompelete === numCPUs.length) {
                    console.log(`[master] finish all work and using ${Date.now() -
     st} ms`);
                    cluster.disconnect();
                }
            });
        })
    } else {
        process.on('message', function({ first, idx, workrow2 }) {
            console.log(`子进程：${process.pid}, filepath: ${first}`)
            mergeManagers({
                target: 'E:\\assay\\data\\x\\age-degree',
                first,
                idx,
                workrow2
            })
        });
    }
}
async function mergeManagers({ target, first, idx, workrow2 }) {
    // let first = 'E:\\assay\\data\\x\\age-degree\\TMT_FigureInfo1_filtered.xlsx'

    // let first =  '/Users/losingyoung/assay/data/m/merged_capitaldebt.xlsx'
    // let second = '/Users/losingyoung/assay/data/m/capital-total/filtered.xlsx'
    var filename = path.join(target, 'merged_' + new Date().toUTCString().replace(/[\s,:]/g, "_") + idx + '.xlsx')
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

    // let worksheet2 = workbook2.getWorksheet(1)

    let existingColumnL = worksheet1.columns.length
    let rowSize1 = worksheet1.rowCount
    let percent = new Percent()
    percent.init()
    worksheet1.eachRow((row, rowNumber) => {
        let values1 = row.values
        if (rowNumber < 4) {
            let row2 = workrow2[rowNumber - 1].slice(1)
            worksheet.addRow(values1.concat(row2)).commit()
        } else {
            // console.log('existingColumnL',existingColumnL)
            let valL = values1.length
                //    console.log('val1', valL)
            if (valL < existingColumnL) {
                values1 = values1.concat(Array(existingColumnL - valL + 1).fill(null))
            }
            //    console.log('values1', values1)
            let id = row.getCell(1).value
            let date1 = row.getCell(2).value
            let personId1 = row.getCell(3).value
            workrow2.forEach((row2, rowNumber2) => {
                let id2 = row2[1]
                let date2 = row2[2]
                let personId2 = row2[3]
                if (id == id2 && date1 == date2 && personId1 == personId2) {
                    values1 = values1.concat(row2.slice(1))
                }
            })
            worksheet.addRow(values1).commit()
        }
        console.log(`number ${idx}: `)
        percent.showPercent(rowNumber, rowSize1)
    })
    worksheet2 = null
    worksheet1 = null
    workbook.commit()
    process.send('finish')
}
async function filterMangers(target) {
    let first = 'E:\\assay\\data\\x\\age-degree\\TMT_FigureInfo4.xlsx'
    var filename = path.join(target, 'filtered_' + new Date().toUTCString().replace(/[\s,:]/g, "_") + '.xlsx')
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
    let rowSize1 = worksheet1.rowCount
    let percent = new Percent()
    percent.init()
    worksheet1.eachRow((row, rowNumber) => {
        if (rowNumber < 4) {
            worksheet.addRow(row.values).commit()
        } else {
            let isG = row.getCell(8) == 1
            let isB = row.getCell(9) == 1
            if (isG || isB) {
                worksheet.addRow(row.values).commit()
            }
        }
        // let taskPercent = (rowNumber / rowSize1 * 100).toFixed(2)
        percent.showPercent(rowNumber, rowSize1)
    })
    workbook.commit()
}


async function mergeTwoTypes(target) {
    let first = 'E:\\assay\\data\\m\\merged_pure-invest-return.xlsx'
    let second = 'E:\\assay\\data\\m\\total-capital-growth\\filtered.xlsx'
        // let first =  '/Users/losingyoung/assay/data/m/merged_capitaldebt.xlsx'
        // let second = '/Users/losingyoung/assay/data/m/capital-total/filtered.xlsx'
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
    let startTime = new Date().getTime()
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
                values1 = values1.concat(Array(existingColumnL - valL + 1).fill(null))
            }
            //    console.log('values1', values1)
            let id = row.getCell(1).value
            let date1 = row.getCell(2).value
            worksheet2.eachRow((row2, rowNumber2) => {
                let id2 = row2.getCell(1).value
                let date2 = row2.getCell(2).value
                if (id == id2 && date1 == date2) {
                    values1 = values1.concat(row2.values.slice(1))
                    worksheet2.spliceRows(rowNumber2, 1)

                }
            })
            worksheet.addRow(values1).commit()
        }

        let taskPercent = (rowNumber / rowSize1 * 100).toFixed(2)
        let now = new Date().getTime()
        let last = now - startTime
        let left = (last / taskPercent) * 100 - last
        console.log('已用时', last / 1000, 's ', '剩余时间', (left / 1000).toFixed(2), 's ', taskPercent + '%')
    })
    worksheet2 = null
    worksheet1 = null
    workbook.commit()
}

async function mergeTwoType(target) {
    let first = '/Users/losingyoung/assay/data/m/capital-debt/filtered.xlsx'
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
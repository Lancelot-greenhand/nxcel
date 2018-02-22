var Excel = require('exceljs');
const path = require('path')
const fs = require('fs')
const cluster = require('cluster');
const numCPUs = require('os').cpus().length;
//ss
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
    // clusterRun()
    // let target = '/Users/losingyoung/assay/data'
    // splitInto
    let target = 'E:\\assay\\data'
        // getCurFirmCode(target)
        // mergeManagers(target)
        // filterMangers(target)
        // handleX(target)
        // findAllyears(target)
        // let filePath = 'E:\\assay\\data\\y\\2015-data.txt'
        // txtToexcel(filePath, target)
    mergeTwoTypes(target)
        // mergeTwoType(target)  IPO and other
        // let filePath = '/Users/losingyoung/assay/data/m/total-capital-growth/merged.xlsx'
        // filterMonth(filePath)
}
async function handleX(target) {
    let first = 'e:\\assay\\data\\filtered_x.xlsx'
    var filename = path.join(target, 'x' + new Date().toUTCString().replace(/[\s,:]/g, "_") + '.xlsx')
    console.log('filename', filename)
    var options = {
        filename
    };
    var workbook = new Excel.stream.xlsx.WorkbookWriter(options);
    var worksheet = workbook.addWorksheet('merged')
    var workbook1 = new Excel.Workbook();
    console.log('workbook1 start reading')
    await workbook1.xlsx.readFile(first)
    console.log('workbook1 end reading')
    let worksheet1 = workbook1.getWorksheet(1)

    let firmMap = {}
    worksheet1.eachRow((row, rowNumber) => {
        let id = row.getCell(1).value
        let date = row.getCell(2).value
        if (typeof firmMap[id] == 'undefined') {
            firmMap[id] = {}
        }
        if (typeof firmMap[id][date] == 'undefined') {
            firmMap[id][date] = {
                age: [],
                degree: [],
                term: [],
                back: []
            }
        }
        let age = parseInt(row.getCell(5).value, 10)
        let term = parseInt(row.getCell(6).value, 10)
        let degree = parseInt(row.getCell(7).value, 10)
        let back = row.getCell(8).value.split(",") //.map(toInt)
        firmMap[id][date].age.push(age)
        firmMap[id][date].term.push(term)
        firmMap[id][date].degree.push(degree)
        firmMap[id][date].back = firmMap[id][date].back.concat(back)
    })
    worksheet1 = null
    workbook1 = null
    worksheet.addRow(["证券id", "日期", "Hage", "Hterm", "Hdeg", "Hback"]).commit()
    let i = 0
    Object.keys(firmMap).forEach(id => {

        Object.keys(firmMap[id]).forEach(date => {
            let data = [id, date]
            let hAge = standAvg(firmMap[id][date].age)
            let hTerm = standAvg(firmMap[id][date].term)
            let hDegree = HHI(firmMap[id][date].degree)
            let hBack = HHI(firmMap[id][date].back)
            data.push(hAge)
            data.push(hTerm)
            data.push(hDegree)
            data.push(hBack)
            console.log(i++)
            worksheet.addRow(data).commit()
        })
    })
    workbook.commit()
}

function standAvg(arr) {
    let len = arr.length
    let avg = Average(arr)
    let marginAll = arr.reduce((sum, cur) => {
        let margin = Math.pow(cur - avg, 2)
        return sum + margin
    }, 0)
    let up = Math.sqrt(marginAll / len)
    return up / avg
}

function HHI(arr) {
    let keyMap = {}
    arr.forEach(val => {
        if (typeof keyMap[val] == "undefined") {
            keyMap[val] = 1
            return
        }
        keyMap[val]++
    })
    let len = arr.length
    let all = 0
    Object.keys(keyMap).forEach(key => {
        let num = Math.pow(keyMap[key] / len, 2)
        all += num
    })
    return 1 - all
}

function Average(arr) {
    let len = arr.length
    let all = arr.reduce((sum, cur) => {
        return sum + cur
    }, 0)
    return all / len
}

function toInt(val) {
    return parseInt(val, 10)
}

async function findAllyears(target) {
    let first = 'e:\\assay\\data\\filtered_y.xlsx'
    var filename = path.join(target, 'filtered' + new Date().toUTCString().replace(/[\s,:]/g, "_") + '.xlsx')
    console.log('filename', filename)
    var options = {
        filename
    };
    var workbook = new Excel.stream.xlsx.WorkbookWriter(options);
    var worksheet = workbook.addWorksheet('merged')
    var workbook1 = new Excel.Workbook();
    console.log('workbook1 start reading')
    await workbook1.xlsx.readFile(first)
    console.log('workbook1 end reading')
    let worksheet1 = workbook1.getWorksheet(1)

    let percent = new Percent()
    percent.init()
    let codeMap = {}
    let rowSize1 = worksheet1.rowCount
    worksheet1.eachRow((row, rowNumber) => {
        let id = row.getCell(1).value
        let date = row.getCell(2).value
        if (typeof codeMap[id] == 'undefined') {
            codeMap[id] = {}
        }
        codeMap[id][date] = ""
    })
    let fourCode = []
    Object.keys(codeMap).forEach(id => {
        if (Object.keys(codeMap[id]).length === 4) {
            fourCode.push(id)
        }
    })

    worksheet1.eachRow((row, rowNumber) => {
        let values1 = row.values
        if (rowNumber < 4) {
            worksheet.addRow(values1).commit()
        } else {
            if (fourCode.indexOf(row.getCell(1).value) > -1) {
                worksheet.addRow(values1).commit()
            }
        }
        percent.showPercent(rowNumber, rowSize1)
    })
    console.log('all firm', fourCode.length)
    workbook.commit()
}
async function getCurFirmCode(target) {
    let first = 'e:\\assay\\data\\x_rate.xlsx'
    var filename = path.join(target, 'codes_' + new Date().toUTCString().replace(/[\s,:]/g, "_") + '.xlsx')
    var options = {
        filename
    };
    var workbook = new Excel.stream.xlsx.WorkbookWriter(options);
    var worksheet = workbook.addWorksheet('merged')

    var workbook1 = new Excel.Workbook();
    console.log('workbook1 start reading')
    await workbook1.xlsx.readFile(first)
    console.log('workbook1 end reading')
    let worksheet1 = workbook1.getWorksheet(1)
    let percent = new Percent()
    percent.init()
    let codeA = []
    worksheet1.eachRow((row, rowNumber) => {
        let id = row.getCell(1).value
        codeA.push(id)
    })
    console.log('end each row')

    let temSet = new Set(codeA)
    let A = Array.from(temSet)
    console.log("A", A)
    A.forEach(id => {
        worksheet.addRow([id]).commit()
    })
    workbook.commit()
}
async function txtToexcel(filepath, target) {
    var filename = path.join(target, 'merged_' + new Date().toUTCString().replace(/[\s,:]/g, "_") + '.xlsx')
    fs.readFile(filepath, "utf-8", (err, data) => {
        if (err) {
            console.log('err', err)
            return
        }
        console.log('data', typeof data)
        let arr = data.split(/\n/)
        console.log(arr)
        let spArr = arr.map(val => {
            val = val.trim()
            val = val.replace(/(drs|irs|-)/g)
            return val.split(/\s+/g)
        })

        var options = {
            filename
        };
        var workbook = new Excel.stream.xlsx.WorkbookWriter(options);
        var worksheet = workbook.addWorksheet('merged')
        spArr.forEach(row => {
            worksheet.addRow(row).commit()
        })

        workbook.commit()
            // spArr.forEach(row => {

        // })
    })
}
async function mergeMember(target) {
    let first = '/Users/losingyoung/assay/data/y/merged_unm.xlsx'
    let second = '/Users/losingyoung/assay/data/y/filtered_m.xlsx'
    var filename = path.join(target, 'merged_' + new Date().toUTCString().replace(/[\s,:]/g, "_") + '.xlsx')
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

    let percent = new Percent()
    percent.init()
    let rowSize1 = worksheet1.rowCount
    let dateMap = {
        "2012-12-31": 2,
        "2013-12-31": 3,
        "2014-12-31": 4,
        "2015-12-31": 5
    }
    worksheet1.eachRow((row, rowNumber) => {
        let values1 = row.values
        if (rowNumber < 4) {
            worksheet.addRow(values1).commit()
        } else {

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
                if (id == id2) {
                    console.log('date1', date1)
                    console.log(values1)
                    let data2 = row2.getCell(dateMap[date1]).value
                    console.log('date2', data2)
                    console.log(values1)
                    values1.push(data2)
                    worksheet.addRow(values1).commit()
                }
            })

        }
        percent.showPercent(rowNumber, rowSize1)


    })
    worksheet2 = null
    worksheet1 = null
    workbook.commit()
}


async function clusterRun() {
    if (cluster.isMaster) {
        let fileId = ["3", "4"]
        let second = 'E:\\assay\\data\\x\\age-degree\\TMT_Position_filtered.xlsx' // 'E:\\assay\\data\\x\\term\\TMT_Position.xlsx'
        var workbook2 = new Excel.Workbook();
        console.log('workbook2 start reading')
        await workbook2.xlsx.readFile(second)
        let worksheet2 = workbook2.getWorksheet(1)
        console.log('workbook2读取完毕')
        let workrow2 = []
        worksheet2.eachRow((row2, rowNumber2) => {
            workrow2.push(row2.values)
        })
        let work2ColumnL = worksheet2.columns.length
        for (let i = 0; i < numCPUs - 1; i++) {
            if (!fileId[i]) { continue }
            let first = `E:\\assay\\data\\x\\age-degree\\filtered${fileId[i]}.xlsx`
            let wk = cluster.fork();
            wk.send({ first, idx: fileId[i], workrow2, work2ColumnL })
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
        process.on('message', function(obj) {
            console.log(`子进程：${process.pid}, filepath: ${obj.first}`)
            mergeManagers(Object.assign({}, {
                target: 'E:\\assay\\data\\x\\age-degree'
            }, obj))
        });
    }
}

async function mergeManagers({ target, first, idx, workrow2, work2ColumnL }) {

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
            if (valL < existingColumnL) {
                values1 = values1.concat(Array(existingColumnL - valL + 1).fill(null))
            }
            let id = row.getCell(1).value
            let date1 = row.getCell(2).value
            let personId1 = row.getCell(3).value
            workrow2.forEach((row2, rowNumber2) => {
                let id2 = row2[1]
                let date2 = row2[2]
                let personId2 = row2[3]
                if (id == id2 && date1 == date2 && personId1 == personId2) {
                    values1 = values1.concat(row2.slice(1))
                    if (row2.length - 1 < work2ColumnL) {
                        values1 = values1.concat(Array(work2ColumnL - (row2.length - 1)).fill(null))
                    }
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
    let first = 'E:\\assay\\data\\merged_x.xlsx'
    var filename = path.join(target, 'filtered_x_' + new Date().toUTCString().replace(/[\s,:]/g, "_") + '.xlsx')
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
    let termMap = {}
    worksheet1.eachRow((row, rowNumber) => {
            if (rowNumber < 4) {
                worksheet.addRow(row.values).commit()
            } else {
                // let date = row.getCell(2)
                // let personId = row.getCell(3)
                // let key = date + personId
                // if (typeof termMap[key] == 'undefined') {
                //     termMap[key] = row.values
                //     return
                // }
                // let oldVal = termMap[key][9]
                // let curVal = row.getCell(9).value
                // if (parseInt(curVal, 10) > parseInt(oldVal, 10)) {
                //     termMap[key][9] = curVal
                // }

                let isH = row.getCell(7)
                if (isH == 1) {
                    worksheet.addRow(row.values).commit()
                }

            }
            percent.showPercent(rowNumber, rowSize1)
        })
        // Object.keys(termMap).forEach(key => {
        //     let val = termMap[key]
        //     worksheet.addRow(val).commit()
        // })
    workbook.commit()
}


async function mergeTwoTypes(target) {
    let first = 'E:\\assay\\data\\merged_filtered_m_y.xlsx'
    let second = 'E:\\assay\\data\\x_rate.xlsx'
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
                        // worksheet2.spliceRows(rowNumber2, 1)
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

module.exports = run
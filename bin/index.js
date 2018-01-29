#!/usr/bin/env node
 // 引入index的实例， 传递参数，触发事件


const prog = require('caporal')
const app = require('../index')


prog.version('1.0.0')
    .description('A simple program that says "biiiip"')
    .command('merge', 'merge 2 files with same columns')
    .argument('<first>', 'first file, format xlsx', /\.(xlsx|xls)$/)
    .argument('<second>', 'second file, format xlsx', /\.(xlsx|xls)$/)
    .option('--name <name>', 'merged name')
    .action(function(args, options, logger) {
        app.merge(Object.assign({}, args, options))
    })

prog.command('run', 'run custom action')
    .action(function() {
        app.run()
    })
prog.parse(process.argv)


// app.emit('tt', 'TT')
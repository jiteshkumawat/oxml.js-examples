const oxml = require("oxmljs")

var workbook1 = oxml.xlsx();
workbook1.sheet('sheet1');
workbook1.download(__dirname + '/basicWorkbook1.xlsx').then(function () {
    console.log('Downloaded Workbook1.');
});

var workbook2 = oxml.xlsx();
var worksheet2 = workbook2.sheet('sheet1');
worksheet2.cell(2, 3).set('Hello World');
workbook2.download(__dirname + '/basicWorkbook2.xlsx').then(function () {
    console.log('Downloaded Workbook2.');
});

var workbook3 = oxml.xlsx();
var worksheet3 = workbook3.sheet('sheet1');
worksheet3.cell(2, 3).set({ value: 'Hello World', type: 'string' });
workbook3.download(__dirname + '/basicWorkbook3.xlsx').then(function () {
    console.log('Downloaded Workbook3.');
});

var workbook4 = oxml.xlsx();
var worksheet4 = workbook4.sheet('sheet1');
worksheet4.cell(2, 3, 'Hello World');
workbook4.download(__dirname + '/basicWorkbook4.xlsx').then(function () {
    console.log('Downloaded Workbook4.');
});

var workbook5 = oxml.xlsx();
var worksheet5 = workbook5.sheet('sheet1');
worksheet5.cell(2, 3, { value: 'Hello World', type: 'string' });
workbook5.download(__dirname + '/basicWorkbook5.xlsx').then(function () {
    console.log('Downloaded Workbook5.');
});

var workbook6 = oxml.xlsx();
var worksheet6 = workbook6.sheet('sheet1');
worksheet6.row(2, 3, ['Cost', 'Sale', 'Profit']);
worksheet6.row(3, 3, [10, 12]);
worksheet6.row(4, 3, [9, 12]);
worksheet6.cell(3, 5, { type: 'formula', formula: '(D3 - C3)', value: 2 });
worksheet6.cell(4, 5, { type: 'formula', formula: '(D4 - C4)', value: 3 });
worksheet6.row(5, 2, [
    { type: 'sharedString', value: 'Total' },
    { type: 'formula', formula: '(C3 + C4)', value: 19 },
    { type: 'formula', formula: '(D3 + D4)', value: 24 }]);
workbook6.download(__dirname + '/basicWorkbook6.xlsx').then(function () {
    console.log('Downloaded Workbook6.');
});

var workbook7 = oxml.xlsx();
var worksheet7 = workbook7.sheet('sheet1');
worksheet7.row(2, 3, ['Hello', 'Wold']);
var row7 = worksheet7.row(2, 3);
row7.set(['Greetings']);
workbook7.download(__dirname + '/basicWorkbook7.xlsx').then(function () {
    console.log('Downloaded Workbook7.');
});

var workbook8 = oxml.xlsx();
var worksheet8 = workbook8.sheet('sheet1');
worksheet8.column(2, 3, ['Cost', 'Sale', 'Profit']);
worksheet8.column(2, 4, [9, 12, 3]);
worksheet8.column(2, 5, [10, 12, 2]);
workbook8.download(__dirname + '/basicWorkbook8.xlsx').then(function () {
    console.log('Downloaded Workbook8.');
});

var workbook9 = oxml.xlsx();
var worksheet9 = workbook9.sheet('sheet1');
worksheet9.column(2, 3, ['Hello', 'Wold']);
var column9 = worksheet9.column(2, 3);
column9.set(['Greetings']);
workbook9.download(__dirname + '/basicWorkbook9.xlsx').then(function () {
    console.log('Downloaded Workbook9.');
});

var workbook10 = oxml.xlsx();
var worksheet10 = workbook10.sheet('sheet1');
worksheet10.grid(2, 3, [
    ['Cost', 'Sales', 'Profit'],
    [10, 12, { type: 'formula', value: 2, formula: '(D3 - C3)' }],
    [9, 12, { type: 'formula', value: 3, formula: '(D4 - C4)' }],
    [11, 12, { type: 'formula', value: 1, formula: '(D5 - C5)' }]
]);
workbook10.download(__dirname + '/basicWorkbook10.xlsx').then(function () {
    console.log('Downloaded Workbook10.');
});

var workbook11 = oxml.xlsx();
var worksheet11 = workbook11.sheet('sheet1');
worksheet11.grid(2, 3, [['Hello', 'World']]);
worksheet11.grid(2, 3).set([['Greetings'], ['Jon']]);
workbook11.download(__dirname + '/basicWorkbook11.xlsx').then(function () {
    console.log('Downloaded Workbook11.');
});

var workbook12 = oxml.xlsx();
var worksheet12 = workbook12.sheet('sheet1');
worksheet12.grid(2, 3, [
    ['Cost', 'Sales', 'Profit'],
    [10, 12],
    [9, 12],
    [11, 12],
    ['Total']
]);
worksheet12.sharedFormula('E3', 'E5', {
    type: 'formula', formula: '(D3 - C3)', value: function (rowIndex, columnIndex) {
        var sale = worksheet12.cell(rowIndex, columnIndex - 1).value;
        var cost = worksheet12.cell(rowIndex, columnIndex - 2).value;
        return sale - cost;
    }
});
worksheet12.sharedFormula('C6', 'D6', {
    type: 'formula', formula: 'SUM(C3:C5)', value: function (rowIndex, columnIndex) {
        var column = worksheet12.column(3, 3), sum = 0;
        for (var index = 0; index < column.cells.length; index++) {
            if (column.cells[index].value && typeof column.cells[index].value === "number") {
                sum += column.cells[index].value;
            }
        }
        return sum;
    }
});
workbook12.download(__dirname + '/basicWorkbook12.xlsx').then(function () {
    console.log('Downloaded Workbook12.');
});
const oxml = require("oxmljs")

var workbook1 = oxml.xlsx();
var worksheet1 = workbook1.sheet('sheet1');
worksheet1.cell(2, 2, 'Hello World!', { fontColor: 'ff0000' });
worksheet1.cell(2, 2).style({ bold: true });
workbook1.download(__dirname + '/stylesWorkbook1.xlsx').then(function () {
    console.log('Downloaded Workbook1.');
});

var workbook2 = oxml.xlsx();
var worksheet2 = workbook2.sheet('sheet1');
worksheet2.row(1, 1, 'Total of Data', {
     fill: {
          gradient: {
               degree: 90,
               stops: [{
                    position: 0,
                    color: 'FF92D050'
               },
               {
                    position: 1,
                    color: 'FF0070C0'
               }]
          }
     },
     fontColor: 'ffffff',
     bold: true,
     underline: true
});
worksheet2.row(2, 2, ['Data 1', 'Data 2', { type: 'sharedString', value: 'Total' }], {
     bold: true,
     italic: true,
     underline: true,
     fontName: 'Calibri Light',
     fontColor: '0000ff'
});
worksheet2.row(3, 2, [5, 9]);
worksheet2.row(4, 2, [7, 3]);
worksheet2.row('D3', 'D4', '(B3 + C3)', { bold: true });
workbook2.download(__dirname + '/stylesWorkbook2.xlsx').then(function () {
    console.log('Downloaded Workbook2.');
});
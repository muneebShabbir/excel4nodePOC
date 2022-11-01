// https://github.com/natergj/excel4node/issues/207
// https://stackoverflow.com/questions/44113196/how-to-create-the-cell-dropdown-list-programmatically-using-excel4node-js-in-nod
// https://www.npmjs.com/package/excel4node
var express = require('express');
var router = express.Router();
var xl = require('excel4node');


/* GET users listing. */
router.get('/', function(req, res, next) {
  console.log('Generating File');


  var wb = new xl.Workbook();

// Add Worksheets to the workbook
var ws = wb.addWorksheet('Sheet 1');
var ws2 = wb.addWorksheet('Sheet 2');

// Create a reusable style
var style = wb.createStyle({
  font: {
    color: '#FF0800',
    size: 12,
  },
  numberFormat: '$#,##0.00; ($#,##0.00); -',
});

ws.cell(1, 1)
  .string('Dropdown section')

ws.cell(1, 2)
  .string('Date Section')

ws.addDataValidation({
  type: 'list',
  allowBlank: true,
  prompt: 'Choose from dropdown',
  errorTitle: 'Invalid Option',
  error: 'Select Option from Dropdown',
  showDropDown: true,
  sqref: 'A2:A200',
  formulas: ['A,B,C,D,E,F'],
});

ws.addDataValidation({
	type: 'date',
	allowBlank: false,
	error: 'My Enter date',
	sqref: 'B2:B100',
	// operator: 'lessThan',
	// formulas: [xl.getExcelTS(new Date("2018-07-07T00:00:00.0000Z"))]
})

// // Set value of cell A1 to 100 as a number type styled with paramaters of style
// ws.cell(1, 1)
//   .number(100)
//   .style(style);

// // Set value of cell B1 to 200 as a number type styled with paramaters of style
// ws.cell(1, 2)
//   .number(200)
//   .style(style);

// // Set value of cell C1 to a formula styled with paramaters of style
// ws.cell(1, 3)
//   .formula('A1 + B1')
//   .style(style);

// // Set value of cell A2 to 'string' styled with paramaters of style
// ws.cell(2, 1)
//   .string('string')
//   .style(style);

// // Set value of cell A3 to true as a boolean type styled with paramaters of style but with an adjustment to the font size.
// ws.cell(3, 1)
//   .bool(true)
//   .style(style)
//   .style({font: {size: 14}});

wb.write('Excel.xlsx');


  res.send('File Generated in the folder');
});

module.exports = router;

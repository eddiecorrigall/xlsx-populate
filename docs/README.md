[![view on npm](http://img.shields.io/npm/v/xlsx-populate.svg)](https://www.npmjs.org/package/xlsx-populate)
[![npm module downloads per month](http://img.shields.io/npm/dm/xlsx-populate.svg)](https://www.npmjs.org/package/xlsx-populate)
[![Build Status](https://travis-ci.org/dtjohnson/xlsx-populate.svg?branch=master)](https://travis-ci.org/dtjohnson/xlsx-populate)
[![Dependency Status](https://david-dm.org/dtjohnson/xlsx-populate.svg)](https://david-dm.org/dtjohnson/xlsx-populate)

# xlsx-populate
Node.js module to populate Excel XLSX templates. This module does not parse Excel workbooks. There are [good modules](https://github.com/SheetJS/js-xlsx) for this already. The purpose of this module is to open existing Excel XLSX workbook templates that have styling in place and populate with data.

## Installation

    $ npm install xlsx-populate

## Usage
Here is a basic example:
```js
var Workbook = require('xlsx-populate');

// Load the input workbook from file.
var workbook = Workbook.fromBlankSync();

// Modify the workbook.
workbook.getSheet(0).getCell("A1").setValue("This is neat!");

// Write to file.
workbook.toFileSync("./out.xlsx");
```

### Getting Sheets
You can get sheets from a Workbook object by either name or index (0-based):
```js
// Get sheet with name "Sheet1".
var sheet = workbook.getSheet("Sheet1");

// Get the first sheet.
var sheet = workbook.getSheet(0);
```

### Getting Cells
You can get a cell from a sheet by either address or row and column:
```js
// Get cell "A5" by address.
var cell = sheet.getCell("A5");

// Get cell "A5" by row and column.
var cell = sheet.getCell(5, 1);
```

You can also get named cells directly from the Workbook:
```js
// Get cell named "Foo".
var cell = sheet.getNamedCell("Foo");
```

### Setting Cell Contents
You can set the cell value or formula:
```js
cell.setValue("foo");
cell.setValue(5.6);
cell.setFormula("SUM(A1:A5)");
```

You can set multiple values along a row:
```js
// Operate on row 4, set values at A4 and D4
sheet.setColumnValues(4, { 'A': 'abc', 'D': 123 });

// Same operation
sheet.setColumnValues(4, { 1: 'abc', 4: 123 });
```

### Share Formula with Cells
You can share a formula at a given cell, along a row or column:
```js
// Share formula at A3 along row ending with C3
sheet.getCell('A3').shareFormulaUntil('C3');

// Share formula at C1 along column ending with C3
sheet.getCell('C1').shareFormulaUntil('C3');
```

### Get Range of Cells
You can get a region of cells:
```js
// The following performs the same operation
sheet.getCellRange('A1:B2');
sheet.getCellRange(['A1', 'B2']);
sheet.getCellRange('A1', 'B2');
sheet.getCellRange(sheet.getCell('A1'), sheet.getCell('B2'));

// Each returns:
// [ sheet.getCell('A1'), sheet.getCell('A2'), sheet.getCell('B1'), sheet.getCell('B2') ]
```

### Serving from Express
You can serve the workbook with [express](http://expressjs.com/) with a route like this:
```js
router.get("/download", function (req, res) {
    // Open the workbook.
    var workbook = Workbook.fromFile("input.xlsx", function (err, workbook) {
        if (err) return res.status(500).send(err);

        // Make edits.
        workbook.getSheet(0).getCell("A1").setValue("foo");

        // Set the output file name.
        res.attachment("output.xlsx");

        // Send the workbook.
        res.send(workbook.output());
    });
});
```

## Development

### Running Tests
Tests are run automatically on [Travis CI](https://travis-ci.org/dtjohnson/xlsx-populate). They can (and should) be triggered locally with:
    
    $ npm test

### Code Linting
[JSHint](http://jshint.com/) and [JSCS](http://jscs.info/) are used to ensure code quality. To run these, run:
    
    $ npm run jshint
    $ npm run jscs

### Generating Documentation
The API reference documentation below is generated by [jsdoc-to-markdown](https://github.com/jsdoc2md/jsdoc-to-markdown). To generate an updated README.md, run:
    
    $ npm run docs

# API Reference

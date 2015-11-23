"use strict";

var Workbook = require('../../');

// Get workbook and sheet
var workbook = Workbook.fromFileSync(__dirname + '/template.xlsx');
var sheet = workbook.getSheet('ClickThroughRateSheet');

// Get header cells
var clicksHeader = sheet.getCell('B2');
var impressionsHeader = sheet.getCell('C2');
var ctrHeader = sheet.getCell('D2');

// Randomly generate 10 rows of data
var r = 0;
while (r < 10) {
	r++; // Skip header
	var clickValue = parseInt(1000*Math.random());
	var impressionValue = parseInt(1000000*Math.random());
	clicksHeader.getRelativeCell(r, 0).setValue(clickValue);
	impressionsHeader.getRelativeCell(r, 0).setValue(impressionValue);
}

// Assign shared formulas from the first 
ctrHeader
	.getRelativeCell(1, 0) // Start from the first cell below header
	.shareFormula(ctrHeader.getRelativeCell(r, 0)) // End at the last modifed row
	;

// Save to file
workbook.toFileSync(__dirname + '/out.xlsx');

"use strict";

var xlsx = require('../../lib/Workbook');

var workbook = xlsx.fromFileSync(__dirname + '/template.xlsx');
var sheet = workbook.getSheet('ClickThroughRateSheet');

var clicks = sheet.getCell('B3');
var impressions = sheet.getCell('C3');
var ctr = sheet.getCell('D3');

var r = 0;
while (r < 10) {
	var clickValue = parseInt(1000*Math.random());
	var impressionValue = parseInt(1000000*Math.random());
	clicks.getRelativeCell(r, 0).setValue(clickValue);
	impressions.getRelativeCell(r, 0).setValue(impressionValue);
	r++;
}

ctr.shareFormula(ctr.getRelativeCell(r-1, 0));

workbook.toFileSync(__dirname + '/out.xlsx');

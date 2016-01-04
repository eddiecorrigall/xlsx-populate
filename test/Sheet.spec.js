/* jshint jasmine: true */

"use strict";

var utils = require('../lib/utils');
var Workbook = require('../lib/Workbook');
var Sheet = require('../lib/Sheet');

describe('Sheet', function () { // Class
    var workbook, sheet;

    beforeEach(function () {
        workbook = Workbook.fromBlankSync();
        sheet = workbook.getSheet(0);
    });

    describe('getWorkbook', function () {
    });

    describe('getName', function () {
    });

    describe('setName', function () {
    });

    describe('getCell', function () {
        describe('deterministic access', function () {
            it('correctly maps to the same cell', function () {
                var upperCaseCell = sheet.getCell('A1');
                var lowerCaseCell = sheet.getCell('a1');
                var rowAndColumnCell = sheet.getCell(1, 1);
                expect(upperCaseCell.getFullAddress()).toBe(lowerCaseCell.getFullAddress());
                expect(upperCaseCell.getFullAddress()).toBe(rowAndColumnCell.getFullAddress());
                var num = Math.random();
                upperCaseCell.setValue(num);
                var upperCaseVNode = upperCaseCell._cellNode.find('./v');
                var lowerCaseVNode = lowerCaseCell._cellNode.find('./v');
                var rowAndColumnVNode = rowAndColumnCell._cellNode.find('./v');
                expect(upperCaseVNode).not.toBeNull('A1 value node should not be null');
                expect(lowerCaseVNode).not.toBeNull('a1 value node should not be null');
                expect(rowAndColumnVNode).not.toBeNull('1,1 value node should not be null');
                expect(parseFloat(upperCaseVNode.text)).toBe(num, 'A1 value set must match value generated');
                expect(parseFloat(lowerCaseVNode.text)).toBe(num, 'a1 value set must match value generated');
                expect(parseFloat(rowAndColumnVNode.text)).toBe(num, '1,1 value set must match value generated');
                expect(upperCaseVNode.text).toBe(lowerCaseVNode.text);
                expect(upperCaseVNode.text).toBe(rowAndColumnVNode.text);
            });
        });
        
        describe('stochastic access', function () {
            var MAX_ROW = 100;
            var MAX_COLUMN = 100;
            var MAX_EDIT = 1000;

            beforeEach(function () {
                // Make random edits to the sheet
                for (var i = 0; i < MAX_EDIT; i++) {
                    var rowNumber = 1 + Math.floor(MAX_ROW * Math.random());
                    var columnNumber = 1 + Math.floor(MAX_COLUMN * Math.random());
                    var cell = sheet.getCell(rowNumber, columnNumber);
                    cell.setValue(Math.random());
                }
            });

            it('does not contain duplicates', function () {
                var addressCounter = {};
                var sheetDataNode = sheet._sheetXML.find('./sheetData');
                var rowNodes = sheetDataNode.findall('./row');
                rowNodes.forEach(function (rowNode) {
                    var cNodes = rowNode.findall('./c');
                    cNodes.forEach(function (cNode) {
                        var address = cNode.get('r');
                        expect(address).not.toBeNull();
                        expect(address).toBeDefined();
                        if ((address in addressCounter) === false) {
                            addressCounter[address] = 0;
                        }
                        addressCounter[address]++;
                        expect(addressCounter[address]).toBeLessThan(2);
                    });
                });
            });

            it('is stored in order', function () {
                // Reload workbook and sheet
                workbook = new Workbook(workbook.output());
                sheet = workbook.getSheet(0);
                // Check order
                var sheetDataNode = sheet._sheetXML.find('./sheetData');
                // ...
                var lastRowNumber = 0;
                var rowNodes = sheetDataNode.findall('./row');
                rowNodes.forEach(function (rowNode) {
                    var rowNumber = parseInt(rowNode.get('r')); // Don't forget to check for NaN
                    expect(isNaN(rowNumber)).toBe(false);
                    expect(rowNumber).toBeGreaterThan(lastRowNumber);
                    lastRowNumber = rowNumber;
                    // ...
                    var lastColumnNumber = 0;
                    var cNodes = rowNode.findall('./c');
                    cNodes.forEach(function (cNode) {
                        var address = cNode.get('r');
                        expect(address).toBeDefined();
                        var ref = utils.addressToRowAndColumn(address);
                        expect(ref.row).toBe(rowNumber);
                        expect(ref.column).toBeGreaterThan(lastColumnNumber);
                        lastColumnNumber = ref.column;
                    });
                });
            });
        });
    });

    describe("getCellRange", function () {
    });

    describe("setValues", function () {
    });
});

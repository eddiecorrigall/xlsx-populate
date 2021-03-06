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
        it("follows the design recipe", function () {
            var stringRange = sheet.getCellRange('A1:C3');
            var stringArrayRange = sheet.getCellRange(['A1', 'C3']);
            var cellArrayRange = sheet.getCellRange([sheet.getCell('A1'), sheet.getCell('C3')]);
            var stringArgumentsRange = sheet.getCellRange('A1', 'C3');
            var cellArgumentsRange = sheet.getCellRange(sheet.getCell('A1'), sheet.getCell('C3'));
            // ...
            expect(stringRange instanceof Array).toBe(true);
            expect(stringArrayRange instanceof Array).toBe(true);
            expect(cellArrayRange instanceof Array).toBe(true);
            expect(stringArgumentsRange instanceof Array).toBe(true);
            expect(cellArgumentsRange instanceof Array).toBe(true);
            // ...
            expect(stringRange.length).toBe(9);
            expect(stringArrayRange.length).toBe(9);
            expect(cellArrayRange.length).toBe(9);
            expect(stringArgumentsRange.length).toBe(9);
            expect(cellArgumentsRange.length).toBe(9);
            // ...
            var cellToAddressMap = function (c) {
                return c.getFullAddress();
            };
            var expectedAddresses = jasmine.arrayContaining([
                "'Sheet1'!B1", "'Sheet1'!B2", "'Sheet1'!B3",
                "'Sheet1'!A1", "'Sheet1'!A2", "'Sheet1'!A3",
                "'Sheet1'!C1", "'Sheet1'!C2", "'Sheet1'!C3"
            ]);
            // ...
            expect(stringRange.map(cellToAddressMap)).toEqual(expectedAddresses);
            expect(stringArrayRange.map(cellToAddressMap)).toEqual(expectedAddresses);
            expect(cellArrayRange.map(cellToAddressMap)).toEqual(expectedAddresses);
            expect(stringArgumentsRange.map(cellToAddressMap)).toEqual(expectedAddresses);
            expect(cellArgumentsRange.map(cellToAddressMap)).toEqual(expectedAddresses);
        });
    });

    describe("setValues", function () {
    });
});

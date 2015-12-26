/* jshint jasmine: true */

"use strict";

var utils = require('../lib/utils'), // Assume utils just works
    Workbook = require('../lib/Workbook'),
    Cell = require('../lib/Cell'),
    etree = require('elementtree'),
    element = etree.Element,
    subelement = etree.SubElement;

describe("Cell", function () {
    var cell;
    var sheetMock = {
        getName: function () {
            return "Foo";
        }
    };

    beforeEach(function () {
        var c = element('c');
        c.attrib.t = "inlineStr";
        c.attrib.r = "C5";
        var isNode = subelement(c, "is");
        var tNode = subelement(isNode, "t");
        tNode.text = "Foo";
        cell = new Cell(sheetMock, 5, 3, c);
    });

    describe("getSheet", function () {
        it("should return the parent sheet object", function () {
            expect(cell.getSheet()).toBe(sheetMock);
        });
    });

    describe("getRow", function () {
        it("should return the row", function () {
            expect(cell.getRow()).toBe(5);
        });
    });

    describe("getColumn", function () {
        it("should return the column", function () {
            expect(cell.getColumn()).toBe(3);
        });
    });

    describe("getAddress", function () {
        it("should return the address", function () {
            expect(cell.getAddress()).toBe("C5");
        });
    });

    describe("getFullAddress", function () {
        it("should return the full address", function () {
            expect(cell.getFullAddress()).toBe("'Foo'!C5");
        });
    });

    describe("setValue", function () {
    });

    describe("setFormula", function () {
    });

    describe("_clearContents", function () {
        it("should clear the node contents", function () {
            expect(cell._cellNode.findall('*').length).toBe(1);
            expect(cell._cellNode.attrib.t).toBeTruthy();
            cell._clearContents();
            expect(cell._cellNode.findall('*').length).toBe(0);
            expect(cell._cellNode.attrib.t).toBeUndefined();
        });
    });
});

describe('Sheet', function () {
    var workbook, sheet;
    beforeEach(function () {
        workbook = Workbook.fromBlankSync();
        sheet = workbook.getSheet(0);
    });
    describe('getCell', function () {
        it('is case insensitive', function () {
            var upperCaseCell = sheet.getCell('A1');
            var lowerCaseCell = sheet.getCell('a1');
            expect(upperCaseCell.getFullAddress()).toBe(lowerCaseCell.getFullAddress());
            upperCaseCell.setValue(Math.random());
            var upperCaseVNode = upperCaseCell._cellNode.find('./v');
            var lowerCaseVNode = lowerCaseCell._cellNode.find('./v');
            expect(upperCaseVNode).not.toBeNull('A1 value node should not be null');
            expect(lowerCaseVNode).not.toBeNull('a1 value node should not be null');
            expect(upperCaseVNode.text).toBe(lowerCaseVNode.text);
        });
    });
});

describe('Randomly populated sheet', function () {
    var MAX_ROW = 100;
    var MAX_COLUMN = 100;
    var MAX_EDIT = 1000;
    var workbook, sheet;
    beforeEach(function () {
        workbook = Workbook.fromBlankSync();
        sheet = workbook.getSheet(0);
        // Make random edits to the sheet
        for (var i = 0; i < MAX_EDIT; i++) {
            var rowNumber = 1 + Math.floor(MAX_ROW * Math.random());
            var columnNumber = 1 + Math.floor(MAX_COLUMN * Math.random());
            sheet
                .getCell(rowNumber, columnNumber)
                .setValue(Math.random())
                ;
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
        /*
        // Writing to file maynot be a good idea when testing
        // Write to temporary file
        var timestamp = (new Date()).getTime();
        var xlsxPath = __dirname + '/' + timestamp + '.xlsx';
        workbook.toFileSync(xlsxPath); // You should be able to open this file in MS Excel
        // Reload
        workbook = Workbook.fromFileSync(xlsxPath);
        */
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

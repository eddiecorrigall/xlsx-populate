/* jshint jasmine: true */

"use strict";

var utils = require('../lib/utils');

describe("utils", function () {
    describe("isInteger", function () {
        it("should return true for ints", function () {
            expect(utils.isInteger(4)).toBe(true);
            expect(utils.isInteger(0)).toBe(true);
            expect(utils.isInteger(-6)).toBe(true);
        });

        it("should return false for non-ints", function () {
            expect(utils.isInteger(4.7)).toBe(false);
            expect(utils.isInteger("foo")).toBe(false);
            expect(utils.isInteger("4")).toBe(false);
            expect(utils.isInteger("")).toBe(false);
            expect(utils.isInteger(null)).toBe(false);
            expect(utils.isInteger(NaN)).toBe(false);
            expect(utils.isInteger(Infinity)).toBe(false);
            expect(utils.isInteger(true)).toBe(false);
            expect(utils.isInteger(undefined)).toBe(false);
        });
    });

    describe("columnNumberToName", function () {
        it("should convert valid column number to name", function () {
            expect(utils.columnNumberToName(1)).toBe("A");
            expect(utils.columnNumberToName(3)).toBe("C");
            expect(utils.columnNumberToName(26)).toBe("Z");
            expect(utils.columnNumberToName(27)).toBe("AA");
            expect(utils.columnNumberToName(30)).toBe("AD");
        });

        it("should convert invalid column number to undefined", function () {
            expect(utils.columnNumberToName(0)).toBe(undefined);
            expect(utils.columnNumberToName(-1)).toBe(undefined);
            expect(utils.columnNumberToName(NaN)).toBe(undefined);
            expect(utils.columnNumberToName(1.5)).toBe(undefined);
            expect(utils.columnNumberToName(false)).toBe(undefined);
            expect(utils.columnNumberToName(true)).toBe(undefined);
            expect(utils.columnNumberToName(null)).toBe(undefined);
            expect(utils.columnNumberToName("")).toBe(undefined);
            expect(utils.columnNumberToName("foo")).toBe(undefined);
            expect(utils.columnNumberToName(undefined)).toBe(undefined);
        });
    });

    describe("columnNameToNumber", function () {
        it("should convert valid column name to number", function () {
            expect(utils.columnNameToNumber("A")).toBe(1);
            expect(utils.columnNameToNumber("C")).toBe(3);
            expect(utils.columnNameToNumber("Z")).toBe(26);
            expect(utils.columnNameToNumber("AA")).toBe(27);
            expect(utils.columnNameToNumber("AD")).toBe(30);
        });

        it("should convert lowercase column name the same as uppercase", function () {
            expect(utils.columnNameToNumber("a")).toBe(utils.columnNameToNumber("A"));
            expect(utils.columnNameToNumber("c")).toBe(utils.columnNameToNumber("C"));
            expect(utils.columnNameToNumber("z")).toBe(utils.columnNameToNumber("Z"));
            expect(utils.columnNameToNumber("aa")).toBe(utils.columnNameToNumber("AA"));
            expect(utils.columnNameToNumber("ad")).toBe(utils.columnNameToNumber("AD"));
        });

        it("should convert invalid column name to undefined", function () {
            expect(utils.columnNameToNumber("")).toBe(undefined);
            expect(utils.columnNameToNumber(null)).toBe(undefined);
            expect(utils.columnNameToNumber(false)).toBe(undefined);
            expect(utils.columnNameToNumber(5)).toBe(undefined);
            expect(utils.columnNameToNumber(undefined)).toBe(undefined);
        });
    });

    describe("rowAndColumnToAddress", function () {
        it("should convert valid row and column to address", function () {
            expect(utils.rowAndColumnToAddress(1, 1)).toBe("A1");
            expect(utils.rowAndColumnToAddress(10, 3)).toBe("C10");
            expect(utils.rowAndColumnToAddress(100, 27)).toBe("AA100");
        });

        it("should convert row, column, and sheet to full address", function () {
            expect(utils.rowAndColumnToAddress(1, 1, "Foo")).toBe("'Foo'!A1");
        });

        it("should convert invalid row and/or column to undefined", function () {
            expect(utils.rowAndColumnToAddress(0, 1)).toBe(undefined);
            expect(utils.rowAndColumnToAddress(1, 0)).toBe(undefined);
            expect(utils.rowAndColumnToAddress(0, 0)).toBe(undefined);
        });
    });

    describe("addressToRowAndColumn", function () {
        it("should convert valid address to row and column", function () {
            expect(utils.addressToRowAndColumn("A1")).toEqual({ row: 1, column: 1 });
            expect(utils.addressToRowAndColumn("C10")).toEqual({ row: 10, column: 3 });
            expect(utils.addressToRowAndColumn("AA100")).toEqual({ row: 100, column: 27 });
        });

        it("should convert full address to row, column, and sheet", function () {
            expect(utils.addressToRowAndColumn("Foo!A1")).toEqual({ row: 1, column: 1, sheet: "Foo" });
            expect(utils.addressToRowAndColumn("'Foo'!A1")).toEqual({ row: 1, column: 1, sheet: "Foo" });
        });

        it("should convert invalid address to row and column", function () {
            expect(utils.addressToRowAndColumn("foo")).toBe(undefined);
            expect(utils.addressToRowAndColumn(null)).toBe(undefined);
            expect(utils.addressToRowAndColumn(undefined)).toBe(undefined);
            expect(utils.addressToRowAndColumn(5)).toBe(undefined);
        });
    });
});

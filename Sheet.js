"use strict";

var utils = require('./utils'),
    etree = require('elementtree'),
    subelement = etree.SubElement,
    Cell = require('./Cell');

/**
 * Initializes a new Sheet.
 * @param {Workbook} workbook
 * @param {string} name
 * @param {etree.Element} sheetNode - The node defining the sheet in the workbook.xml.
 * @param {etree.Element} sheetXML
 * @constructor
 */
var Sheet = function (workbook, sheetNode, sheetXML) {
    this._workbook = workbook;
    this._sheetNode = sheetNode;
    this._sheetXML = sheetXML;
    this._cells = {};
};

/**
 * Gets the parent workbook.
 * @returns {Workbook}
 */
Sheet.prototype.getWorkbook = function () {
    return this._workbook;
};

/**
 * Gets the name of the sheet.
 * @returns {string}
 */
Sheet.prototype.getName = function () {
    return this._sheetNode.attrib.name;
};

Sheet.prototype.setName = function (name) {
    this._sheetNode.attrib.name = name;
};

/**
 * Gets the cell with either the provided row and column or address.
 * @returns {Cell}
 */
Sheet.prototype.getCell = function () {

    var row, column, address;

    if (arguments.length === 1) {
        address = arguments[0];
        var ref = utils.addressToRowAndColumn(address);
        row = ref.row;
        column = ref.column;
    }
    else {
        row = arguments[0];
        column = arguments[1];
        address = utils.rowAndColumnToAddress(row, column);
    }

    var getRowNode = function (sheetDataNode, row) {
        var rowNodes = sheetDataNode.findall('./row');
        for (var i = 0; i < rowNodes.length; i++) {
            if (parseInt(rowNodes[i].get('r')) === row) {
                return rowNodes[i];
            }
        }
        return null;
    };

    var getCellNode = function (rowNode, address) {
        var columnNodes = rowNode.findall('./c');
        for (var i = 0; i < columnNodes.length; i++) {
            if (columnNodes[i].get('r') === address) {
                return columnNodes[i];
            }
        }
        return null;
    }

    if ((address in this._cells) === false) {
        
        var sheetDataNode = this._sheetXML.find('./sheetData');
        
        var rowNode = getRowNode(sheetDataNode, row);
        if (!rowNode) {
            rowNode = subelement(sheetDataNode, 'row');
            rowNode.set('r', row);
        }
        
        var cellNode = getCellNode(rowNode, address);
        if (!cellNode) {
            cellNode = subelement(rowNode, 'c');
            cellNode.set('r', address);
        }

        this._cells[address] = new Cell(this, row, column, cellNode);
    }

    return this._cells[address];
};

/**
 * Given a row, assign column values to the row for eahc column name and value pair provided.
 * Example: sheet.setRow(4, { 'A': 'abc', 'D': 123 });
 */
Sheet.prototype.setColumnValues = function (row, columnValues) {
    for (var columnNameOrIndex in columnValues) {
        var column = null;
        var value = columnValues[columnNameOrIndex];
        switch (typeof(columnNameOrIndex)) {
            case 'string': {
                column = utils.columnNameToNumber(columnNameOrIndex);
            } break;
            case 'number': {
                if (utils.isInteger(value)) {
                    column = columnNameOrIndex;
                }
            } break;
        }
        if (column) {
            this.getCell(row, column).setValue(value);
        }
        else {
            console.error('Column name or index', columnNameOrIndex);
            throw new Error('setColumnValues - invalid column name or index passed');
        }
    }
};

module.exports = Sheet;

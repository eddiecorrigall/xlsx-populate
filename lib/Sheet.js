"use strict";

var utils = require('./utils'),
    etree = require('elementtree'),
    subelement = etree.SubElement,
    Cell = require('./Cell');

/**
 * Initializes a new Sheet.
 * 
 * @param {Workbook} workbook
 * @param {String} name
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
 * 
 * @returns {Workbook}
 */
Sheet.prototype.getWorkbook = function () {
    return this._workbook;
};

/**
 * Gets the name of the sheet.
 * 
 * @returns {String}
 */
Sheet.prototype.getName = function () {
    return this._sheetNode.attrib.name;
};

Sheet.prototype.setName = function (name) {
    this._sheetNode.attrib.name = name;
};

/**
 * Gets the cell with either the provided row and column or address.
 * 
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

    if ((utils.isInteger(row) === false) || (row < 0)) {
        throw new Error('Row must be a positive integer');
    }

    if ((utils.isInteger(column) === false) || (column > 0)) {
        throw new Error('Column must be a positive integer');
    }

    if ((typeof address) === 'string') {
        throw new Error('Address must be of type string');
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

    // Fast lookup

    if ((address in this._cells) === false) {
        
        var sheetDataNode = this._sheetXML.find('./sheetData');
        
        var rowNode = getRowNode(sheetDataNode, row) || subelement(sheetDataNode, 'row');
        rowNode.set('r', row);
        
        var cellNode = getCellNode(rowNode, address) || subelement(rowNode, 'c');
        cellNode.set('r', address);

        this._cells[address] = new Cell(this, row, column, cellNode);
    }

    return this._cells[address];
};

/**
 * Gets the cells within a described cell region.
 * Input parameters are of two cells which are endpoints of a rectangluar region.
 * Returns an array of Cells.
 * 
 * Example:
 *   sheet.getCellRange('A1:B2') => 
 *   sheet.getCellRange(['A1', 'B2']) => 
 *   sheet.getCellRange('A1', 'B2') => 
 *   sheet.getCellRange(sheet.getCell('A1'), sheet.getCell('B2')) => 
 *     [ this.getCell('A1'), this.getCell('A2'), this.getCell('B1'), this.getCell('B2') ]
 * 
 * @returns {Array}
 */
Sheet.prototype.getCellRange = function () {
    var ref = arguments;
    if (ref.length === 1) {
        if (typeof ref[0] === 'string') {
            ref = ref[0].split(':');
        }
        if (ref[0] instanceof Array) {
            ref = ref[0];
        }
    }
    if (typeof ref[0] === 'string') {
        ref[0] = this.getCell(ref[0]);
    }
    if (typeof ref[1] === 'string') {
        ref[1] = this.getCell(ref[1]);
    }
    if ((ref[0] instanceof Cell) === false) {
        throw new Error('Expected first reference to be Cell');
    }
    if ((ref[1] instanceof Cell) === false) {
        throw new Error('Expected second reference to be Cell');
    }
    // ...
    var cells = [];
    var c0 = ref[0].getColumn();
    var r0 = ref[0].getRow();
    var c1 = ref[1].getColumn();
    var r1 = ref[1].getRow();
    var cMin = Math.min(c0, c1);
    var cMax = Math.max(c0, c1);
    var rMin = Math.min(r0, r1);
    var rMax = Math.max(r0, r1);
    for (var r = rMin; r <= rMax; r++) {
        for (var c = cMin; c <= cMax; c++) {
            cells.push(this.getCell(r, c));
        }
    }
    return cells;
};

/**
 * Given a row, assign column values to the row for each column name and value pair provided.
 * 
 * Example: 
 *   sheet.setColumnValues(4, { 'A': 'abc', 'D': 123 });
 *   sheet.setColumnValues(4, { 1: 'abc', 4: 123 });
 * 
 * @returns {Sheet}
 */
Sheet.prototype.setColumnValues = function (row, columnValues) {
    for (var columnNameOrIndex in columnValues) {
        var column = 0;
        var value = columnValues[columnNameOrIndex];
        switch (typeof(columnNameOrIndex)) {
            case 'string': {
                column = utils.columnNameToNumber(columnNameOrIndex);
            } break;
            case 'number': {
                if (utils.isInteger(columnNameOrIndex) === false) {
                    throw new Error('Column number must be integer');
                }
                column = columnNameOrIndex;
            } break;
            default: {
                throw new Error('Column must alphabetical or numerical');
            } break;
        }
        this.getCell(row, column).setValue(value);
    }
    return this;
};

module.exports = Sheet;

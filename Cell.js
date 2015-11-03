"use strict";

var etree = require('elementtree'),
    subelement = etree.SubElement,
    utils = require('./utils');

var assert = require('assert');

/**
 * Initializes a new Cell.
 * @param {Sheet} sheet
 * @param {number} row
 * @param {number} column
 * @param {etree.SubElement} cellNode
 * @constructor
 */
var Cell = function (sheet, row, column, cellNode) {
    this._sheet = sheet;
    this._row = row;
    this._column = column;
    this._cellNode = cellNode;
};

/**
 * Gets the parent sheet.
 * @returns {Sheet}
 */
Cell.prototype.getSheet = function () {
    return this._sheet;
};

/**
 * Gets the row of the cell.
 * @returns {number}
 */
Cell.prototype.getRow = function () {
    return this._row;
};

/**
 * Gets the column of the cell.
 * @returns {number}
 */
Cell.prototype.getColumn = function () {
    return this._column;
};

/**
 * Gets the address of the cell (e.g. "A5").
 * @returns {string}
 */
Cell.prototype.getAddress = function () {
    return utils.rowAndColumnToAddress(this._row, this._column);
};

/**
 * Gets the full address of the cell including sheet (e.g. "Sheet1!A5").
 * @returns {string}
 */
Cell.prototype.getFullAddress = function () {
    return utils.rowAndColumnToAddress(this._row, this._column, this.getSheet().getName());
};

/**
 * Sets the value of the cell.
 * @param {*} value
 * @returns {Cell}
 */
Cell.prototype.setValue = function (value) {
    switch (typeof(value)) {
        case 'number': {
            this._clearContents();
            this._cellNode.set('t', 'n');
            var vNode = subelement(this._cellNode, 'v');
            vNode.text = ('' + value);
        } break;
        case 'string': { // Not tested
            this._clearContents();
            this._cellNode.set('t', 'inlineStr');
            var isNode = subelement(this._cellNode, 'is');
            var tNode = subelement(isNode, 't');
            tNode.text = value;
        } break;
        case 'boolean': { // Not tested
            this._clearContents();
            this._cellNode.set('t', 'b');
            var vNode = subelement(this._cellNode, 'v');
            vNode.text = value ? 'true' : 'false';
        } break;
        default: {
            if (value instanceof Date) {
                this._clearContents();
                this._cellNode.set('t', 'd');
                this._cellNode.set('s', '2'); // hack
                var vNode = subelement(this._cellNode, 'v');
                vNode.text = ('' + value.toISOString());
            }
        } break;
    }
    return this;
};

/**
 * Sets the formula for a cell (with optional precalculated value).
 * @param {string} formula
 * @param {*} calculatedValue
 * @param {Integer} sharedIndex
 * @param {Cell Array} sharedRef
 * @returns {Cell}
 */
Cell.prototype.setFormula = function (formula, calculatedValue, sharedIndex, sharedRef) {
    this._clearContents();
    var fNode = subelement(this._cellNode, 'f'); // formula node
    if (typeof formula === 'string') {
        fNode.text = formula;
    }
    if (calculatedValue !== undefined) {
        var vNode = subelement(this._cellNode, 'v'); // value node
        vNode.text = calculatedValue;
    }
    if (utils.isInteger(sharedIndex) && (0 <= sharedIndex)) {
        // TODO: Check that sharedIndex is unique
        fNode.set('t', 'shared');
        fNode.set('si', '' + sharedIndex);
    }
    if (typeof sharedRef === 'string') {
        fNode.set('ref', sharedRef);
    }
    return this;
};

/**
 * Check if the current cell is a shared formula.
 * @returns {boolean}
 */
Cell.prototype.isSharedFormula = function (isSource) {
    isSource = isSource || false;
    /*
    <sheetData>
        <row ...>
            <c ...>
                <f ref="F2:F519" si="0" t="shared">C2/B2</f>
                ...
            </c>
        </row>
    </sheetData>
    */
    var fNode = this._cellNode.find('./f');
    if (!fNode) { // check for formula
        console.log(this.getAddress(), 'fNode not found', fNode);
        return false;
    }
    if (!fNode.text.length) { // check for formula content
        console.log(this.getAddress(), 'fNode formula is empty', fNode.text);
        return false;
    }
    var fNodeType = fNode.get('t');
    if (fNodeType !== 'shared') { // check for type
        console.log(this.getAddress(), 'fNode type not shared', fNodeType);
        return false;
    }
    var fNodeSharedIndex = fNode.get('si');
    if (!fNodeSharedIndex.length) { // check for shared/string index
        console.log(this.getAddress(), 'fNode si is empty', fNodeSharedIndex);
        return false;
    }
    var fNodeRef = fNode.get('ref');
    if (isSource && !fNodeRef.length) { // check for reference
        console.log(this.getAddress(), 'fNode ref is empty', fNodeRef);
        return false;
    }
    return true;
};

Cell.prototype.hasSameRow = function () {
    var otherCell = (arguments[0] instanceof Cell)
        ? arguments[0]
        : this.getSheet().getCell(arguments)
        ;
    return this.getRow() === otherCell.getRow();
};

Cell.prototype.hasSameColumn = function () {
    var otherCell = (arguments[0] instanceof Cell)
        ? arguments[0]
        : this.getSheet().getCell(arguments)
        ;
    return this.getColumn() === otherCell.getColumn();
};

Cell.prototype.isEqual = function () {
    var otherCell = (arguments[0] instanceof Cell)
        ? arguments[0]
        : this.getSheet().getCell(arguments)
        ;
    return this.hasSameRow(otherCell) && this.hasSameColumn(otherCell);
}

// Examples:
// sheet.getCell('C1').setSharedFormula('C3')
// sheet.getCell('A3').setSharedFormula('C3')
Cell.prototype.shareFormula = function (lastSharedCell) {
    assert(this.isSharedFormula(true), 'Expected cell to be a shared formula');
    var fNode = this._cellNode.find('./f');
    assert(fNode !== null, 'Expected <f> node to be defined');
    var sharedIndex = parseInt(fNode.get('si'));
    assert(utils.isInteger(sharedIndex) && (0 <= sharedIndex), 'Expected shared index to be a non-negative integer');
    // ...
    if (typeof lastSharedCell === 'string') {
        lastSharedCell = this.getCell(lastSharedCell);
    }
    assert(lastSharedCell instanceof Cell, 'Expected shared reference to be a cell');
    // ...
    if (this.hasSameRow(lastSharedCell)) {
        for (var c = 1+this.getColumn(); c <= lastSharedCell.getColumn(); c++) {
            var cell = this.getSheet().getCell(this.getRow(), c);
            cell.setFormula(undefined, undefined, sharedIndex);
        }
    }
    else if (this.hasSameColumn(lastSharedCell)) {
        for (var r = 1+this.getRow(); r <= lastSharedCell.getRow(); r++) {
            var cell = this.getSheet().getCell(r, this.getColumn());
            assert(cell instanceof Cell, 'Expected valid cell');
            cell.setFormula(undefined, undefined, sharedIndex);
        }
    }
    else {
        throw new Error('Expected last shared forumla cell to align either row-wise or column-wise with shared formula source');
    }
    // ...
    var sharedRef = this.getAddress() + ':' + lastSharedCell.getAddress();
    fNode.set('ref', sharedRef);
    // ...
    return this;
};

/**
 * Returns a cell with a relative position to the offsets provided.
 * @param {integer} rowOffset
 * @param {integer} columnOffset
 * @returns {Cell}
 */
Cell.prototype.getRelativeCell = function (rowOffset, columnOffset) {
    if (!utils.isInteger(rowOffset)) return null;
    if (!utils.isInteger(columnOffset)) return null;
    var absoluteRow = this.getRow()+rowOffset;
    assert(absoluteRow >= 0, 'Expected relative row to be a non-negative');
    var absoluteColumn = this.getColumn()+columnOffset;
    assert(absoluteColumn >= 0, 'Expected relative column to be a non-negative');
    return this.getSheet().getCell(absoluteRow, absoluteColumn);
};

/**
 * Clears the contents from the cell.
 * @private
 */
Cell.prototype._clearContents = function () {
    var self = this;
    this._cellNode.getchildren().forEach(function (childNode) {
        self._cellNode.remove(childNode);
    });
    delete this._cellNode.attrib.t;
};

module.exports = Cell;

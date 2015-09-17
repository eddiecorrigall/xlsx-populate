"use strict";

var etree = require('elementtree'),
    subelement = etree.SubElement,
    utils = require('./utils');

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
    if (formula) {
        fNode.text = formula;
    }
    if (utils.isInteger(sharedIndex) && (0 <= sharedIndex)) {
        // TODO: Check that sharedIndex is unique
        fNode.set('t', 'shared');
        fNode.set('si', sharedIndex);
    }
    // Example:
    // [sheet.getCell('B2'), sheet.getCell('B10')] => 'B2:B10'
    // var cell = sheet.getCell('B2');
    // [cell, cell.getRelativeCell(10)] => 'B2:B10'
    if (sharedRef instanceof Array) {
        // TODO: consider moving functionality from Sheet.setShareFormulaColumns here
        if (sharedRef.length === 2) {
            if ((sharedRef[0] instanceof Cell) && (sharedRef[1] instanceof Cell)) {
                var sharedRefString = sharedRef
                    .map(function (c) { return c.getAddress(); })
                    .join(':')
                    ;
                fNode.set('ref', sharedRefString);
            }
        }
    }
    if (calculatedValue) {
        var vNode = subelement(this._cellNode, 'v'); // value node
        vNode.text = calculatedValue;
    }
    return this;
};

/**
 * Check if the current cell is a shared formula.
 * @returns {boolean}
 */
Cell.prototype.isSharedFormula = function () {
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
        console.log('fNode not found');
        return false;
    }
    if (!fNode.text.length) { // check for formula content
        console.log('fNode formula is empty');
        return false;
    }
    if (fNode.get('t') !== 'shared') { // check for type
        console.log('fNode type not shared');
        return false;
    }
    if (!fNode.get('ref').length) { // check for reference
        console.log('fNode ref is empty');
        return false;
    }
    if (!fNode.get('si').length) { // check for shared/string index
        console.log('fNode si is empty');
        return false;
    }
    return true;
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
    var relativeRow = this.getRow()+rowOffset;
    if (relativeRow <= 0) return null;
    var relativeColumn = this.getColumn()+columnOffset;
    if (relativeColumn <= 0) return null;
    return this.getSheet().getCell(relativeRow, relativeColumn);
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

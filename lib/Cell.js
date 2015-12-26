"use strict";

var etree = require('elementtree'),
    subelement = etree.SubElement,
    utils = require('./utils');

var Sheet = require('./Sheet');

/**
 * Initializes a new Cell.
 * 
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
 * 
 * @returns {Sheet}
 */
Cell.prototype.getSheet = function () {
    return this._sheet;
};

/**
 * Gets the row of the cell.
 * 
 * @returns {number}
 */
Cell.prototype.getRow = function () {
    return this._row;
};

/**
 * Gets the column of the cell.
 * 
 * @returns {number}
 */
Cell.prototype.getColumn = function () {
    return this._column;
};

/**
 * Gets the address of the cell (e.g. "A5").
 * 
 * @returns {string}
 */
Cell.prototype.getAddress = function () {
    return utils.rowAndColumnToAddress(this.getRow(), this.getColumn());
};

/**
 * Gets the full address of the cell including sheet (e.g. "Sheet1!A5").
 * 
 * @returns {string}
 */
Cell.prototype.getFullAddress = function () {
    return utils.rowAndColumnToAddress(this.getRow(), this.getColumn(), this.getSheet().getName());
};

/**
 * Sets the value of the cell.
 * Returns itself.
 * 
 * @param {*} value
 * @returns {Cell}
 */
Cell.prototype.setValue = function (value) {
    switch (typeof value) {
        case 'number': {
            this._clearContents();
            this._cellNode.set('t', 'n');
            var vNodeNumber = subelement(this._cellNode, 'v');
            vNodeNumber.text = ('' + value);
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
            var vNodeBoolean = subelement(this._cellNode, 'v');
            vNodeBoolean.text = value ? 'true' : 'false';
        } break;
        default: {
            if (value instanceof Date) {
                this._clearContents();
                this._cellNode.set('t', 'd');
                this._cellNode.set('s', '2'); // hack
                var vNodeDate = subelement(this._cellNode, 'v');
                vNodeDate.text = ('' + value.toISOString());
            }
        } break;
    }
    return this;
};

/**
 * Sets the formula for a cell (with optional precalculated value).
 * Returns itself.
 * 
 * @param {string} formula
 * @param {*} calculatedValue
 * @param {number} sharedIndex
 * @param {string} sharedRef
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
 * When isSource is set true, the function will check if the formula is defined.
 * 
 * @param {boolean} isSource
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
        console.log(this.getAddress(), 'fNode not found');
        console.log(JSON.stringify(this._cellNode));
        return false;
    }
    if (isSource) {
        if (!fNode.text || !fNode.text.length) { // check for formula content
            console.log(this.getAddress(), 'fNode formula is empty');
            console.log(JSON.stringify(fNode));
            return false;
        }
        var fNodeRef = fNode.get('ref');
        if (!fNodeRef || !fNodeRef.length) { // check for reference
            console.log(this.getAddress(), 'fNode ref is empty');
            console.log(JSON.stringify(fNode));
            return false;
        }
    }
    var fNodeType = fNode.get('t');
    if (fNodeType !== 'shared') { // check for type
        console.log(this.getAddress(), 'fNode type not shared');
        console.log(JSON.stringify(fNode));
        return false;
    }
    var fNodeSharedIndex = fNode.get('si');
    if (!fNodeSharedIndex || !fNodeSharedIndex.length) { // check for shared/string index
        console.log(this.getAddress(), 'fNode si is empty');
        console.log(JSON.stringify(fNode));
        return false;
    }
    return true;
};

/**
 * Check if both cells share the same row.
 * Return true if rows are equal to otherCell, otherwise false.
 * 
 * @param {Cell} otherCell
 * @returns {boolean}
 */
Cell.prototype.hasSameRow = function () {
    var otherCell = Sheet.prototype.getCell.apply(this, arguments);
    if (!otherCell) return false;
    return this.getRow() === otherCell.getRow();
};

/**
 * Check if both cells share the same column.
 * Return true if columns are equal to otherCell, otherwise false.
 * 
 * @param {Cell} otherCell
 * @returns {boolean}
 */
Cell.prototype.hasSameColumn = function () {
    var otherCell = Sheet.prototype.getCell.apply(this, arguments);
    if (!otherCell) return false;
    return this.getColumn() === otherCell.getColumn();
};

/**
 * Check if both cells share the same row and column.
 * Return true if row and column are equal to otherCell, otherwise false.
 * 
 * @param {Cell} otherCell
 * @returns {boolean}
 */
Cell.prototype.isSame = function () {
    var otherCell = (arguments[0] instanceof Cell)
        ? arguments[0]
        : Sheet.prototype.getCell.apply(this, arguments)
        ;
    if (!otherCell) return false;
    return this.getFullAddress() === otherCell.getFullAddress();
};

/**
 * If this cell is the source of a shared formula,
 * then assign all cells from this cell to lastSharedCell.
 * Note that lastSharedCell must share the same row or column.
 * Returns itself.
 * 
 * @param {Cell} lastSharedCell
 * @returns {Cell}
 */
Cell.prototype.shareFormulaUntil = function (lastSharedCell) {
    if (this.isSharedFormula(true) === false) {
        throw new Error('Expected cell to be a shared formula');
    }
    var fNode = this._cellNode.find('./f');
    if (fNode === null) {
        throw new Error('Expected <f> node to be defined');
    }
    var sharedIndex = parseInt(fNode.get('si'));
    if ((utils.isInteger(sharedIndex) === false) || (sharedIndex < 0)) {
        throw new Error('Expected shared index to be a non-negative integer');
    }
    // ...
    if ((typeof lastSharedCell) === 'string') {
        lastSharedCell = this.getCell(lastSharedCell);
    }
    if ((lastSharedCell instanceof Cell) === false) {
        throw new Error('Expected shared reference to be a cell');
    }
    // ...
    if (this.hasSameRow(lastSharedCell)) {
        for (var c = 1 + this.getColumn(); c <= lastSharedCell.getColumn(); c++) {
            this
                .getSheet()
                .getCell(this.getRow(), c)
                .setFormula(undefined, undefined, sharedIndex)
                ;
        }
    } else if (this.hasSameColumn(lastSharedCell)) {
        for (var r = 1 + this.getRow(); r <= lastSharedCell.getRow(); r++) {
            var cell = this.getSheet().getCell(r, this.getColumn());
            if ((cell instanceof Cell) === false) {
                throw new Error('Expected valid cell');
            }
            cell.setFormula(undefined, undefined, sharedIndex);
        }
    } else {
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
 * 
 * @param {number} rowOffset
 * @param {number} columnOffset
 * @returns {Cell}
 */
Cell.prototype.getRelativeCell = function (rowOffset, columnOffset) {
    if (!utils.isInteger(rowOffset)) return null;
    if (!utils.isInteger(columnOffset)) return null;
    var absoluteRow = rowOffset + this.getRow();
    if (absoluteRow < 0) {
        throw new Error('Expected relative row to be a non-negative');
    }
    var absoluteColumn = columnOffset + this.getColumn();
    if (absoluteColumn < 0) {
        throw new Error('Expected relative column to be a non-negative');
    }
    return this.getSheet().getCell(absoluteRow, absoluteColumn);
};

/**
 * Clears the contents from the cell.
 * 
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

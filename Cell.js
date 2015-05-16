"use strict";

var etree = require('elementtree'),
    subelement = etree.SubElement,
    utils = require('./utils'),
    Style = require('./Style');

/**
 * Initializes a new Cell.
 * @param {Sheet} sheet
 * @param {number} row
 * @param {number} column
 * @param {etree.SubElement} cellNode
 * @constructor
 */
var Cell = function (sheet, row, column, cellNode, stylesXML) {
    this._sheet = sheet;
    this._row = row;
    this._column = column;
    this._cellNode = cellNode;
    this._stylesXML = stylesXML;
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
    var vNode;
    this._clearContents();

    if (typeof value === "string") {
        this._cellNode.attrib.t = "inlineStr";
        var isNode = subelement(this._cellNode, "is");
        var tNode = subelement(isNode, "t");
        tNode.text = value;
    } else if (typeof value === "number") {
        vNode = subelement(this._cellNode, "v");
        vNode.text = value;
    } else if (typeof value === "boolean") {
        this._cellNode.attrib.t = "b";
        vNode = subelement(this._cellNode, "v");
        vNode.text = value ? 1 : 0;
    } else if (value instanceof Date) {
        // TODO: Style
        vNode = subelement(this._cellNode, "v");
        vNode.text = utils.dateToExcelNumber(value);
    }

    return this;
};

/**
 * Sets the formula for a cell (with optional precalculated value).
 * @param {string} formula
 * @param {*=} calculatedValue
 * @returns {Cell}
 */
Cell.prototype.setFormula = function (formula, calculatedValue) {
    this._clearContents();

    var fNode = subelement(this._cellNode, "f");
    fNode.text = formula;

    if (arguments.length > 1) {
        var vNode = subelement(this._cellNode, "v");
        vNode.text = calculatedValue;
    }

    return this;
};

Cell.prototype.getStyle = function () {
    var nodeAndIndex = utils.copyOrAddNode(this._stylesXML, "cellXfs", "xf", this._cellNode.attrib.s);
    this._cellNode.attrib.s = nodeAndIndex.index;
    return new Style(nodeAndIndex.node, this._stylesXML);
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

"use strict";

var Alignment = function (alignmentNode) {
    this._alignmentNode = alignmentNode;
};

Alignment.prototype.setHorizontal = function (value) {
    this._alignmentNode.attrib.horizontal = value;
};

Alignment.prototype.setVertical = function (value) {
    this._alignmentNode.attrib.vertical = value;
};

Alignment.prototype.setWrapText = function (value) {
    this._alignmentNode.attrib.wrapText = value ? 1 : 0;
};

module.exports = Alignment;

"use strict";

var etree = require('elementtree'),
    subelement = etree.SubElement;

var Font = function (fontNode) {
    this._fontNode = fontNode;
};

Font.prototype._setValue = function (value, nodeName, attribName) {
    var current = this._fontNode.find(nodeName);
    if (current) this._fontNode.remove(current);

    if (value) {
        var node = subelement(this._fontNode, nodeName);
        if (attribName) node.attrib[attribName] = value;
    }

    return this;
};

Font.prototype.setBold = function (value) {
    return this._setValue(value, "b");
};

Font.prototype.setItalics = function (value) {
    return this._setValue(value, "i");
};

Font.prototype.setUnderline = function (value) {
    return this._setValue(value, "u", value === "double" && "val");
};

Font.prototype.setColor = function (value) {
    return this._setValue(value, "color", "rgb");
};

Font.prototype.setName = function (value) {
    return this._setValue(value, "name", "val");
};

Font.prototype.setSize = function (value) {
    return this._setValue(value, "sz", "val");
};

module.exports = Font;

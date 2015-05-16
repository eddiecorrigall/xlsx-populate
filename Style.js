"use strict";

var utils = require('./utils'),
    Font = require('./Font'),
    Alignment = require('./Alignment'),
    etree = require('elementtree'),
    subelement = etree.SubElement;

var Style = function (styleNode, stylesXML) {
    this._styleNode = styleNode;
    this._stylesXML = stylesXML;
};

Style.prototype.getFont = function () {
    var nodeAndIndex = utils.copyOrAddNode(this._stylesXML, "fonts", "font", this._styleNode.attrib.fontId);
    this._styleNode.attrib.fontId = nodeAndIndex.index;
    return new Font(nodeAndIndex.node, this._stylesXML);
};

Style.prototype.getAlignment = function () {
    var alignmentNode = this._styleNode.find("alignment");
    if (!alignmentNode) {
        alignmentNode = subelement(this._styleNode, "alignment");
    }

    return new Alignment(alignmentNode);
};

module.exports = Style;

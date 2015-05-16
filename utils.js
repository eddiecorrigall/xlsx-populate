"use strict";

var etree = require('elementtree'),
    subelement = etree.SubElement;

/**
 * Regex to parse Excel addresses.
 */
var addressRegex = /^\s*(?:'?(.+?)'?\!)?\$?([A-Z]+)\$?(\d+)\s*$/i;

/**
 * The base date to use for Excel conversion.
 */
var dateBase = new Date(1900, 0, 0);

/**
 * The last day in February 1900.
 */
var incorrectLeapDate = new Date(1900, 1, 28);

/**
 * Number of milliseconds in a day.
 */
var millisecondsInDay = 1000 * 60 * 60 * 24;

module.exports = {
    /**
     * Checks if a number is an integer. Taken from here:
     * http://stackoverflow.com/questions/3885817/how-to-check-if-a-number-is-float-or-integer#3885844
     * @param {*} value
     * @returns {boolean}
     */
    isInteger: function (value) {
        return value === +value && value === (value | 0);
    },

    /**
     * Converts a column number to column name (e.g. 2 -> "B").
     * @param {number} number
     * @returns {string}
     */
    columnNumberToName: function (number) {
        if (!this.isInteger(number) || number <= 0) return;

        var dividend = number;
        var name = '';
        var modulo = 0;

        while (dividend > 0) {
            modulo = (dividend - 1) % 26;
            name = String.fromCharCode('A'.charCodeAt(0) + modulo) + name;
            dividend = Math.floor((dividend - modulo) / 26);
        }

        return name;
    },

    /**
     * Converts a column name to column number (e.g. "B" -> 2).
     * @param {string} name
     * @returns {number}
     */
    columnNameToNumber: function (name) {
        if (!name || typeof name !== "string") return;

        name = name.toUpperCase();
        var sum = 0;
        for (var i = 0; i < name.length; i++) {
            sum *= 26;
            sum += (name[i].charCodeAt(0) - 'A'.charCodeAt(0) + 1);
        }

        return sum;
    },

    /**
     * Converts a row and column (and option sheet) to an Excel address.
     * @param {number} row
     * @param {number} column
     * @param {string=} sheet
     * @returns {string}
     */
    rowAndColumnToAddress: function (row, column, sheet) {
        if (!this.isInteger(row) || !this.isInteger(column) || row <= 0 || column <= 0) return;
        var address = this.columnNumberToName(column) + row;
        if (sheet) address = "'" + sheet + "'!" + address;
        return address;
    },

    /**
     * Converts an address to row and column (and sheet if present).
     * @param {string} address
     * @returns {*}
     */
    addressToRowAndColumn: function (address) {
        var match = addressRegex.exec(address);
        if (!match) return;

        var ref = {
            row: parseInt(match[3]),
            column: this.columnNameToNumber(match[2])
        };

        if (match[1]) ref.sheet = match[1];

        return ref;
    },

    /**
     * Converts a date to an Excel number. (The number of days since 1/1/1900.
     * @param {Date} date
     * @returns {number}
     */
    dateToExcelNumber: function (date) {
        var num = (date - dateBase) / millisecondsInDay;

        // "Bug" in Excel that treats 1900 as a leap year.
        if (date > incorrectLeapDate) num += 1;

        return num;
    },

    copyOrAddNode: function (xml, parentPath, childName, currentIndex) {
        var parentNode = xml.find(parentPath);
        var childNodes = parentNode.findall(childName);

        var childNode;
        var index = childNodes.length;

        if (currentIndex === undefined) {
            childNode = subelement(parentNode, childName);
        } else {
            currentIndex = parseInt(currentIndex);
            var currentNode = childNodes[currentIndex];
            childNode = JSON.parse(JSON.stringify(currentNode));
            parentNode.append(childNode);
        }

        return {
            node: childNode,
            index: index
        };
    }
};

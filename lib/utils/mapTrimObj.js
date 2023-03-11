"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.mapTrimObj = void 0;
const mapTrimObj = (row, colsToKeep) => {
    const filteredRow = {};
    colsToKeep.forEach((colName) => (filteredRow[colName] = Object.keys(row).includes(colName)
        ? row[colName]
        : ""));
    // if (Object.keys(row).includes("rowIndex"))
    // return {
    //   ...filteredRow,
    //   rowIndex: row.rowIndex,
    //   a1Range: row.a1Range,
    // } as ExtRow;
    return filteredRow;
};
exports.mapTrimObj = mapTrimObj;

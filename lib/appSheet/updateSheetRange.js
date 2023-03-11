"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.updateSheetRange = void 0;
const alphanumeric_encoder_1 = __importDefault(require("alphanumeric-encoder"));
const google_1 = require("../google");
const updateSheetRange = async ({ sheetId, tabName, startCoords, data, }) => {
    const sheetApp = (0, google_1.appSheet)();
    const encoder = new alphanumeric_encoder_1.default();
    const encodedStartCol = encoder.encode(startCoords[1] || 1);
    const encodedEndCol = encoder.encode((startCoords[1] || 1) - 1 + data[0].length);
    const rangeA1notation = `'${tabName}'!${encodedStartCol}${startCoords[0] || 1}:${encodedEndCol}${data.length + (startCoords[0] || 1) - 1}`;
    await sheetApp.spreadsheets.values.update({
        spreadsheetId: sheetId,
        range: rangeA1notation,
        valueInputOption: "USER_ENTERED",
        requestBody: {
            values: data,
        },
    });
};
exports.updateSheetRange = updateSheetRange;

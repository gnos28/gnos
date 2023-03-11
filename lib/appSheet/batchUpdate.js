"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.batchUpdate = void 0;
const google_1 = require("../google");
const protectedRangeBatchBuffer = {};
exports.batchUpdate = {
    addProtectedRange: ({ spreadsheetId, editors, namedRangeId, sheetId, startColumnIndex, startRowIndex, endColumnIndex, endRowIndex, }) => {
        if (!protectedRangeBatchBuffer[spreadsheetId])
            protectedRangeBatchBuffer[spreadsheetId] = [];
        const namedRange = namedRangeId
            ? namedRangeId
            : Math.random().toString().split(".")[1];
        protectedRangeBatchBuffer[spreadsheetId].push({
            addProtectedRange: {
                protectedRange: {
                    editors: { users: editors },
                    description: namedRange,
                    range: {
                        sheetId,
                        startColumnIndex,
                        startRowIndex,
                        endColumnIndex: endColumnIndex + 1,
                        endRowIndex: endRowIndex + 1,
                    },
                },
            },
        });
    },
    runProtectedRange: async (spreadsheetId) => {
        const sheetApp = (0, google_1.appSheet)();
        const requests = protectedRangeBatchBuffer[spreadsheetId];
        if (requests) {
            console.log("requests count : ", requests.length);
            // console.log("requests", requests);
            await sheetApp.spreadsheets.batchUpdate({
                spreadsheetId,
                requestBody: {
                    requests,
                },
            });
            protectedRangeBatchBuffer[spreadsheetId] = [];
        }
    },
};

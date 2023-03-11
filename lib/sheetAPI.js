"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.sheetAPI = void 0;
const batchUpdate_1 = require("./appSheet/batchUpdate");
const getSheetTabIds_1 = require("./appSheet/getSheetTabIds");
const google_1 = require("./google");
const importSheetData_1 = require("./appSheet/importSheetData");
const updateSheetRange_1 = require("./appSheet/updateSheetRange");
const clearSheetRows_1 = require("./appSheet/clearSheetRows");
let tabCache = {};
let tabIdsCache = {};
let lastReadRequestTime = undefined;
let lastWriteRequestTime = undefined;
let nbInQueueRead = 0;
let nbInQueueWrite = 0;
let readCatchCount = 0;
let writeCatchCount = 0;
let AUTH_JSON_PATH = "./";
const DELAY = 200; // in ms
const CATCH_DELAY_MULTIPLIER = 10;
const MAX_CATCH_COUNT = 60;
const handleReadTryCatch = async (callback, delayMultiplier) => {
    let res = undefined;
    try {
        res = await callback();
        lastReadRequestTime = new Date().getTime();
        nbInQueueRead -= delayMultiplier || 1;
    }
    catch (e) {
        console.log(`inside catch 💩 #${readCatchCount}`, e.message);
        readCatchCount++;
        lastReadRequestTime = new Date().getTime();
        nbInQueueRead -= delayMultiplier || 1;
        if (readCatchCount < MAX_CATCH_COUNT)
            res = await handleReadDelay(callback, CATCH_DELAY_MULTIPLIER);
    }
    finally {
        readCatchCount = 0;
        return res;
    }
};
const handleReadDelay = async (callback, delayMultiplier) => {
    const currentTime = new Date().getTime();
    nbInQueueRead += delayMultiplier || 1;
    if (lastReadRequestTime &&
        currentTime < lastReadRequestTime + DELAY * nbInQueueRead) {
        console.log("*** force DELAY [READ] ", nbInQueueRead, lastReadRequestTime
            ? lastReadRequestTime + DELAY * nbInQueueRead - currentTime
            : 0);
        await new Promise((resolve) => setTimeout(() => resolve(null), lastReadRequestTime
            ? lastReadRequestTime + DELAY * nbInQueueRead - currentTime
            : 0));
    }
    const res = await handleReadTryCatch(callback, delayMultiplier);
    return res;
};
const handleWriteTryCatch = async (callback, delayMultiplier) => {
    let res = undefined;
    try {
        res = await callback();
        lastWriteRequestTime = new Date().getTime();
        nbInQueueWrite -= delayMultiplier || 1;
    }
    catch (e) {
        console.log(`inside catch 💩 #${writeCatchCount}`, e.message);
        writeCatchCount++;
        lastWriteRequestTime = new Date().getTime();
        nbInQueueWrite -= delayMultiplier || 1;
        if (writeCatchCount < MAX_CATCH_COUNT)
            res = await handleWriteDelay(callback, CATCH_DELAY_MULTIPLIER);
    }
    finally {
        writeCatchCount = 0;
        return res;
    }
};
const handleWriteDelay = async (callback, delayMultiplier) => {
    const currentTime = new Date().getTime();
    nbInQueueWrite += delayMultiplier || 1;
    if (lastWriteRequestTime &&
        currentTime < lastWriteRequestTime + DELAY * nbInQueueWrite) {
        console.log("*** force DELAY [WRITE]", nbInQueueWrite, lastWriteRequestTime
            ? lastWriteRequestTime + DELAY * nbInQueueWrite - currentTime
            : 0);
        await new Promise((resolve) => setTimeout(() => resolve(null), lastWriteRequestTime
            ? lastWriteRequestTime + DELAY * nbInQueueWrite - currentTime
            : 0));
    }
    const res = await handleWriteTryCatch(callback, delayMultiplier);
    return res;
};
exports.sheetAPI = {
    setAuthJsonPath: (path) => {
        AUTH_JSON_PATH = path;
    },
    getTabIds: async (sheetId) => {
        console.log("*** sheetAPI.getTabIds", sheetId);
        if (sheetId) {
            const cacheKey = sheetId;
            if (tabIdsCache[cacheKey] === undefined) {
                await handleReadDelay(async () => {
                    tabIdsCache[cacheKey] = await (0, getSheetTabIds_1.getSheetTabIds)(sheetId);
                });
            }
            else
                console.log("*** using cache 👍");
            return tabIdsCache[cacheKey];
        }
        return [];
    },
    getTabData: async (sheetId, tabList, tabName, headerRowIndex) => {
        var _a;
        console.log("*** sheetAPI.getTabData", tabName);
        const tabId = (_a = tabList.filter((tab) => tab.sheetName === tabName)[0]) === null || _a === void 0 ? void 0 : _a.sheetId;
        if (tabId === undefined)
            throw new Error(`tab ${tabName} not found`);
        const cacheKey = sheetId + ":" + tabId;
        if (tabCache[cacheKey] === undefined) {
            await handleReadDelay(async () => {
                tabCache[cacheKey] = await (0, importSheetData_1.importSheetData)(sheetId, tabId, headerRowIndex);
            });
        }
        else
            console.log("*** using cache 👍");
        return tabCache[cacheKey];
    },
    getTabMetaData: async ({ spreadsheetId, fields, ranges, }) => {
        console.log("*** sheetAPI.getTabMetaData");
        const metaData = await handleReadDelay(async () => {
            const sheetApp = (0, google_1.appSheet)();
            return await sheetApp.spreadsheets.get({
                spreadsheetId,
                fields,
                ranges,
            });
        });
        return metaData;
    },
    clearCache: () => {
        tabCache = {};
        tabIdsCache = {};
    },
    updateRange: async ({ sheetId, tabName, startCoords, data, }) => {
        console.log("*** sheetAPI.updateRange");
        await handleWriteDelay(async () => {
            await (0, updateSheetRange_1.updateSheetRange)({
                sheetId,
                tabName,
                startCoords,
                data,
            });
        });
    },
    addBatchProtectedRange: ({ spreadsheetId, editors, namedRangeId, sheetId, startColumnIndex, startRowIndex, endColumnIndex, endRowIndex, }) => {
        batchUpdate_1.batchUpdate.addProtectedRange({
            spreadsheetId,
            editors,
            namedRangeId,
            sheetId,
            startColumnIndex,
            startRowIndex,
            endColumnIndex,
            endRowIndex,
        });
    },
    runBatchProtectedRange: async (spreadsheetId) => {
        console.log("*** sheetAPI.runBatchProtectedRange");
        await handleWriteDelay(async () => {
            await batchUpdate_1.batchUpdate.runProtectedRange(spreadsheetId);
        });
    },
    clearTabData: async ({ sheetId, tabList, tabName, headerRowIndex, }) => {
        console.log("*** sheetAPI.clearTabData");
        await handleWriteDelay(async () => {
            await (0, clearSheetRows_1.clearTabData)(sheetId, tabList, tabName, headerRowIndex);
        });
    },
};

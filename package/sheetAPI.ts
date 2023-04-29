// import path from "path";
// import callerPath from "caller-path";
import { batchUpdate } from "./appSheet/batchUpdate";
import { getSheetTabIds } from "./appSheet/getSheetTabIds";
import { appSheet } from "./google";
import { importSheetData } from "./appSheet/importSheetData";
import { updateSheetRange } from "./appSheet/updateSheetRange";
import { clearTabData } from "./appSheet/clearSheetRows";
import { DataRowWithId, TabDataItem, TabListItem } from "./interfaces";
import { getTabSize } from "./appSheet/getTabSize";
import { exportToSheet } from "./appSheet/exportToSheet";

type TabCache = {
  [key: string]: ({
    [key: string]: string;
  } & {
    rowIndex: number;
    a1Range: string;
  })[];
};

type TabIdsCache = {
  [key: string]: {
    tabId: string;
    tabName: string;
  }[];
};

type TabSizeCache = {
  [key: string]: TabSize;
};

type SetAuthJsonPathProps = {
  path: string;
};

type GetTabIdsProps = {
  sheetId: string | undefined;
};

type GetTabDataProps = {
  sheetId: string;
  tabName: string;
  headerRowIndex?: number;
  tabList?: TabListItem[];
};

export type AddProtectedRangeProps = {
  sheetId: string;
  editors: string[];
  namedRangeId?: string;
  tabId: number;
  startColumnIndex: number;
  startRowIndex: number;
  endColumnIndex: number;
  endRowIndex: number;
};

type GetProtectedRangeIdsProps = {
  sheetId: string;
  tabName: string;
};

type DeleteProtectedRangeProps = {
  sheetId: string;
  protectedRangeIds: number[];
};

type RunBatchProtectedRangeProps = {
  sheetId: string;
};

type GetTabMetaDataProps = {
  sheetId: string;
  fields: string;
  ranges: string[];
};

type GetTabSizeProps = {
  sheetId: string;
  tabName: string;
};

type TabSize = {
  nbRows: number | undefined;
  nbColumns: number | undefined;
};

type UpdateSheetRangeProps = {
  sheetId: string;
  tabName: string;
  startCoords: [number, number];
  data: any[][];
};

type AppendToSheetProps = {
  sheetId: string;
  tabName: string;
  data: DataRowWithId[];
};

type ClearTabDataProps = {
  sheetId: string;
  tabName: string;
  headerRowIndex?: number;
  tabList?: TabListItem[];
};

let tabCache: TabCache = {};
let tabIdsCache: TabIdsCache = {};
let tabSizesCache: TabSizeCache = {};
let lastReadRequestTime: number | undefined = undefined;
let lastWriteRequestTime: number | undefined = undefined;
let nbInQueueRead = 0;
let nbInQueueWrite = 0;
let readCatchCount = 0;
let writeCatchCount = 0;

let AUTH_JSON_PATH = "./auth.json";
let VERBOSE_MODE = false;

const DELAY = 200; // in ms
const CATCH_DELAY_MULTIPLIER = 10;
const MAX_CATCH_COUNT = 60;

const handleReadTryCatch = async <T>(
  callback: () => Promise<T>,
  delayMultiplier?: number
) => {
  let res: T | undefined = undefined;

  try {
    res = await callback();
    lastReadRequestTime = new Date().getTime();
    nbInQueueRead -= delayMultiplier || 1;
  } catch (e: any) {
    if (VERBOSE_MODE)
      console.log(`inside catch 💩 #${readCatchCount}`, e.message);
    readCatchCount++;
    lastReadRequestTime = new Date().getTime();
    nbInQueueRead -= delayMultiplier || 1;

    if (readCatchCount < MAX_CATCH_COUNT)
      res = await handleReadDelay(callback, CATCH_DELAY_MULTIPLIER);
  } finally {
    readCatchCount = 0;
    return res as T;
  }
};

const handleReadDelay = async <T>(
  callback: () => Promise<T>,
  delayMultiplier?: number
) => {
  const currentTime = new Date().getTime();
  nbInQueueRead += delayMultiplier || 1;

  if (
    lastReadRequestTime &&
    currentTime < lastReadRequestTime + DELAY * nbInQueueRead
  ) {
    if (VERBOSE_MODE)
      console.log(
        "*** force DELAY [READ] ",
        nbInQueueRead,
        lastReadRequestTime
          ? lastReadRequestTime + DELAY * nbInQueueRead - currentTime
          : 0
      );
    await new Promise((resolve) =>
      setTimeout(
        () => resolve(null),
        lastReadRequestTime
          ? lastReadRequestTime + DELAY * nbInQueueRead - currentTime
          : 0
      )
    );
  }

  const res: T = await handleReadTryCatch(callback, delayMultiplier);

  return res;
};

const handleWriteTryCatch = async <T>(
  callback: () => Promise<T>,
  delayMultiplier?: number
) => {
  let res: T | undefined = undefined;

  try {
    res = await callback();
    lastWriteRequestTime = new Date().getTime();
    nbInQueueWrite -= delayMultiplier || 1;
  } catch (e: any) {
    if (VERBOSE_MODE)
      console.log(`inside catch 💩 #${writeCatchCount}`, e.message);
    writeCatchCount++;
    lastWriteRequestTime = new Date().getTime();
    nbInQueueWrite -= delayMultiplier || 1;

    if (writeCatchCount < MAX_CATCH_COUNT)
      res = await handleWriteDelay(callback, CATCH_DELAY_MULTIPLIER);
  } finally {
    writeCatchCount = 0;
    return res as T;
  }
};

const handleWriteDelay = async <T>(
  callback: () => Promise<T>,
  delayMultiplier?: number
) => {
  const currentTime = new Date().getTime();
  nbInQueueWrite += delayMultiplier || 1;

  if (
    lastWriteRequestTime &&
    currentTime < lastWriteRequestTime + DELAY * nbInQueueWrite
  ) {
    if (VERBOSE_MODE)
      console.log(
        "*** force DELAY [WRITE]",
        nbInQueueWrite,
        lastWriteRequestTime
          ? lastWriteRequestTime + DELAY * nbInQueueWrite - currentTime
          : 0
      );
    await new Promise((resolve) =>
      setTimeout(
        () => resolve(null),
        lastWriteRequestTime
          ? lastWriteRequestTime + DELAY * nbInQueueWrite - currentTime
          : 0
      )
    );
  }

  const res: T = await handleWriteTryCatch(callback, delayMultiplier);

  return res;
};

export const sheetAPI = {
  /**
   * Set your auth.json path (default : "./auth.json")
   * use something like path.join(__dirname + "/" + <RELATIVE PATH TO YOUR FILE>)
   */
  setAuthJsonPath: ({ path: authPath }: SetAuthJsonPathProps) => {
    // const callerPath_ = callerPath({depth:2});
    // console.log(callerPath_);
    // if (callerPath_) {
    //   const dirname = path.dirname(callerPath_);
    //   console.log("dirname", dirname);

    //   const joinedPath = path.join(dirname, authPath);
    //   console.log("joinedPath", joinedPath);

    //   const normPath = path.normalize(joinedPath);
    //   console.log("normPath", normPath);
    // }
    AUTH_JSON_PATH = authPath;
  },

  /**
   * Turns on logs with detailed infos over cache usage & delay between requests
   */
  enableConsoleLog: () => {
    VERBOSE_MODE = true;
  },

  /**
   * Return list of sheet tabs with their respective IDs
   */
  getTabIds: async ({ sheetId }: GetTabIdsProps): Promise<TabListItem[]> => {
    if (VERBOSE_MODE) console.log("*** sheetAPI.getTabIds", sheetId);
    if (sheetId) {
      const cacheKey = sheetId;
      if (tabIdsCache[cacheKey] === undefined) {
        await handleReadDelay(async () => {
          tabIdsCache[cacheKey] = await getSheetTabIds({
            sheetId,
            AUTH_JSON_PATH,
            VERBOSE_MODE,
          });
        });
      } else if (VERBOSE_MODE) console.log("*** using cache 👍");

      return tabIdsCache[cacheKey];
    }
    return [];
  },

  /**
   * Return tab data in the form of an array of objects built according to header values
   */
  getTabData: async ({
    sheetId,
    tabName,
    headerRowIndex,
    tabList,
  }: GetTabDataProps): Promise<TabDataItem[]> => {
    if (VERBOSE_MODE) console.log("*** sheetAPI.getTabData", tabName);

    let iTabList = tabList;
    if (!iTabList) iTabList = await sheetAPI.getTabIds({ sheetId });

    const tabId = iTabList.filter((tab) => tab.tabName === tabName)[0]?.tabId;
    if (tabId === undefined) throw new Error(`tab ${tabName} not found`);

    const cacheKey = sheetId + ":" + tabId;

    if (tabCache[cacheKey] === undefined) {
      await handleReadDelay(async () => {
        tabCache[cacheKey] = await importSheetData({
          sheetId,
          tabId,
          headerRowIndex,
          AUTH_JSON_PATH,
          VERBOSE_MODE,
        });
      });
    } else if (VERBOSE_MODE) console.log("*** using cache 👍");

    return tabCache[cacheKey];
  },

  // TOO MUCH COMPLEX, PUT IT AWAY FROM DOCUMENTATION AND DEPRECATE IT IN 3.0.0 (REPLACED BY getTabSize)
  /**
   * Will be deprecated in 3.0.0 - replaced by getTabSize
   */
  getTabMetaData: async ({ sheetId, fields, ranges }: GetTabMetaDataProps) => {
    if (VERBOSE_MODE) console.log("*** sheetAPI.getTabMetaData");

    const metaData = await handleReadDelay(async () => {
      const sheetApp = appSheet(AUTH_JSON_PATH);

      return await sheetApp.spreadsheets.get({
        spreadsheetId: sheetId,
        fields,
        ranges,
      });
    });

    return metaData;
  },

  /**
   * Return columns & rows count for a tab
   */
  getTabSize: async ({
    sheetId,
    tabName,
  }: GetTabSizeProps): Promise<TabSize> => {
    if (VERBOSE_MODE) console.log("*** sheetAPI.getTabSize", tabName);

    const cacheKey = "getTabSize:" + sheetId + ":" + tabName;

    if (tabSizesCache[cacheKey] === undefined) {
      await handleReadDelay(async () => {
        tabSizesCache[cacheKey] = await getTabSize({
          sheetId,
          tabName,
          AUTH_JSON_PATH,
        });
      });
    } else if (VERBOSE_MODE) console.log("*** using cache 👍");

    return tabSizesCache[cacheKey];
  },

  /**
   * Clear all read data in cache
   */
  clearCache: () => {
    tabCache = {};
    tabIdsCache = {};
    tabSizesCache = {};
  },

  /**
   * Udpate values of a range of cells
   * startCoords index start at 1
   */
  updateRange: async ({
    sheetId,
    tabName,
    startCoords,
    data,
  }: UpdateSheetRangeProps) => {
    if (VERBOSE_MODE) console.log("*** sheetAPI.updateRange");

    await handleWriteDelay(async () => {
      await updateSheetRange({
        sheetId,
        tabName,
        startCoords,
        data,
        AUTH_JSON_PATH,
      });
    });
  },

  /**
   * Append data to a sheet
   */
  appendToSheet: async ({ sheetId, tabName, data }: AppendToSheetProps) => {
    if (VERBOSE_MODE) console.log("*** sheetAPI.updateRange");

    const iTabList = await sheetAPI.getTabIds({ sheetId });

    const tabId = iTabList.filter((tab) => tab.tabName === tabName)[0]?.tabId;
    if (tabId === undefined) throw new Error(`tab ${tabName} not found`);

    await handleWriteDelay(async () => {
      await exportToSheet({
        datas: data,
        sheetId: tabId,
        exportSheetId: sheetId,
        VERBOSE_MODE: false,
        AUTH_JSON_PATH,
      });
    });
  },

  /**
   * get all protectedRange ids of a defined tab
   */
  getProtectedRangeIds: async ({
    sheetId,
    tabName,
  }: GetProtectedRangeIdsProps) => {
    if (VERBOSE_MODE) console.log("*** sheetAPI.getProtectedRangeIds");

    const iTabList = await sheetAPI.getTabIds({ sheetId });

    const tabId = iTabList.filter((tab) => tab.tabName === tabName)[0]?.tabId;
    if (tabId === undefined) throw new Error(`tab ${tabName} not found`);

    let protectedRangesIds: number[] = [];

    await handleReadDelay(async () => {
      protectedRangesIds = await batchUpdate.getProtectedRangeIds({
        spreadsheetId: sheetId,
        sheetId: parseInt(tabId, 10),
        AUTH_JSON_PATH,
      });
    });

    return protectedRangesIds;
  },

  /**
   * delete a list of protected range ids
   */
  deleteProtectedRange: async ({
    sheetId,
    protectedRangeIds,
  }: DeleteProtectedRangeProps) => {
    if (VERBOSE_MODE) console.log("*** sheetAPI.deleteProtectedRange");

    await handleWriteDelay(async () => {
      await batchUpdate.deleteProtectedRange({
        spreadsheetId: sheetId,
        protectedRangeIds,
        AUTH_JSON_PATH,
        VERBOSE_MODE,
      });
    });
  },

  addBatchProtectedRange: ({
    sheetId,
    editors,
    namedRangeId,
    tabId,
    startColumnIndex,
    startRowIndex,
    endColumnIndex,
    endRowIndex,
  }: AddProtectedRangeProps) => {
    batchUpdate.addProtectedRange({
      sheetId,
      editors,
      namedRangeId,
      tabId,
      startColumnIndex,
      startRowIndex,
      endColumnIndex,
      endRowIndex,
    });
  },

  runBatchProtectedRange: async ({ sheetId }: RunBatchProtectedRangeProps) => {
    if (VERBOSE_MODE) console.log("*** sheetAPI.runBatchProtectedRange");

    await handleWriteDelay(async () => {
      await batchUpdate.runProtectedRange({
        spreadsheetId: sheetId,
        AUTH_JSON_PATH,
        VERBOSE_MODE,
      });
    });
  },

  /**
   * Clear all values of a defined tab
   */
  clearTabData: async ({
    sheetId,
    tabName,
    headerRowIndex,
    tabList,
  }: ClearTabDataProps) => {
    if (VERBOSE_MODE) console.log("*** sheetAPI.clearTabData");

    let iTabList = tabList;
    if (!iTabList) iTabList = await sheetAPI.getTabIds({ sheetId });

    await handleWriteDelay(async () => {
      await clearTabData({
        sheetId,
        tabList: iTabList,
        tabName,
        headerRowIndex,
        AUTH_JSON_PATH,
        VERBOSE_MODE,
      });
    });
  },
};

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
  disableCache?: boolean;
};

type GetTabDataProps = {
  sheetId: string;
  tabName: string;
  headerRowIndex?: number;
  tabList?: TabListItem[];
  disableCache?: boolean;
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
  onTimeoutCallback?: () => Promise<void>;
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
let writeCatchCount = 0;

let AUTH_JSON_PATH = "./auth.json";
let VERBOSE_MODE = false;

const DELAY = 200; // in ms
const CATCH_DELAY_MULTIPLIER = 10;
const MAX_CATCH_COUNT = 60;
const MAX_AWAITING_TIME = 120_000;

export class MaxAwaitingTimeError extends Error {}

type HandleReadTryCatchProps<T> = {
  callback: () => Promise<T>;
  readCatchCount: number;
  delayMultiplier?: number;
  onTimeoutCallback?: () => Promise<void>;
};

const handleReadTryCatch = async <T>({
  callback,
  readCatchCount,
  delayMultiplier,
  onTimeoutCallback,
}: HandleReadTryCatchProps<T>) => {
  let res: T | undefined = undefined;
  let timeout: NodeJS.Timeout | undefined = undefined;

  try {
    timeout = setTimeout(async () => {
      if (VERBOSE_MODE) console.log("[READ] MAX_AWAITING_TIME reached 💀");
      if (onTimeoutCallback !== undefined) {
        await onTimeoutCallback();
        readCatchCount = MAX_CATCH_COUNT;
      }
      throw new MaxAwaitingTimeError();
    }, MAX_AWAITING_TIME);

    res = await callback();

    clearTimeout(timeout);

    lastReadRequestTime = new Date().getTime();
    nbInQueueRead -= delayMultiplier || 1;
  } catch (e: any) {
    if (VERBOSE_MODE)
      console.log(`inside catch 💩 #${readCatchCount}`, e.message);
    readCatchCount++;
    lastReadRequestTime = new Date().getTime();
    nbInQueueRead -= delayMultiplier || 1;
    clearTimeout(timeout);

    if (readCatchCount < MAX_CATCH_COUNT)
      res = await handleReadDelay({
        callback,
        readCatchCount,
        delayMultiplier: CATCH_DELAY_MULTIPLIER,
        onTimeoutCallback,
      });
  } finally {
    readCatchCount = 0;
    return res as T;
  }
};

type HandleReadDelayProps<T> = {
  callback: () => Promise<T>;
  readCatchCount?: number;
  delayMultiplier?: number;
  onTimeoutCallback?: () => Promise<void>;
};

const handleReadDelay = async <T>({
  callback,
  readCatchCount = 0,
  delayMultiplier,
  onTimeoutCallback,
}: HandleReadDelayProps<T>) => {
  const currentTime = new Date().getTime();
  nbInQueueRead += delayMultiplier || 1;

  if (
    lastReadRequestTime &&
    currentTime < lastReadRequestTime + DELAY * nbInQueueRead
  ) {
    if (VERBOSE_MODE)
      console.log("*** force DELAY [READ] ", {
        nbInQueueRead: nbInQueueRead / (delayMultiplier || 1),
        timeout: lastReadRequestTime
          ? lastReadRequestTime + DELAY * nbInQueueRead - currentTime
          : 0,
      });
    await new Promise((resolve) =>
      setTimeout(
        () => resolve(null),
        lastReadRequestTime
          ? lastReadRequestTime + DELAY * nbInQueueRead - currentTime
          : 0
      )
    );
  }

  const res: T = await handleReadTryCatch({
    callback,
    readCatchCount,
    delayMultiplier,
    onTimeoutCallback,
  });

  return res;
};

type HandleWriteTryCatchProps<T> = {
  callback: () => Promise<T>;
  writeCatchCount: number;
  delayMultiplier?: number;
  onTimeoutCallback?: () => Promise<void>;
};

const handleWriteTryCatch = async <T>({
  callback,
  writeCatchCount,
  delayMultiplier,
  onTimeoutCallback,
}: HandleWriteTryCatchProps<T>) => {
  let res: T | undefined = undefined;
  let timeout: NodeJS.Timeout | undefined = undefined;

  try {
    timeout = setTimeout(async () => {
      if (VERBOSE_MODE) console.log("[WRITE] MAX_AWAITING_TIME reached 💀");
      if (onTimeoutCallback !== undefined) {
        await onTimeoutCallback();
        writeCatchCount = MAX_CATCH_COUNT;
      }
      throw new MaxAwaitingTimeError();
    }, MAX_AWAITING_TIME);

    res = await callback();
    clearTimeout(timeout);

    lastWriteRequestTime = new Date().getTime();
    nbInQueueWrite -= delayMultiplier || 1;
  } catch (e: any) {
    if (VERBOSE_MODE)
      console.log(`inside catch 💩 #${writeCatchCount}`, e.message);
    writeCatchCount++;
    lastWriteRequestTime = new Date().getTime();
    nbInQueueWrite -= delayMultiplier || 1;
    clearTimeout(timeout);

    if (writeCatchCount < MAX_CATCH_COUNT)
      res = await handleWriteDelay({
        callback,
        writeCatchCount,
        delayMultiplier: CATCH_DELAY_MULTIPLIER,
        onTimeoutCallback,
      });
  } finally {
    writeCatchCount = 0;
    return res as T;
  }
};

type HandleWriteDelayProps<T> = {
  callback: () => Promise<T>;
  writeCatchCount?: number;
  delayMultiplier?: number;
  onTimeoutCallback?: () => Promise<void>;
};

const handleWriteDelay = async <T>({
  callback,
  writeCatchCount = 0,
  delayMultiplier,
  onTimeoutCallback,
}: HandleWriteDelayProps<T>) => {
  const currentTime = new Date().getTime();
  nbInQueueWrite += delayMultiplier || 1;

  if (
    lastWriteRequestTime &&
    currentTime < lastWriteRequestTime + DELAY * nbInQueueWrite
  ) {
    if (VERBOSE_MODE)
      console.log("*** force DELAY [WRITE]", {
        nbInQueueWrite: nbInQueueWrite / (delayMultiplier || 1),
        timeout: lastWriteRequestTime
          ? lastWriteRequestTime + DELAY * nbInQueueWrite - currentTime
          : 0,
      });
    await new Promise((resolve) =>
      setTimeout(
        () => resolve(null),
        lastWriteRequestTime
          ? lastWriteRequestTime + DELAY * nbInQueueWrite - currentTime
          : 0
      )
    );
  }

  const res: T = await handleWriteTryCatch({
    callback,
    writeCatchCount,
    delayMultiplier,
    onTimeoutCallback,
  });

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
  getTabIds: async ({
    sheetId,
    disableCache,
  }: GetTabIdsProps): Promise<TabListItem[]> => {
    if (VERBOSE_MODE) console.log("*** sheetAPI.getTabIds", sheetId);
    if (sheetId) {
      const cacheKey = sheetId;
      if (disableCache || tabIdsCache[cacheKey] === undefined) {
        await handleReadDelay({
          callback: async () => {
            tabIdsCache[cacheKey] = await getSheetTabIds({
              sheetId,
              AUTH_JSON_PATH,
              VERBOSE_MODE,
            });
          },
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
    disableCache,
  }: GetTabDataProps): Promise<TabDataItem[]> => {
    if (VERBOSE_MODE) console.log("*** sheetAPI.getTabData", tabName);

    let iTabList = tabList;
    if (iTabList === undefined)
      iTabList = await sheetAPI.getTabIds({ sheetId, disableCache });

    const tabId = iTabList.filter((tab) => tab.tabName === tabName)[0]?.tabId;
    if (tabId === undefined) throw new Error(`tab ${tabName} not found`);

    const cacheKey = sheetId + ":" + tabId;

    if (disableCache || tabCache[cacheKey] === undefined) {
      await handleReadDelay({
        callback: async () => {
          tabCache[cacheKey] = await importSheetData({
            sheetId,
            tabId,
            headerRowIndex,
            AUTH_JSON_PATH,
            VERBOSE_MODE,
          });
        },
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

    const metaData = await handleReadDelay({
      callback: async () => {
        const sheetApp = appSheet(AUTH_JSON_PATH);

        return await sheetApp.spreadsheets.get({
          spreadsheetId: sheetId,
          fields,
          ranges,
        });
      },
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
      await handleReadDelay({
        callback: async () => {
          tabSizesCache[cacheKey] = await getTabSize({
            sheetId,
            tabName,
            AUTH_JSON_PATH,
          });
        },
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

    await handleWriteDelay({
      callback: async () => {
        await updateSheetRange({
          sheetId,
          tabName,
          startCoords,
          data,
          AUTH_JSON_PATH,
        });
      },
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

    await handleWriteDelay({
      callback: async () => {
        await exportToSheet({
          datas: data,
          sheetId: tabId,
          exportSheetId: sheetId,
          VERBOSE_MODE: false,
          AUTH_JSON_PATH,
        });
      },
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

    await handleReadDelay({
      callback: async () => {
        protectedRangesIds = await batchUpdate.getProtectedRangeIds({
          spreadsheetId: sheetId,
          sheetId: parseInt(tabId, 10),
          AUTH_JSON_PATH,
        });
      },
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

    await handleWriteDelay({
      callback: async () => {
        await batchUpdate.deleteProtectedRange({
          spreadsheetId: sheetId,
          protectedRangeIds,
          AUTH_JSON_PATH,
          VERBOSE_MODE,
        });
      },
    });
  },

  /**
   * Add a request to batch buffer
   */
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

  /**
   * Clear all requests batch buffer
   */
  clearBuffer: async () => {
    if (VERBOSE_MODE) console.log("*** sheetAPI.clearBuffer");
    batchUpdate.clearBuffer();
  },

  /**
   * Return all requests in batch buffer
   */
  getBatchProtectedRange: async () => {
    if (VERBOSE_MODE) console.log("*** sheetAPI.getBatchProtectedRange");
    return batchUpdate.getBatchProtectedRange();
  },

  /**
   * Run all requests in batch buffer
   */
  runBatchProtectedRange: async ({
    sheetId,
    onTimeoutCallback,
  }: RunBatchProtectedRangeProps) => {
    if (VERBOSE_MODE) console.log("*** sheetAPI.runBatchProtectedRange");

    await handleWriteDelay({
      callback: async () => {
        await batchUpdate.runProtectedRange({
          spreadsheetId: sheetId,
          AUTH_JSON_PATH,
          VERBOSE_MODE,
        });
      },
      onTimeoutCallback,
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

    await handleWriteDelay({
      callback: async () => {
        await clearTabData({
          sheetId,
          tabList: iTabList,
          tabName,
          headerRowIndex,
          AUTH_JSON_PATH,
          VERBOSE_MODE,
        });
      },
    });
  },
};

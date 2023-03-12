import { batchUpdate } from "./appSheet/batchUpdate";
import { getSheetTabIds } from "./appSheet/getSheetTabIds";
import { appSheet } from "./google";
import { importSheetData } from "./appSheet/importSheetData";
import { updateSheetRange } from "./appSheet/updateSheetRange";
import { clearTabData } from "./appSheet/clearSheetRows";
import { TabDataItem, TabListItem } from "./interfaces";
import { getTabSize } from "./appSheet/getTabSize";

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
  setAuthJsonPath: ({ path }: SetAuthJsonPathProps) => {
    AUTH_JSON_PATH = path;
  },

  enableConsoleLog: () => {
    VERBOSE_MODE = true;
  },

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

  getTabData: async ({
    sheetId,
    tabList,
    tabName,
    headerRowIndex,
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

  clearCache: () => {
    tabCache = {};
    tabIdsCache = {};
    tabSizesCache = {};
  },

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

  clearTabData: async ({
    sheetId,
    tabList,
    tabName,
    headerRowIndex,
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

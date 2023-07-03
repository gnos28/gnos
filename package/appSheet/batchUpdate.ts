import { sheets_v4 } from "googleapis";
import { appSheet } from "../google";
import { AddProtectedRangeProps } from "../sheetAPI";

type RunProtectedRangeProps = {
  spreadsheetId: string;
  AUTH_JSON_PATH: string;
  VERBOSE_MODE: boolean;
};

type GetProtectedRangeIdsProps = {
  spreadsheetId: string;
  sheetId: number;
  AUTH_JSON_PATH: string;
};

type DeleteProtectedRangeProps = {
  spreadsheetId: string;
  protectedRangeIds: number[];
  AUTH_JSON_PATH: string;
  VERBOSE_MODE: boolean;
};

const protectedRangeBatchBuffer: {
  [key: string]: sheets_v4.Schema$Request[];
} = {};

export const batchUpdate = {
  addProtectedRange: ({
    sheetId,
    editors,
    namedRangeId,
    tabId,
    startColumnIndex,
    startRowIndex,
    endColumnIndex,
    endRowIndex,
  }: AddProtectedRangeProps) => {
    if (!protectedRangeBatchBuffer[sheetId])
      protectedRangeBatchBuffer[sheetId] = [];

    const namedRange = namedRangeId
      ? namedRangeId
      : Math.random().toString().split(".")[1];

    protectedRangeBatchBuffer[sheetId].push({
      addProtectedRange: {
        protectedRange: {
          editors: { users: editors },
          description: namedRange,
          range: {
            sheetId: tabId,
            startColumnIndex,
            startRowIndex,
            endColumnIndex: endColumnIndex + 1,
            endRowIndex: endRowIndex + 1,
          },
        },
      },
    });
  },

  clearBuffer: () => {
    Object.keys(protectedRangeBatchBuffer).forEach((key) => {
      protectedRangeBatchBuffer[key] = [];
    });
  },

  getBatchProtectedRange: () => {
    return protectedRangeBatchBuffer;
  },

  runProtectedRange: async ({
    spreadsheetId,
    AUTH_JSON_PATH,
    VERBOSE_MODE,
  }: RunProtectedRangeProps) => {
    const sheetApp = appSheet(AUTH_JSON_PATH);

    const requests = protectedRangeBatchBuffer[spreadsheetId];
    if (requests.length > 0) {
      if (VERBOSE_MODE)
        console.log(
          spreadsheetId.substring(0, 8),
          "[runProtectedRange] requests count : ",
          requests.length
        );
      const startTime = new Date().getTime();

      await sheetApp.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: {
          requests,
        },
      });
      if (VERBOSE_MODE)
        console.log(
          spreadsheetId.substring(0, 8),
          "[runProtectedRange] time taken (ms) : ",
          new Date().getTime() - startTime
        );
      protectedRangeBatchBuffer[spreadsheetId] = [];
    }
  },

  getProtectedRangeIds: async ({
    spreadsheetId,
    sheetId,
    AUTH_JSON_PATH,
  }: GetProtectedRangeIdsProps) => {
    const sheetApp = appSheet(AUTH_JSON_PATH);

    const getResult = await sheetApp.spreadsheets.get({ spreadsheetId });

    const sheets = getResult.data.sheets;
    if (sheets !== undefined) {
      const sheet = sheets.filter(
        (sheet) => sheet.properties?.sheetId === sheetId
      )[0];

      const protectedRanges = sheet.protectedRanges;

      if (protectedRanges !== undefined) {
        return protectedRanges
          .map((protectedRange) => protectedRange.protectedRangeId)
          .filter((id) => id !== null && id !== undefined) as number[];
      }
    }

    return [];
  },

  deleteProtectedRange: async ({
    spreadsheetId,
    protectedRangeIds,
    AUTH_JSON_PATH,
    VERBOSE_MODE,
  }: DeleteProtectedRangeProps) => {
    const sheetApp = appSheet(AUTH_JSON_PATH);

    const requests: sheets_v4.Schema$Request[] = protectedRangeIds.map(
      (protectedRangeId) => ({
        deleteProtectedRange: { protectedRangeId },
      })
    );
    if (VERBOSE_MODE)
      console.log(
        spreadsheetId.substring(0, 8),
        "[deleteProtectedRange] requests count : ",
        requests.length
      );

    const startTime = new Date().getTime();

    await sheetApp.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests,
      },
    });

    if (VERBOSE_MODE)
      console.log(
        spreadsheetId.substring(0, 8),
        "[deleteProtectedRange] time taken (ms) : ",
        new Date().getTime() - startTime
      );
  },
};

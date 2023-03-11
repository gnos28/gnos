import { sheets_v4 } from "googleapis";
import { appSheet } from "../google";
import { AddProtectedRangeProps } from "../sheetAPI";

type RunProtectedRangeProps = {
  spreadsheetId: string;
  AUTH_JSON_PATH: string;
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

  runProtectedRange: async ({
    spreadsheetId,
    AUTH_JSON_PATH,
  }: RunProtectedRangeProps) => {
    const sheetApp = appSheet(AUTH_JSON_PATH);

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

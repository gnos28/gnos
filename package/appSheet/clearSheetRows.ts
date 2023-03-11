import { GoogleSpreadsheet } from "google-spreadsheet";
import { TabListItem } from "../interfaces";
import { getAuthJson } from "../utils/getAuthJson";

type ClearSheetRowsProps = {
  sheetId: string;
  tabId: string;
  headerRowIndex?: number;
  AUTH_JSON_PATH: string;
};

export const clearSheetRows = async ({
  sheetId,
  tabId,
  headerRowIndex,
  AUTH_JSON_PATH,
}: ClearSheetRowsProps) => {
  console.log("clearSheetRows");

  const { GOOGLE_SERVICE_ACCOUNT_EMAIL, GOOGLE_PRIVATE_KEY } =
    getAuthJson(AUTH_JSON_PATH);

  if (GOOGLE_SERVICE_ACCOUNT_EMAIL && GOOGLE_PRIVATE_KEY) {
    const doc = new GoogleSpreadsheet(sheetId);

    await doc.useServiceAccountAuth({
      // env var values are copied from service account credentials generated by google
      // see "Authentication" section in docs for more info
      client_email: GOOGLE_SERVICE_ACCOUNT_EMAIL,
      private_key: GOOGLE_PRIVATE_KEY,
    });

    await doc.loadInfo(); // loads document properties and worksheets

    const sheet = doc.sheetsById[tabId];

    if (headerRowIndex) await sheet.loadHeaderRow(headerRowIndex);

    await sheet.clearRows();
  }
};

type ClearTabDataProps = {
  sheetId: string;
  tabList: TabListItem[];
  tabName: string;
  headerRowIndex?: number;
  AUTH_JSON_PATH: string;
};

export const clearTabData = async ({
  sheetId,
  tabList,
  tabName,
  headerRowIndex,
  AUTH_JSON_PATH,
}: ClearTabDataProps) => {
  const tabId = tabList.filter((tab) => tab.sheetName === tabName)[0]?.sheetId;
  if (tabId === undefined) throw new Error(`tab ${tabName} not found`);

  return await clearSheetRows({
    sheetId,
    tabId,
    headerRowIndex,
    AUTH_JSON_PATH,
  });
};
